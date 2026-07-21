# -*- coding: utf-8 -*-
"""App 主类。

负责：
  - 创建根窗口布局（工具栏 / 表格区 / 状态栏）
  - 组装 KeywordPopup、ResultTable、SearchController
  - 处理用户交互（搜索触发、索引操作、右键菜单动作）
  - 轮询队列，把后台消息转为 UI 更新
"""
import os
import queue
import threading

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import xls_search.excel_actions as excel_actions
import xls_search.ime as ime
import xls_search.search_excel as search_excel
from xls_search.paths import col_letter
from xls_search.storage import (load_settings, save_settings,
                                load_keywords, save_keyword,
                                load_sources, save_sources)

from xls_search.keyword_popup import KeywordPopup
from xls_search.search_controller import SearchController
from xls_search.table_widget import ResultTable


class App:
    def __init__(self, root, scale=1.0):
        self.root = root
        self.scale = scale
        root.title("xls_search")
        root.geometry(f"{int(1000 * scale)}x{int(640 * scale)}")
        root.minsize(int(760 * scale), int(480 * scale))

        self.q = queue.Queue()          # 后台线程 -> 主线程 的消息队列
        self.busy = False               # 是否有任务在跑
        self.settings = load_settings() # 记忆的偏好（模式等）
        self._hl_keyword = ""           # 当前用于高亮的关键字
        self.cancel_event = threading.Event()   # 置位表示请求取消当前后台任务

        self._controller = SearchController(self.q, self.cancel_event)

        self._build_ui()
        self._refresh_sources()
        self.root.after(80, self._poll_queue)
        self.root.after(1500, self._check_sync)   # 定时比对目录/索引文件数
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ================================================================== #
    #  UI 构建                                                            #
    # ================================================================== #

    def _build_ui(self):
        pad = {"padx": 6, "pady": 4}
        s = self.scale
        px = lambda n: -int(round(n * s))          # 像素字号（负值=像素高）

        style = ttk.Style()
        self.ui_font   = ("Microsoft YaHei UI", px(17))
        self.head_font = ("Microsoft YaHei UI", px(17), "bold")
        entry_font     = ("Microsoft YaHei UI", px(17), "bold")     # 输入框内文字加粗更显眼
        for st in (".", "TLabel", "TButton", "TRadiobutton", "TCheckbutton"):
            style.configure(st, font=self.ui_font)
        style.configure("TEntry",    font=entry_font)
        style.configure("TCombobox", font=entry_font)
        kw_font = ("Microsoft YaHei UI", px(20))   # 关键字框：经典 tk.Entry 才吃 font
        self.root.option_add("*TCombobox*Listbox.font", self.ui_font)  # 下拉历史也放大

        self._build_toolbar(pad, kw_font, px, s)
        self._build_table(pad, s)
        self._build_statusbar(pad)

    # ---------- 工具栏 ----------

    def _build_toolbar(self, pad, kw_font, px, s):
        # 目录行
        top = ttk.Frame(self.root)
        top.pack(fill="x", **pad)
        ttk.Label(top, text="xls 目录:").grid(row=0, column=0, sticky="w")
        self.dir_var = tk.StringVar()
        self.dir_combo = ttk.Combobox(top, textvariable=self.dir_var, width=70)
        self.dir_combo.grid(row=0, column=1, sticky="we", padx=4)
        self.dir_combo.bind("<Return>", lambda e: self._on_dir_return(e))
        self.dir_combo.bind("<<ComboboxSelected>>", lambda e: self._on_dir_selected())
        # 点击别处时目录输入框失去焦点（bind_all 在事件链最末执行，
        # 不影响 Combobox 下拉选择等内部处理）
        self.root.bind_all("<Button-1>", self._on_global_dir_click, add="+")
        ttk.Button(top, text="浏览…", command=self._browse).grid(row=0, column=2)
        top.columnconfigure(1, weight=1)

        # 模式行
        mode_frame = ttk.Frame(self.root)
        mode_frame.pack(fill="x", **pad)
        ttk.Label(mode_frame, text="模式:").pack(side="left")
        saved_mode = self.settings.get("mode", "2")
        self.mode_var = tk.StringVar(
            value=saved_mode if saved_mode in ("1", "2") else "2")
        for val, text in [("1", "直接读取文件"), ("2", "查询索引")]:
            ttk.Radiobutton(mode_frame, text=text, value=val,
                            variable=self.mode_var).pack(side="left", padx=6)
        self.mode_var.trace_add("write", lambda *a: self._save_mode())

        # 索引操作下拉：选择「更新变动 / 重建全部」后弹确认再后台执行
        self._index_placeholder = "索引操作 ▾"
        self.index_action_var = tk.StringVar(value=self._index_placeholder)
        self.index_combo = ttk.Combobox(
            mode_frame, textvariable=self.index_action_var,
            state="readonly", width=14,
            values=["更新变动索引", "重建全部索引"])
        self.index_combo.pack(side="left", padx=(14, 6))
        self.index_combo.bind("<<ComboboxSelected>>", self._on_index_action)

        # 模式行最右：文件/索引不同步时的绿字提醒
        self.sync_var = tk.StringVar(value="")
        ttk.Label(mode_frame, textvariable=self.sync_var, foreground="#0a9a0a",
                  font=("Microsoft YaHei UI", px(10), "bold")
                  ).pack(side="right", padx=8)

        # 关键字行
        kw = ttk.Frame(self.root)
        kw.pack(fill="x", **pad)
        ttk.Label(kw, text="关键字:").grid(row=0, column=0, sticky="w")
        self.kw_var = tk.StringVar()
        self.kw_entry = tk.Entry(kw, textvariable=self.kw_var, font=kw_font,
                                 relief="solid", bd=1)
        self.kw_entry.grid(row=0, column=1, sticky="we", padx=4,
                           ipady=int(round(4 * self.scale)))
        self.kw_entry.bind("<Return>", lambda e: self._start_search())
        self.kw_entry.bind("<FocusIn>", lambda e: self._set_kw_ime_font())

        # 历史下拉：点击输入框弹出、输入时实时筛选（自定义弹层，不抢输入焦点）
        self._kw_popup = KeywordPopup(
            root=self.root,
            entry=self.kw_entry,
            kw_var=self.kw_var,
            ui_font=self.ui_font,
            on_pick=self._on_kw_pick,
            is_busy=lambda: self.busy,
        )
        self._refresh_keywords()

        self.exact_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(kw, text="精确匹配",
                        variable=self.exact_var).grid(row=0, column=2, padx=4)
        ttk.Label(kw, text="文件名含:").grid(row=0, column=3, sticky="w")
        self.filter_var = tk.StringVar()
        ttk.Entry(kw, textvariable=self.filter_var,
                  width=14).grid(row=0, column=4, padx=4)
        ttk.Label(kw, text="限定列:").grid(row=0, column=5, sticky="w")
        self.col_var = tk.StringVar()
        ttk.Entry(kw, textvariable=self.col_var,
                  width=5).grid(row=0, column=6, padx=4)
        self.search_btn = ttk.Button(kw, text="搜索", command=self._start_search)
        self.search_btn.grid(row=0, column=7, padx=6)
        kw.columnconfigure(1, weight=1)

    # ---------- 表格区 ----------

    def _build_table(self, pad, s):
        saved_px = self.settings.get("col_px", {})   # 记忆的列宽（像素）

        def _w(key, default):
            # 记忆值钳制在合理范围，避免异常大/小的值把布局撑坏
            lo, hi = int(round(40 * s)), int(round(900 * s))
            v = saved_px.get(key)
            if not isinstance(v, (int, float)):
                v = default * s
            return int(round(max(lo, min(hi, v))))

        # 列定义：(key, 标题, 固定像素宽 或 None=填充剩余, 对齐)
        col_spec = [
            ["#",     "#",     _w("#", 55),      "center"],
            ["file",  "文件",  _w("file", 300),  "w"],
            ["sheet", "Sheet", _w("sheet", 150), "center"],
            ["row",   "行",    _w("row", 70),    "center"],
            ["col",   "列",    _w("col", 80),    "center"],
            ["value", "值",    None,             "w"],
        ]

        table_frame = ttk.Frame(self.root)
        table_frame.pack(fill="both", expand=True, **pad)

        self.table = ResultTable(
            parent=table_frame,
            col_spec=col_spec,
            scale=s,
            ui_font=self.ui_font,
            head_font=self.head_font,
            hl_color="#e8590c",                # 关键字命中：橙色
            on_open_file=self._ctx_open_file,
            on_view_value=self._ctx_view_value,
            on_copy_name=self._ctx_copy_name,
            on_open_dir=self._ctx_open_dir,
            on_copy_row=self._ctx_copy_row,
        )
        self.table.set_col_resize_callback(
            lambda: self.table.save_col_px(self.settings, save_settings))
        self.table.set_page_size_callback(
            lambda n: self._save_page_size(n))
        # 从记忆恢复每页条数
        saved_ps = self.settings.get("page_size", 50)
        self.table.set_page_size(saved_ps)

    # ---------- 状态栏 ----------

    def _build_statusbar(self, pad):
        bottom = ttk.Frame(self.root)
        bottom.pack(fill="x", **pad)
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=4)
        # 取消按钮先创建但不 pack —— 无任务时隐藏，busy 时才显示（见 _set_busy）
        self.cancel_btn = ttk.Button(bottom, text="取消", command=self._cancel)
        self.status_var = tk.StringVar(value="就绪")
        self.status_lbl = ttk.Label(bottom, textvariable=self.status_var,
                                    width=40, anchor="w")
        self.status_lbl.pack(side="right", padx=6)

    # ================================================================== #
    #  目录 / 关键字历史                                                  #
    # ================================================================== #

    def _refresh_sources(self):
        srcs = load_sources()
        # 最近用的排最前
        self.dir_combo["values"] = list(reversed(srcs))
        if srcs and not self.dir_var.get():
            self.dir_var.set(srcs[-1])

    def _on_dir_return(self, event):
        """目录输入框回车：不存在弹提示框保留光标，存在则失焦并记历史。"""
        d = self.dir_var.get().strip().strip('"')
        if d and not os.path.isdir(d):
            messagebox.showwarning("提示", "目录不存在")
            return
        if d and os.path.isdir(d):
            self._save_dir_history(d)
        self.root.focus_set()

    def _on_dir_selected(self):
        """从下拉列表选了一个历史目录后更新顺序。"""
        d = self.dir_var.get().strip().strip('"')
        if d and os.path.isdir(d):
            self._save_dir_history(d)

    def _save_dir_history(self, d):
        """异步写盘 + 立即更新内存下拉列表（不重读文件，避免异步写盘竞态）。"""
        threading.Thread(target=lambda: save_sources([d]), daemon=True).start()
        # 立即更新下拉列表：最新的放最上面
        vals = list(self.dir_combo["values"])
        vals = [v for v in vals if v != d]   # 去重
        vals.insert(0, d)                     # 最新→最前（顶部）
        self.dir_combo["values"] = vals

    def _on_global_dir_click(self, event):
        """全局点击：如果点击在目录 Combobox 外，让它失去焦点。"""
        try:
            w = event.widget
            while w is not None:
                if w is self.dir_combo:
                    return  # 点击在 combobox 内部（含下拉列表），不动
                w = w.master
        except Exception:
            pass
        # 点击在 combobox 外，且输入框非空 → 校验目录
        if self.root.focus_get() is self.dir_combo:
            d = self.dir_var.get().strip().strip('"')
            if d and not os.path.isdir(d):
                messagebox.showwarning("提示", "目录不存在")
                return   # 目录不存在，保留光标
            if d and os.path.isdir(d):
                self._save_dir_history(d)
            self.root.focus_set()

    def _refresh_keywords(self):
        # 最近用的排最前；供点击/输入时筛选
        kws = list(reversed(load_keywords()))
        self._kw_popup.set_keywords(kws)

    def _set_kw_ime_font(self):
        # 让正在输入的拼音字体跟随关键字框（雅黑 20px 常规）
        ime.set_composition_font(self.kw_entry, 20 * self.scale,
                                 family="Microsoft YaHei UI", weight=400)

    def _on_kw_pick(self, keyword):
        """关键字下拉选中后直接触发搜索。"""
        self._start_search()

    def _browse(self):
        start = self.dir_var.get() if os.path.isdir(self.dir_var.get()) else None
        d = filedialog.askdirectory(initialdir=start, title="选择 xls 目录")
        if d:
            d = os.path.normpath(d)
            self.dir_var.set(d)
            self._save_dir_history(d)

    # ================================================================== #
    #  搜索 / 建索引                                                      #
    # ================================================================== #

    def _start_search(self):
        if self.busy:
            return
        self._kw_popup.hide()
        xls_dir = self.dir_var.get().strip().strip('"')
        keyword  = self.kw_var.get().strip()
        if not xls_dir or not os.path.isdir(xls_dir):
            messagebox.showwarning("提示", "请选择有效的 xls 目录")
            return
        if not keyword:
            messagebox.showwarning("提示", "请输入关键字")
            return

        mode = self.mode_var.get()
        exact = self.exact_var.get()
        filter_str = self.filter_var.get().strip() or None
        col_txt = self.col_var.get().strip().upper()
        if col_txt.isdigit():
            col_filter = int(col_txt)
        elif len(col_txt) == 1 and 'A' <= col_txt <= 'Z':
            col_filter = ord(col_txt) - 64
        elif len(col_txt) == 2 and col_txt.isalpha() and col_txt.isascii():
            col_filter = (ord(col_txt[0]) - 64) * 26 + (ord(col_txt[1]) - 64)
        else:
            col_filter = None

        # 异步写盘；UI 侧直接推入内存，不等待磁盘 I/O
        threading.Thread(target=lambda: save_keyword(keyword), daemon=True).start()
        threading.Thread(target=lambda: save_sources([xls_dir]), daemon=True).start()
        self._refresh_sources()
        # 直接在内存中追加新词，不要重新读文件（异步写盘可能还没完成）
        self._kw_popup.push_keyword(keyword)

        # 模式2 但索引不存在 -> 询问是否建立
        if mode == "2" and not search_excel.index_exists(xls_dir):
            if messagebox.askyesno("索引不存在", "该目录还没有索引，现在建立吗？"):
                mode = "3"
            else:
                return

        self._hl_keyword = keyword      # 供结果渲染时高亮命中字
        self.cancel_event.clear()
        self._set_busy(True)
        self.table.clear()
        self._controller.run_search(
            xls_dir=xls_dir, keyword=keyword, exact=exact,
            filter_str=filter_str, col_filter=col_filter, mode=mode)

    def _on_index_action(self, event=None):
        action = self.index_action_var.get()
        self.index_action_var.set(self._index_placeholder)   # 复位下拉显示
        self.index_combo.selection_clear()
        if self.busy:
            return
        if action == "重建全部索引":
            if messagebox.askyesno(
                    "确认重建",
                    "将清空并重新建立全部索引，耗时较久。\n确定重建全部索引吗？"):
                self._start_index_build("4")
        elif action == "更新变动索引":
            if messagebox.askyesno(
                    "确认更新",
                    "将只重建有变动的文件索引。\n确定更新变动索引吗？"):
                self._start_index_build("3")

    def _start_index_build(self, mode):
        if self.busy:
            return
        xls_dir = self.dir_var.get().strip().strip('"')
        if not xls_dir or not os.path.isdir(xls_dir):
            messagebox.showwarning("提示", "请选择有效的 xls 目录")
            return
        save_sources([xls_dir])
        self._refresh_sources()
        self.cancel_event.clear()
        self._set_busy(True)
        self.table.clear()
        self._controller.run_index_build(xls_dir=xls_dir, mode=mode)

    # ================================================================== #
    #  队列轮询（主线程更新 UI）                                          #
    # ================================================================== #

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "progress":
                    i, total, rel = payload
                    self.progress["maximum"] = max(total, 1)
                    self.progress["value"] = i
                    self.status_var.set(f"[{i}/{total}] {rel[:40]}")
                elif kind == "status":
                    self.status_var.set(payload)
                elif kind == "mode":
                    self.mode_var.set(payload)
                elif kind == "results":
                    self._show_results(payload)
                elif kind == "error":
                    messagebox.showerror("出错", payload)
                    self.status_var.set("出错: " + payload)
                elif kind == "cancelled":
                    self.status_var.set(payload or "已取消")
                elif kind == "sync":
                    self.sync_var.set(payload)
                elif kind == "done":
                    self._set_busy(False)
        except queue.Empty:
            pass
        self.root.after(80, self._poll_queue)

    # ================================================================== #
    #  结果展示                                                           #
    # ================================================================== #

    def _show_results(self, rows):
        self.table.set_rows(rows, hl_keyword=self._hl_keyword)
        self.progress["value"] = self.progress["maximum"]
        total = len(self.table.all_rows)
        self.status_var.set("没有匹配" if total == 0 else f"找到 {total} 条")

    # ================================================================== #
    #  右键菜单动作                                                       #
    # ================================================================== #

    def _selected_full(self):
        data = self.table._selected_data()
        if not data:
            return None
        xls_dir = self.dir_var.get().strip().strip('"')
        return os.path.normpath(os.path.join(xls_dir, data[0]))

    def _ctx_open_file(self, data, jump_cell=False):
        full = self._selected_full()
        if not full or not data:
            return
        if not os.path.exists(full):
            messagebox.showwarning("提示", f"文件不存在：\n{full}")
            return
        _, sheet, row, col, _ = data
        sheet = sheet or None
        row = row if isinstance(row, int) else None
        col = col if isinstance(col, int) else None
        # COM 启动 Excel 可能较慢，放后台线程；失败则退回默认打开
        threading.Thread(
            target=self._open_in_excel,
            args=(full, sheet, row, col, jump_cell),
            daemon=True).start()

    def _open_in_excel(self, full, sheet, row, col, jump_cell=False):
        excel_actions.open_in_excel(
            full, sheet, row, col, jump_cell=jump_cell,
            on_status=lambda msg: self.q.put(("status", msg)))

    def _ctx_view_value(self, data, idx, table):
        """弹出可滚动窗口显示某行的完整值（不受表格两行片段限制）。"""
        if data is None:
            return
        file, sheet, row, col, val = data
        seq = idx + 1   # 序号（1-based，与表格 # 列一致）
        xls_dir = self.dir_var.get().strip().strip('"')
        full = os.path.normpath(os.path.join(xls_dir, file))

        # 已有窗口则复用（更新内容），不销毁重建，避免窗口位置跳动
        old = getattr(self, "_view_val_win", None)
        if old is not None and old.winfo_exists():
            old.title(f"完整值 — #{seq} {os.path.basename(file)}  {col_letter(col)}{row}")
            # 更新工具栏中的文件路径信息（老窗口可能无工具栏，用 hasattr 防御）
            old._val_full = full
            old._val_sheet = sheet
            old._val_row = row
            old._val_col = col
            if hasattr(old, "_val_info_lbl"):
                old._val_info_lbl.configure(
                    text=f"#{seq} {os.path.basename(file)} | Sheet={sheet} | {col_letter(col)}{row}")
            txt = old._val_text
            txt.configure(state="normal")
            txt.delete("1.0", "end")
            txt.insert("1.0", val)
            kw = self._hl_keyword
            if kw:
                txt.tag_delete("hit")
                txt.tag_configure("hit", foreground="#e8590c", font=self.head_font)
                start, n = "1.0", len(kw)
                while True:
                    pos = txt.search(kw, start, stopindex="end", nocase=1)
                    if not pos:
                        break
                    txt.tag_add("hit", pos, f"{pos}+{n}c")
                    start = f"{pos}+{n}c"
            txt.configure(state="disabled")
            old.deiconify()
            old.lift()
            return

        win = tk.Toplevel(self.root)
        self._view_val_win = win
        win.title(f"完整值 — #{seq} {os.path.basename(file)}  {col_letter(col)}{row}")
        s = self.scale
        win.geometry(f"{int(720 * s)}x{int(480 * s)}")
        # 保存上下文信息到窗口对象，供工具栏按钮回调使用
        win._val_full = full
        win._val_sheet = sheet
        win._val_row = row
        win._val_col = col

        # 顶部工具栏：打开文件按钮（带下拉选项）
        toolbar = ttk.Frame(win)
        toolbar.pack(fill="x", padx=6, pady=(6, 2))
        win._val_info_lbl = ttk.Label(toolbar,
            text=f"#{seq} {os.path.basename(file)} | Sheet={sheet} | {col_letter(col)}{row}",
            font=self.ui_font)
        win._val_info_lbl.pack(side="left")
        btn_frame = ttk.Frame(toolbar)
        btn_frame.pack(side="right")

        # 用一个 MenuButton 实现"打开文件" + 两个选项
        open_btn = ttk.Menubutton(btn_frame, text="打开文件 ▾")
        open_menu = tk.Menu(open_btn, tearoff=0, font=self.ui_font)
        open_menu.add_command(
            label="打开文件",
            command=lambda w=win: self._open_val_win_file(w, False))
        open_menu.add_command(
            label="打开并跳转单元格",
            command=lambda w=win: self._open_val_win_file(w, True))
        open_btn.configure(menu=open_menu)
        open_btn.pack(side="right")

        txt = tk.Text(win, wrap="word", font=self.ui_font, padx=8, pady=6)
        vsb = ttk.Scrollbar(win, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)
        win._val_text = txt   # 保存引用供后续复用
        txt.insert("1.0", val)
        # 高亮命中关键字（与表格一致：橙色加粗）
        kw = self._hl_keyword
        if kw:
            txt.tag_configure("hit", foreground="#e8590c", font=self.head_font)
            start, n = "1.0", len(kw)
            while True:
                pos = txt.search(kw, start, stopindex="end", nocase=1)
                if not pos:
                    break
                txt.tag_add("hit", pos, f"{pos}+{n}c")
                start = f"{pos}+{n}c"
        txt.configure(state="disabled")

    def _open_val_win_file(self, win, jump_cell=False):
        """从「完整值」窗口的打开按钮触发打开文件。"""
        full = getattr(win, "_val_full", None)
        if not full or not os.path.exists(full):
            messagebox.showwarning("提示", f"文件不存在：\n{full}")
            return
        sheet = getattr(win, "_val_sheet", None) or None
        row = getattr(win, "_val_row", None)
        row = row if isinstance(row, int) else None
        col = getattr(win, "_val_col", None)
        col = col if isinstance(col, int) else None
        threading.Thread(
            target=self._open_in_excel,
            args=(full, sheet, row, col, jump_cell),
            daemon=True).start()

    def _ctx_copy_name(self, data):
        if not data:
            return
        name = os.path.splitext(os.path.basename(data[0]))[0]   # 去掉 .xlsx 后缀
        self.root.clipboard_clear()
        self.root.clipboard_append(name)
        self._flash_status(f"已复制文件名：{name}")

    def _ctx_copy_row(self, data):
        if not data:
            return
        file, sheet, row, col, val = data
        text = "\t".join(str(x) for x in
                         (file, sheet, row, f"{col_letter(col)}({col})", val))
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self._flash_status("已复制该行到剪贴板")

    def _ctx_open_dir(self, data):
        full = self._selected_full()
        if not full:
            return
        if not os.path.exists(full):
            messagebox.showwarning("提示", f"文件不存在：\n{full}")
            return
        try:
            excel_actions.reveal_in_explorer(full)
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

    # ================================================================== #
    #  定时同步检查                                                       #
    # ================================================================== #

    def _check_sync(self):
        # 任务运行时跳过（索引正在写、DB 计数无意义），否则后台比对文件数
        if not self.busy:
            xls_dir = self.dir_var.get().strip().strip('"')
            self._controller.run_sync_probe(xls_dir)
        self.root.after(5000, self._check_sync)

    # ================================================================== #
    #  状态管理                                                           #
    # ================================================================== #

    def _save_page_size(self, n):
        self.settings["page_size"] = n
        save_settings(self.settings)

    def _save_mode(self):
        self.settings["mode"] = self.mode_var.get()
        save_settings(self.settings)

    def _cancel(self):
        if self.busy:
            self.cancel_event.set()
            self.cancel_btn.configure(state="disabled")
            self.status_var.set("正在取消…")

    def _on_close(self):
        """关闭窗口。停止所有后台活动后退出。"""
        # 停掉所有 after 回调，避免关闭过程中他们再启动新任务
        try:
            for after_id in self.root.tk.call("after", "info"):
                self.root.after_cancel(after_id)
        except Exception:
            pass
        # 通知取消
        self.cancel_event.set()
        # 清空结果表，释放大量 widget（避免 destroy 逐个回收卡顿）
        self.table.clear()
        # 有建索引任务在跑时用 os._exit 直接退出
        if self.busy:
            try:
                self.root.destroy()
            except Exception:
                pass
            os._exit(0)
        else:
            self.root.destroy()

    def _set_busy(self, busy):
        self.busy = busy
        self.search_btn.configure(state="disabled" if busy else "normal")
        if busy:
            # 任务运行时显示取消按钮（放在状态标签右侧，即最右）
            self.cancel_btn.configure(state="normal")
            self.cancel_btn.pack(side="right", padx=6, before=self.status_lbl)
            self.sync_var.set("")   # 更新/建索引期间隐藏同步提醒（此时索引正在写）
            self.kw_entry.configure(state="disabled")   # 锁定关键字输入
            self.index_combo.configure(state="disabled")  # 锁定索引操作下拉
            self._kw_popup.hide()
        else:
            self.cancel_btn.pack_forget()   # 无任务时隐藏
            self.kw_entry.configure(state="normal")
            self.index_combo.configure(state="readonly")
            self.progress["value"] = 0

    def _flash_status(self, msg, timeout=5000):
        """显示一条临时状态，timeout 毫秒后若未被覆盖则清空。"""
        self.status_var.set(msg)
        self._flash_msg = msg
        if getattr(self, "_status_after", None):
            self.root.after_cancel(self._status_after)
        self._status_after = self.root.after(timeout, self._clear_flash)

    def _clear_flash(self):
        self._status_after = None
        if self.status_var.get() == getattr(self, "_flash_msg", None):
            self.status_var.set("就绪")

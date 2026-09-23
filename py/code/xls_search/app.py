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
import xls_search.theme as theme
from xls_search.paths import col_letter, col_name_to_num, ASSETS_DIR
from xls_search.storage import (load_settings, save_settings,
                                load_keywords, save_keyword,
                                load_sources, save_sources)

from xls_search.close_dialog import CloseDialog
from xls_search.keyword_popup import KeywordPopup
from xls_search.search_controller import SearchController
from xls_search.settings_dialog import SettingsDialog
from xls_search.table_widget import ResultTable
from xls_search.tray import TrayIcon


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
        self._index_dirty = False       # sync probe 检测到索引过期
        self._stale_dismissed = False   # 本次会话用户已选「否跳过」，不再弹提示

        close_action = self.settings.get("close_action", "ask")
        self._close_action = close_action if close_action in ("ask", "exit", "tray") else "ask"

        self._controller = SearchController(self.q, self.cancel_event)

        self._build_ui()
        self._refresh_sources()
        self.root.after(80, self._poll_queue)
        self.root.after(1500, self._check_sync)   # 定时比对目录/索引文件数
        self._init_tray()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ================================================================== #
    #  UI 构建                                                            #
    # ================================================================== #

    def _build_ui(self):
        self.theme = theme.setup(self.root, self.scale)
        self.ui_font   = self.theme.font
        self.head_font = self.theme.font_bold

        self._build_toolbar()
        self._build_table()
        self._build_statusbar()

    def _open_settings(self):
        SettingsDialog(self)

    # ---------- 工具栏 ----------
    #
    # 三行（目录 / 关键字 / 模式）共用一个 grid：第 0 列是等宽的行首标签，
    # 第 1 列是会拉伸的输入区，最后一列是右对齐的按钮区。这样三行的左右
    # 边缘都在同一条竖线上，不再各行各自 pack 出参差不齐的缩进。

    def _build_toolbar(self):
        t = self.theme
        px = t.px
        bar = ttk.Frame(self.root, padding=(px(14), px(12), px(14), px(10)))
        bar.pack(fill="x")
        bar.columnconfigure(1, weight=1)
        gap_y = px(8)

        def label(text, row):
            ttk.Label(bar, text=text, style="Field.TLabel").grid(
                row=row, column=0, sticky="w", padx=(0, px(10)),
                pady=(0, gap_y))

        # ---- 第 1 行：目录 ----
        label("目录", 0)
        self.dir_var = tk.StringVar()
        self.dir_combo = ttk.Combobox(bar, textvariable=self.dir_var,
                                      font=t.font)
        self.dir_combo.grid(row=0, column=1, sticky="we", pady=(0, gap_y))
        self.dir_combo.bind("<Return>", lambda e: self._on_dir_return(e))
        self.dir_combo.bind("<<ComboboxSelected>>", lambda e: self._on_dir_selected())
        # 点击别处时目录输入框失去焦点（bind_all 在事件链最末执行，
        # 不影响 Combobox 下拉选择等内部处理）
        self.root.bind_all("<Button-1>", self._on_global_dir_click, add="+")

        dir_btns = ttk.Frame(bar)
        dir_btns.grid(row=0, column=2, sticky="e", padx=(px(8), 0),
                      pady=(0, gap_y))
        ttk.Button(dir_btns, text="浏览…", command=self._browse).pack(
            side="left")
        ttk.Button(dir_btns, text="设置", command=self._open_settings).pack(
            side="left", padx=(px(6), 0))

        # ---- 第 2 行：关键字 + 搜索 ----
        label("关键字", 1)
        kw_box, kw_inner = t.field_box(bar)
        kw_box.grid(row=1, column=1, sticky="we", pady=(0, gap_y))
        self.kw_var = tk.StringVar()
        self.kw_entry = tk.Entry(kw_inner, textvariable=self.kw_var,
                                 font=t.font_kw, relief="flat", bd=0,
                                 bg=theme.CARD, fg=theme.TEXT,
                                 insertbackground=theme.TEXT,
                                 selectbackground=theme.ACCENT_SOFT,
                                 selectforeground=theme.TEXT,
                                 highlightthickness=0)
        self.kw_entry.pack(fill="both", expand=True, padx=px(6), pady=px(5))
        self.kw_entry.bind("<Return>", lambda e: self._start_search())
        self.kw_entry.bind("<FocusIn>", lambda e: self._on_kw_focus(kw_box, True))
        self.kw_entry.bind("<FocusOut>", lambda e: self._on_kw_focus(kw_box, False))

        # 历史下拉：点击输入框弹出、输入时实时筛选（自定义弹层，不抢输入焦点）
        self._kw_popup = KeywordPopup(
            root=self.root,
            entry=self.kw_entry,
            kw_var=self.kw_var,
            ui_font=self.ui_font,
            on_pick=self._on_kw_pick,
            is_busy=lambda: self.busy,
            anchor=kw_box,
        )
        self._refresh_keywords()

        self.search_btn = ttk.Button(bar, text="搜索", style="Accent.TButton",
                                     command=self._start_search)
        self.search_btn.grid(row=1, column=2, sticky="nsew", padx=(px(8), 0),
                             pady=(0, gap_y))

        # ---- 第 3 行：模式 + 过滤条件 ----
        label("模式", 2)
        opts = ttk.Frame(bar)
        opts.grid(row=2, column=1, columnspan=2, sticky="we")

        saved_mode = self.settings.get("mode", "2")
        self.mode_var = tk.StringVar(
            value=saved_mode if saved_mode in ("0", "1", "2") else "2")
        seg = tk.Frame(opts, bg=theme.BORDER, bd=0, highlightthickness=0)
        seg.grid(row=0, column=0, sticky="w")
        seg_in = tk.Frame(seg, bg=theme.BG, bd=0, highlightthickness=0)
        seg_in.pack(padx=1, pady=1)
        for val, text in [("0", "文件名"), ("1", "内容·逐文件"), ("2", "内容·索引")]:
            ttk.Radiobutton(seg_in, text=text, value=val,
                            variable=self.mode_var, style="Toolbutton",
                            takefocus=False).pack(side="left")
        self.mode_var.trace_add("write", lambda *a: self._save_mode())

        # 索引操作：菜单按钮（原来的 readonly Combobox 会把选项写进框里，
        # 还得手动复位占位文字，换成 Menubutton 更贴合"执行动作"的语义）
        self.index_btn = ttk.Menubutton(opts, text="索引操作")
        index_menu = tk.Menu(self.index_btn, tearoff=0, font=t.font)
        index_menu.add_command(label="更新变动索引",
                               command=lambda: self._on_index_action("3"))
        index_menu.add_command(label="重建全部索引",
                               command=lambda: self._on_index_action("4"))
        self.index_btn.configure(menu=index_menu)
        self.index_btn.grid(row=0, column=1, sticky="w", padx=(px(10), px(16)))

        self.exact_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="精确匹配", variable=self.exact_var,
                        takefocus=False).grid(row=0, column=2, sticky="w")

        ttk.Label(opts, text="文件名含", style="Field.TLabel").grid(
            row=0, column=3, sticky="w", padx=(px(16), px(6)))
        self.filter_var = tk.StringVar()
        ttk.Entry(opts, textvariable=self.filter_var, font=t.font,
                  width=12).grid(row=0, column=4, sticky="w")
        ttk.Label(opts, text="限定列", style="Field.TLabel").grid(
            row=0, column=5, sticky="w", padx=(px(14), px(6)))
        self.col_var = tk.StringVar()
        ttk.Entry(opts, textvariable=self.col_var, font=t.font,
                  width=5).grid(row=0, column=6, sticky="w")

        t.hline(self.root).pack(fill="x")

    def _on_kw_focus(self, box, focused):
        """关键字框聚焦时把外框边线染成主题蓝，和 ttk.Entry 的表现保持一致。"""
        box.configure(bg=theme.ACCENT if focused else theme.BORDER)
        if focused:
            self._set_kw_ime_font()

    # ---------- 表格区 ----------

    def _build_table(self):
        s = self.scale
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
            ["#",     "#",     _w("#", 52),      "center"],
            ["file",  "文件",  _w("file", 300),  "w"],
            ["sheet", "Sheet", _w("sheet", 150), "center"],
            ["row",   "行",    _w("row", 64),    "center"],
            ["col",   "列",    _w("col", 86),    "center"],
            ["value", "值",    None,             "w"],
        ]

        px = self.theme.px
        table_frame = ttk.Frame(self.root, padding=(px(14), px(10),
                                                    px(14), px(4)))
        table_frame.pack(fill="both", expand=True)

        self.table = ResultTable(
            parent=table_frame,
            col_spec=col_spec,
            scale=s,
            ui_font=self.ui_font,
            head_font=self.head_font,
            hl_color=theme.HL,                 # 关键字命中
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

    def _build_statusbar(self):
        px = self.theme.px
        self.theme.hline(self.root).pack(fill="x")
        bottom = ttk.Frame(self.root, padding=(px(14), px(7), px(14), px(7)))
        bottom.pack(fill="x")
        # 状态文字放左边（读起来顺），进度条占中间的弹性空间，取消按钮在最右
        self.status_var = tk.StringVar(value="就绪")
        self.status_lbl = ttk.Label(bottom, textvariable=self.status_var,
                                    style="Hint.TLabel", anchor="w")
        self.status_lbl.pack(side="left")

        # 索引过期提醒：原先挤在过滤条件行的最右端，会把「限定列」输入框顶出
        # 窗口；这里挪到状态栏右侧，横向空间充裕，也更符合"提示信息"的定位
        self.sync_var = tk.StringVar(value="")
        ttk.Label(bottom, textvariable=self.sync_var, style="Warn.TLabel",
                  anchor="e").pack(side="right", padx=(px(12), 0))
        # 取消按钮先创建但不 pack —— 无任务时隐藏，busy 时才显示（见 _set_busy）
        self.cancel_btn = ttk.Button(bottom, text="取消", command=self._cancel)
        self.progress = ttk.Progressbar(bottom, mode="determinate",
                                        style="Slim.Horizontal.TProgressbar")
        self._progress_pack = dict(side="left", fill="x", expand=True,
                                   padx=(px(14), px(4)), pady=px(5))
        # 进度条同样只在 busy 时出现，空闲时状态栏就是一行干净的提示文字

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
        # 切换到新目录，重置索引脏状态和弹窗抑制标志
        self._index_dirty = False
        self._stale_dismissed = False
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
        # 让正在输入的拼音字体跟随关键字框，避免上屏前后字号跳变
        ime.set_composition_font(self.kw_entry, self.theme.kw_px,
                                 family=theme.FAMILY, weight=400)

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
        col_txt = self.col_var.get().strip()
        col_filter = col_name_to_num(col_txt) if col_txt else None

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

        # 模式2 且索引有更新（定时器检测到）-> 询问是否先更新
        if mode == "2" and self._index_dirty and not self._stale_dismissed:
            ans = messagebox.askyesno(
                "索引已过期",
                "检测到文件有变动，索引可能不完整。\n"
                "是否先更新索引再搜索？\n\n"
                "（选「否」本次不再提示，下次启动程序后恢复提醒）")
            if ans:
                mode = "3"      # 更新变动索引再搜索
            else:
                self._stale_dismissed = True    # 本次会话不再弹

        self._hl_keyword = keyword      # 供结果渲染时高亮命中字
        self.cancel_event.clear()
        self._set_busy(True)
        self.table.clear()
        self.table.set_empty_text("正在搜索…")
        self._controller.run_search(
            xls_dir=xls_dir, keyword=keyword, exact=exact,
            filter_str=filter_str, col_filter=col_filter, mode=mode)

    def _on_index_action(self, mode):
        """索引菜单项：mode "3"=更新变动，"4"=重建全部。确认后后台执行。"""
        if self.busy:
            return
        if mode == "4":
            ok = messagebox.askyesno(
                "确认重建",
                "将清空并重新建立全部索引，耗时较久。\n确定重建全部索引吗？")
        else:
            ok = messagebox.askyesno(
                "确认更新",
                "将只重建有变动的文件索引。\n确定更新变动索引吗？")
        if ok:
            self._start_index_build(mode)

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
        self.table.set_empty_text("正在处理索引…")
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
                    # 索引刚建完（或重建完），脏状态清零
                    self._index_dirty = False
                    self._stale_dismissed = False
                elif kind == "results":
                    self._show_results(payload)
                elif kind == "error":
                    messagebox.showerror("出错", payload)
                    self.status_var.set("出错: " + payload)
                elif kind == "cancelled":
                    self.status_var.set(payload or "已取消")
                elif kind == "sync":
                    self.sync_var.set(payload)
                    # probe 有结果时同步维护脏标志（payload 非空 = 有变动）
                    self._index_dirty = bool(payload)
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
        self.status_var.set("没有匹配的内容" if total == 0 else f"找到 {total} 条结果")

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
        row = row if isinstance(row, int) and row > 0 else None
        col = col if isinstance(col, int) and col > 0 else None
        # COM 启动 Excel 可能较慢，放后台线程；失败则退回默认打开
        threading.Thread(
            target=self._open_in_excel,
            args=(full, sheet, row, col, jump_cell),
            daemon=True).start()

    def _open_in_excel(self, full, sheet, row, col, jump_cell=False):
        excel_actions.open_in_excel(
            full, sheet, row, col, jump_cell=jump_cell,
            on_status=lambda msg: self.q.put(("status", msg)))

    @staticmethod
    def _val_info_text(seq, file, sheet, row, col):
        """「完整值」窗口工具栏里的那行上下文信息。"""
        return (f"#{seq} {os.path.basename(file)}\n"
                f"Sheet={sheet} | {col_letter(col)}{row}")

    def _wrap_val_info(self, event, win):
        """按工具栏实际宽度设定信息标签的换行宽度。

        ttk.Label 的 wraplength 只认固定像素，窗口可缩放就得跟着变；减去
        「打开文件」按钮和内边距占掉的宽度，剩下多少才是文字能用的。
        """
        lbl = getattr(win, "_val_info_lbl", None)
        if lbl is None or not lbl.winfo_exists():
            return
        btn = lbl.master.grid_slaves(row=0, column=1)
        reserved = (btn[0].winfo_reqwidth() if btn else 0) + self.theme.px(36)
        lbl.configure(wraplength=max(self.theme.px(80), event.width - reserved))

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
                    text=self._val_info_text(seq, file, sheet, row, col))
            txt = old._val_text
            txt.configure(state="normal")
            txt.delete("1.0", "end")
            txt.insert("1.0", val)
            kw = self._hl_keyword
            if kw:
                txt.tag_delete("hit")
                txt.tag_configure("hit", foreground=theme.HL, font=self.head_font)
                start, n = "1.0", len(kw)
                while True:
                    pos = txt.search(kw, start, stopindex="end", nocase=1)
                    if not pos:
                        break
                    txt.tag_add("hit", pos, f"{pos}+{n}c")
                    start = f"{pos}+{n}c"
            txt.configure(state="disabled")
            # 主窗口挪过位置就跟着重新贴边；没挪过则保留用户自己拖的位置
            theme.reposition_if_parent_moved(old, self.root,
                                             theme.place_beside)
            old.deiconify()
            old.lift()
            return

        win = tk.Toplevel(self.root)
        self._view_val_win = win
        win.configure(bg=theme.BG)
        win.title(f"完整值 — #{seq} {os.path.basename(file)}  {col_letter(col)}{row}")
        s = self.scale
        px = self.theme.px
        # 只给尺寸不给位置时，窗口落点由系统决定（看着像随机）。这里贴在主窗口
        # 左侧并排、与主窗口等高，方便一边看列表一边看完整值。
        theme.place_beside(win, self.root, int(360 * s),
                           self.root.winfo_height())
        # 保存上下文信息到窗口对象，供工具栏按钮回调使用
        win._val_full = full
        win._val_sheet = sheet
        win._val_row = row
        win._val_col = col

        # 顶部工具栏：信息 + 打开文件按钮（带下拉选项）。
        # 用 grid 而不是 pack —— 侧边窗口窄，pack 的 side="left" 标签会按自身
        # 请求宽度占满，把右侧按钮挤到不显示；grid 里按钮独占一列才挤不掉。
        toolbar = ttk.Frame(win, padding=(px(12), px(10), px(12), px(8)))
        toolbar.pack(fill="x")
        toolbar.columnconfigure(0, weight=1)

        # 用一个 MenuButton 实现"打开文件" + 两个选项
        open_btn = ttk.Menubutton(toolbar, text="打开文件")
        open_menu = tk.Menu(open_btn, tearoff=0, font=self.ui_font)
        open_menu.add_command(
            label="打开文件",
            command=lambda w=win: self._open_val_win_file(w, False))
        open_menu.add_command(
            label="打开并跳转单元格",
            command=lambda w=win: self._open_val_win_file(w, True))
        open_btn.configure(menu=open_menu)
        open_btn.grid(row=0, column=1, sticky="e", padx=(px(8), 0))

        # 信息文字放不下就换行（窄窗口下文件名往往比一行长）
        win._val_info_lbl = ttk.Label(
            toolbar, text=self._val_info_text(seq, file, sheet, row, col),
            style="Field.TLabel", justify="left", anchor="w")
        win._val_info_lbl.grid(row=0, column=0, sticky="we")
        toolbar.bind("<Configure>", lambda e, w=win: self._wrap_val_info(e, w))

        self.theme.hline(win).pack(fill="x")
        body = ttk.Frame(win, padding=(px(12), px(10), px(4), px(12)))
        body.pack(fill="both", expand=True)
        txt = tk.Text(body, wrap="word", font=self.ui_font, padx=px(10),
                      pady=px(8), bg=theme.CARD, fg=theme.TEXT, relief="flat",
                      bd=0, highlightthickness=1,
                      highlightbackground=theme.BORDER,
                      highlightcolor=theme.BORDER,
                      selectbackground=theme.ACCENT_SOFT,
                      selectforeground=theme.TEXT)
        vsb = ttk.Scrollbar(body, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", padx=(px(4), 0))
        txt.pack(side="left", fill="both", expand=True)
        win._val_text = txt   # 保存引用供后续复用
        txt.insert("1.0", val)
        # 高亮命中关键字（与表格一致）
        kw = self._hl_keyword
        if kw:
            txt.tag_configure("hit", foreground=theme.HL, font=self.head_font)
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
        row = row if isinstance(row, int) and row > 0 else None
        col = getattr(win, "_val_col", None)
        col = col if isinstance(col, int) and col > 0 else None
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
    #  系统托盘                                                           #
    # ================================================================== #

    def _init_tray(self):
        ico_path = os.path.join(ASSETS_DIR, "app.ico")
        self._tray = TrayIcon(ico_path, "xls_search",
                              on_show=self._restore_from_tray,
                              on_quit=self._quit_from_tray)
        if not self._tray.ok:
            return  # 没有 pywin32 等情况下保持原生关闭即退出行为

        # 图标在程序启动时就挂上，整个运行期间常驻，不随窗口显隐增删。
        self._tray.show()

        # Tk 没有"最小化"事件，只能监听 <Unmap>：窗口被最小化时会触发，
        # 此时 state() 是 "iconic"。注意子控件也会冒泡 Unmap，需按 widget 过滤。
        self.root.bind("<Unmap>", self._on_unmap)
        self._pump_tray()

    def _pump_tray(self):
        """在 Tk 主循环里抽取托盘消息，避免另起线程碰 Tk 控件。"""
        self._tray.pump()
        self._tray_job = self.root.after(100, self._pump_tray)

    def _on_unmap(self, event):
        if event.widget is not self.root:
            return
        if self.root.state() == "iconic":
            self._minimize_to_tray()

    def _restore_from_tray(self):
        """点托盘图标唤出窗口。图标常驻，这里只管窗口，不动托盘。"""
        self.root.deiconify()
        self.root.state("normal")
        self.root.lift()
        self.root.focus_force()

    def _quit_from_tray(self):
        self._shutdown()

    def _minimize_to_tray(self):
        self.root.withdraw()          # 从任务栏移除，托盘图标本来就在

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
        """点 × 时：托盘不可用直接退出；否则按记忆的选择，或弹窗询问退出/最小化。"""
        if not getattr(self, "_tray", None) or not self._tray.ok:
            self._shutdown()
            return

        action = self._close_action
        if action == "ask":
            dlg = CloseDialog(self.root, scale=self.scale, font=self.ui_font)
            if dlg.result is None:
                return                      # 用户取消，什么都不做
            action = dlg.result
            if dlg.remember:
                self._close_action = action
                self.settings["close_action"] = action
                save_settings(self.settings)

        if action == "tray":
            self._minimize_to_tray()
        else:
            self._shutdown()

    def _shutdown(self):
        """停止所有后台活动、注销托盘图标后退出。"""
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
        if getattr(self, "_tray", None):
            self._tray.destroy()
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
            # 任务运行时才显示进度条和取消按钮
            self.progress.pack(**self._progress_pack)
            self.cancel_btn.configure(state="normal")
            self.cancel_btn.pack(side="left")
            self.sync_var.set("")   # 更新/建索引期间隐藏同步提醒（此时索引正在写）
            self.kw_entry.configure(state="disabled")   # 锁定关键字输入
            self.index_btn.configure(state="disabled")  # 锁定索引操作
            self._kw_popup.hide()
        else:
            self.progress["value"] = 0
            self.progress.pack_forget()
            self.cancel_btn.pack_forget()   # 无任务时隐藏
            self.kw_entry.configure(state="normal")
            self.index_btn.configure(state="normal")

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

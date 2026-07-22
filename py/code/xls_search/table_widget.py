# -*- coding: utf-8 -*-
"""分页结果表格。

每页只渲染当前页的行（grid 布局），翻页按钮切换页面，
因此无论结果多少条都不会耗尽 GDI / 控件资源。

外部接口
--------
ResultTable(parent, col_spec, scale, ui_font, head_font,
            hl_color, on_open_file, on_view_value,
            on_copy_name, on_open_dir, on_copy_row)
    .set_rows(rows, hl_keyword="")   —— 加载新结果并渲染
    .clear()                         —— 清空
    .sort_by(col_key)                —— 按列排序
    .save_col_px(settings, save_fn)  —— 把当前列宽写入设置
    .set_col_resize_callback(fn)     —— 注册列宽拖动完成回调
    .sel_index                       —— 当前选中行下标（只读）
    .all_rows                        —— 当前全部结果行（只读）
"""
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

from xls_search.paths import col_letter


class ResultTable:
    VAL_LINES = 2       # 「值」列最多显示行数

    def __init__(self, parent, col_spec, scale, ui_font, head_font,
                 hl_color,
                 on_open_file, on_view_value,
                 on_copy_name, on_open_dir, on_copy_row):
        """
        参数
        ----
        parent      : ttk.Frame     —— 父容器（由 App 创建）
        col_spec    : list          —— [key, title, width_px|None, anchor]
        scale       : float         —— DPI 缩放倍数
        ui_font     : tuple         —— 正文字体
        head_font   : tuple         —— 表头字体（加粗）
        hl_color    : str           —— 关键字高亮颜色
        on_*        : callable      —— 右键菜单各项的回调，均接收 (row_data) 元组
        """
        s = scale
        self._scale = s
        self._ui_font = ui_font
        self._head_font = head_font
        self._hl_color = hl_color
        self._hl_keyword = ""
        self._hl_kw = ""
        self._hl_idx = 0
        self._hl_tag_name = "hit"

        self.col_spec = col_spec
        self._value_ci = len(col_spec) - 1
        self._sep_w = 1
        self._min_col = int(round(40 * s))

        self._bg_even, self._bg_odd = "#ffffff", "#f3f6f9"
        self._bg_sel, self._fg_sel = "#3399ff", "#ffffff"

        self.all_rows = []
        self.sel_index = None
        self.sort_state = {}

        self._on_open_file = on_open_file
        self._on_view_value = on_view_value
        self._on_copy_name = on_copy_name
        self._on_open_dir = on_open_dir
        self._on_copy_row = on_copy_row

        # --- 测量字体 ---
        self._ui_font_obj = tkfont.Font(font=ui_font)
        self._line_h = self._ui_font_obj.metrics("linespace")
        self._char_w = max(1, self._ui_font_obj.measure("中"))

        # 分页状态
        self._page_size = 50          # 每页条数（用户可选）
        self._cur_page = 1           # 当前页码（1-based）
        self._total_pages = 1        # 总页数
        self._pending_page = None    # 待渲染的目标页（合并快速连击）

        self._resize_margin = max(10, int(round(12 * s)))
        self._edge_col = None
        self._resizing = None
        self._drag_x0 = 0
        self._drag_w0 = 0

        self._parent = parent
        self._on_col_resize_done = None
        self._on_page_size_save = None

        # 当前页渲染的行控件：(widget_row_frame, labels[], value_text_widget)
        self._page_widgets = []

        self._build(parent)

    # ------------------------------------------------------------------ #
    #  构建控件                                                            #
    # ------------------------------------------------------------------ #

    def _build(self, parent):
        s = self._scale

        def _config_cols(frame):
            for i, (_, _, w, _) in enumerate(self.col_spec):
                if w is None:
                    frame.columnconfigure(i * 2, weight=1,
                                          minsize=int(round(160 * s)))
                else:
                    frame.columnconfigure(i * 2, weight=0, minsize=w)
            for i in range(len(self.col_spec) - 1):
                frame.columnconfigure(i * 2 + 1, weight=0, minsize=self._sep_w)

        # 表头（固定）—— 用 tk.Frame 与 rows_frame 保持一致，避免 ttk 主题 padding 造成列错位
        self.head_frame = tk.Frame(parent, bd=0, highlightthickness=0)
        self.head_frame.grid(row=0, column=0, sticky="we")
        _config_cols(self.head_frame)
        self._head_labels = []
        self._tooltip = None
        for i, (key, title, w, anchor) in enumerate(self.col_spec):
            hl = tk.Label(self.head_frame, text=title, font=self._head_font,
                          anchor="center", padx=6, pady=4, cursor="hand2")
            hl.grid(row=0, column=i * 2, sticky="we")
            hl.bind("<Motion>",          lambda e, idx=i: self._head_motion(e, idx))
            hl.bind("<Leave>",           lambda e, idx=i: self._head_leave(idx))
            hl.bind("<Button-1>",        lambda e, idx=i, k=key: self._head_press(e, idx, k))
            hl.bind("<B1-Motion>",       lambda e, idx=i: self._head_drag(e, idx))
            hl.bind("<ButtonRelease-1>", lambda e, idx=i: self._head_release(e, idx))
            self._head_labels.append(hl)

        # 表头竖线（1px）
        for i in range(len(self.col_spec) - 1):
            tk.Frame(self.head_frame, bg="#c8c8c8", width=1).grid(
                row=0, column=i * 2 + 1, sticky="ns")

        # 行区域：Canvas + 滚动条，内容超出时可滚动
        self._rows_canvas = tk.Canvas(parent, highlightthickness=0, bg=self._bg_even)
        vsb = ttk.Scrollbar(parent, orient="vertical", command=self._rows_canvas_yview)
        self._rows_canvas.configure(yscrollcommand=vsb.set)
        self._rows_canvas.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)
        self._rows_canvas.bind("<Configure>", self._on_rows_canvas_configure)
        self._rows_canvas.bind("<MouseWheel>", self._on_wheel_canvas)
        # 行控件填充了 canvas，滚轮事件会被子控件拦截，因此 bind_all 兜底；
        # 但只当鼠标在 canvas 区域内才滚动，不影响输入框等其他区域
        self._rows_canvas.bind_all("<MouseWheel>", self._on_wheel_global)

        self.rows_frame = tk.Frame(self._rows_canvas, bg=self._bg_even,
                                   bd=0, highlightthickness=0)
        _config_cols(self.rows_frame)
        self._rows_window = self._rows_canvas.create_window(
            (0, 0), window=self.rows_frame, anchor="nw")

        # 底部翻页栏
        nav = ttk.Frame(parent)
        nav.grid(row=2, column=0, sticky="ew", pady=(2, 0))
        # 左侧：页码导航
        left = ttk.Frame(nav)
        left.pack(side="left")
        self._nav_prev = ttk.Button(left, text="«", width=3,
                                    command=lambda: self._jump_page(self._cur_page - 1))
        self._nav_prev.pack(side="left", padx=1)
        # 8 个固定槽位始终 pack，保持宽度不变，不用的设空文本（不可见但占位）
        self._page_btns = []
        self._ellipsis_left = tk.Label(left, text="", font=self._ui_font,
                                        bg="#e8e8e8", width=1)
        self._ellipsis_left.pack(side="left", padx=0)
        for _ in range(6):
            btn = tk.Button(left, text="", width=4, font=self._ui_font,
                           relief="flat", bd=0)
            btn.pack(side="left", padx=1)
            self._page_btns.append(btn)
        self._ellipsis_right = tk.Label(left, text="", font=self._ui_font,
                                         bg="#e8e8e8", width=1)
        self._ellipsis_right.pack(side="left", padx=0)
        self._nav_next = ttk.Button(left, text="»", width=3,
                                    command=lambda: self._jump_page(self._cur_page + 1))
        self._nav_next.pack(side="left", padx=1)
        # 右侧：每页条数 + 跳转 + 总页数
        right = ttk.Frame(nav)
        right.pack(side="right")
        ttk.Label(right, text="每页", font=self._ui_font).pack(side="left", padx=(8, 2))
        self._page_size_var = tk.StringVar(value="50")
        self._page_size_combo = ttk.Combobox(
            right, textvariable=self._page_size_var,
            values=["10", "20", "50", "100", "200", "500"],
            state="readonly", width=5, font=self._ui_font)
        self._page_size_combo.pack(side="left", padx=2)
        self._page_size_combo.bind("<<ComboboxSelected>>", self._on_page_size_change)
        ttk.Label(right, text="跳转", font=self._ui_font).pack(side="left", padx=(8, 2))
        self._jump_var = tk.StringVar()
        self._jump_entry = ttk.Entry(right, textvariable=self._jump_var, width=5,
                                     font=self._ui_font)
        self._jump_entry.pack(side="left", padx=2)
        self._jump_entry.bind("<Return>", lambda e: self._jump_by_input())
        ttk.Button(right, text="Go", width=3,
                   command=self._jump_by_input).pack(side="left", padx=2)
        ttk.Label(right, text="总", font=self._ui_font).pack(side="left", padx=(8, 2))
        self._total_pages_var = tk.StringVar(value="0 页")
        ttk.Label(right, textvariable=self._total_pages_var,
                  font=self._ui_font, width=6, anchor="center").pack(side="left")

        # 右键菜单
        self.ctx_menu = tk.Menu(parent, tearoff=0, font=self._ui_font)
        self.ctx_menu.add_command(
            label="打开文件",
            command=lambda: self._on_open_file(self._selected_data(), jump_cell=False))
        self.ctx_menu.add_command(
            label="打开并跳转单元格",
            command=lambda: self._on_open_file(self._selected_data(), jump_cell=True))
        self.ctx_menu.add_command(
            label="查看完整值",
            command=lambda: self._on_view_value(self._selected_data(),
                                                self.sel_index, self))
        self.ctx_menu.add_command(
            label="复制文件名",
            command=lambda: self._on_copy_name(self._selected_data()))
        self.ctx_menu.add_command(
            label="打开所在目录",
            command=lambda: self._on_open_dir(self._selected_data()))

    # ------------------------------------------------------------------ #
    #  公共接口                                                            #
    # ------------------------------------------------------------------ #

    def set_rows(self, rows, hl_keyword=""):
        self._hl_keyword = hl_keyword
        # 先清空旧页控件再设新数据，避免新旧 widget 混在一起
        self._destroy_page_widgets()
        self.all_rows = list(rows)
        self._cur_page = 1
        self._render()

    def clear(self):
        self.all_rows = []
        self.sel_index = None
        self._destroy_page_widgets()
        self._total_pages = 1
        self._total_pages_var.set("0 页")

    def sort_by(self, col):
        _COL_IDX = {"#": None, "file": 0, "sheet": 1, "row": 2, "col": 3, "value": 4}
        idx = _COL_IDX[col]
        if idx is None:
            return
        desc = not self.sort_state.get(col, False)
        self.sort_state[col] = desc

        def key(r):
            return r[idx] if col in ("row", "col") else str(r[idx]).lower()

        self.all_rows.sort(key=key, reverse=desc)
        self._cur_page = 1
        self._render()

    def save_col_px(self, settings, save_fn):
        d = {key: w for key, _, w, _ in self.col_spec if w}
        if d != settings.get("col_px"):
            settings["col_px"] = d
            save_fn(settings)

    def set_col_resize_callback(self, fn):
        self._on_col_resize_done = fn

    def set_page_size_callback(self, fn):
        """注册每页条数变更回调（App 调用，用于持久化）。"""
        self._on_page_size_save = fn

    def set_page_size(self, n):
        """设置每页条数（从持久化设置恢复）。只接受标准值，其余用 50。"""
        if n in (10, 20, 50, 100, 200, 500):
            self._page_size = n
            self._page_size_var.set(str(n))
        else:
            self._page_size = 50
            self._page_size_var.set("50")

    # ------------------------------------------------------------------ #
    #  选中行                                                              #
    # ------------------------------------------------------------------ #

    def _selected_data(self):
        if self.sel_index is None or self.sel_index >= len(self.all_rows):
            return None
        return self.all_rows[self.sel_index]

    def _select_row(self, global_idx):
        """选中全局行号，更新当前页高亮。"""
        if global_idx is None:
            return
        prev = self.sel_index
        self.sel_index = global_idx
        # 刷新旧选中和新选中的行样式
        start = (self._cur_page - 1) * self._page_size
        end = min(len(self.all_rows), start + self._page_size)
        for idx in (prev, global_idx):
            if idx is not None and start <= idx < end:
                wi = idx - start
                self._style_row(wi, idx)

    def _row_style(self, idx):
        if idx == self.sel_index:
            return self._bg_sel, self._fg_sel
        return (self._bg_odd if idx % 2 else self._bg_even), "black"

    # ------------------------------------------------------------------ #
    #  渲染                                                                #
    # ------------------------------------------------------------------ #

    def _render(self):
        """渲染当前页。复用已有控件，只更新文本/绑定/样式。"""
        self.sel_index = None

        if not self.all_rows:
            self._total_pages = 1
            self._update_page_bar()
            return

        self._total_pages = max(1, (len(self.all_rows) + self._page_size - 1) // self._page_size)
        self._cur_page = min(self._cur_page, self._total_pages)

        start = (self._cur_page - 1) * self._page_size
        end = min(len(self.all_rows), start + self._page_size)
        new_count = end - start
        old_count = len(self._page_widgets)

        # --- 销毁多余的旧行 ---
        if new_count < old_count:
            for i in range(new_count, old_count):
                cells, _ = self._page_widgets[i]
                for c in cells:
                    c.grid_forget()
                    c.destroy()
            self._page_widgets = self._page_widgets[:new_count]
            for child in list(self.rows_frame.winfo_children()):
                if not isinstance(child, tk.Frame):
                    continue
                try:
                    r = child.grid_info().get("row")
                except Exception:
                    continue
                if isinstance(r, int) and r >= new_count:
                    child.destroy()

        # --- 更新已有行（只改文本/样式/绑定，不重建控件） ---
        for i in range(min(new_count, old_count)):
            self._update_row(i, start + i)

        # --- 创建缺少的新行 ---
        for i in range(old_count, new_count):
            self._make_row(i, start + i)

        # 立即更新 scrollregion，避免 after_idle 延迟导致的闪跳
        self.rows_frame.update_idletasks()
        self._update_scrollregion()
        self._rows_canvas.yview_moveto(0)   # 翻页后滚回顶部
        self._update_page_bar()

        # 异步补齐高亮，避免首次渲染卡顿
        self._apply_hl_async()

    def _bind_row_events(self, c, ci, global_idx, file, val, texts, col_w):
        """为单个单元格控件绑定行交互事件（Button-1/Double-1/Button-3/Enter/Leave）。"""
        for seq in ("<Button-1>", "<Double-1>", "<Button-3>", "<Enter>", "<Leave>"):
            c.unbind(seq)

        c.bind("<Button-1>", lambda e, gi=global_idx: self._select_row(gi))
        c.bind("<Button-3>", lambda e, gi=global_idx: self._on_right_click(e, gi))

        if ci == self._value_ci:
            c.bind("<Double-1>", lambda e, gi=global_idx: self._on_dbl_value(gi))
            tip_val = val if len(val) <= 300 else val[:300] + "......"
            c.bind("<Enter>", lambda e, t="— 双击显示全部 —", h=tip_val: self._show_tooltip(e, t, header=h))
            c.bind("<Leave>", lambda e: self._hide_tooltip())
        elif ci == 1:
            c.bind("<Double-1>", lambda e, gi=global_idx: self._on_dbl_file(gi))
            c.bind("<Enter>", lambda e, t="— 双击打开文件 —", h=file: self._show_tooltip(e, t, header=h))
            c.bind("<Leave>", lambda e: self._hide_tooltip())
        else:
            c.bind("<Double-1>", lambda e, gi=global_idx: self._on_dbl(gi))
            if col_w is not None:
                displayed = self._fit_col_text(texts[ci], col_w, middle=False)
                if displayed != texts[ci]:
                    c.bind("<Enter>", lambda e, t=texts[ci]: self._show_tooltip(e, t))
                    c.bind("<Leave>", lambda e: self._hide_tooltip())

    def _update_row(self, page_row, global_idx):
        """更新已有行的文本、样式、事件绑定（不销毁重建）。"""
        file, sheet, row, col, val = self.all_rows[global_idx]
        val_disp = self._value_snippet(val, self._col_width(self._value_ci))
        texts = [str(global_idx + 1), file, sheet, str(row),
                 f"{col_letter(col)}({col})", val_disp]
        bg, fg = self._row_style(global_idx)

        cells, _ = self._page_widgets[page_row]
        self._page_widgets[page_row] = (cells, global_idx)

        for ci, (key, _, w, anchor) in enumerate(self.col_spec):
            c = cells[ci]
            if isinstance(c, tk.Text):
                c.configure(state="normal", bg=bg, fg=fg)
                c.delete("1.0", "end")
                c.insert("1.0", texts[ci])
                c.configure(state="disabled")
                c.configure(wrap="word")
            else:
                c.configure(text=self._fit_col_text(texts[ci], w, middle=(ci == 1)),
                            bg=bg, fg=fg)

            self._bind_row_events(c, ci, global_idx, file, val, texts, w)

    def _make_row(self, page_row, global_idx):
        """创建一行控件（grid 在 rows_frame 的 page_row 行）。"""
        file, sheet, row, col, val = self.all_rows[global_idx]
        val_disp = self._value_snippet(val, self._col_width(self._value_ci))
        texts = [str(global_idx + 1), file, sheet, str(row),
                 f"{col_letter(col)}({col})", val_disp]
        bg, fg = self._row_style(global_idx)

        cells = []
        for ci, (key, _, w, anchor) in enumerate(self.col_spec):
            if ci == self._value_ci:
                c = tk.Text(self.rows_frame, wrap="word", font=self._ui_font,
                            bd=0, highlightthickness=0, padx=6, pady=3,
                            cursor="arrow", height=self.VAL_LINES, width=1,
                            bg=bg, fg=fg)
                c.insert("1.0", texts[ci])
                c.configure(state="disabled")
                c.configure(wrap="word")
            else:
                just = {"w": "left", "e": "right", "center": "center"}[anchor]
                anc  = {"w": "nw",   "e": "ne",    "center": "n"}[anchor]
                c = tk.Label(self.rows_frame, font=self._ui_font,
                             anchor=anc, justify=just, padx=6, pady=3,
                             bg=bg, fg=fg)
                c.configure(text=self._fit_col_text(texts[ci], w, middle=(ci == 1)))
            c.grid(row=page_row, column=ci * 2, sticky="we")
            self._bind_row_events(c, ci, global_idx, file, val, texts, w)
            cells.append(c)

        # 竖线分隔
        for ci in range(len(self.col_spec) - 1):
            sep = tk.Frame(self.rows_frame, bg="#e0e0e0", width=1, height=1)
            sep.grid(row=page_row, column=ci * 2 + 1, sticky="ns")

        self._page_widgets.append((cells, global_idx))

    def _style_row(self, page_row, global_idx):
        """刷新某行的样式（选中/取消选中后调用）。"""
        bg, fg = self._row_style(global_idx)
        for c in self._page_widgets[page_row][0]:
            c.configure(bg=bg, fg=fg)
            if isinstance(c, tk.Text):
                c.configure(state="normal", bg=bg, fg=fg)
                # 刷新 value snippet (选中可能改变高亮)
                c.configure(state="disabled")

    def _fit_col_text(self, text, col_px, middle=False):
        """把固定列文本裁到列宽内（超出加 …），保证内容列不撑宽于表头列。

        普通列尾部裁剪；middle=True（文件列）保留头尾、中间省略，
        以便文件名后缀（.xlsx）仍可见。用字体实测宽度，任意字符串都可靠。
        """
        text = str(text)
        f = self._ui_font_obj
        inner = max(10, col_px - 14)          # 减去 padx(6*2) 再留 2px 余量
        if f.measure(text) <= inner:
            return text
        ell = "…"
        avail = max(0, inner - f.measure(ell))
        if not middle:
            lo, hi = 0, len(text)
            while lo < hi:
                mid = (lo + hi + 1) // 2
                if f.measure(text[:mid]) <= avail:
                    lo = mid
                else:
                    hi = mid - 1
            return text[:lo] + ell
        # 中间省略：头尾对半分
        half = avail // 2
        h = 0
        while h < len(text) and f.measure(text[:h + 1]) <= half:
            h += 1
        t = len(text)
        while t > h and f.measure(text[t - 1:]) <= half:
            t -= 1
        return text[:h] + ell + text[t:]

    def _col_width(self, ci):
        """列 ci 的像素宽（估算，用于 Text wraplength）。"""
        spec = self.col_spec[ci]
        if spec[2] is None:
            # 填充列：用 rows_frame 当前宽减去其它列
            total = self.rows_frame.winfo_width()
            for i, (_, _, fw, _) in enumerate(self.col_spec):
                if i != ci and fw is not None:
                    total -= fw + self._sep_w
            return max(40, total)
        return spec[2]

    def _destroy_page_widgets(self):
        """销毁当前页所有行控件及分隔线。"""
        for cells, _ in self._page_widgets:
            for c in cells:
                c.destroy()
        self._page_widgets = []
        # 清掉 rows_frame 里所有残留子控件（分隔线等）
        for child in list(self.rows_frame.winfo_children()):
            child.destroy()

    def _value_snippet(self, val, col_px):
        """取以关键字为中心、约 VAL_LINES 行长度的片段。"""
        inner = max(20, col_px - 14)
        cpl = max(4, int(inner // self._char_w))
        budget = max(4, cpl * self.VAL_LINES - 2)
        if len(val) <= budget:
            return val
        kw = self._hl_keyword
        lo = val.lower().find(kw.lower()) if kw else -1
        if lo < 0:
            return val[:budget] + "…"
        half = max(0, (budget - len(kw)) // 2)
        start = max(0, lo - half)
        end   = min(len(val), start + budget)
        start = max(0, end - budget)
        s = val[start:end]
        if start > 0:
            s = "…" + s
        if end < len(val):
            s = s + "…"
        return s

    # ------------------------------------------------------------------ #
    #  分页逻辑                                                            #
    # ------------------------------------------------------------------ #

    def _jump_page(self, page):
        """跳到指定页码（1-based）。合并快速连击。"""
        if not self.all_rows:
            return
        page = max(1, min(self._total_pages, page))
        if page == self._cur_page:
            return
        # 渲染期间锁住导航按钮，防止事件堆积
        self._nav_prev.configure(state="disabled")
        self._nav_next.configure(state="disabled")
        for btn in self._page_btns:
            btn.configure(state="disabled")
        self._pending_page = page
        # after_idle 合并：连续点击之间只有最后一次起效
        if getattr(self, "_render_after_id", None):
            self._parent.after_cancel(self._render_after_id)
        self._render_after_id = self._parent.after(1, self._do_jump)

    def _do_jump(self):
        """真正执行跳转和渲染（延迟合并后调用）。"""
        self._render_after_id = None
        if self._pending_page is None:
            return
        page = self._pending_page
        self._pending_page = None
        if page < 1 or page > self._total_pages:
            return
        self._cur_page = page
        self._render()

    def _jump_by_input(self):
        """从输入框读取页码并跳转。"""
        try:
            pg = int(self._jump_var.get().strip())
        except ValueError:
            return
        self._jump_var.set("")
        self._jump_page(pg)

    def _update_page_bar(self):
        """更新页码按钮。8 个槽位始终在 « » 之间，不用的设空文本，宽度不变。"""
        tp = self._total_pages
        cp = self._cur_page
        self._total_pages_var.set(f"{tp} 页")

        if tp <= 1:
            # 只有一页也显示 1
            self._ellipsis_left.configure(text="")
            self._ellipsis_right.configure(text="")
            self._page_btns[0].configure(text="1", state="disabled",
                                         disabledforeground="#ffffff",
                                         bg=self._bg_sel, relief="flat")
            for btn in self._page_btns[1:]:
                btn.configure(text="", state="normal", bg="#e8e8e8")
            self._nav_prev.configure(state="disabled")
            self._nav_next.configure(state="disabled")
            return

        # ≤ 6 页：全部显示，无省略号
        if tp <= 6:
            pages = list(range(1, tp + 1))
            show_left = show_right = False
        else:
            # > 6 页：6 个槽位只显示 6 个连续页码，当前页尽量居中
            lo = cp - 2
            hi = cp + 3
            if lo < 1:
                lo, hi = 1, min(6, tp)
            if hi > tp:
                lo, hi = max(1, tp - 5), tp
            pages = list(range(lo, hi + 1))
            show_left  = (lo > 1)
            show_right = (hi < tp)

        self._ellipsis_left.configure(text="…" if show_left else "")
        self._ellipsis_right.configure(text="…" if show_right else "")
        for i, btn in enumerate(self._page_btns):
            if i < len(pages):
                p = pages[i]
                if p == cp:
                    btn.configure(text=str(p), state="disabled",
                                  disabledforeground="#ffffff",
                                  bg=self._bg_sel, relief="flat")
                else:
                    btn.configure(text=str(p), state="normal",
                                  bg="#e8e8e8", fg="black")
                btn.configure(command=lambda pg=p: self._jump_page(pg))
            else:
                btn.configure(text="", state="normal", bg="#e8e8e8")

        self._nav_prev.configure(state="disabled" if cp <= 1 else "normal")
        self._nav_next.configure(state="disabled" if cp >= tp else "normal")

    def _on_page_size_change(self, event=None):
        """用户切换每页条数时重新渲染。"""
        try:
            new_size = int(self._page_size_var.get())
        except ValueError:
            return
        if not self.all_rows:
            return
        if new_size == self._page_size:
            return
        self._page_size = new_size
        if self._on_page_size_save:
            self._on_page_size_save(new_size)
        # 取消飞行中的跳转
        if getattr(self, "_render_after_id", None):
            self._parent.after_cancel(self._render_after_id)
            self._render_after_id = None
        self._pending_page = None
        self._cur_page = 1
        # 清掉旧页控件再渲染，避免新旧数量不一致导致的错乱
        self._destroy_page_widgets()
        self._render()

    # ------------------------------------------------------------------ #
    #  滚动                                                                #
    # ------------------------------------------------------------------ #

    def _rows_canvas_yview(self, *args):
        self._rows_canvas.yview(*args)

    def _on_wheel_canvas(self, event):
        """鼠标直接在 canvas 空白区（非行控件上方）时滚动。"""
        self._rows_canvas.yview_scroll(int(-event.delta / 120), "units")

    def _on_wheel_global(self, event):
        """bind_all 兜底：鼠标在行控件上方时滚轮事件被子控件拦截，
        全局捕获后只当鼠标在 canvas 区域内才滚动表格。"""
        # 如果焦点在弹窗（值查看窗口等）上，不处理
        try:
            w = event.widget
            while w is not None:
                if isinstance(w, tk.Toplevel) and w is not self._rows_canvas.winfo_toplevel():
                    return
                w = w.master
        except Exception:
            pass
        x, y = event.x_root, event.y_root
        cx1 = self._rows_canvas.winfo_rootx()
        cy1 = self._rows_canvas.winfo_rooty()
        cx2 = cx1 + self._rows_canvas.winfo_width()
        cy2 = cy1 + self._rows_canvas.winfo_height()
        if cx1 <= x <= cx2 and cy1 <= y <= cy2:
            self._rows_canvas.yview_scroll(int(-event.delta / 120), "units")

    def _on_rows_canvas_configure(self, event):
        """Canvas 宽度变化时更新 rows_frame 宽度和 scrollregion。"""
        self._rows_canvas.itemconfigure(self._rows_window, width=event.width)
        self._update_scrollregion()

    def _update_scrollregion(self):
        """根据 rows_frame 实际尺寸更新 canvas 的 scrollregion，保证至少等于视口高度。"""
        self.rows_frame.update_idletasks()
        view_h = self._rows_canvas.winfo_height()
        frame_w = self.rows_frame.winfo_reqwidth()
        frame_h = self.rows_frame.winfo_reqheight()
        self._rows_canvas.itemconfigure(self._rows_window,
                                        width=max(frame_w, self._rows_canvas.winfo_width()),
                                        height=frame_h)
        self._rows_canvas.configure(
            scrollregion=(0, 0, max(frame_w, 100), max(frame_h, view_h)))

    def _apply_hl_async(self):
        """分批异步补齐当前页所有值列的高亮，避免同步逐行搜索阻塞 UI。
        每批 20 行，间隔 1ms 让出主线程。"""
        kw = self._hl_keyword
        if not kw or not self._page_widgets:
            return
        self._hl_kw = kw
        self._hl_idx = 0
        self._hl_tag_name = "hit"
        self._do_hl_batch()

    def _do_hl_batch(self):
        kw = getattr(self, "_hl_kw", "")
        if not kw:
            return
        tag = self._hl_tag_name
        batch_size = 20
        total = len(self._page_widgets)
        i = getattr(self, "_hl_idx", 0)
        end = min(i + batch_size, total)
        while i < end:
            cells, _ = self._page_widgets[i]
            c = cells[self._value_ci]
            try:
                c.configure(state="normal")
                c.tag_delete(tag)
                c.tag_configure(tag, foreground=self._hl_color,
                                font=self._head_font)
                pos = "1.0"
                n = len(kw)
                while True:
                    found = c.search(kw, pos, stopindex="end", nocase=1)
                    if not found:
                        break
                    c.tag_add(tag, found, f"{found}+{n}c")
                    pos = f"{found}+{n}c"
                c.configure(state="disabled")
            except Exception:
                pass
            i += 1
        self._hl_idx = i
        if i < total:
            self._parent.after(1, self._do_hl_batch)

    # ------------------------------------------------------------------ #
    #  事件：列宽拖动 / 悬浮提示                                            #
    # ------------------------------------------------------------------ #

    def _show_tooltip(self, event, text, header=None):
        """在鼠标下方显示一个半透明悬浮提示。

        如果 header 不为 None，显示两行：header（正常字体）+ text（小字）。
        """
        self._hide_tooltip()
        tw = tk.Toplevel(self._parent)
        tw.wm_overrideredirect(True)
        try:
            tw.attributes("-topmost", True)
        except Exception:
            pass
        if header is not None:
            # 两行提示：第一行内容，第二行提示文字
            txt = header + "\n" + text
            max_w = min(800, self._parent.winfo_screenwidth() - 100)
            lbl = tk.Label(tw, text=txt, font=self._ui_font,
                           bg="#ffffcc", fg="#333333", bd=1, relief="solid",
                           padx=8, pady=3, justify="left", wraplength=max_w)
            lbl.pack()
        else:
            lbl = tk.Label(tw, text=text, font=self._ui_font,
                           bg="#ffffcc", fg="#333333", bd=1, relief="solid",
                           padx=8, pady=3)
            lbl.pack()
        tw.update_idletasks()
        wx = self._parent.winfo_pointerx() + 14
        wy = self._parent.winfo_pointery() + 14
        tw.wm_geometry(f"+{wx}+{wy}")
        self._tooltip = tw

    def _hide_tooltip(self):
        tw = getattr(self, "_tooltip", None)
        if tw is not None:
            try:
                tw.destroy()
            except Exception:
                pass
            self._tooltip = None

    def _head_motion(self, event, i):
        hl = self._head_labels[i]
        if self.col_spec[i][2] is not None and \
                event.x >= hl.winfo_width() - self._resize_margin:
            hl.configure(cursor="sb_h_double_arrow")
            self._edge_col = i
        else:
            hl.configure(cursor="hand2")
            if self._edge_col == i:
                self._edge_col = None

    def _head_leave(self, i):
        if self._edge_col == i:
            self._edge_col = None

    def _head_press(self, event, i, key):
        # 按下时按实际位置重新判定是否在列边缘，不依赖上一次 <Motion>
        # 留下的 _edge_col —— 否则按下前鼠标轻微左移就会误判成排序点击
        hl = self._head_labels[i]
        near_edge = (self.col_spec[i][2] is not None and
                     event.x >= hl.winfo_width() - self._resize_margin)
        if near_edge:
            self._edge_col = i
            self._resizing = i
            self._drag_x0 = event.x_root
            self._drag_w0 = self.col_spec[i][2]
        else:
            self._resizing = None
            self.sort_by(key)

    def _head_drag(self, event, i):
        if self._resizing is None:
            return
        j = self._resizing
        # 下限取 _min_col 与表头 Label 自然宽度的较大值：
        # tkinter 列宽 = max(minsize, 控件 reqwidth)，若 minsize < header_reqwidth，
        # head_frame 列会比 rows_frame 列宽，导致两帧错位。
        hdr_min = self._head_labels[j].winfo_reqwidth()
        new_w = max(self._min_col, hdr_min, self._drag_w0 + (event.x_root - self._drag_x0))
        self.col_spec[j][2] = new_w
        self.head_frame.columnconfigure(j * 2, minsize=new_w)
        self.rows_frame.columnconfigure(j * 2, minsize=new_w)

    def _head_release(self, event, i):
        if self._resizing is None:
            return
        self._resizing = None
        # 列宽变了：重渲染
        self._render()
        if self._on_col_resize_done:
            self._on_col_resize_done()

    # ------------------------------------------------------------------ #
    #  事件：行交互                                                        #
    # ------------------------------------------------------------------ #

    def _on_dbl(self, i):
        self._select_row(i)
        self._on_copy_row(self._selected_data())

    def _on_dbl_file(self, i):
        self._select_row(i)
        self._on_open_file(self._selected_data(), jump_cell=False)

    def _on_dbl_value(self, i):
        self._select_row(i)
        self._on_view_value(self._selected_data(), i, self)

    def _on_right_click(self, event, i):
        self._select_row(i)
        try:
            self.ctx_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.ctx_menu.grab_release()

# -*- coding: utf-8 -*-
"""关键字历史下拉弹层。

独立管理弹层的创建、显示、隐藏、键盘导航和条目选取，
不依赖 App 的其他状态，通过回调与外部通信。
"""
import tkinter as tk
from tkinter import ttk


class KeywordPopup:
    """关键字输入框下方的历史下拉弹层。

    参数
    ----
    root        : tk.Tk / tk.Toplevel —— 主窗口，用于 after() 调度
    entry       : tk.Entry            —— 关键字输入框
    kw_var      : tk.StringVar        —— 输入框绑定的变量
    ui_font     : tuple               —— 列表字体
    on_pick     : callable            —— 用户选中某条关键字后的回调，接收 (keyword)
    is_busy     : callable            —— 返回 bool，当前是否有任务在跑
    """

    def __init__(self, root, entry, kw_var, ui_font, on_pick, is_busy):
        self._root = root
        self._entry = entry
        self._kw_var = kw_var
        self._ui_font = ui_font
        self._on_pick = on_pick
        self._is_busy = is_busy

        self._all_keywords = []      # 全部历史（最近的排最前）
        self._shown = []             # 当前弹层里展示的条目
        self._pop = None             # tk.Toplevel
        self._listbox = None         # tk.Listbox
        self._nav_syncing = False    # _nav 同步文字到输入框时不触发 trace 刷新
        self._bind_all_id = None     # bind_all 绑定 ID
        self._kw_after_id = None     # 输入变化后的延迟刷新 after ID

        # 绑定事件到 entry
        entry.bind("<Button-1>", self._show_all)
        entry.bind("<Down>", self._nav)
        entry.bind("<Up>", self._nav)
        entry.bind("<Escape>", self._clear_focus)

        # 窗口移动/大小变化时隐藏弹层
        root.bind("<Configure>", lambda e: self.hide(), add="+")

        # 用 trace 监听关键字变量变化，50ms 防抖，确保 Entry 内部处理完再刷新
        self._kw_var.trace_add("write", lambda *a: self._schedule_refresh())

    # ---------- 公共接口 ----------

    def set_keywords(self, keywords):
        """更新全部历史关键字（最近的排最前）。"""
        self._all_keywords = list(keywords)

    def push_keyword(self, keyword):
        """在内存中追加一个关键字到列表末尾（最近使用），不重读文件。"""
        kw = keyword.strip()
        if not kw:
            return
        try:
            self._all_keywords.remove(kw)
        except ValueError:
            pass
        self._all_keywords.insert(0, kw)

    def destroy(self):
        """清理。"""
        self.hide()
        self._unbind_all()

    def hide(self):
        self._unbind_all()
        if self._pop is not None:
            try:
                self._pop.withdraw()
            except Exception:
                pass

    # ---------- bind_all 管理 ----------

    def _bind_all(self):
        """弹层显示时，bind_all 抓取全局点击。"""
        self._unbind_all()
        self._bind_all_id = self._root.bind_all(
            "<Button-1>", self._on_global_click, add="+")

    def _unbind_all(self):
        if self._bind_all_id is not None:
            try:
                self._root.unbind_all(self._bind_all_id)
            except Exception:
                pass
            self._bind_all_id = None

    def _on_global_click(self, event):
        """bind_all 回调：用屏幕坐标判断是否在弹层/输入框外。"""
        if self._pop is None or not self._pop.winfo_ismapped():
            self._unbind_all()
            return
        self._check_click_pos(event.x_root, event.y_root)

    # ---------- 事件处理 ----------

    def _clear_focus(self, event=None):
        """ESC 隐藏 popup 并让焦点离开输入框。"""
        self.hide()
        self._entry.focus_set()
        self._entry.icursor("end")

    def _check_click_pos(self, x_root, y_root):
        """检查屏幕坐标：弹层内/输入框内不动，其它隐藏。"""
        # 点击在弹层内 → 不关
        try:
            pw = self._pop
            if (pw.winfo_rootx() <= x_root <= pw.winfo_rootx() + pw.winfo_width() and
                    pw.winfo_rooty() <= y_root <= pw.winfo_rooty() + pw.winfo_height()):
                return
        except Exception:
            pass
        # 点击在输入框内 → 交给 entry 自己的 handler，不管
        e = self._entry
        try:
            if (e.winfo_rootx() <= x_root <= e.winfo_rootx() + e.winfo_width() and
                    e.winfo_rooty() <= y_root <= e.winfo_rooty() + e.winfo_height()):
                return
        except Exception:
            pass
        # 点击在其他位置 → 隐藏
        self.hide()
        self._entry.focus_set()
        self._entry.icursor("end")

    def _on_popup_click(self, event):
        """弹层自身的 <Button-1>：阻止事件冒泡到 bind_all。"""
        self._check_click_pos(event.x_root, event.y_root)
        return "break"   # 阻止冒泡，不让 bind_all 再处理

    # ---------- 输入变化调度 ----------

    def _schedule_refresh(self):
        """50ms 防抖：快速连续输入时只刷新一次。"""
        if self._kw_after_id is not None:
            self._root.after_cancel(self._kw_after_id)
        self._kw_after_id = self._root.after(50, self._on_kw_changed)

    def _on_kw_changed(self):
        """输入框文字有变化时刷新弹层（50ms after 回调）。"""
        self._kw_after_id = None
        if self._is_busy():
            return
        if self._nav_syncing:
            self._nav_syncing = False   # 清除 _nav/_pick 设置的标志，跳过本次刷新
            return
        text = self._kw_var.get()
        if not text.strip():
            self._show(filter_text=False)
        else:
            self._show(filter_text=True)

    # ---------- 内部 ----------

    def _matches(self, text):
        text = text.strip().lower()
        if not text:
            return list(self._all_keywords)
        return [k for k in self._all_keywords if text in k.lower()]

    def _build(self):
        pop = tk.Toplevel(self._root)
        pop.wm_overrideredirect(True)       # 无标题栏裸弹层
        # 弹层自身 bind <Button-1>，返回 "break" 阻止冒泡到 bind_all
        pop.bind("<Button-1>", self._on_popup_click)
        lb = tk.Listbox(pop, activestyle="dotbox", font=self._ui_font,
                        height=10, relief="solid", bd=1,
                        highlightthickness=0, exportselection=False,
                        selectbackground="#3399ff", selectforeground="#ffffff",
                        selectmode="single")
        sb = ttk.Scrollbar(pop, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        lb.pack(side="left", fill="both", expand=True)
        lb.bind("<ButtonRelease-1>", self._pick)
        lb.bind("<MouseWheel>", self._on_wheel)
        pop.bind("<MouseWheel>", self._on_wheel)
        self._pop = pop
        self._listbox = lb
        pop.withdraw()

    def _show(self, event=None, filter_text=True):
        if self._is_busy():
            self.hide()
            return
        if filter_text and self._kw_var.get():
            matches = self._matches(self._kw_var.get())
        else:
            matches = list(self._all_keywords)
        if not matches:
            self.hide()
            return
        if self._pop is None:
            self._build()
        e = self._entry
        self._pop.update_idletasks()
        x = e.winfo_rootx()
        y = e.winfo_rooty() + e.winfo_height()
        w = e.winfo_width()
        lb = self._listbox
        lb.delete(0, "end")
        for m in matches:
            lb.insert("end", m)
        self._shown = matches
        lb["height"] = min(len(matches), 10)
        self._pop.update_idletasks()
        h = self._pop.winfo_reqheight()
        self._pop.wm_geometry(f"{w}x{h}+{x}+{y}")
        self._pop.deiconify()
        self._pop.lift()
        # 弹层显示后，通过 bind_all 捕获所有点击
        self._bind_all()

    def _show_all(self, event=None):
        """直接点击输入框：已显示则隐藏，未显示则显示全部历史。"""
        if self._pop is not None and self._pop.winfo_ismapped():
            self.hide()
            self._entry.focus_set()
            self._entry.icursor("end")
        else:
            self._show(filter_text=False)

    def _nav(self, event):
        if self._pop is None or not self._pop.winfo_ismapped():
            if event.keysym == "Down":
                self._show(filter_text=False)
            return "break"
        lb = self._listbox
        size = lb.size()
        if size == 0:
            return "break"
        cur = lb.curselection()
        idx = cur[0] if cur else -1
        idx += 1 if event.keysym == "Down" else -1
        idx = max(0, min(size - 1, idx))
        lb.selection_clear(0, "end")
        lb.selection_set(idx)
        lb.activate(idx)
        lb.see(idx)
        self._nav_syncing = True
        self._kw_var.set(self._shown[idx])
        # _nav_syncing 由 _on_kw_changed (50ms after) 清除
        self._entry.icursor("end")
        return "break"

    def _on_wheel(self, event):
        self._listbox.yview_scroll(-1 if event.delta > 0 else 1, "units")
        return "break"

    def _pick(self, event):
        lb = self._listbox
        sel = lb.curselection()
        idx = sel[0] if sel else lb.nearest(event.y)
        picked = 0 <= idx < len(self._shown)
        if picked:
            self._nav_syncing = True
            self._kw_var.set(self._shown[idx])
            self._entry.icursor("end")
        self.hide()
        self._entry.focus_set()
        self._entry.icursor("end")
        if picked:
            self._on_pick(self._shown[idx])

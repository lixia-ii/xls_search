# -*- coding: utf-8 -*-
"""点关闭按钮时的"退出还是最小化"询问框。"""
import tkinter as tk
from tkinter import ttk


class CloseDialog(tk.Toplevel):
    """模态询问框。

    结果放在 self.result:
        "exit" / "tray" / None(取消或直接关掉对话框)
    self.remember 为 True 表示用户勾了"不再提示"。
    """

    def __init__(self, parent, scale=1.0, font=None):
        super().__init__(parent)
        self.title("关闭 xls_search")
        self.resizable(False, False)
        self.transient(parent)

        self.result = None
        self.remember = False
        self._remember_var = tk.BooleanVar(value=False)

        px = lambda n: int(round(n * scale))
        ui_font = font or ("Microsoft YaHei UI", -int(round(13 * scale)))

        frame = ttk.Frame(self, padding=(px(22), px(20), px(22), px(16)))
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="要退出程序，还是最小化到托盘？", font=ui_font).grid(
            row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(frame, text="最小化后可在右下角托盘图标恢复窗口。",
                  font=ui_font, foreground="#666666").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(px(6), px(16)))

        ttk.Checkbutton(frame, text="不再提示，以后都这样处理",
                        variable=self._remember_var).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(0, px(16)))

        btns = ttk.Frame(frame)
        btns.grid(row=3, column=0, columnspan=2, sticky="e")
        ttk.Button(btns, text="最小化到托盘",
                   command=lambda: self._choose("tray")).pack(
            side=tk.LEFT, padx=(0, px(8)))
        exit_btn = ttk.Button(btns, text="退出",
                              command=lambda: self._choose("exit"))
        exit_btn.pack(side=tk.LEFT)

        frame.columnconfigure(0, weight=1, minsize=px(300))

        # 关掉对话框本身 = 取消，什么都不做
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<Escape>", lambda e: self._cancel())
        self.bind("<Return>", lambda e: self._choose("exit"))

        self._center_on(parent)
        self.grab_set()          # 放在最后，避免抢焦点时窗口还没定位好
        exit_btn.focus_set()
        self.wait_window(self)

    def _center_on(self, parent):
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry("+%d+%d" % (max(x, 0), max(y, 0)))

    def _choose(self, action):
        self.result = action
        self.remember = bool(self._remember_var.get())
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()

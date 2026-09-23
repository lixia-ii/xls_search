# -*- coding: utf-8 -*-
"""设置面板：关闭行为 / 开机自启动 / 清除缓存与历史。"""
import tkinter as tk
from tkinter import ttk, messagebox

import xls_search.theme as theme
from xls_search import autostart
from xls_search.storage import clear_cache, save_settings

CLOSE_CHOICES = [
    ("ask", "每次询问"),
    ("tray", "最小化到托盘"),
    ("exit", "直接退出"),
]


class SettingsDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self._app = app
        self.title("设置")
        self.resizable(False, False)
        self.transient(app.root)
        self.configure(bg=theme.BG)

        scale = app.scale
        px = lambda n: int(round(n * scale))
        ui_font = app.ui_font

        frame = ttk.Frame(self, padding=(px(22), px(20), px(22), px(16)))
        frame.pack(fill=tk.BOTH, expand=True)

        # --- 关闭按钮行为（这里可以改回"每次询问"）---
        ttk.Label(frame, text="点击关闭按钮时", style="Field.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, px(6)))

        self._close_var = tk.StringVar(
            value=dict(CLOSE_CHOICES).get(app._close_action, "每次询问"))
        close_cb = ttk.Combobox(frame, textvariable=self._close_var,
                                values=[label for _, label in CLOSE_CHOICES],
                                state="readonly", font=ui_font)
        close_cb.grid(row=1, column=0, sticky="ew", pady=(0, px(16)))
        close_cb.bind("<<ComboboxSelected>>", self._on_close_action_changed)

        # --- 开机自启动 ---
        self._autostart_var = tk.BooleanVar(value=autostart.is_enabled())
        ttk.Checkbutton(frame, text="开机时自动启动",
                        variable=self._autostart_var,
                        command=self._on_autostart_toggled).grid(
            row=2, column=0, sticky="w", pady=(0, px(16)))

        app.theme.hline(frame).grid(row=3, column=0, sticky="ew",
                                    pady=(0, px(16)))

        ttk.Button(frame, text="清除缓存 / 历史",
                   command=self._clear_cache).grid(
            row=4, column=0, sticky="w")

        self._result_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self._result_var, style="Hint.TLabel",
                  wraplength=px(320)).grid(
            row=5, column=0, sticky="w", pady=(px(10), px(18)))

        ttk.Button(frame, text="关闭", command=self.destroy).grid(
            row=6, column=0, sticky="e")

        frame.columnconfigure(0, weight=1, minsize=px(300))
        self.bind("<Escape>", lambda e: self.destroy())

        theme.center_over(self, app.root)
        self.grab_set()   # 放在最后，避免抢焦点时窗口还没定位好

    def _on_close_action_changed(self, event=None):
        label = self._close_var.get()
        for value, text in CLOSE_CHOICES:
            if text == label:
                self._app._close_action = value
                self._app.settings["close_action"] = value
                save_settings(self._app.settings)
                break
        self.focus_set()   # 清掉 readonly combobox 的选中高亮

    def _on_autostart_toggled(self):
        want = bool(self._autostart_var.get())
        ok, msg = autostart.set_enabled(want)
        self._result_var.set(msg)
        if not ok:
            # 没写成注册表就把勾选框拨回真实状态，别让界面骗人
            self._autostart_var.set(autostart.is_enabled())

    def _clear_cache(self):
        """清空索引缓存和关键字/目录历史。不可撤销，先确认。"""
        if not messagebox.askyesno(
                "清除缓存/历史",
                "将清除索引缓存、关键字历史和目录历史。\n\n此操作不可撤销，确定继续？",
                icon="warning", parent=self):
            return

        ok, msg = clear_cache()
        self._result_var.set(msg)
        if ok:
            self._app._refresh_sources()
            self._app._refresh_keywords()

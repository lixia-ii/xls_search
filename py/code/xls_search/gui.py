#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xls_search 图形界面入口。

用 pythonw.exe 启动时不会弹控制台窗口。
"""
import os
import sys

# pythonw.exe 下 sys.stdout / sys.stderr 为 None，而被 import 的模块顶部会
# 访问 sys.stdout.encoding、核心函数里也有 print，这里先兜底，避免崩溃。
class _NullOut:
    encoding = "utf-8"
    def write(self, *a, **k):
        pass
    def flush(self):
        pass

if sys.stdout is None:
    sys.stdout = _NullOut()
if sys.stderr is None:
    sys.stderr = _NullOut()

import tkinter as tk

from xls_search import autostart
from xls_search.app import App
from xls_search.paths import ASSETS_DIR


def enable_dpi_awareness():
    """声明进程 DPI 感知，避免在缩放显示器上被系统位图拉伸（文字发虚/重影）。"""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        # PROCESS_PER_MONITOR_DPI_AWARE = 2（Win8.1+），失败再退回旧 API
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _set_window_icon(root, ico_path):
    """给标题栏 / 任务栏挂图标。

    Tk 的 iconbitmap 固定按 16/32px 取图，DPI 缩放下标题栏/任务栏实际
    要的尺寸（比如 150% 缩放下是 24/48px）它不会精确匹配，于是拿小图
    拉伸上去——这就是任务栏图标发虚的原因。改用 Win32 的 LoadImage 按
    系统实际尺寸精确取出内嵌帧，再 WM_SETICON 挂上去。
    """
    if sys.platform != "win32" or not os.path.exists(ico_path):
        return
    try:
        root.iconbitmap(default=ico_path)  # 兜底 + 让子窗口（Toplevel）继承
    except Exception:
        pass
    try:
        import ctypes
        root.update_idletasks()
        user32 = ctypes.windll.user32
        hwnd = user32.GetParent(root.winfo_id()) or root.winfo_id()

        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x0010
        WM_SETICON = 0x0080
        user32.LoadImageW.restype = ctypes.c_void_p

        root._icon_handles = []  # 保持引用，避免被 GC 掉后图标变空白
        for wparam, metric in ((0, 49), (1, 11)):  # 0=小图标, 1=大图标
            n = user32.GetSystemMetrics(metric)     # SM_CXSMICON / SM_CXICON
            h = user32.LoadImageW(None, ico_path, IMAGE_ICON, n, n,
                                  LR_LOADFROMFILE)
            if h:
                root._icon_handles.append(h)
                user32.SendMessageW(hwnd, WM_SETICON, wparam,
                                    ctypes.c_void_p(h))
    except Exception:
        pass


def main():
    # 程序目录可能被移动过，校正自启动项里记的旧路径
    try:
        autostart.sync()
    except Exception:
        pass

    enable_dpi_awareness()
    root = tk.Tk()
    _set_window_icon(root, os.path.join(ASSETS_DIR, "app.ico"))
    # DPI 感知后，按真实缩放倍数放大字体/行高，避免控件过小。
    # winfo_fpixels('1i') 在 DPI 感知下返回真实 DPI（96=100%, 144=150%…）
    try:
        scale = root.winfo_fpixels("1i") / 96.0
    except Exception:
        scale = 1.0
    if scale < 1.0:
        scale = 1.0
    App(root, scale)
    root.mainloop()


if __name__ == "__main__":
    main()

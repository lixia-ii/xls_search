#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xls_search 图形界面入口。

用 pythonw.exe 启动时不会弹控制台窗口。
"""
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

from xls_search.app import App


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


def main():
    enable_dpi_awareness()
    root = tk.Tk()
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

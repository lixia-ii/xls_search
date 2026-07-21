#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Win32 输入法（IME）辅助。

把某个输入框输入法上下文的「合成字体」设成与控件一致，否则 Windows 下
正在输入的拼音会用系统默认小字，与框里已上屏的中文不一致。
非 Windows 或 API 失败时静默跳过。
"""
import sys


def set_composition_font(widget, height_px, family="Microsoft YaHei UI", weight=400):
    """height_px: 字符像素高（正数即可，内部转成 LOGFONT 的负值）。
    weight: 400=常规, 700=加粗。"""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        from ctypes import wintypes

        class LOGFONTW(ctypes.Structure):
            _fields_ = [
                ("lfHeight", ctypes.c_long),
                ("lfWidth", ctypes.c_long),
                ("lfEscapement", ctypes.c_long),
                ("lfOrientation", ctypes.c_long),
                ("lfWeight", ctypes.c_long),
                ("lfItalic", ctypes.c_byte),
                ("lfUnderline", ctypes.c_byte),
                ("lfStrikeOut", ctypes.c_byte),
                ("lfCharSet", ctypes.c_byte),
                ("lfOutPrecision", ctypes.c_byte),
                ("lfClipPrecision", ctypes.c_byte),
                ("lfQuality", ctypes.c_byte),
                ("lfPitchAndFamily", ctypes.c_byte),
                ("lfFaceName", ctypes.c_wchar * 32),
            ]

        imm32 = ctypes.WinDLL("imm32", use_last_error=True)
        imm32.ImmGetContext.restype = wintypes.HANDLE
        imm32.ImmGetContext.argtypes = [wintypes.HWND]
        imm32.ImmSetCompositionFontW.argtypes = [wintypes.HANDLE,
                                                 ctypes.POINTER(LOGFONTW)]
        imm32.ImmReleaseContext.argtypes = [wintypes.HWND, wintypes.HANDLE]

        hwnd = widget.winfo_id()
        himc = imm32.ImmGetContext(hwnd)
        if not himc:
            return
        lf = LOGFONTW()
        lf.lfHeight = -int(round(height_px))   # 负值=字符像素高
        lf.lfWeight = weight
        lf.lfCharSet = 1                        # DEFAULT_CHARSET
        lf.lfFaceName = family
        imm32.ImmSetCompositionFontW(himc, ctypes.byref(lf))
        imm32.ImmReleaseContext(hwnd, himc)
    except Exception:
        pass

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Excel / 资源管理器 相关的系统动作。

open_in_excel 走 COM 打开工作簿并跳转/高亮；失败退回默认打开。
这些函数应在后台线程调用（COM 启动 Excel 可能较慢）。
"""
import os
import subprocess


def open_in_excel(full, sheet, row, col, jump_cell=False, on_status=None):
    """用 COM 打开 full 指向的 xlsx，激活 sheet、高亮 row（可选跳到单元格）。

    on_status(msg): 可选回调，跳转失败时报告一句（线程安全由调用方保证）。
    win32com/pythoncom 不可用时退回 os.startfile。
    """
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        try:
            os.startfile(full)
        except Exception:
            pass
        return
    pythoncom.CoInitialize()
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(full)     # 已打开则返回现有工作簿
        wb.Activate()
        target = None
        for ws in wb.Worksheets:
            if ws.Name == sheet:
                target = ws
                break
        if target is not None:
            target.Activate()
            if row:
                target.Rows(row).Select()   # 选中整行 -> 高亮所在行
                # 纵向滚动把该行挪到视图中间（可视行数一半做偏移）；横向不动
                try:
                    win = excel.ActiveWindow
                    try:
                        vis = int(win.VisibleRange.Rows.Count)
                    except Exception:
                        vis = 0
                    win.ScrollRow = max(1, row - (vis // 2 if vis else 5))
                except Exception:
                    pass
        # 把 Excel 主窗口切到前台，选中行才显示成醒目的激活态（不改窗口大小/位置）
        try:
            import win32gui
            hwnd = int(excel.Hwnd)
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            try:
                excel.ActiveWindow.Activate()
            except Exception:
                pass
        # 先高亮整行；仅「打开并跳转单元格」时停顿后再高亮对应单元格
        if jump_cell and target is not None and row and col:
            try:
                import time
                time.sleep(0.5)
                target.Cells(row, col).Select()
            except Exception:
                pass
    except Exception as e:
        if on_status:
            on_status(f"跳转失败，已直接打开：{e}")
        try:
            os.startfile(full)
        except Exception:
            pass
    finally:
        pythoncom.CoUninitialize()


def reveal_in_explorer(full):
    """在资源管理器中打开所在目录并选中该文件。"""
    subprocess.Popen(f'explorer /select,"{full}"')

# -*- coding: utf-8 -*-
"""开机自动启动。

写 HKCU\\...\\Run 注册表项 —— 只影响当前用户，不需要管理员权限。
"""
import os
import sys
import winreg

from xls_search.paths import CODE_DIR

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
VALUE_NAME = "XlsSearch"

# search.bat 启动用的同一个 GUI 入口 shim
_ENTRY = os.path.join(CODE_DIR, "gui.py")


def _pythonw():
    """优先用 pythonw.exe，避免开机时闪出一个控制台黑窗。"""
    exe = os.path.abspath(sys.executable)
    cand = os.path.join(os.path.dirname(exe), "pythonw.exe")
    return cand if os.path.exists(cand) else exe


def command() -> str:
    """开机要执行的命令行。

    两段都必须是绝对路径并加引号：注册表 Run 项不设置工作目录，相对路径
    开机找不到；路径含空格不加引号会被拆成两个参数。
    """
    return '"%s" "%s"' % (_pythonw(), os.path.abspath(_ENTRY))


def registered_command():
    """读取注册表里实际存的命令行；没设置过则返回 None。"""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
            value, _ = winreg.QueryValueEx(key, VALUE_NAME)
            return value
    except OSError:
        return None


def is_enabled() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
            value, _ = winreg.QueryValueEx(key, VALUE_NAME)
            return bool(value)
    except OSError:
        return False


def enable() -> tuple:
    """写入自启动项。返回 (成功与否, 提示文本)。"""
    if not os.path.exists(_ENTRY):
        return False, "缺少启动脚本: %s" % _ENTRY
    try:
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, RUN_KEY, 0,
                                winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, VALUE_NAME, 0, winreg.REG_SZ, command())
        return True, "已设置开机自动启动"
    except OSError as e:
        return False, "设置失败: %s" % e


def disable() -> tuple:
    """移除自启动项。已经不存在也算成功。"""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0,
                            winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, VALUE_NAME)
        return True, "已取消开机自动启动"
    except FileNotFoundError:
        return True, "已取消开机自动启动"
    except OSError as e:
        return False, "取消失败: %s" % e


def set_enabled(on: bool) -> tuple:
    return enable() if on else disable()


def sync() -> bool:
    """程序启动时校正自启动路径。

    注册表里存的是设置那一刻的绝对路径。如果之后整个程序目录被移动或
    改名，那条命令就指向了不存在的位置，开机时会静默失败。这里在每次
    启动时比对一次：只要自启动是开着的、且记录的命令和当前实际路径不
    一致，就用新路径覆盖掉。返回是否发生了修正。
    """
    current = registered_command()
    if not current:
        return False            # 没开自启动，不关我们的事
    expected = command()
    if current == expected:
        return False            # 路径没变
    if not os.path.exists(_ENTRY):
        return False            # 当前位置本身就不完整，别乱写
    try:
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, RUN_KEY, 0,
                                winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, VALUE_NAME, 0, winreg.REG_SZ, expected)
        return True
    except OSError:
        return False

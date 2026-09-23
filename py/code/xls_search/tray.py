# -*- coding: utf-8 -*-
"""Windows 系统托盘图标。

tkinter 没有托盘能力，这里用 Win32 的 Shell_NotifyIcon 实现。托盘回调需要
一个窗口来接收消息，但 Tk 的窗口过程被 Tk 自己占着，所以另开一个隐藏窗口
专收托盘消息，再由 Tk 的 after 循环调用 pump() 把消息取出来派发 —— 全程
单线程，不用担心跨线程操作 Tk 控件。
"""
import os

try:
    import win32api
    import win32con
    import win32gui
    AVAILABLE = True
except ImportError:  # 没装 pywin32 时优雅降级为普通最小化
    AVAILABLE = False

WM_TRAY = 0x0400 + 20  # WM_USER + 20

NIM_ADD, NIM_DELETE = 0, 2
NIF_MESSAGE, NIF_ICON, NIF_TIP = 0x01, 0x02, 0x04

_ID_SHOW = 1023
_ID_QUIT = 1024


class TrayIcon:
    """托盘图标：启动时 show() 挂上，退出时 destroy() 摘掉，中间常驻。

    失败时静默降级，绝不让主程序起不来。
    """

    def __init__(self, ico_path, tooltip, on_show, on_quit):
        self._ico_path = str(ico_path)
        self._tooltip = tooltip
        self._on_show = on_show
        self._on_quit = on_quit
        self._hwnd = None
        self._hicon = None
        self._visible = False
        self.ok = False

        if not AVAILABLE or not os.path.exists(self._ico_path):
            return
        try:
            self._create_window()
            self._load_icon()
            self.ok = True
        except Exception:
            self.ok = False

    # ---------- 初始化 ----------

    def _create_window(self):
        hinst = win32api.GetModuleHandle(None)
        wc = win32gui.WNDCLASS()
        wc.hInstance = hinst
        wc.lpszClassName = "XlsSearchTrayHost"
        wc.lpfnWndProc = {
            WM_TRAY: self._on_tray_message,
            win32con.WM_DESTROY: self._on_destroy,
        }
        try:
            cls = win32gui.RegisterClass(wc)
        except win32gui.error:
            cls = wc.lpszClassName  # 已注册过（同进程内重开）
        self._hwnd = win32gui.CreateWindow(
            cls, "xls_search Tray", 0, 0, 0, 0, 0, 0, 0, hinst, None)
        win32gui.UpdateWindow(self._hwnd)

    def _load_icon(self):
        """取一个明显偏大的图源尺寸（64px，ico 里有精确匹配的帧）。

        任务栏托盘条只需要 SM_CXSMICON（约 16px），但 Windows 10/11
        "显示隐藏的图标"溢出面板会把同一个 HICON 在更大的方块里放大
        展示，具体放大到多少物理像素并不固定（随 DPI/系统版本变化）。
        缩小显示总是清晰的，放大显示才会糊——所以图源尺寸只嫌小不嫌
        大，直接给一个明显够用的 128px（本身就是 ico 里的一帧，不需要
        再插值），两处显示都有足够细节。
        """
        n = 128
        self._hicon = win32gui.LoadImage(
            0, self._ico_path, win32con.IMAGE_ICON, n, n,
            win32con.LR_LOADFROMFILE)

    # ---------- 显示 / 隐藏 ----------

    def show(self):
        """注册托盘图标。程序启动时调用一次，整个运行期间常驻。"""
        if not self.ok or self._visible:
            return
        nid = (self._hwnd, 0, NIF_MESSAGE | NIF_ICON | NIF_TIP,
               WM_TRAY, self._hicon, self._tooltip)
        try:
            win32gui.Shell_NotifyIcon(NIM_ADD, nid)
            self._visible = True
        except Exception:
            pass

    def hide(self):
        """注销托盘图标。只在退出时调用 —— 图标平时是常驻的。"""
        if not self.ok or not self._visible:
            return
        try:
            win32gui.Shell_NotifyIcon(NIM_DELETE, (self._hwnd, 0))
        except Exception:
            pass
        self._visible = False

    def destroy(self):
        self.hide()
        if self._hwnd:
            try:
                win32gui.DestroyWindow(self._hwnd)
            except Exception:
                pass
            self._hwnd = None

    # ---------- 消息 ----------

    def pump(self):
        """由 Tk 的 after 循环定期调用，取出并派发托盘消息。"""
        if not self.ok:
            return
        try:
            while True:
                rc, msg = win32gui.PeekMessage(
                    self._hwnd, 0, 0, win32con.PM_REMOVE)
                if not rc:
                    break
                win32gui.TranslateMessage(msg)
                win32gui.DispatchMessage(msg)
        except Exception:
            pass

    def _on_tray_message(self, hwnd, msg, wparam, lparam):
        if lparam in (win32con.WM_LBUTTONDBLCLK, win32con.WM_LBUTTONUP):
            self._on_show()
        elif lparam == win32con.WM_RBUTTONUP:
            self._popup_menu()
        return True

    def _popup_menu(self):
        menu = win32gui.CreatePopupMenu()
        win32gui.AppendMenu(menu, win32con.MF_STRING, _ID_SHOW, "显示主窗口")
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        win32gui.AppendMenu(menu, win32con.MF_STRING, _ID_QUIT, "退出")
        pos = win32gui.GetCursorPos()
        # 不先 SetForegroundWindow 的话，菜单点别处不会消失（Win32 老问题）
        win32gui.SetForegroundWindow(self._hwnd)
        cmd = win32gui.TrackPopupMenu(
            menu, win32con.TPM_LEFTALIGN | win32con.TPM_RIGHTBUTTON
            | win32con.TPM_RETURNCMD | win32con.TPM_NONOTIFY,
            pos[0], pos[1], 0, self._hwnd, None)
        win32gui.PostMessage(self._hwnd, win32con.WM_NULL, 0, 0)
        win32gui.DestroyMenu(menu)

        if cmd == _ID_SHOW:
            self._on_show()
        elif cmd == _ID_QUIT:
            self._on_quit()

    def _on_destroy(self, hwnd, msg, wparam, lparam):
        self.hide()
        return True

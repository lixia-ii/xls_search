# -*- coding: utf-8 -*-
"""统一视觉：配色、字体、ttk 控件样式。

界面里所有颜色/字体都从这里取，别再在各模块里写死色值——原先
每个控件各自一套 relief / 字号 / 底色，叠在一起就显得杂乱。

用法
----
    theme = theme.setup(root, scale)
    theme.font / theme.font_bold / theme.font_kw   —— 字体元组
    ttk 控件通过 style 名取样式：
        "Field.TLabel"      行首标签（灰色）
        "Hint.TLabel"       状态栏提示（小字灰色）
        "Accent.TButton"    主按钮（搜索）
        "Toolbutton"        扁平切换按钮（模式分段）
        "Warn.TLabel"       提醒（索引过期等）
        "Slim.Horizontal.TProgressbar"
    theme.field_box(parent) —— 给经典 tk 控件套一圈 1px 边框
    theme.hline(parent)    —— 1px 分隔线
"""
import tkinter as tk
from tkinter import ttk

FAMILY = "Microsoft YaHei UI"

# ---- 配色 ----
BG          = "#f4f5f7"    # 窗口底色
CARD        = "#ffffff"    # 输入框 / 卡片
BORDER      = "#d7dbe2"
BORDER_HOVER= "#b9c0cc"
TEXT        = "#1f2328"
TEXT_MUTED  = "#6b7280"
ACCENT      = "#2f6feb"
ACCENT_DARK = "#1f58c8"
ACCENT_LIT  = "#4a83ef"
ACCENT_SOFT = "#e7efff"
HL          = "#d9480f"    # 关键字命中
WARN        = "#b54708"

# ---- 表格 ----
ROW_EVEN = "#ffffff"
ROW_ODD  = "#f7f9fc"
SEL_BG   = "#cfe2ff"       # 浅底深字：命中关键字的橙色仍然看得清
SEL_FG   = "#10305c"
HEAD_BG  = "#eef0f4"
HEAD_FG  = "#3d444d"
GRID     = "#e5e8ec"       # 行分隔线
GRID_HEAD= "#d7dbe2"       # 表头分隔线

# ---- 悬浮提示 ----
TIP_BG   = "#32373f"
TIP_FG   = "#f2f4f7"
TIP_HINT = "#a9b1bd"


class Theme:
    """持有按 DPI 换算好的字体/尺寸，并已把 ttk 样式配置好。"""

    def __init__(self, root, scale=1.0):
        self.root = root
        self.scale = scale
        s = scale
        self.px = lambda n: int(round(n * s))          # 正数像素
        fpx = lambda n: -int(round(n * s))             # 字号（负值=像素高）

        self.font       = (FAMILY, fpx(15))
        self.font_bold  = (FAMILY, fpx(15), "bold")
        self.font_small = (FAMILY, fpx(13))
        self.font_kw    = (FAMILY, fpx(17))            # 关键字输入框
        self.kw_px      = 17 * s                       # 供 IME 合成字体用

        root.configure(bg=BG)
        root.option_add("*Menu.font", self.font)
        root.option_add("*Menu.background", CARD)
        root.option_add("*Menu.foreground", TEXT)
        root.option_add("*Menu.activeBackground", ACCENT_SOFT)
        root.option_add("*Menu.activeForeground", ACCENT)
        root.option_add("*Menu.relief", "solid")
        root.option_add("*Menu.borderWidth", 1)
        root.option_add("*Menu.activeBorderWidth", 0)
        root.option_add("*TCombobox*Listbox.font", self.font)
        root.option_add("*TCombobox*Listbox.background", CARD)
        root.option_add("*TCombobox*Listbox.selectBackground", ACCENT_SOFT)
        root.option_add("*TCombobox*Listbox.selectForeground", ACCENT)

        self._configure_styles()

    # ------------------------------------------------------------------ #

    def _configure_styles(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")    # 只有 clam/alt 才吃自定义配色
        except tk.TclError:
            pass

        pad_x, pad_y = self.px(8), self.px(4)

        st.configure(".", background=BG, foreground=TEXT, font=self.font,
                     focuscolor=BG)
        st.configure("TFrame", background=BG)
        st.configure("TLabel", background=BG, foreground=TEXT)
        st.configure("Field.TLabel", foreground=TEXT_MUTED)
        st.configure("Hint.TLabel", foreground=TEXT_MUTED, font=self.font_small)
        st.configure("Warn.TLabel", foreground=WARN, font=self.font_small)

        # clam 给按钮类控件预置了 width=-11（至少 11 字符宽），中文按钮只有
        # 两三个字时会被撑得又空又长——统一归零，宽度只由 padding 决定。
        for name in ("TButton", "TMenubutton", "Toolbutton"):
            st.configure(name, width=0)

        # 普通按钮：白底 1px 灰边，hover 变深
        st.configure("TButton", background=CARD, foreground=TEXT,
                     bordercolor=BORDER, lightcolor=CARD, darkcolor=CARD,
                     relief="flat", padding=(pad_x + self.px(4), pad_y))
        st.map("TButton",
               background=[("pressed", "#e9edf3"), ("active", "#f4f6f9"),
                           ("disabled", "#f2f3f5")],
               foreground=[("disabled", "#adb2ba")],
               bordercolor=[("active", BORDER_HOVER), ("disabled", "#e6e8eb")])

        # 主按钮：搜索
        st.configure("Accent.TButton", background=ACCENT, foreground="#ffffff",
                     bordercolor=ACCENT, lightcolor=ACCENT, darkcolor=ACCENT,
                     relief="flat", padding=(self.px(18), pad_y))
        st.map("Accent.TButton",
               background=[("pressed", ACCENT_DARK), ("active", ACCENT_LIT),
                           ("disabled", "#b9cdf6")],
               bordercolor=[("pressed", ACCENT_DARK), ("active", ACCENT_LIT),
                            ("disabled", "#b9cdf6")],
               foreground=[("disabled", "#eef3fd")])

        # 扁平切换按钮：模式分段 / 精确匹配（选中时淡蓝底 + 蓝字）
        st.configure("Toolbutton", background=BG, foreground=TEXT_MUTED,
                     bordercolor=BG, lightcolor=BG, darkcolor=BG,
                     relief="flat", padding=(self.px(12), pad_y),
                     anchor="center")
        st.map("Toolbutton",
               background=[("selected", ACCENT_SOFT), ("active", "#e9edf3"),
                           ("disabled", BG)],
               foreground=[("selected", ACCENT), ("disabled", "#b6bac1")],
               bordercolor=[("selected", ACCENT_SOFT)])

        # 菜单按钮（索引操作）：外观对齐普通按钮
        st.configure("TMenubutton", background=CARD, foreground=TEXT,
                     bordercolor=BORDER, lightcolor=CARD, darkcolor=CARD,
                     arrowcolor=TEXT_MUTED, relief="flat",
                     padding=(pad_x + self.px(2), pad_y))
        st.map("TMenubutton",
               background=[("pressed", "#e9edf3"), ("active", "#f4f6f9"),
                           ("disabled", "#f2f3f5")],
               foreground=[("disabled", "#adb2ba")],
               bordercolor=[("active", BORDER_HOVER)])

        # 输入框 / 下拉：白底细边，聚焦时边框变蓝
        for name in ("TEntry", "TCombobox"):
            st.configure(name, fieldbackground=CARD, background=CARD,
                         foreground=TEXT, bordercolor=BORDER,
                         lightcolor=CARD, darkcolor=CARD,
                         insertcolor=TEXT, arrowcolor=TEXT_MUTED,
                         selectbackground=ACCENT_SOFT, selectforeground=TEXT,
                         padding=(self.px(6), pad_y))
            st.map(name,
                   bordercolor=[("focus", ACCENT), ("hover", BORDER_HOVER)],
                   fieldbackground=[("readonly", CARD), ("disabled", "#f2f3f5")],
                   foreground=[("disabled", "#adb2ba")],
                   arrowcolor=[("disabled", "#c3c7ce")])
        st.map("TCombobox", background=[("readonly", CARD)])

        # clam 的勾选框默认 indicatorsize 不随 DPI 变，高缩放下会小成一个点
        st.configure("TCheckbutton", background=BG, foreground=TEXT,
                     indicatorcolor=CARD, bordercolor=BORDER,
                     lightcolor=CARD, darkcolor=CARD, focuscolor=BG,
                     indicatorsize=self.px(11),
                     indicatormargin=(0, 0, self.px(7), 0),
                     padding=(0, pad_y))
        st.map("TCheckbutton",
               background=[("active", BG)],
               indicatorcolor=[("selected", ACCENT), ("pressed", ACCENT_SOFT),
                               ("disabled", "#f2f3f5")],
               bordercolor=[("selected", ACCENT), ("active", BORDER_HOVER)],
               foreground=[("disabled", "#b6bac1")])

        st.configure("Slim.Horizontal.TProgressbar",
                     background=ACCENT, troughcolor="#e4e7ec",
                     bordercolor="#e4e7ec", lightcolor=ACCENT,
                     darkcolor=ACCENT, thickness=self.px(6))

        # 去掉滚动条两端的箭头按钮，只留滑块（clam 默认带箭头，很有年代感）
        try:
            st.layout("Vertical.TScrollbar", [
                ("Vertical.Scrollbar.trough", {
                    "sticky": "ns",
                    "children": [("Vertical.Scrollbar.thumb",
                                  {"expand": 1, "sticky": "nswe"})]})])
        except tk.TclError:
            pass
        st.configure("Vertical.TScrollbar", background="#c4cad4",
                     troughcolor="#f0f2f5", bordercolor="#f0f2f5",
                     lightcolor="#c4cad4", darkcolor="#c4cad4",
                     relief="flat", gripcount=0, width=self.px(10))
        st.map("Vertical.TScrollbar",
               background=[("pressed", "#939dac"), ("active", "#adb5c1")])

        st.configure("Sep.TFrame", background=BORDER)
        return st

    # ------------------------------------------------------------------ #

    def hline(self, parent, color=BORDER):
        """1px 水平分隔线（调用方自己 pack/grid）。"""
        return tk.Frame(parent, bg=color, height=1, bd=0, highlightthickness=0)

    def field_box(self, parent):
        """返回 (外框, 内容区)：1px 边框 + 白底，用来包经典 tk 控件。

        ttk.Entry 自带边框，但 tk.Entry（关键字框要用它才能挂 IME 字体和
        自定义历史弹层）默认边框很土，这里用嵌套 Frame 画一条干净的边。
        """
        outer = tk.Frame(parent, bg=BORDER, bd=0, highlightthickness=0)
        inner = tk.Frame(outer, bg=CARD, bd=0, highlightthickness=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        return outer, inner


def setup(root, scale=1.0):
    return Theme(root, scale)


def _deco(win):
    """窗口装饰（标题栏/边框）的厚度 (dx, dy)。

    geometry() 定的是窗口外框左上角，winfo_rootx/y 报的却是客户区左上角，
    两者差值就是装饰尺寸。不把它算进去，摆位会整体偏右下。
    """
    return (max(0, win.winfo_rootx() - win.winfo_x()),
            max(0, win.winfo_rooty() - win.winfo_y()))


def _frame_size(win, width=None, height=None):
    """返回 ((客户区宽, 高), (外框宽, 高))。

    已显示的窗口用它的实际尺寸；reqwidth/reqheight 是控件想要的理想尺寸，
    可能比窗口实际大（比如没设 width 的 tk.Text 默认按 80 字符要宽度），
    拿它算位置会偏。只有还没映射到屏幕的新窗口才退回 req 值。
    """
    def size(actual, req):
        return actual if win.winfo_ismapped() and actual > 1 else req

    w = int(width) if width else size(win.winfo_width(), win.winfo_reqwidth())
    h = int(height) if height else size(win.winfo_height(), win.winfo_reqheight())
    dx, dy = _deco(win)
    return (w, h), (w + dx * 2, h + dy + dx)   # 下边框按侧边框估


def _parent_frame(parent):
    """父窗口的外框 (x, y, 宽, 高)。客户区不含它自己的标题栏，摆位要按外框算。"""
    pdx, pdy = _deco(parent)
    return (parent.winfo_rootx() - pdx, parent.winfo_rooty() - pdy,
            parent.winfo_width() + pdx * 2,
            parent.winfo_height() + pdy + pdx)


def _apply(win, w, h, x, y, fw, fh, parent, ref, explicit_size):
    """钳进 ref 点所在显示器的工作区后写入 geometry，并记下父窗口状态。

    不要用 max(x, 0) 去兜负数——副屏在主屏左侧时屏幕坐标本来就是负的，
    钳到 0 会把窗口硬拽回主屏。
    """
    left, top, right, bottom = monitor_bounds(parent, *ref)
    x = max(left, min(x, right - fw))
    y = max(top, min(y, bottom - fh))
    if explicit_size:
        win.geometry(f"{w}x{h}+{int(x)}+{int(y)}")
    else:
        win.geometry(f"+{int(x)}+{int(y)}")
    # 记下当时父窗口的位置/大小，供 reposition_if_parent_moved 判断要不要跟随
    win._placed_over = parent_geom(parent)


def center_over(win, parent, width=None, height=None):
    """把 win 摆到 parent 正中，并保证整窗留在所在显示器的工作区内。

    width/height 给了就一并设定窗口尺寸（新建窗口用），否则只挪位置、
    保留控件算出来的自适应尺寸（自动布局的小对话框用）。
    """
    win.update_idletasks()
    (w, h), (fw, fh) = _frame_size(win, width, height)
    px, py, pw, ph = _parent_frame(parent)
    cx, cy = px + pw // 2, py + ph // 2
    _apply(win, w, h, cx - fw // 2, cy - fh // 2, fw, fh,
           parent, (cx, cy), bool(width or height))


def place_beside(win, parent, width=None, height=None, gap=0):
    """把 win 贴在 parent 左侧并排，顶边对齐。

    左边放不下（主窗口靠屏幕左缘）就翻到右侧；两侧都放不下才退回居中——
    否则窗口会被钳得压在主窗口上，并排的意义就没了。
    """
    win.update_idletasks()
    (w, h), (fw, fh) = _frame_size(win, width, height)
    px, py, pw, ph = _parent_frame(parent)
    left, top, right, bottom = monitor_bounds(parent, px + pw // 2, py + ph // 2)

    x = px - gap - fw
    if x < left:                      # 左边塞不下 -> 试右侧
        right_x = px + pw + gap
        if right_x + fw <= right:
            x = right_x
        else:                         # 两侧都不够 -> 居中兜底
            center_over(win, parent, width, height)
            return
    _apply(win, w, h, x, py, fw, fh, parent, (px + pw // 2, py + ph // 2),
           bool(width or height))


def parent_geom(parent):
    """父窗口当前的 (x, y, 宽, 高)，用于比对它是否被移动/缩放过。"""
    return (parent.winfo_rootx(), parent.winfo_rooty(),
            parent.winfo_width(), parent.winfo_height())


def reposition_if_parent_moved(win, parent, place=center_over):
    """复用已有子窗口时：父窗口挪过/缩放过就重新摆位，否则保持不动。

    一直重新摆位会覆盖用户手动拖到的位置；一直不动则在主窗口挪走后，子窗口
    仍停在旧地方（甚至跑到屏幕外）。折中：只跟随主窗口的移动。
    """
    if getattr(win, "_placed_over", None) != parent_geom(parent):
        place(win, parent)


def monitor_bounds(widget, x, y):
    """返回屏幕点 (x, y) 所在显示器的工作区 (left, top, right, bottom)。

    多屏时不能用 winfo_screenwidth/height 判边界——Tk 只报主屏（或第一块屏）
    的尺寸，副屏上的弹层会被按错误的边界推走。这里问 Win32 要光标所在
    显示器的工作区（已排除任务栏）；非 Windows 或调用失败时退回 Tk 的值。
    """
    try:
        import ctypes
        from ctypes import wintypes

        class RECT(ctypes.Structure):
            _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                        ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

        class MONITORINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_ulong), ("rcMonitor", RECT),
                        ("rcWork", RECT), ("dwFlags", ctypes.c_ulong)]

        user32 = ctypes.windll.user32
        user32.MonitorFromPoint.restype = ctypes.c_void_p
        MONITOR_DEFAULTTONEAREST = 2
        hmon = user32.MonitorFromPoint(wintypes.POINT(int(x), int(y)),
                                       MONITOR_DEFAULTTONEAREST)
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(MONITORINFO)
        if hmon and user32.GetMonitorInfoW(ctypes.c_void_p(hmon),
                                           ctypes.byref(mi)):
            w = mi.rcWork
            return w.left, w.top, w.right, w.bottom
    except Exception:
        pass
    return 0, 0, widget.winfo_screenwidth(), widget.winfo_screenheight()

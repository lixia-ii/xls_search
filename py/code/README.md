# xls_search 代码说明

## 项目概览

在指定 xlsx 目录中快速搜索关键字的桌面工具，支持 GUI（tkinter）和命令行两种交互方式。搜索可以直读文件（openpyxl）或走 SQLite 索引（毫秒级），索引支持全量重建和增量更新。

## 文件结构

```
code/
  gui.py                  — GUI 入口 shim（search.bat 启动目标，把 code/ 加入 sys.path 后调起 xls_search.gui.main）
  run.py                  — 命令行入口 shim（保留 `python run.py` 旧习惯，同理会调起 xls_search.cli.main）
  xls_search/             — 源码包
    __init__.py            — 包说明
    gui.py                 — GUI 真入口：启用 DPI 感知、pythonw 下 stdout/stderr 兜底、计算缩放倍数后启动 App
    app.py                 — GUI 主类：构建窗口布局、组装各模块、处理用户交互、轮询后台消息队列更新 UI
    storage.py             — 持久化：纯函数，无 tkinter 依赖。管理三个文件：
                             · gui_settings.json — 界面偏好（搜索模式、列宽、每页条数）
                             · keywords.txt      — 关键字历史（最多 200 条，命中的置顶）
                             · sources.txt        — 目录历史（最多 20 条，大小写不敏感去重）
    paths.py               — 公共路径与工具函数：
                             · SCRIPT_DIR / CODE_DIR / DATA_DIR — 三级路径常量
                             · get_index_path(xls_dir) — 索引路径 = py/cache/<md5前8位>/index.db
                             · col_letter(n)           — 列号 → 字母（1→A, 27→AA）
                             · collect_files(xls_dir)  — 递归收集 .xlsx，排除 ~$ 临时锁文件
                             · run_module(module, args) — 以 `python -m` 启动子进程并注入 PYTHONPATH
    build_index.py          — 索引引擎：
                             · build()               — 全量重建（清空索引目录后重新解析全部文件）
                             · build_incremental()   — 增量重建（只处理新增/修改/删除的文件）
                             · scan_stale()           — 轻量比对：只看文件数量/修改时间/大小，不含 hash
                             变动判断三重短路：修改时间 → 文件大小 → 内容 MD5
                             解析走多进程（ProcessPoolExecutor），进程池起不来自动退回单线程
                             索引格式：SQLite，含 cells 表和 file_meta 表（SCHEMA_VERSION=2）
    search_excel.py         — 搜索逻辑：
                             · search_index() — 走 SQLite 索引搜索，支持 cancel 回调中断
                             · search_files() — 直读 xlsx（openpyxl），逐个文件/Sheet/行/列匹配
                             均支持：子串/精确匹配、文件名过滤、限定列号
    search_controller.py    — 后台线程控制器，把耗时操作从主线程剥离：
                             通过 queue.Queue 向主线程发送消息：
                               ("progress", ...) / ("status", ...) / ("results", ...) /
                               ("error", ...) / ("cancelled", ...) / ("done", ...) / ("sync", ...)
                             支持 threading.Event 取消信号
    excel_actions.py        — Excel 自动化：
                             · open_in_excel()   — COM 打开 xlsx → 激活 Sheet → 选中行高亮 →
                                                   纵向滚动到视口中间 → 可选跳转单元格 →
                                                   SetForegroundWindow 切到前台
                                                   失败则退回 os.startfile 默认打开
                             · reveal_in_explorer() — 资源管理器定位文件
    table_widget.py         — 自绘分页结果表格（ResultTable）：
                             每页只渲染可见行控件，结果再多也不耗尽 GDI
                             功能：列头点击排序、列宽拖拽、虚拟滚动（Canvas+Scrollbar）、
                                   关键字橙色高亮（tk.Text tag）、翻页栏（页码按钮+省略号）、
                                   右键菜单（打开文件/跳转单元格/查看完整值/复制文件名/打开目录）、
                                   双击文件列→打开文件、双击值列→查看完整值、双击其他列→复制行
    keyword_popup.py        — 关键字历史下拉弹层（KeywordPopup）：
                             overrideredirect 无标题栏裸弹层，贴着输入框下方弹出
                             输入时实时筛选匹配的历史关键字
                             键盘 ↑↓ 导航，回车或点击触发搜索
    ime.py                  — Win32 输入法适配：
                             通过 ImmSetCompositionFontW 把拼音候选字体设为与输入框一致
                             解决 Windows 拼音输入时候选字过小的问题
    cli.py                  — 命令行交互版：选目录 → 选模式 → 输关键字 → 打印结果，循环

上级目录（py/）数据文件：
  sources.txt              — 目录历史（每行一个路径，最多 20 条，最近用的在最后）
  keywords.txt             — 关键字历史（每行一个，最多 200 条，命中的移到末尾）
  gui_settings.json        — 界面偏好（mode、col_px、page_size）
  cache/<目录hash>/index.db — SQLite 索引（不同 xls 目录各自独立）
  python-3.8.2/            — 内置 Python 环境
```

## 数据流

```
用户操作 (app.py)
  │
  ├─ 搜索: _start_search()
  │    └─ SearchController.run_search()  ──后台线程──→  build_index / search_excel
  │                                                         │
  │                                                    queue.Queue
  │                                                         │
  │    ┌─ _poll_queue() 轮询 ←──────────────────────────────┘
  │    ├─ "progress" → 进度条 + 状态栏
  │    ├─ "results"  → ResultTable.set_rows() → 分页渲染
  │    └─ "done"     → _set_busy(False) 恢复 UI
  │
  ├─ 建索引: _start_index_build()
  │    └─ SearchController.run_index_build()  ──后台线程──→  build_index.build / build_incremental
  │                                                              │
  │                                                         queue.Queue → _poll_queue()
  │
  └─ 定时同步: _check_sync() [每 5 秒]
       └─ SearchController.run_sync_probe()  ──后台线程──→  scan_stale() → "sync" 消息
```

## 入口

| 方式 | 命令 | 说明 |
|------|------|------|
| GUI | 双击 `search.bat` | pythonw 启动，无命令行窗口 |
| GUI | `python py/code/gui.py` | 直接启动，可见控制台输出 |
| CLI | `python py/code/run.py` | 命令行交互 |

## 依赖

- **Python 3.8.2**（内置于 `py/python-3.8.2/`）
- **openpyxl** — 读取 xlsx
- **pywin32** — Excel COM 自动化（win32com.client、win32gui）
- 标准库：tkinter、sqlite3、json、hashlib、concurrent.futures、threading、queue、subprocess、ctypes

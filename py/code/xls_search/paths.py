# -*- coding: utf-8 -*-
"""公共路径与工具函数。

集中原来散落在 build_index / search_excel / run / storage 里重复定义的：
  - SCRIPT_DIR / CODE_DIR / DATA_DIR
  - get_index_path(xls_dir)
  - col_letter(n)
  - collect_files(xls_dir)
  - run_module(module, args)   —— 供以 `python -m` 启动本包内子模块
  - ensure_utf8_stdout()        —— 确保 stdout/stderr 为 UTF-8 编码

路径关系
--------
  py/                      DATA_DIR   （cache/、*.txt、*.json 所在）
    code/                  CODE_DIR   （需在 sys.path / PYTHONPATH 里，才能 import xls_search）
      xls_search/          SCRIPT_DIR （本文件所在，即包目录）
"""
import hashlib
import io
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))      # .../py/code/xls_search
CODE_DIR   = os.path.dirname(SCRIPT_DIR)                     # .../py/code
DATA_DIR   = os.path.dirname(CODE_DIR)                       # .../py


def ensure_utf8_stdout():
    """确保 stdout/stderr 为 UTF-8 编码，供子进程入口模块调用。

    Windows 下 CLI 子进程的默认编码可能不是 UTF-8，导致含中文路径/内容的
    print 报 UnicodeEncodeError。
    """
    if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


def get_index_path(xls_dir):
    """不同 xls 目录各自一份索引：md5(目录路径)[:8] 作为缓存子目录名。"""
    d = hashlib.md5(xls_dir.encode()).hexdigest()[:8]
    return os.path.join(DATA_DIR, "cache", d, "index.db")


def col_letter(n):
    """列号 -> 字母（1->A, 26->Z, 27->AA …）。"""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def col_name_to_num(s):
    """列字母 -> 列号（A->1, Z->26, AA->27, XFD->16384）。
    纯数字字符串原样转 int。非法输入返回 None。
    """
    s = s.strip().upper()
    if s.isdigit():
        n = int(s)
        return n if 1 <= n <= 16384 else None
    if not s or not s.isalpha() or not s.isascii() or len(s) > 3:
        return None
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n if 1 <= n <= 16384 else None


def collect_files(xls_dir):
    """递归收集目录下所有 .xlsx，排除 Excel 临时锁文件 ~$*。"""
    result = []
    for root, _, files in os.walk(xls_dir):
        for f in files:
            if f.endswith(".xlsx") and not f.startswith("~$"):
                result.append(os.path.join(root, f))
    return sorted(result)


def run_module(module, args, **kw):
    """以 `python -m <module> <args...>` 启动子进程，并注入 PYTHONPATH=CODE_DIR，
    保证子进程能 import 本包（xls_search）。

    替代原先「按 .py 文件路径 spawn」的写法——模块移进 package 后，
    直接跑 .py 文件其 sys.path 不含 code/，会 import 失败；-m + PYTHONPATH 才稳。
    """
    import subprocess
    env = dict(os.environ)
    existing = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = os.pathsep.join(
        p for p in [CODE_DIR, existing] if p)
    return subprocess.run([sys.executable, "-m", module, *args], env=env, **kw)

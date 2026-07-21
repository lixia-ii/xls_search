#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""命令行交互版入口。

复用 build_index / search_excel 的核心逻辑，提供「选目录 -> 选模式 -> 搜关键字」
的交互循环。GUI（app.py）复用 storage 里的目录历史，不依赖本模块。
"""
import sys, io
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import os

from xls_search.paths import run_module
from xls_search.storage import load_sources, save_sources, SOURCES_FILE

RESET  = "\033[0m"
BOLD   = "\033[1m"
CYAN   = "\033[36m"
YELLOW = "\033[33m"
GREEN  = "\033[32m"
RED    = "\033[31m"


def select_dir(sources):
    print(f"{CYAN}sources file: {SOURCES_FILE}{RESET}")
    print()
    if sources:
        for i, s in enumerate(reversed(sources), 1):
            exists = GREEN + "ok" + RESET if os.path.isdir(s) else RED + "not found" + RESET
            print(f"  {i}. {s}  [{exists}]")
    print()
    print(f"  or type a path directly")
    print()

    while True:
        choice = input("choose: ").strip()
        if choice.isdigit():
            idx = int(choice)
            reversed_sources = list(reversed(sources))
            if 1 <= idx <= len(reversed_sources):
                return reversed_sources[idx - 1]
            else:
                print(f"{RED}invalid number, please try again.{RESET}")
        else:
            # 非数字直接当路径处理
            path = choice.strip('"')
            if not path:
                print(f"{RED}invalid input, please try again.{RESET}")
                continue
            if not os.path.isdir(path):
                print(f"{RED}path not found: {path}, please try again.{RESET}")
                continue
            if path not in sources:
                sources.append(path)
                save_sources(sources)
                print(f"{GREEN}saved to sources.txt{RESET}")
            return path


def run_search(xls_dir, mode):
    print()
    print(f"xls dir: {xls_dir}")
    keyword = input("keyword: ").strip()
    if not keyword:
        print("no keyword input.")
        return mode

    if mode is None:
        print()
        print("1. search xlsx files directly")
        print("2. search by index")
        print("3. rebuild index then search")
        print()
        mode = input("choose [1/2/3]: ").strip()
        if mode not in ("1", "2", "3"):
            print("invalid choice.")
            return None

    if mode == "1":
        run_module("xls_search.search_excel",
                   [keyword, "--xls-dir", xls_dir, "--no-index"])
    elif mode == "2":
        run_module("xls_search.search_excel",
                   [keyword, "--xls-dir", xls_dir])
    elif mode == "3":
        run_module("xls_search.build_index", ["--xls-dir", xls_dir])
        run_module("xls_search.search_excel",
                   [keyword, "--xls-dir", xls_dir])
        mode = "2"

    save_sources([xls_dir])

    print()
    print("-" * 40)
    print("done.")
    print("-" * 40)
    return mode


def main():
    sources = load_sources()
    xls_dir = select_dir(sources)
    if not xls_dir:
        return

    mode = None
    while True:
        mode = run_search(xls_dir, mode)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import argparse
import os
import sqlite3
import warnings
warnings.filterwarnings("ignore")

from xls_search.paths import col_letter, ensure_utf8_stdout, get_index_path, run_module
ensure_utf8_stdout()

RESET  = "\033[0m"
BOLD   = "\033[1m"
CYAN   = "\033[36m"
YELLOW = "\033[33m"
GREEN  = "\033[32m"
RED    = "\033[31m"


def _escape_like(s):
    """转义 LIKE 通配符 % _ \\，使关键字按字面子串匹配（配合 SQL 的 ESCAPE '\\'）。"""
    return s.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def index_exists(xls_dir):
    return os.path.exists(get_index_path(xls_dir))


def index_file_count(xls_dir):
    """索引里记录的文件数量（file_meta 行数）。索引不存在返回 0。"""
    index_path = get_index_path(xls_dir)
    if not os.path.exists(index_path):
        return 0
    conn = sqlite3.connect(index_path)
    try:
        try:
            row = conn.execute("SELECT COUNT(*) FROM file_meta").fetchone()
        except sqlite3.OperationalError:
            # 旧格式索引没有 file_meta，退而统计 cells 里去重后的文件数
            row = conn.execute("SELECT COUNT(DISTINCT file) FROM cells").fetchone()
        return row[0] if row else 0
    except sqlite3.OperationalError:
        return 0
    finally:
        conn.close()


def search_index(xls_dir, keyword, exact=False, filter_str=None, col_filter=None, cancel=None):
    index_path = get_index_path(xls_dir)
    if not os.path.exists(index_path):
        print(f"{YELLOW}index not found: {index_path}{RESET}")
        ans = input("build index now? [y/n]: ").strip().lower()
        if ans == "y":
            run_module("xls_search.build_index", ["--xls-dir", xls_dir], check=True)
        else:
            print(f"{RED}search cancelled.{RESET}")
            return []

    conn = sqlite3.connect(index_path)
    if cancel:
        # cancel() 返回 True 时中断查询（fetchall 会抛 OperationalError）
        conn.set_progress_handler(lambda: 1 if cancel() else 0, 2000)

    if exact:
        sql    = "SELECT file,sheet,row,col,value FROM cells WHERE value=?"
        params = [keyword]
    else:
        sql    = "SELECT file,sheet,row,col,value FROM cells WHERE value LIKE ? ESCAPE '\\'"
        params = [f"%{_escape_like(keyword)}%"]

    if filter_str:
        sql += " AND LOWER(file) LIKE ? ESCAPE '\\'"
        params.append(f"%{_escape_like(filter_str.lower())}%")
    if col_filter:
        sql += " AND col=?"
        params.append(col_filter)

    sql += " ORDER BY file,sheet,row,col"
    try:
        rows = conn.execute(sql, params).fetchall()
    except sqlite3.OperationalError:
        conn.close()
        if cancel and cancel():
            return None          # 被用户取消
        raise
    conn.close()
    return rows


def search_files(xls_dir, keyword, exact=False, filter_str=None, col_filter=None, progress=None):
    try:
        import openpyxl  # noqa: F401 — 仅作可用性检查
    except ImportError:
        print("missing: pip install openpyxl")
        sys.exit(1)

    from xls_search.build_index import _iter_parsed

    all_files = []
    for root, _, files in os.walk(xls_dir):
        for f in sorted(files):
            if not f.endswith(".xlsx") or f.startswith("~$"):
                continue
            full = os.path.join(root, f)
            if filter_str and filter_str.lower() not in full.lower():
                continue
            all_files.append(full)

    items = [(os.path.relpath(p, xls_dir), p) for p in all_files]
    total = len(items)
    kw_lower = keyword.lower()
    results = []

    cli_prog = progress or (lambda i, t, rel:
                            print(f"\r  scanning [{i}/{t}] {rel[:50]:<50}", end="", flush=True))

    for rel, rows, _mt, _sz, _h, err in _iter_parsed(items, cli_prog, on_status=lambda _: None):
        if err:
            print(f"\n  {RED}skip {rel}: {err}{RESET}")
            continue
        if not rows:
            continue
        file_hits = []
        for _file, sheet, row_idx, col_idx, cell_str in rows:
            if col_filter and col_idx != col_filter:
                continue
            matched = (cell_str == keyword) if exact else (kw_lower in cell_str.lower())
            if matched:
                file_hits.append((sheet, row_idx, col_idx, cell_str))
        if file_hits:
            print(f"\r{BOLD}{GREEN}[{rel}]{RESET}" + " " * 20)
            for sheet, row, col, val in file_hits:
                display = val if len(val) <= 60 else val[:57] + "..."
                print(f"  Sheet={YELLOW}{sheet}{RESET}  "
                      f"行={BOLD}{row}{RESET}  "
                      f"列={col_letter(col)}({col})  "
                      f"值={display}")
            results.extend([(rel, s, r, c, v) for s, r, c, v in file_hits])

    print(f"\r{'':70}")
    return results


def print_results(rows):
    cur_file = None
    for file, sheet, row, col, val in rows:
        if file != cur_file:
            print(f"\n{BOLD}{GREEN}[{file}]{RESET}")
            cur_file = file
        display = val if len(val) <= 60 else val[:57] + "..."
        print(f"  Sheet={YELLOW}{sheet}{RESET}  "
              f"行={BOLD}{row}{RESET}  "
              f"列={col_letter(col)}({col})  "
              f"值={display}")
    print("\n" + "-" * 70)
    print(f"{GREEN}found {len(rows)} matches{RESET}" if rows else f"{RED}no matches{RESET}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("keyword")
    parser.add_argument("--xls-dir", required=True)
    parser.add_argument("-f", "--filter", default=None)
    parser.add_argument("--exact", action="store_true")
    parser.add_argument("--col", type=int, default=None)
    parser.add_argument("--no-index", action="store_true")
    args = parser.parse_args()

    print(f"xls dir: {args.xls_dir}")
    print(f"{CYAN}keyword: {BOLD}{args.keyword}{RESET}  "
          f"exact: {args.exact}  filter: {args.filter or 'none'}  col: {args.col or 'none'}")
    print("-" * 70)

    if args.no_index:
        rows = search_files(args.xls_dir, args.keyword, exact=args.exact,
                            filter_str=args.filter, col_filter=args.col)
        print("-" * 70)
        print(f"{GREEN}found {len(rows)} matches{RESET}" if rows else f"{RED}no matches{RESET}")
    else:
        rows = search_index(args.xls_dir, args.keyword, exact=args.exact,
                            filter_str=args.filter, col_filter=args.col)
        print_results(rows)


if __name__ == "__main__":
    main()

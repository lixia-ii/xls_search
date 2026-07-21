#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import argparse
import concurrent.futures as _cf
import hashlib
import os
import shutil
import sqlite3
import warnings
warnings.filterwarnings("ignore")

try:
    import openpyxl
except ImportError:
    print("missing: pip install openpyxl")
    sys.exit(1)

from xls_search.paths import DATA_DIR, ensure_utf8_stdout, get_index_path, collect_files
ensure_utf8_stdout()

SCHEMA_VERSION = 2      # file_meta 含 size/hash 的版本；写入 PRAGMA user_version


def _parse_file(path, rel):
    """解析单个 xlsx，返回 [(rel, sheet, row, col, value), ...]。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = []
    for sheet in wb.worksheets:
        for row_idx, row in enumerate(sheet.iter_rows(values_only=True), 1):
            for col_idx, cell in enumerate(row, 1):
                if cell is not None:
                    rows.append((rel, sheet.title, row_idx, col_idx, str(cell)))
    wb.close()
    return rows


def _file_hash(path):
    """整文件 md5，作为内容指纹。"""
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _parse_job(rel, path):
    """解析单个文件并算好 meta，供线程池调用。失败时 err 带错误信息。
    返回 (rel, rows, mtime, size, hash, err)。"""
    try:
        rows = _parse_file(path, rel)
        st = os.stat(path)
        return (rel, rows, st.st_mtime, st.st_size, _file_hash(path), None)
    except Exception as e:
        return (rel, None, None, None, None, str(e))


def _ping():
    return 1


def _make_process_pool(items_count):
    """尝试建一个可用的进程池并预热探测；不可用（如 pythonw 下起不来）返回 None。"""
    workers = min(os.cpu_count() or 4, 8)
    if workers < 2 or items_count < 4:
        return None
    try:
        pool = _cf.ProcessPoolExecutor(max_workers=workers)
        pool.submit(_ping).result(timeout=5)   # 确认子进程能起来
        return pool
    except Exception:
        try:
            pool.shutdown(wait=False)
        except Exception:
            pass
        return None


def _make_status(on_status, prefix=""):
    """返回 _status(msg) 闭包：同时走 print（CLI）和 on_status 回调（GUI）。
    prefix 只加在 print 里（如 "parse: "），on_status 消息不加前缀保持状态栏干净。"""
    def _status(msg):
        print(f"{prefix}{msg}")
        if on_status:
            on_status(msg)
    return _status


def _iter_parsed(items, progress, on_status=None):
    """并行解析 items=[(rel, path), ...]，按完成顺序产出
    (rel, rows, mtime, size, hash, err)。

    优先多进程（真正并行，openpyxl 是 CPU/GIL 密集，线程提速有限）；
    进程池起不来时退回单线程。progress(i, total, rel) 可抛异常以取消。
    on_status(msg): 可选，用于向 GUI 报告状态（替代 print）。
    """
    total = len(items)

    _status = _make_status(on_status, prefix="parse: ")

    def _sequential():
        for i, (rel, path) in enumerate(items, 1):
            if progress:
                progress(i, total, rel)
            yield _parse_job(rel, path)

    if total == 0:
        return
    if total == 1:
        yield from _sequential()
        return

    pool = _make_process_pool(total)
    if pool is None:
        _status("单进程模式（进程池不可用）")
        yield from _sequential()
        return

    _status(f"多进程模式 x{pool._max_workers}")
    futures = [pool.submit(_parse_job, rel, path) for rel, path in items]
    try:
        for i, fut in enumerate(_cf.as_completed(futures), 1):
            res = fut.result()
            if progress:
                progress(i, total, res[0])      # 可能抛异常（取消）
            yield res
    except BaseException:
        for f in futures:
            f.cancel()
        pool.shutdown(wait=False)
        raise
    pool.shutdown(wait=True)      # 正常结束：等子进程干净退出


def _ensure_schema(conn):
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=OFF")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS cells (
            file  TEXT,
            sheet TEXT,
            row   INTEGER,
            col   INTEGER,
            value TEXT
        )
    """)
    # file_meta 记录每个文件的 修改时间/大小/内容hash，用于增量重建判断
    conn.execute("""
        CREATE TABLE IF NOT EXISTS file_meta (
            file  TEXT PRIMARY KEY,
            mtime REAL,
            size  INTEGER,
            hash  TEXT
        )
    """)
    conn.execute(f"PRAGMA user_version={SCHEMA_VERSION}")   # 标记索引结构版本，便于将来迁移判断


def _meta_is_current(conn):
    """索引里是否已有完整的 file_meta（含 size/hash）。旧索引缺列时返回 False。"""
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='file_meta'"
    ).fetchone()
    if not row:
        return False
    cols = {r[1] for r in conn.execute("PRAGMA table_info(file_meta)")}
    return {"file", "mtime", "size", "hash"}.issubset(cols)


def scan_stale(xls_dir):
    """轻量比对磁盘与索引，判断索引是否过期。

    只看 数量 / 修改时间 / 大小（不算 hash），足够轻，可定时轮询。
    返回 (n_disk, n_index, dirty)：dirty=True 表示有新增/删除/内容改动。
    索引不存在时返回 (n_disk, 0, False)，交给上层处理。
    """
    index_path = get_index_path(xls_dir)
    n_disk_files = collect_files(xls_dir)
    n_disk = len(n_disk_files)
    if not os.path.exists(index_path):
        return (n_disk, 0, False)

    conn = sqlite3.connect(index_path)
    try:
        try:
            old = {f: (mt, sz) for f, mt, sz in
                   conn.execute("SELECT file, mtime, size FROM file_meta")}
        except sqlite3.OperationalError:
            # 旧格式索引没有 file_meta，只能比数量
            row = conn.execute("SELECT COUNT(DISTINCT file) FROM cells").fetchone()
            n_index = row[0] if row else 0
            return (n_disk, n_index, n_disk != n_index)
    finally:
        conn.close()

    cur = {os.path.relpath(p, xls_dir): p for p in n_disk_files}
    n_index = len(old)

    if set(cur) != set(old):                 # 有新增或删除
        return (n_disk, n_index, True)
    for rel, path in cur.items():            # 数量一致：逐个比 修改时间/大小
        try:
            st = os.stat(path)
        except OSError:
            return (n_disk, n_index, True)
        omt, osz = old[rel]
        if abs(omt - st.st_mtime) > 1e-6 or osz != st.st_size:
            return (n_disk, n_index, True)
    return (n_disk, n_index, False)


def build(xls_dir, progress=None, on_status=None):
    """全量建立索引（清空重建）。

    progress: 可选回调 progress(i, total, rel)，用于 GUI 显示进度。
              为 None 时走命令行的 print 输出。
    on_status: 可选回调 on_status(msg)，用于向 GUI 报告状态消息。
    """
    index_path = get_index_path(xls_dir)
    index_dir  = os.path.dirname(index_path)

    _status = _make_status(on_status)

    if os.path.exists(index_dir):
        shutil.rmtree(index_dir)
    os.makedirs(index_dir)

    conn = sqlite3.connect(index_path)
    done = 0            # 已成功建好的文件数
    pending = None      # 被取消/中断时先存异常，提交已建部分后再抛给上层
    try:
        _ensure_schema(conn)

        files = collect_files(xls_dir)
        total = len(files)
        _status(f"building index: {total} files -> {index_path}")

        items = [(os.path.relpath(p, xls_dir), p) for p in files]
        cli_prog = progress or (lambda i, t, rel:
                                print(f"\r  [{i}/{t}] {rel[:60]:<60}", end="", flush=True))
        try:
            for rel, rows, mt, sz, h, err in _iter_parsed(items, cli_prog, on_status):
                if err:
                    print(f"\n  skip {rel}: {err}")
                    continue
                if rows:
                    conn.executemany("INSERT INTO cells VALUES (?,?,?,?,?)", rows)
                conn.execute("INSERT OR REPLACE INTO file_meta VALUES (?,?,?,?)",
                             (rel, mt, sz, h))
                done += 1
                if done % 500 == 0:      # 阶段性提交，避免 WAL 无限增长
                    conn.commit()
        except Exception as e:
            pending = e      # 取消/中断：跳出循环，把已处理的文件提交保留

        # 无论是否被中断，都为已建部分建索引并提交，保留已生成的成果
        conn.execute("CREATE INDEX IF NOT EXISTS idx_value ON cells(value)")
        conn.commit()
        if pending is None:
            _status(f"done. index saved: {index_path}")
        else:
            _status(f"interrupted. kept {done} files indexed.")
    finally:
        conn.close()   # 先释放文件锁，Windows 下才能删掉 db

    if pending is not None:
        if done == 0:
            # 一条都没建成就取消 -> 删掉空索引目录，避免空 db 被当成有效索引
            shutil.rmtree(index_dir, ignore_errors=True)
        raise pending      # 让上层知道被取消（已建部分已保留）
    return done


def build_incremental(xls_dir, progress=None, on_status=None):
    """增量重建：只处理新增 / 修改 / 删除的文件。

    变动判断：修改时间、文件大小、内容 hash —— 任一不同即视为变动。
    （短路：时间或大小先不一致直接重建，无需算 hash；两者都一致时再比 hash。）

    返回本次实际处理（重建 + 删除）的文件数。
    索引不存在或为旧格式（file_meta 缺列）时，自动退化为全量重建。
    """
    index_path = get_index_path(xls_dir)
    if not os.path.exists(index_path):
        return build(xls_dir, progress, on_status)

    conn = sqlite3.connect(index_path)
    if not _meta_is_current(conn):
        conn.close()
        return build(xls_dir, progress, on_status)

    _status = _make_status(on_status)

    try:
        _ensure_schema(conn)
        old = {}
        for f, mt, sz, h in conn.execute("SELECT file, mtime, size, hash FROM file_meta"):
            old[f] = (mt, sz, h)

        cur = {}  # rel -> abs path
        for path in collect_files(xls_dir):
            cur[os.path.relpath(path, xls_dir)] = path

        # 磁盘上已删除的文件：从索引移除
        removed = [rel for rel in old if rel not in cur]
        for rel in removed:
            conn.execute("DELETE FROM cells WHERE file=?", (rel,))
            conn.execute("DELETE FROM file_meta WHERE file=?", (rel,))

        # 新增 / 修改的文件；precomp_hash 为判断阶段已算出的 hash（避免重复计算）
        changed = []
        for rel, path in cur.items():
            try:
                st = os.stat(path)
            except OSError:
                continue
            precomp_hash = None
            if rel not in old:
                is_changed = True                      # 新增
            else:
                omt, osz, oh = old[rel]
                if abs(omt - st.st_mtime) > 1e-6 or osz != st.st_size:
                    is_changed = True                  # 时间或大小不同，直接重建
                else:
                    precomp_hash = _file_hash(path)     # 时间+大小都同，比内容 hash
                    is_changed = (precomp_hash != oh)
            if is_changed:
                changed.append((rel, path, st.st_mtime, st.st_size, precomp_hash))

        total = len(changed)
        _status(f"incremental: {total} changed, {len(removed)} removed -> {index_path}")

        items = [(rel, path) for (rel, path, mt, sz, ph) in changed]
        cli_prog = progress or (lambda i, t, rel:
                                print(f"\r  [{i}/{t}] {rel[:60]:<60}", end="", flush=True))
        done = 0
        pending = None      # 被取消/中断时先存异常，提交已处理部分后再抛
        try:
            for rel, rows, mt, sz, h, err in _iter_parsed(items, cli_prog, on_status):
                if err:
                    _status(f"skip {rel}: {err}")
                    continue
                conn.execute("DELETE FROM cells WHERE file=?", (rel,))
                if rows:
                    conn.executemany("INSERT INTO cells VALUES (?,?,?,?,?)", rows)
                conn.execute("INSERT OR REPLACE INTO file_meta VALUES (?,?,?,?)",
                             (rel, mt, sz, h))
                done += 1
                if done % 500 == 0:
                    conn.commit()
        except Exception as e:
            pending = e      # 取消/中断：跳出循环，把已处理的文件提交保留

        # 已处理部分（含删除项）建索引并提交，保留成果
        conn.execute("CREATE INDEX IF NOT EXISTS idx_value ON cells(value)")
        conn.commit()
        if pending is None:
            _status(f"done. changed={total} removed={len(removed)}")
            return total + len(removed)
        _status(f"interrupted. kept {done}/{total} changed, {len(removed)} removed.")
        raise pending      # 让上层知道被取消（已处理部分已保留）
    finally:
        conn.close()   # 释放文件锁


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--xls-dir", required=True)
    parser.add_argument("--incremental", action="store_true",
                        help="只重建变动的文件")
    args = parser.parse_args()
    if args.incremental:
        build_incremental(args.xls_dir)
    else:
        build(args.xls_dir)

# -*- coding: utf-8 -*-
"""后台搜索 / 建索引控制器。

所有耗时操作都在子线程里跑，通过 queue.Queue 把进度/结果发回主线程。
不持有任何 tkinter 控件引用，UI 与逻辑彻底解耦。

消息格式（放入 q 的元组）
--------------------------
("progress",  (i, total, rel))   进度更新
("status",    msg)               状态文字
("mode",      mode)              切换模式（建完索引后切 "2"）
("results",   rows)              搜索结果
("error",     msg)               出错
("cancelled", msg)               已取消
("done",      None)              任务结束（无论成功/取消/出错）
("sync",      hint)              同步检查结果
"""
import os
import threading

import xls_search.build_index as build_index
import xls_search.search_excel as search_excel


class _Cancelled(Exception):
    """用于从内层回调中断后台任务。"""


class SearchController:
    """封装搜索与建索引的后台线程逻辑。

    参数
    ----
    q            : queue.Queue      —— 与主线程通信的消息队列
    cancel_event : threading.Event  —— 置位表示请求取消
    """

    def __init__(self, q, cancel_event):
        self.q = q
        self.cancel_event = cancel_event

    # ---------- 公共启动接口 ----------

    def run_search(self, *, xls_dir, keyword, exact, filter_str, col_filter, mode):
        """启动搜索（或先建索引再搜索）线程。"""
        threading.Thread(
            target=self._worker,
            kwargs=dict(xls_dir=xls_dir, keyword=keyword, exact=exact,
                        filter_str=filter_str, col_filter=col_filter, mode=mode),
            daemon=True,
        ).start()

    def run_index_build(self, *, xls_dir, mode):
        """启动纯建索引线程（不搜索）。"""
        threading.Thread(
            target=self._index_worker,
            kwargs=dict(xls_dir=xls_dir, mode=mode),
            daemon=True,
        ).start()

    def run_sync_probe(self, xls_dir):
        """后台检查磁盘文件数与索引是否同步。"""
        threading.Thread(
            target=self._sync_probe,
            args=(xls_dir,),
            daemon=True,
        ).start()

    # ---------- 内部线程函数 ----------

    def _mkprog(self, prefix=""):
        """返回进度回调，自动检测取消信号。"""
        def _prog(i, total, rel):
            if self.cancel_event.is_set():
                raise _Cancelled()
            self.q.put(("progress", (i, total, prefix + rel)))
        return _prog

    def _mkstatus(self):
        """返回状态回调，把 build_index 的内部消息发到主线程状态栏。"""
        def _status(msg):
            self.q.put(("status", msg))
        return _status

    def _worker(self, xls_dir, keyword, exact, filter_str, col_filter, mode):
        building = False
        try:
            if mode in ("3", "4"):
                building = True
                if mode == "3":
                    self.q.put(("status", "正在更新变动的索引…"))
                    n = build_index.build_incremental(xls_dir, progress=self._mkprog("更新 "),
                                                      on_status=self._mkstatus())
                    if not n:
                        self.q.put(("status", "索引已是最新，无变动"))
                else:
                    self.q.put(("status", "正在重建全部索引…"))
                    build_index.build(xls_dir, progress=self._mkprog("重建 "),
                                      on_status=self._mkstatus())
                building = False
                mode = "2"
                self.q.put(("mode", "2"))

            if mode == "1":
                self.q.put(("status", "正在扫描 xlsx 文件…"))
                rows = search_excel.search_files(
                    xls_dir, keyword, exact=exact,
                    filter_str=filter_str, col_filter=col_filter,
                    progress=self._mkprog())
            else:  # mode == "2"
                self.q.put(("status", "正在查询索引…"))
                rows = search_excel.search_index(
                    xls_dir, keyword, exact=exact,
                    filter_str=filter_str, col_filter=col_filter,
                    cancel=lambda: self.cancel_event.is_set())
                if rows is None:
                    raise _Cancelled()

            self.q.put(("results", rows))

        except _Cancelled:
            if building:
                # 建索引中途取消 -> build/build_incremental 已提交并保留已建好的部分
                if search_excel.index_exists(xls_dir):
                    self.q.put(("cancelled",
                                "已取消（已保留已建好的部分，可稍后用「更新变动索引」补齐）"))
                else:
                    self.q.put(("cancelled", "已取消（尚未建成，未保留）"))
            else:
                self.q.put(("cancelled", "已取消"))
        except Exception as e:
            self.q.put(("error", str(e)))
        finally:
            self.q.put(("done", None))

    def _index_worker(self, xls_dir, mode):
        try:
            if mode == "3":
                self.q.put(("status", "正在更新变动的索引…"))
                n = build_index.build_incremental(xls_dir, progress=self._mkprog("更新 "),
                                                  on_status=self._mkstatus())
                self.q.put(("status", "索引已是最新，无变动" if not n
                            else f"索引更新完成（处理 {n} 个文件）"))
            else:
                self.q.put(("status", "正在重建全部索引…"))
                n = build_index.build(xls_dir, progress=self._mkprog("重建 "),
                                      on_status=self._mkstatus())
                self.q.put(("status", f"索引重建完成（{n} 个文件）"))
            self.q.put(("mode", "2"))     # 建完切到「查询索引」
        except _Cancelled:
            if search_excel.index_exists(xls_dir):
                self.q.put(("cancelled",
                            "已取消（已保留已建好的部分，可稍后再「更新变动索引」补齐）"))
            else:
                self.q.put(("cancelled", "已取消（尚未建成，未保留）"))
        except Exception as e:
            self.q.put(("error", str(e)))
        finally:
            self.q.put(("done", None))

    def _sync_probe(self, xls_dir):
        hint = ""
        try:
            if xls_dir and os.path.isdir(xls_dir) and search_excel.index_exists(xls_dir):
                n_disk, n_index, dirty = build_index.scan_stale(xls_dir)
                if dirty:
                    if n_disk != n_index:
                        hint = (f"⚠ 文件有更新（磁盘 {n_disk} / 索引 {n_index}），"
                                "建议更新变动或重建索引")
                    else:
                        hint = "⚠ 文件内容有改动，建议更新变动或重建索引"
        except Exception:
            hint = ""
        self.q.put(("sync", hint))

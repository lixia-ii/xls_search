#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xls_search GUI 入口 shim（search.bat 启动目标）。

把 code/ 加入 sys.path 后调起 xls_search 包里的真入口。
保留这个文件是为了让 search.bat 里写死的 `py\code\gui.py` 路径继续可用。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from xls_search.gui import main

if __name__ == "__main__":
    main()

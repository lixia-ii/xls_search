#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xls_search 命令行入口 shim。

把 code/ 加入 sys.path 后调起 xls_search.cli。保留这个文件是为了让
`python run.py` 的旧习惯继续可用。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from xls_search.cli import main

if __name__ == "__main__":
    main()

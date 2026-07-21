# -*- coding: utf-8 -*-
"""xls_search 包。

模块布局
--------
paths              公共路径/工具（DATA_DIR、get_index_path、col_letter、collect_files、run_module）
storage            界面偏好 / 关键字历史 / 目录历史 持久化
build_index        索引引擎（全量 + 增量）
search_excel       搜索逻辑（直读 / 走索引）
excel_actions      Excel COM 跳转高亮、资源管理器定位
ime                Win32 输入法字体适配
keyword_popup      关键字历史下拉弹层
search_controller  后台搜索 / 建索引线程控制器
table_widget       自绘虚拟滚动结果表格
app                图形界面主类
gui                GUI 入口（DPI 感知、pythonw 兜底、main）
cli                命令行交互版入口
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""持久化：界面偏好、关键字历史、目录历史。

纯函数、无 tkinter 依赖。文件都放在 py/（DATA_DIR）下：
  gui_settings.json   界面偏好（记忆的搜索模式、列宽等）
  keywords.txt        关键字历史
  sources.txt         目录历史
"""
import json
import os

from xls_search.paths import DATA_DIR

SETTINGS_FILE = os.path.join(DATA_DIR, "gui_settings.json")
KEYWORDS_FILE = os.path.join(DATA_DIR, "keywords.txt")
SOURCES_FILE  = os.path.join(DATA_DIR, "sources.txt")

MAX_KEYWORDS = 200
MAX_SOURCES = 20


# ---------- 界面偏好 ----------

DEFAULT_SETTINGS = {
    "mode": "2",
    "col_px": {
        "#": 83,
        "file": 254,
        "sheet": 169,
        "row": 75,
        "col": 105,
    },
    "page_size": 100,
}


def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        save_settings(DEFAULT_SETTINGS)
        return dict(DEFAULT_SETTINGS)


def save_settings(d):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ---------- 关键字历史 ----------

def load_keywords():
    """读取关键字历史（去重、去空，保持"最近用的在最后"的顺序）。"""
    if not os.path.exists(KEYWORDS_FILE):
        open(KEYWORDS_FILE, "w", encoding="utf-8").close()
        return []
    try:
        with open(KEYWORDS_FILE, "r", encoding="utf-8") as f:
            lines = [l.strip() for l in f]
    except Exception:
        return []
    seen, result = set(), []
    for l in lines:
        if not l or l in seen:      # 关键字区分大小写，精确去重
            continue
        seen.add(l)
        result.append(l)
    return result


def save_keyword(kw):
    """把一个关键字写入历史：已存在则置顶（挪到最后），上限 MAX_KEYWORDS 条。"""
    kw = kw.strip()
    if not kw:
        return
    existing = load_keywords()
    if kw in existing:
        existing.remove(kw)
    existing.append(kw)
    existing = existing[-MAX_KEYWORDS:]
    try:
        with open(KEYWORDS_FILE, "w", encoding="utf-8") as f:
            for k in existing:
                f.write(k + "\n")
    except Exception:
        pass


# ---------- 目录历史 ----------

def load_sources():
    """读取目录历史（大小写不敏感去重，保持"最近用的在最后"）。"""
    if not os.path.exists(SOURCES_FILE):
        open(SOURCES_FILE, "w", encoding="utf-8").close()
        return []
    with open(SOURCES_FILE, "r", encoding="utf-8") as f:
        lines = [l.strip() for l in f if l.strip()]
    seen, result = set(), []
    for l in lines:
        key = l.lower()
        if key not in seen:
            seen.add(key)
            result.append(l)
    if len(result) != len(lines):
        save_sources(result)
    return result


def save_sources(new_sources):
    """把目录写入历史：已存在则置顶（挪到最后），上限 MAX_SOURCES 条。"""
    existing = []
    if os.path.exists(SOURCES_FILE):
        try:
            with open(SOURCES_FILE, "r", encoding="utf-8") as f:
                existing = [l.strip() for l in f if l.strip()]
        except Exception:
            pass

    merged = list(existing)
    for s in new_sources:
        if s in merged:
            merged.remove(s)
        merged.append(s)

    seen, deduped = set(), []
    for s in merged:
        key = s.lower()
        if key not in seen:
            seen.add(key)
            deduped.append(s)

    deduped = deduped[-MAX_SOURCES:]

    with open(SOURCES_FILE, "w", encoding="utf-8") as f:
        for s in deduped:
            f.write(s + "\n")

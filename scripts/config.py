# -*- coding: utf-8 -*-
"""配置读写：一次填好申请人/软件名/版本号，后续脚本自动读取。

配置文件 00_基本信息.json 放在工作目录（脚本会向上找两级），字段：
  申请人名称、软件全称、版本号、统一社会信用代码、成立日期、项目目录、输出目录
"""
import json
import os

CONFIG_NAME = '00_基本信息.json'

DEFAULTS = {
    '申请人名称': '',
    '软件全称': '',
    '版本号': 'V1.0',
    '统一社会信用代码': '',
    '成立日期': '',
    '项目目录': '',
    '输出目录': './软著材料',
    '联系人邮箱': '',
}


def find_config(path=None, start=None):
    """按优先级找配置：显式路径 → 当前目录 → 上两级。"""
    if path:
        return path if os.path.isfile(path) else None
    d = os.path.abspath(start or os.getcwd())
    for _ in range(3):
        p = os.path.join(d, CONFIG_NAME)
        if os.path.isfile(p):
            return p
        parent = os.path.dirname(d)
        if parent == d:
            break
        d = parent
    return None


def find_up(filename, start=None, levels=3):
    """文件不在当前目录时向上找，兼容「表单/统计在父目录、材料在子目录」的布局。"""
    d = os.path.abspath(start or os.getcwd())
    for _ in range(levels):
        p = os.path.join(d, filename)
        if os.path.isfile(p):
            return p
        parent = os.path.dirname(d)
        if parent == d:
            break
        d = parent
    return None


def load(path=None, start=None):
    p = find_config(path, start)
    if not p:
        return {}
    try:
        with open(p, encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save(data, path=None, start=None):
    p = path or find_config(None, start) or os.path.join(os.getcwd(), CONFIG_NAME)
    merged = dict(DEFAULTS)
    merged.update({k: v for k, v in (data or {}).items() if v not in (None, '')})
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)
    return p


def get(data, key, fallback=None):
    v = (data or {}).get(key)
    if v in (None, ''):
        return fallback
    return v

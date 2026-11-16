"""教学设计 plan JSON 的共享解析逻辑。

HTML 版（make_lesson_html.py）与 Markdown 版（make_lesson_md.py）共用本模块，
确保两条渲染路径对「角色句」「行内标记」的解释完全一致，不会各写一套而漂移。

零依赖，仅用标准库。
"""

import re

# 行内标记：**加粗**、__着重__
# HTML 版映射为 <strong> / <span class='hl'>，MD 版映射为 **加粗** / ==着重==
INLINE_PATTERNS = [
    (re.compile(r"\*\*(.+?)\*\*", re.S), r"<strong>\1</strong>"),
    (re.compile(r"__(.+?)__", re.S), r"<span class='hl'>\1</span>"),
]

MD_INLINE_PATTERNS = [
    (re.compile(r"\*\*(.+?)\*\*", re.S), r"**\1**"),
    (re.compile(r"__(.+?)__", re.S), r"==\1=="),
]

# 「教师活动：……」「学生活动：……」这类角色句的前缀标签
ROLE_LABELS = ("教师活动", "学生活动", "设计意图", "评价方式", "预设评价",
               "教学活动", "师生活动", "组织形式", "关键一问")


def is_role(text):
    """判断段落是否为「教师活动：……」这类角色句。

    返回 (标签, 正文) 二元组；不是角色句则返回 None。
    """
    t = str(text).strip()
    for lab in ROLE_LABELS:
        if t.startswith(lab):
            rest = t[len(lab):].lstrip("：: ")
            if rest:
                return lab, rest
    return None


def info_pairs(info_rows):
    """把 info_rows 规整成 [(标签, 内容), ...]，兼容非二元组的写法。"""
    pairs = []
    for item in info_rows or []:
        if isinstance(item, (list, tuple)):
            pairs.append((str(item[0]), str(item[1]) if len(item) > 1 else ""))
        else:
            pairs.append(("信息", str(item)))
    return pairs


def chunk_pairs(pairs, per_row=2):
    """把键值对按每行 per_row 个切分，不足处补空串，便于渲染成表格。"""
    out = []
    for i in range(0, len(pairs), per_row):
        chunk = list(pairs[i:i + per_row])
        while len(chunk) < per_row:
            chunk.append(("", ""))
        out.append(chunk)
    return out

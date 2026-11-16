# -*- coding: utf-8 -*-
"""将教学设计 JSON 渲染为 Markdown 教案（零依赖）。

用法:
    python make_lesson_md.py <plan.json> <output.md>

定位:
    HTML 版用于打印与正式存档，MD 版用于快速传阅、粘贴到公众号/群聊/文档、
    版本比对与再加工。因此 MD 版追求「结构清晰、纯文本友好、无外部依赖」，
    不追求视觉排版。

与HTML 版共用 plan_common 的角色句与行内标记解析，两条渲染路径不会漂移。

JSON 结构见技能 SKILL.md 的「JSON 结构」一节。
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from plan_common import MD_INLINE_PATTERNS, is_role, info_pairs  # noqa: E402


def inline(text):
    """处理行内标记：**加粗** 保留为 Markdown 加粗，__着重__ 转为 ==着重==。"""
    out = str(text)
    for pat, rep in MD_INLINE_PATTERNS:
        out = pat.sub(rep, out)
    return out


def esc_cell(text):
    """表格单元格转义：竖线与换行会破坏表格结构。"""
    return inline(text).replace("|", "\\|").replace("\n", "<br>")


def render_info(rows):
    """基本信息表：渲染为两列表格，比 HTML 的并排布局更适合纯文本阅读。"""
    pairs = info_pairs(rows)
    if not pairs:
        return ""
    out = ["| 项目 | 内容 |", "| --- | --- |"]
    for k, v in pairs:
        out.append("| %s | %s |" % (esc_cell(k), esc_cell(v)))
    return "\n".join(out)


def render_para(text):
    role = is_role(text)
    if role:
        return "**%s：** %s" % (role[0], inline(role[1]))
    return inline(text)


def render_activities(acts):
    """环节卡片：MD 里用四级标题 + 列表，比缩进段落更易扫读。"""
    out = []
    for act in acts:
        label = act.get("label", "活动")
        out.append("#### %s" % inline(label))
        out.append("")
        for it in act.get("items") or []:
            out.append("- %s" % render_para(it))
        out.append("")
    return "\n".join(out)


def render_table(spec):
    head = spec.get("header") or []
    rows = spec.get("rows") or []
    out = []
    if head:
        out.append("| " + " | ".join(esc_cell(h) for h in head) + " |")
        out.append("| " + " | ".join("---" for _ in head) + " |")
        for r in rows:
            cells = r if isinstance(r, (list, tuple)) else [r]
            cells = list(cells) + [""] * (len(head) - len(cells))
            out.append("| " + " | ".join(esc_cell(c) for c in cells[:len(head)]) + " |")
        return "\n".join(out)
    # 无表头时退化为两列键值表，避免输出不合法的表格
    out.append("| 项目 | 内容 |")
    out.append("| --- | --- |")
    for r in rows:
        cells = r if isinstance(r, (list, tuple)) else [r]
        cells = list(cells)
        while len(cells) < 2:
            cells.append("")
        out.append("| %s | %s |" % (esc_cell(cells[0]), esc_cell(cells[1])))
    return "\n".join(out)


def render_kv(pairs):
    """键值行：用定义式列表，比表格更适合纯文本粘贴。"""
    return "\n".join("- **%s：** %s" % (inline(k), inline(v)) for k, v in pairs)


def render_body(obj):
    out = []
    for p in obj.get("paragraphs") or []:
        out.append(render_para(p))
        out.append("")
    for key, marker in (("bullets", "-"), ("numbered", "1.")):
        items = obj.get(key) or []
        if not items:
            continue
        for i, it in enumerate(items, 1):
            prefix = marker if marker == "-" else "%d." % i
            out.append("%s %s" % (prefix, render_para(it)))
        out.append("")
    acts = obj.get("activities") or []
    if acts:
        out.append(render_activities(acts))
    if obj.get("table"):
        out.append(render_table(obj["table"]))
        out.append("")
    if obj.get("kv"):
        out.append(render_kv(obj["kv"]))
        out.append("")
    if obj.get("note"):
        text = inline(obj["note"])
        if "\n" in text.strip():
            # 多行 note（如板书设计）降级为围栏代码块。
            # 早先的实现给每一行都加 ">"，下游逐行提取内容时会把 ">" 当成正文，
            # 表现为课件/文本里「每句话前面挂一个 >」。代码块不带任何前缀标记。
            out.append("```text")
            out.append(text.rstrip())
            out.append("```")
        else:
            out.append("> " + text)
        out.append("")
    if obj.get("block"):
        # 显式声明的结构化块文本（板书设计、思维导图、流程步骤等）。
        # 用围栏代码块承载：缩进与层级原样保留，且不含任何 Markdown 标记，
        # 下游提取时拿到的是干净纯文本。
        out.append("```text")
        out.append(inline(obj["block"]).rstrip())
        out.append("```")
        out.append("")
    return "\n".join(out)


def render_section(sec, level=2):
    """小节标题层级：level 1 -> ##，逐级加深，MD 里最多用到 #####。"""
    out = []
    heading = sec.get("heading")
    if heading:
        depth = min(level + 1, 5)
        out.append("%s %s" % ("#" * depth, inline(heading)))
        out.append("")
    out.append(render_body(sec))
    for sub in sec.get("subsections") or []:
        out.append(render_section(sub, level + 1))
    return "\n".join(out)


def build(plan):
    title = plan.get("title", "教学设计")
    subtitle = plan.get("subtitle", "")
    meta = plan.get("meta") or {}
    school = meta.get("school") or meta.get("left") or "呼和浩特市实验学校"
    tag = meta.get("tag", "")

    out = ["# %s" % inline(title), ""]

    # 抬头信息压成一行纯文本。不加">" 等块级标记：
    # 本文件的下游用途之一是逐行提取文案去做课件，任何行首标记都可能被当成正文。
    bits = [school]
    if tag:
        bits.append(tag)
    out.append(" ｜ ".join(inline(b) for b in bits))
    out.append("")

    if subtitle:
        out.append("**%s**" % inline(subtitle))
        out.append("")

    info = render_info(plan.get("info_rows") or [])
    if info:
        out.append(info)
        out.append("")

    for sec in plan.get("sections") or []:
        out.append(render_section(sec))
        out.append("")

    footer = plan.get("footer") or title
    out.append("---")
    out.append("")
    out.append("*%s*" % inline(footer))

    text = "\n".join(out)
    # 收敛连续空行，最多保留一个空行（段落间隔），让渲染结果干净
    return re.sub(r"\n{3,}", "\n\n", text).strip() + "\n"


def main():
    if len(sys.argv) != 3:
        print("用法: python make_lesson_md.py <plan.json> <output.md>")
        sys.exit(1)
    with open(sys.argv[1], "r", encoding="utf-8") as f:
        plan = json.load(f)
    out = build(plan)
    with open(sys.argv[2], "w", encoding="utf-8") as f:
        f.write(out)
    print("OK:", sys.argv[2])


if __name__ == "__main__":
    main()

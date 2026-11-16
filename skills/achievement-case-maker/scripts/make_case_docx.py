# -*- coding: utf-8 -*-
"""由成品目录中的 案例定稿.md 生成 案例定稿.docx

用法：
    python make_case_docx.py <成品目录>

读取 <成品目录>/案例定稿.md，解析出：
    - 填写人信息（姓名/联系方式/提交日期）
    - 第一部分：元信息卡（YAML 式代码块）→ 渲染为两列表格
    - 第二部分：正文七栏目 → 渲染为各级标题与段落
并输出 <成品目录>/案例定稿.docx。
支持 **加粗**、> 引用、- 列表项等轻量 Markdown 标记。
"""
import os
import re
import sys

from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn

FIELD_ORDER = [
    "案例编号", "案例名称", "副标题/金句", "分类", "学科/领域", "适用学段",
    "核心AI工具", "负责人", "团队成员", "实施时间", "涉及校区/班级",
    "一句话简介", "关键词", "成果级别",
]


def set_cn(run, font="宋体"):
    run.font.name = "Times New Roman"
    run.element.rPr.rFonts.set(qn("w:eastAsia"), font)


def add_runs(p, text):
    """把 **加粗** 标记转换为加粗 run"""
    for i, part in enumerate(text.split("**")):
        if part == "":
            continue
        r = p.add_run(part)
        r.bold = (i % 2 == 1)
        set_cn(r)
    return p


def parse_meta(lines):
    """从 YAML 式代码块解析出 {字段: 值}，字段值去掉行尾注释
    兼容半角冒号 : 与全角冒号 ："""
    meta = {}
    for ln in lines:
        ln = ln.strip()
        if not ln:
            continue
        m = re.match(r"^([^:：]+)[:：]\s*(.*)$", ln)
        if not m:
            continue
        key = m.group(1).strip()
        if key not in FIELD_ORDER:
            continue
        val = m.group(2).split("#")[0].strip()  # 去掉行尾注释
        meta[key] = val
    return meta


def parse_md(md):
    """把 案例定稿.md 解析为结构化内容"""
    doc = {"title": "", "person": [], "meta": {}, "sections": []}
    lines = md.splitlines()

    # 标题（首个一级标题）
    for ln in lines:
        if ln.startswith("# "):
            doc["title"] = ln[2:].strip()
            break

    in_meta_block = False
    cur_section = None

    i = 0
    while i < len(lines):
        raw = lines[i]
        ln = raw.strip()

        # 填写人信息（- 开头、含冒号的列表项，且在元信息卡之前）
        # 注意：中文冒号「：」与半角「:」都要支持
        if ln.startswith("- ") and doc["meta"] == {} and cur_section is None:
            m = re.match(r"^-\s*([^:：]+)[:：]\s*(.*)$", ln)
            if m:
                k, v = m.group(1).strip(), m.group(2).strip()
                if k in ("姓名（含部门/职务）", "姓名", "联系方式（电话或企业微信）", "联系方式", "提交日期"):
                    doc["person"].append((k, v))
                    i += 1
                    continue

        # 元信息卡代码块
        if ln.startswith("```"):
            if not in_meta_block:
                in_meta_block = True
                i += 1
                block = []
                while i < len(lines) and not lines[i].strip().startswith("```"):
                    block.append(lines[i])
                    i += 1
                doc["meta"] = parse_meta(block)
                i += 1  # 跳过结束的 ```
            continue

        # 章节标题：## 跳过（## 第一部分 等），### 视为正文栏目
        if ln.startswith("### "):
            cur_section = {"title": ln[4:].strip(), "items": []}
            doc["sections"].append(cur_section)
            i += 1
            continue

        if ln.startswith("## "):
            cur_section = None
            i += 1
            continue

        # 正文内容
        if cur_section is not None and ln:
            if ln.startswith("> "):
                cur_section["items"].append(("quote", ln[2:].strip()))
            elif ln.startswith("- "):
                cur_section["items"].append(("bullet", ln[2:].strip()))
            else:
                cur_section["items"].append(("para", ln))
        i += 1

    return doc


def build_docx(parsed, out_path):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    for hname, size in [("Heading 1", 15), ("Heading 2", 13)]:
        st = doc.styles[hname]
        st.font.name = "Times New Roman"
        st.font.size = Pt(size)
        st.font.color.rgb = RGBColor(0x1F, 0x3B, 0x5C)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")

    # 标题
    title = parsed["title"] or "人工智能教育应用案例 · 成果定稿"
    t = doc.add_heading(title, level=0)
    for r in t.runs:
        set_cn(r, "黑体")

    # 填写人信息
    if parsed["person"]:
        doc.add_heading("填写人信息", level=1)
        for k, v in parsed["person"]:
            p = doc.add_paragraph()
            rk = p.add_run(f"{k}：")
            rk.bold = True
            set_cn(rk)
            rv = p.add_run(v)
            set_cn(rv)

    # 元信息卡 → 表格
    if parsed["meta"]:
        doc.add_heading("第一部分：元信息卡", level=1)
        table = doc.add_table(rows=1, cols=2)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        hdr[0].text = "字段"
        hdr[1].text = "内容"
        for c in hdr:
            for p in c.paragraphs:
                for r in p.runs:
                    r.bold = True
                    set_cn(r)
        for key in FIELD_ORDER:
            if key not in parsed["meta"]:
                continue
            cells = table.add_row().cells
            cells[0].text = key
            cells[1].text = parsed["meta"][key]
            for c in cells:
                for p in c.paragraphs:
                    for r in p.runs:
                        set_cn(r)

    # 正文七栏目
    if parsed["sections"]:
        doc.add_heading("第二部分：正文七栏目", level=1)
        for sec in parsed["sections"]:
            doc.add_heading(sec["title"], level=2)
            for kind, text in sec["items"]:
                if kind == "quote":
                    p = doc.add_paragraph(style="Intense Quote")
                    add_runs(p, text)
                elif kind == "bullet":
                    p = doc.add_paragraph(style="List Bullet")
                    add_runs(p, text)
                else:
                    add_runs(doc.add_paragraph(), text)

    doc.save(out_path)
    return doc


def main():
    if len(sys.argv) < 2:
        print("用法: python make_case_docx.py <成品目录>")
        sys.exit(1)
    outdir = sys.argv[1]
    md_path = os.path.join(outdir, "案例定稿.md")
    docx_path = os.path.join(outdir, "案例定稿.docx")

    if not os.path.isfile(md_path):
        print("ERROR: 未找到", md_path)
        sys.exit(1)

    with open(md_path, encoding="utf-8") as f:
        md = f.read()

    parsed = parse_md(md)
    build_docx(parsed, docx_path)

    # 回读验证
    d2 = Document(docx_path)
    print("SAVED:", docx_path)
    print("title:", parsed["title"])
    print("person fields:", len(parsed["person"]))
    print("meta fields:", len(parsed["meta"]))
    print("sections:", len(parsed["sections"]))
    print("paragraphs:", len(d2.paragraphs), "tables:", len(d2.tables))


if __name__ == "__main__":
    main()
# -*- coding: utf-8 -*-
"""将教学设计 JSON 渲染为 A4 打印级静态 HTML 教案（零依赖）。

用法:
    python make_lesson_html.py <plan.json> <output.html>

设计要点:
  * 页面固定 A4（210mm x 297mm），屏幕预览带纸张投影，打印时自动裁掉装饰
  * 顶部操作条提供「一键生成 PDF」（调起浏览器打印 → 另存为 PDF）
  * 打印样式表已按 A4 配置页边距，浏览器打印对话框选 A4、勾选「背景图形」即得佳效
  * 正文/标题使用中文常见字体栈，兼顾屏幕阅读与黑白打印
  * 校徽置于「信笺抬头」层（左徽 + 校名中英文 + 模型标），不与标题争夺视觉重心
  * 配色取自校徽「呼和浩特市实验学校」：校徽蓝 #2458A5（主色）+ 校徽绿 #52B05A（辅色）
  * 校徽内嵌为 base64（优先 WebP，回退 PNG），HTML 单文件即可分发，不依赖外部图片

JSON 结构见技能 SKILL.md 的「JSON 结构」一节。
"""
import base64
import html
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from plan_common import INLINE_PATTERNS, is_role, chunk_pairs, info_pairs  # noqa: E402

# ---------------------------------------------------------------- 校徽（内嵌）

_SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_ASSETS = os.path.join(_SKILL_DIR, "assets")


def badge_data_uri():
    """读取校徽并返回 data URI；校徽缺失时返回空串（模板降级为无校徽排版）。"""
    for name, mime in (("badge.webp", "image/webp"), ("badge.png", "image/png")):
        path = os.path.join(_ASSETS, name)
        if os.path.isfile(path):
            with open(path, "rb") as f:
                return "data:%s;base64,%s" % (
                    mime, base64.b64encode(f.read()).decode("ascii"))
    return ""


# ---------------------------------------------------------------- 样式

CSS = """
:root{
  --fs: 10.5pt;
  --ink:#1f2328;
  --ink-soft:#3d444d;
  --muted:#6b7280;
  --line:#d8dde3;
  /* 主色：校徽蓝 */
  --accent:#2458a5;
  --accent-deep:#1b427c;
  --accent-soft:#e8eef8;
  /* 辅色：校徽绿 */
  --accent-2:#52b05a;
  --accent-2-deep:#3d8a44;
  --accent-2-soft:#eaf5ec;
  --paper:#ffffff;
  --bg:#eef1f4;
}
*{box-sizing:border-box;}
html,body{margin:0;padding:0;}
body{
  background:var(--bg);
  color:var(--ink);
  font-family:"Source Han Serif SC","Noto Serif SC","Songti SC","SimSun",
              "Microsoft YaHei",-apple-system,serif;
  font-size:var(--fs);
  line-height:1.78;
  -webkit-font-smoothing:antialiased;
}
.page{
  width:210mm;
  min-height:297mm;
  margin:10mm auto;
  padding:16mm 15mm 16mm;
  background:var(--paper);
  box-shadow:0 2px 6px rgba(16,24,40,.10),0 12px 32px rgba(16,24,40,.10);
  position:relative;
}

/* ---------- 页脚 ---------- */
.running-foot{
  position:fixed;left:0;right:0;bottom:6mm;
  display:none;justify-content:space-between;
  padding:0 22mm;
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-size:8pt;color:var(--muted);
}

/* ---------- 顶部品牌色条 ---------- */
.brandbar{
  display:flex;align-items:stretch;gap:1.4mm;
  height:2.4mm;margin:0 0 4.5mm;
}
.brandbar i{display:block;border-radius:2pt;}
.brandbar i:nth-child(1){width:26mm;background:var(--accent);}
.brandbar i:nth-child(2){width:9mm;background:var(--accent-2);}
.brandbar i:nth-child(3){flex:1 1 auto;background:var(--line);}

/* ---------- 信笺抬头（校徽层） ---------- */
.letterhead{
  display:flex;align-items:center;gap:4.5mm;
  margin:0 0 5mm;
}
/* 白底圆托盘：让透明底校徽在任何底色上都干净 */
.lh-badge{
  flex:0 0 auto;width:18mm;height:18mm;border-radius:50%;
  background:#fff;border:.7pt solid var(--accent-soft);
  box-shadow:0 .6pt 2pt rgba(27,66,124,.10);
  display:flex;align-items:center;justify-content:center;overflow:hidden;
}
.lh-badge img{width:15.4mm;height:15.4mm;object-fit:contain;display:block;}
.lh-id{flex:1 1 auto;min-width:0;}
.lh-cn{
  display:block;
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-size:13pt;font-weight:700;letter-spacing:.16em;
  color:var(--accent-deep);line-height:1.3;
}
.lh-en{
  display:block;margin-top:.8mm;
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-size:6.6pt;font-weight:600;letter-spacing:.2em;
  color:var(--muted);text-transform:uppercase;line-height:1.3;
}
/* 右上角模型标：蓝底小胶囊，与主标题呼应（避开校徽绿，避免与徽章绿色糊在一起） */
.lh-tag{
  flex:0 0 auto;align-self:center;
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-size:8.4pt;font-weight:700;letter-spacing:.12em;
  color:#fff;background:var(--accent);
  padding:1.1mm 3mm;border-radius:20pt;white-space:nowrap;
}
.lh-rule{
  height:.7pt;background:var(--line);
  margin:0 0 6mm;
  background:linear-gradient(90deg,var(--accent) 0 18%,var(--line) 18% 100%);
}

/* ---------- 标题区 ---------- */
.titleblock{text-align:center;margin:0 0 6.5mm;}
.doc-title{
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-size:19.5pt;font-weight:700;letter-spacing:.05em;
  line-height:1.4;color:var(--ink);margin:0;
}
.doc-subtitle{
  margin:2mm 0 0;font-size:10pt;color:var(--muted);
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  letter-spacing:.06em;
}
/* 双色饰线：左段校徽蓝、右段校徽绿，呼应校徽上下分色 */
.head-rule{display:flex;justify-content:center;align-items:center;margin:3.6mm auto 0;}
.head-rule i{display:block;height:2.2pt;border-radius:2pt;}
.head-rule i:first-child{width:15mm;background:var(--accent);}
.head-rule i:last-child{width:5mm;background:var(--accent-2);}

/* ---------- 基本信息表 ---------- */
.info-table{
  width:100%;border-collapse:collapse;margin:0 0 6.5mm;
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-size:9.5pt;
  border:.6pt solid var(--line);border-radius:2pt;overflow:hidden;
  break-inside:avoid;
}
.info-table td{
  border:.6pt solid var(--line);
  padding:1.9mm 2.6mm;vertical-align:middle;line-height:1.55;
}
.info-table td.k{
  width:22mm;background:var(--accent-soft);color:var(--accent-deep);
  font-weight:700;text-align:center;white-space:nowrap;
}
.info-table td.v{width:calc(50% - 22mm);}

/* ---------- 章节 ---------- */
.section{margin:0 0 5mm;break-inside:auto;}
.section.brk{break-before:page;}
h2.sec{
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-size:12.5pt;font-weight:700;color:var(--ink);
  margin:0 0 2.6mm;padding:0 0 1.6mm 3mm;
  border-bottom:.8pt solid var(--line);
  border-left:3pt solid var(--accent);
  line-height:1.5;break-after:avoid;
}
h3.sub{
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-size:11pt;font-weight:700;color:var(--accent-deep);
  margin:3.6mm 0 1.8mm;padding-left:2.4mm;
  border-left:2pt solid var(--accent-2);line-height:1.5;
  break-after:avoid;
}
h4.sub2{
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-size:10.5pt;font-weight:700;color:var(--ink-soft);
  margin:3mm 0 1.4mm;break-after:avoid;
}
p{margin:0 0 2.2mm;text-align:justify;text-indent:2em;}
p.nt,.bullet,.num,.kv-v,td,.cell{text-indent:0;}

/* ---------- 列表 ---------- */
ul.bullets{margin:0 0 2.6mm;padding-left:1.4em;}
ul.bullets li{margin:0 0 1.2mm;}
ol.nums{margin:0 0 2.6mm;padding-left:1.6em;}
ol.nums li{margin:0 0 1.2mm;}

/* ---------- 角色卡（教师活动 / 学生活动 / 设计意图…） ---------- */
.acts{margin:0 0 3mm;display:flex;flex-direction:column;gap:2mm;break-inside:avoid;}
.act{
  border:.6pt solid var(--line);border-radius:2pt;
  background:#fcfdfd;overflow:hidden;
}
.act-label{
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-size:9pt;font-weight:700;letter-spacing:.08em;
  color:var(--accent-2-deep);background:var(--accent-2-soft);
  padding:1.1mm 2.6mm;border-bottom:.6pt solid var(--line);
}
.act-body{padding:1.8mm 2.6mm 2mm;font-size:9.8pt;line-height:1.72;}
.act-body p{margin:0 0 1.4mm;text-indent:0;}
.act-body p:last-child{margin-bottom:0;}
.act-body ul.bullets{margin:0 0 1.4mm;}

/* ---------- 通用表格 ---------- */
table.grid{
  width:100%;border-collapse:collapse;margin:0 0 3.4mm;
  font-size:9.6pt;break-inside:avoid;
}
table.grid th,table.grid td{
  border:.6pt solid var(--line);padding:1.8mm 2.4mm;
  text-align:left;vertical-align:top;line-height:1.62;
}
table.grid thead th{
  background:var(--accent-soft);color:var(--accent-deep);font-weight:700;
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-size:9.4pt;
}

/* ---------- 键值行 ---------- */
.kv{
  display:flex;gap:2.4mm;margin:0 0 2.2mm;
  font-size:10pt;break-inside:avoid;
}
.kv-k{
  flex:0 0 auto;min-width:20mm;
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-weight:700;color:var(--accent-deep);font-size:9.4pt;padding-top:.4mm;
}
.kv-v{flex:1 1 auto;}

/* ---------- 提示框 ---------- */
.note{
  margin:0 0 3.4mm;padding:2.2mm 3mm;
  border-left:2.6pt solid var(--accent);background:var(--accent-soft);
  border-radius:0 2pt 2pt 0;font-size:9.6pt;color:var(--ink-soft);
  break-inside:avoid;
}
.note b{color:var(--accent-deep);}
/* 板书等结构化块内容：保留原始换行与全角空格缩进，层级一眼可辨 */
.note.block{
  white-space:normal;
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
  font-size:9.2pt;line-height:1.95;
}

/* ---------- 顶部操作条（仅屏幕显示） ---------- */
.topbar{
  position:sticky;top:0;z-index:99;
  display:flex;justify-content:space-between;align-items:center;gap:16px;
  max-width:210mm;margin:0 auto;padding:10px 4px;
  background:rgba(238,241,244,.92);backdrop-filter:blur(6px);
  font-family:"Source Han Sans SC","Noto Sans SC","Microsoft YaHei",sans-serif;
}
.topbar .tip{font-size:12px;color:#5b6472;line-height:1.55;}
.topbar .tip b{color:var(--accent);}
.topbar .btns{display:flex;gap:8px;align-items:center;flex:0 0 auto;}
.toolbar button{
  font:inherit;font-size:12.5px;font-weight:600;cursor:pointer;white-space:nowrap;
  border:1px solid var(--accent);background:var(--accent);color:#fff;
  border-radius:7px;padding:7px 14px;
}
.toolbar button.ghost{background:#fff;color:var(--accent);}
.toolbar button.ghost.on{background:var(--accent);color:#fff;}
.toolbar button.hidden{display:none;}
.toolbar button:hover{filter:brightness(1.08);}
.toolbar code{
  background:#fff;border:1px solid var(--line);border-radius:4px;padding:0 4px;
}

/* ---------- 可编辑正文区 ----------
   editable：编辑模式开关下的可编辑区域
   用 plaintext-only 而非 true：只允许纯文本，粘贴进来的 HTML/样式会被剥除，
   因此外部富文本无法破坏既定版式。 */
.editable{
  cursor:text;border-radius:2px;
  transition:background .12s ease,box-shadow .12s ease;
  outline:none;
}
.editing .editable:hover{background:rgba(82,176,90,.07);}
.editing .editable:focus{
  background:rgba(36,88,165,.05);
  box-shadow:0 0 0 2px rgba(36,88,165,.28);
}
/* 空段落占位：contenteditable 里删空后浏览器会留 <br>，:empty 失效，
   故由 JS 打 .is-empty 类，再用属性选择器兜底 */
.editable.is-empty::before{
  content:attr(data-ph);color:#a8b2bd;font-style:italic;
}

/* 打印时务必去掉编辑态的所有痕迹 */
@media print{
  .editable{background:none !important;box-shadow:none !important;outline:none !important;}
  .editable.is-empty::before{content:none !important;}
}

/* ---------- 打印 ---------- */
@page{ size:A4 portrait; margin:15mm 13mm 14mm; }
@media print{
  :root{ --fs:10.5pt; }
  body{background:#fff;font-size:var(--fs);line-height:1.72;}
  .page{
    width:auto;min-height:0;margin:0;padding:0;box-shadow:none;
    background:#fff;
  }
  .toolbar,.hint,.screen-only,.topbar{display:none !important;}
  .running-foot{display:flex;}
  h2.sec,h3.sub,h4.sub2,.act,.acts,table.grid,table.info-table,.note,.kv,
  .letterhead,.brandbar,.titleblock,.info-table{
    break-inside:avoid;page-break-inside:avoid;
  }
  h2.sec,h3.sub,h4.sub2{break-after:avoid;page-break-after:avoid;}
  .section.brk{break-before:page;page-break-before:always;}
  a{color:inherit;text-decoration:none;}
  *{-webkit-print-color-adjust:exact;print-color-adjust:exact;}
}
"""

TOOLBAR = """
<div class="topbar screen-only">
  <div class="tip" id="tip">
    <b>编辑正文：</b>点「编辑正文」后直接在页面里改字，改完点「保存副本」。
    <b>导出 PDF：</b>点「一键生成 PDF」→ 目标打印机选「另存为 PDF」→ 纸张 <code>A4</code> →
    边距「默认」→ 勾选「背景图形」→ 保存。版式已按 A4 排好，不会被编辑打乱。
  </div>
  <div class="toolbar btns">
    <button class="ghost" type="button" onclick="zoom(-1)" title="缩小字号">A-</button>
    <button class="ghost" type="button" onclick="zoom(1)" title="放大字号">A+</button>
    <button class="ghost" type="button" id="btn-edit" onclick="toggleEdit()">编辑正文</button>
    <button class="ghost hidden" type="button" id="btn-save" onclick="saveCopy()">保存副本</button>
    <button type="button" onclick="window.print()">一键生成 PDF</button>
  </div>
</div>
"""

JS = """
var BASE = 10.5, cur = BASE;
function zoom(d){
  cur = Math.min(14, Math.max(8.5, cur + d));
  document.documentElement.style.setProperty('--fs', cur + 'pt');
}

/* ---- 编辑模式：只开 contenteditable，不动任何样式表，版式天然不受影响 ---- */
var editing = false;

function refreshEmpty(root){
  var list = (root || document).querySelectorAll('.editable');
  for (var i = 0; i < list.length; i++) {
    var el = list[i];
    // textContent 去掉 <br> 占位后为空，即视为删空
    var t = (el.textContent || '').replace(/\\u00a0/g, ' ').trim();
    if (t === '') { el.classList.add('is-empty'); }
    else { el.classList.remove('is-empty'); }
  }
}

function toggleEdit(){
  editing = !editing;
  document.body.classList.toggle('editing', editing);
  var btn = document.getElementById('btn-edit');
  var save = document.getElementById('btn-save');
  btn.textContent = editing ? '完成编辑' : '编辑正文';
  btn.classList.toggle('on', editing);
  save.classList.toggle('hidden', !editing);
  if (editing) {
    refreshEmpty();
  } else {
    // 退出时移除浏览器留下的 spellcheck 红线与选区高亮
    Array.prototype.forEach.call(
      document.querySelectorAll('[contenteditable]'),
      function(el){ el.removeAttribute('spellcheck'); el.blur(); }
    );
  }
}

document.addEventListener('input', function(e){
  var el = e.target;
  if (el && el.classList && el.classList.contains('editable')) refreshEmpty(el.parentNode || document);
});

/* ---- 粘贴净化：只取纯文本，杜绝外部 HTML/CSS 污染排版 ---- */
document.addEventListener('paste', function(e){
  var el = e.target;
  if (!el || !el.classList || !el.classList.contains('editable')) return;
  e.preventDefault();
  if (!e.clipboardData) return;
  var text = e.clipboardData.getData('text/plain');
  if (!text) {
    /* 从 Word/网页复制时可能没有 text/plain，退而用 HTML 源剥标签取文字，
       仍然不执行任何 HTML，避免样式污染 */
    var html = e.clipboardData.getData('text/html');
    if (!html) return;
    var box = document.createElement('div');
    box.innerHTML = html;
    text = (box.textContent || '').replace(/\\s+/g, ' ').trim();
    if (!text) return;
  }
  document.execCommand('insertText', false, text);
});

/* ---- 导出当前编辑结果为新的 HTML 副本（保留排版与校徽 base64） ---- */
function saveCopy(){
  var clone = document.documentElement.cloneNode(true);
  var body = clone.querySelector('body');
  if (body) body.classList.remove('editing');
  Array.prototype.forEach.call(
    clone.querySelectorAll('[contenteditable]'),
    function(el){
      el.removeAttribute('contenteditable');
      el.removeAttribute('spellcheck');
      el.classList.remove('is-empty');
    }
  );
  var html = '<!DOCTYPE html>\\n' + clone.outerHTML;
  var blob = new Blob([html], {type:'text/html;charset=utf-8'});
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = (document.title || '教学设计') + '.html';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(function(){ URL.revokeObjectURL(a.href); }, 2000);
}
"""

# ---------------------------------------------------------------- 工具

# INLINE_PATTERNS / ROLE_LABELS / is_role 见 plan_common（与 MD 版共用）

def inline(text):
    """转义并处理行内标记：**加粗**、__着重__。"""
    out = html.escape(str(text), quote=False)
    for pat, rep in INLINE_PATTERNS:
        out = pat.sub(rep, out)
    return out


# ---------------------------------------------------------------- 渲染

# 可编辑标记。plaintext-only 只接受纯文本，粘贴进来的 HTML/CSS 会被浏览器剥除，
# 因此用户改字不会把既定版式（字号/颜色/间距）带乱。
# spellcheck=false 关闭浏览器红色波浪线，避免误以为排版出错。
def ce(cls="", ph=""):
    """生成可编辑节点的属性串（**不含结尾的 `>`**）。

    统一走这一个出口，避免各处漏写 data-ph（漏了就删空不显示提示）。
    cls 会自动补上 editable 类；ph 为删空后的灰色占位文案。

    注意：返回值以空格开头、以 `data-ph` 结尾，**故意不带 `>`**。
    调用处统一写成 `<p%s>文字</p>`，闭合的 `>` 由模板提供。
    早先本函数末尾拼了一个 `>`，而调用处模板里还有一个，
    于是每个可编辑节点都多出一个 `>` 字符，
    逐行提取文案时表现为「每行前面挂一个 >」。
    """
    cls = (cls + " editable").strip()
    out = ' class="%s" contenteditable="plaintext-only" spellcheck="false"' % cls
    if ph:
        out += ' data-ph="%s"' % ph
    return out


def ed_attr(placeholder=""):
    """可编辑属性；带占位提示，用于删空后的视觉兜底。"""
    return ce(ph=placeholder)


# ---------------------------------------------------------------- 渲染

def render_para(text):
    role = is_role(text)
    if role:
        lab, rest = role
        return ('<p%s><b>%s：</b>%s</p>'
                % (ce("nt", "点击输入文字"), lab, inline(rest)))
    return ('<p%s>%s</p>' % (ed_attr("点击输入段落文字"), inline(text)))


def render_activities(acts):
    out = ['<div class="acts">']
    for act in acts:
        label = act.get("label", "活动")
        items = act.get("items") or []
        body = []
        for it in items:
            role = is_role(it)
            if role:
                body.append('<p%s><b>%s：</b>%s</p>'
                            % (ce(ph="点击输入文字"), role[0], inline(role[1])))
            else:
                body.append('<p%s>%s</p>' % (ed_attr("点击输入"), inline(it)))
        out.append('<div class="act"><div%s>%s</div>'
                   '<div class="act-body">%s</div></div>'
                   % (ce("act-label", "阶段标签"), inline(label), "".join(body)))
    out.append('</div>')
    return "".join(out)


def render_table(spec):
    head = spec.get("header") or []
    rows = spec.get("rows") or []
    out = ['<table class="grid">']
    if head:
        out.append('<thead><tr>' + "".join(
            '<th%s>%s</th>' % (ce(ph="表头"), inline(h)) for h in head) + '</tr></thead>')
    out.append('<tbody>')
    for r in rows:
        cells = r if isinstance(r, (list, tuple)) else [r]
        out.append('<tr>' + "".join(
            '<td%s>%s</td>' % (ce(ph="点击输入"), inline(c)) for c in cells) + '</tr>')
    out.append('</tbody></table>')
    return "".join(out)


def render_info(rows):
    """info_rows 每项为 [标签, 内容]，两两并排成一行。"""
    pairs = info_pairs(rows)
    if not pairs:
        return ""
    out = ['<table class="info-table"><tbody>']
    for chunk in chunk_pairs(pairs, 2):
        cells = []
        for k, v in chunk:
            cells.append('<td%s>%s</td>' % (ce("k", "标签"), inline(k)))
            cells.append('<td%s>%s</td>' % (ce("v", "内容"), inline(v)))
        out.append("<tr>" + "".join(cells) + "</tr>")
    out.append("</tbody></table>")
    return "".join(out)


def render_kv(pairs):
    return "".join(
        '<div class="kv"><div%s>%s</div><div%s>%s</div></div>'
        % (ce("kv-k", "字段名"), inline(k), ce("kv-v", "内容"), inline(v))
        for k, v in pairs)


def render_body(obj):
    """渲染一个小节的正文（不含标题），支持多种块类型。"""
    out = []
    for p in obj.get("paragraphs", []) or []:
        out.append(render_para(p))
    items = obj.get("bullets", []) or []
    if items:
        out.append('<ul class="bullets">'
                   + "".join('<li%s>%s</li>' % (ed_attr("点击输入列表项"), inline(i))
                             for i in items) + '</ul>')
    items = obj.get("numbered", []) or []
    if items:
        out.append('<ol class="nums">'
                   + "".join('<li%s>%s</li>' % (ed_attr("点击输入列表项"), inline(i))
                             for i in items) + '</ol>')
    acts = obj.get("activities", []) or []
    if acts:
        out.append(render_activities(acts))
    if obj.get("table"):
        out.append(render_table(obj["table"]))
    if obj.get("kv"):
        out.append(render_kv(obj["kv"]))
    if obj.get("note"):
        out.append('<div%s>%s</div>'
                   % (ce("note", "提示内容"), inline(obj["note"])))
    if obj.get("block"):
        # 结构化多行内容（板书设计、思维导图、流程步骤）。
        # 与 MD 版的 block 字段对应：这里保留换行与缩进层级，
        # 全角空格（U+3000）原样输出，缩进感与纯文本稿一致。
        out.append('<div%s>%s</div>'
                   % (ce("note block", "块内容"), inline(obj["block"])
                      .replace("\n", "<br>")))
    return "".join(out)


def render_section(sec, level=1):
    heading = sec.get("heading")
    out = ['<section class="section%s">' % (" brk" if sec.get("page_break") else "")]
    if heading:
        cls = {1: "sec", 2: "sub"}.get(level, "sub2")
        out.append('<h%s%s>%s</h%s>'
                   % (level + 1, ce(cls, "小节标题"), inline(heading), level + 1))
    out.append(render_body(sec))
    for sub in sec.get("subsections", []) or []:
        out.append(render_section(sub, level + 1))
    out.append("</section>")
    return "".join(out)


def build(plan, base_css_note=""):
    title = plan.get("title", "教学设计")
    subtitle = plan.get("subtitle", "")
    meta = plan.get("meta") or {}
    footer = plan.get("footer") or title

    # 校徽放在信笺抬头层；JSON 未显式关闭时默认显示
    badge = plan.get("badge", True)
    badge_uri = badge_data_uri() if badge else ""
    school_cn = meta.get("school") or meta.get("left") or "呼和浩特市实验学校"
    school_en = meta.get("school_en") or "Hohhot Experimental Middle School"
    tag = meta.get("tag", "")

    body = ['<div class="page">']

    # 顶部品牌色条
    body.append('<div class="brandbar"><i></i><i></i><i></i></div>')

    # 信笺抬头：校徽 + 校名 + 模型标
    lh = ['<header class="letterhead">']
    if badge_uri:
        lh.append('<span class="lh-badge"><img src="%s" alt="校徽"></span>' % badge_uri)
    lh.append('<span class="lh-id">'
              '<b%s>%s</b>'
              '<i%s>%s</i></span>'
              % (ce("lh-cn", "学校名称"), inline(school_cn),
                 ce("lh-en", "School Name"), inline(school_en)))
    if tag:
        lh.append('<span%s>%s</span>' % (ce("lh-tag", "模型"), inline(tag)))
    lh.append('</header>')
    body.append("".join(lh))
    body.append('<div class="lh-rule"></div>')

    # 标题区
    body.append('<div class="titleblock">')
    body.append('<h1%s>%s</h1>' % (ce("doc-title", "课题标题"), inline(title)))
    if subtitle:
        body.append('<p%s>%s</p>' % (ce("doc-subtitle", "副标题"), inline(subtitle)))
    body.append('<div class="head-rule"><i></i><i></i></div></div>')

    body.append(render_info(plan.get("info_rows") or []))
    for sec in plan.get("sections", []) or []:
        body.append(render_section(sec))
    body.append("</div>")
    body.append('<div class="running-foot"><span>%s</span></div>' % inline(footer))

    doc = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%s</title>
<style>%s
/* 行内标记 __着重__：校徽绿浅底，保持与主色系统一致 */
.hl{background:linear-gradient(transparent 60%%,rgba(82,176,90,.30) 0);padding:0 1px;}
</style>
</head>
<body>
%s
<script>%s</script>
</body>
</html>
""" % (inline(title), CSS, TOOLBAR + "".join(body), JS)

    # ---- 出厂自检 ----
    # 每个可编辑节点的属性串后必须紧跟标签的闭合 >，不能多出第二个。
    # 早先 ce() 末尾拼了 ">" 而模板里还有一个，导致每个节点多一个 ">" 字符，
    # 下游逐行提取文案时表现为「每行前面挂一个 >」。这里做兜底拦截。
    if '">>' in doc or ">>" in doc.replace("&gt;", ""):
        import re as _re
        bad = _re.search(r'.{0,60}">>.{0,40}', doc)
        raise AssertionError(
            "HTML 自检失败：检测到多余的 '>'（应为标签闭合符）。\n"
            "疑似 ce() 又返回了带 '>' 的属性串。\n上下文：%s"
            % (bad.group(0) if bad else ""))
    return doc


def main():
    if len(sys.argv) != 3:
        print("用法: python make_lesson_html.py <plan.json> <output.html>")
        sys.exit(1)
    with open(sys.argv[1], "r", encoding="utf-8") as f:
        plan = json.load(f)
    out = build(plan)
    with open(sys.argv[2], "w", encoding="utf-8") as f:
        f.write(out)
    print("OK:", sys.argv[2])


if __name__ == "__main__":
    main()

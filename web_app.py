from __future__ import annotations

import html
import io
import json
import os
import re

import streamlit as st
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openai import (
    APIConnectionError,
    APITimeoutError,
    AuthenticationError,
    OpenAI,
    RateLimitError,
)


SYSTEM_PROMPT = """你是一位有 10 年经验的 HR 与简历排版顾问。请把求职者的口语化素材整理成结构化简历数据，程序会据此排版成 A4 Word 简历。

【内容规则】
1. 只使用素材中出现的事实。严禁编造公司、学校、时间、数字、奖项、证书、课程成绩。
2. 素材里缺失的信息写成带方括号的占位符，例如 [手机号]、[学校名称]、[起止时间]，绝不用猜测值填充。
3. 素材里出现过数字才写数字；没有就用定性描述，不要自行补“提升 30%”“排名前 10%”这类数据。
4. 每条要点采用「领起词 + 具体动作 + 方法 + 结果」的结构，单条不超过 80 字，删掉“负责一些”“参与了”这类空话。
5. 每条要点开头用 **双星号** 标出 2 到 6 个字的领起词（如 **项目背景**、**核心技术**、**学分绩点**），同一句里其余文字不要再加粗。

【板块规则】
- 板块顺序固定：教育经历 → 项目经历 → 实习经历 → 在校经历。素材没有的板块直接省略，禁止硬凑。
- 板块标题固定为 4 个字，并配一个图标：教育经历 🎓、项目经历 📁、实习经历 💼、在校经历 🏫。

【条目规则】
- 教育经历：period 写起止时间（如 2022.9-2026.6），org 写学校名，role 写「专业 | 学历」（如 计算机科学与技术 | 本科），title 留空字符串。
- 项目经历：title 写项目名称，role 写担任角色（如 项目负责人），period 与 org 留空字符串。
- 实习经历：period 写起止时间，org 写公司或单位名，role 写岗位名，title 留空字符串。
- 在校经历：title 写职务名（如 班级团支部书记），period、org、role 留空字符串。
- 同一个板块有多条内容时，按时间倒序排列。

【输出规则】
只输出一个 JSON 对象，不要 Markdown 代码块，不要任何解释文字。字段结构如下：
{
  "name": "姓名，缺失时写 [姓名]",
  "intent": "求职意向",
  "contact": [
    {"label": "电话", "value": "..."},
    {"label": "邮箱", "value": "..."},
    {"label": "性别", "value": "..."},
    {"label": "年龄", "value": "..."},
    {"label": "最高学历", "value": "..."},
    {"label": "政治面貌", "value": "..."}
  ],
  "sections": [
    {
      "title": "教育经历",
      "icon": "🎓",
      "entries": [
        {
          "period": "2022.9-2026.6",
          "org": "[学校名称]",
          "role": "计算机科学与技术 | 本科",
          "title": "",
          "bullets": ["**学分绩点** 3.04/5.0，班级排名前 30%"]
        }
      ]
    }
  ]
}

contact 固定 6 项、顺序不变，缺失的值写占位符。sections 里每个 entry 必须包含 period、org、role、title、bullets 五个键，没有内容的键给空字符串或空数组。
"""


# ---------- 主题 ----------
NAVY = "1F3864"       # 板块色块 / 主分隔线
INK = "1A1A1A"        # 正文
MUTED = "595959"      # 字段名
HAIRLINE = "8EA9DB"   # 板块下方的细分隔线
FONT = "微软雅黑"

PAGE_W_CM = 21.0
PAGE_H_CM = 29.7
MARGIN_CM = 2.0
USABLE_CM = PAGE_W_CM - 2 * MARGIN_CM  # 17.0

BODY_PT = 10.5
TITLE_PT = 12
NAME_PT = 20

_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
_TWIPS_PER_CM = 567.0

_PPR_SUCCESSORS = (
    "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
    "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)
_TCPR_SUCCESSORS = (
    "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
    "w:hideMark", "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge",
    "w:tcPrChange",
)


# ---------- 底层小工具 ----------
def _set_font(run, size=None, bold=None, color=None, name=FONT):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:eastAsia"), name)
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)


def _run(paragraph, text, size=BODY_PT, color=INK, bold=False):
    if not text:
        return None
    run = paragraph.add_run(text)
    _set_font(run, size=size, bold=bold, color=color)
    return run


def add_rich(paragraph, text, size=BODY_PT, color=INK, bold=False):
    """按 **双星号** 切分，把标记内文字加粗。"""
    text = text or ""
    pos = 0
    for match in _BOLD_RE.finditer(text):
        if match.start() > pos:
            _run(paragraph, text[pos:match.start()], size, color, bold)
        _run(paragraph, match.group(1), size, color, True)
        pos = match.end()
    if pos < len(text):
        _run(paragraph, text[pos:], size, color, bold)


def _p_border(paragraph, edge="bottom", sz=6, color=NAVY, space=1):
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = ppr.find(qn("w:pBdr"))
    if pbdr is None:
        pbdr = OxmlElement("w:pBdr")
        ppr.insert_element_before(pbdr, *_PPR_SUCCESSORS)
    el = OxmlElement(f"w:{edge}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    pbdr.append(el)


def _cell_border(cell, edge="bottom", sz=8, color=NAVY):
    tcpr = cell._tc.get_or_add_tcPr()
    tcborders = tcpr.find(qn("w:tcBorders"))
    if tcborders is None:
        tcborders = OxmlElement("w:tcBorders")
        tcpr.insert_element_before(tcborders, *_TCPR_SUCCESSORS)
    el = tcborders.find(qn(f"w:{edge}"))
    if el is None:
        el = OxmlElement(f"w:{edge}")
        tcborders.append(el)
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), color)


def _shade(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = tcpr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcpr.insert_element_before(shd, *_TCPR_SUCCESSORS)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def _cell_margins(cell, top=0.0, bottom=0.0, left=0.0, right=0.0):
    tcpr = cell._tc.get_or_add_tcPr()
    tcm = tcpr.find(qn("w:tcMar"))
    if tcm is None:
        tcm = OxmlElement("w:tcMar")
        tcpr.insert_element_before(
            tcm, "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark",
            "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange",
        )
    for tag, cm in (("w:top", top), ("w:left", left), ("w:bottom", bottom), ("w:right", right)):
        el = tcm.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            tcm.append(el)
        el.set(qn("w:w"), str(int(round(cm * _TWIPS_PER_CM))))
        el.set(qn("w:type"), "dxa")


def _no_borders(table):
    tblpr = table._tbl.tblPr
    existing = tblpr.find(qn("w:tblBorders"))
    if existing is not None:
        tblpr.remove(existing)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "FFFFFF")
        borders.append(el)
    tblpr.insert_element_before(
        borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
        "w:tblCaption", "w:tblDescription", "w:tblPrChange",
    )


def _set_col_width(table, index, cm):
    table.columns[index].width = Cm(cm)
    for cell in table.columns[index].cells:
        cell.width = Cm(cm)


def _blank(paragraph):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0


def _spacer(doc, pt):
    p = doc.add_paragraph()
    _blank(p)
    run = p.add_run("")
    run.font.size = Pt(pt)
    return p


def _rule(doc, sz=6, color=HAIRLINE, before=2, after=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.0
    run = p.add_run("")
    run.font.size = Pt(1)
    _p_border(p, "bottom", sz=sz, color=color, space=1)
    return p


# ---------- 页面 ----------
def _setup_page(doc):
    section = doc.sections[0]
    section.page_width = Cm(PAGE_W_CM)
    section.page_height = Cm(PAGE_H_CM)
    section.left_margin = Cm(MARGIN_CM)
    section.right_margin = Cm(MARGIN_CM)
    section.top_margin = Cm(1.7)
    section.bottom_margin = Cm(1.7)
    section.header_distance = Cm(1.0)
    section.footer_distance = Cm(1.0)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(BODY_PT)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), FONT)
    rfonts.set(qn("w:ascii"), FONT)
    rfonts.set(qn("w:hAnsi"), FONT)
    pf = normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.15


# ---------- 版块 ----------
def _section_header(doc, title, icon=""):
    title = (title or "经历").strip()
    label = f"{icon}  {title}" if icon else title
    box_w = 0.42 + 0.44 * len(title) + (0.85 if icon else 0.0)
    box_w = max(2.6, min(8.0, box_w))

    table = doc.add_table(rows=1, cols=2)
    _no_borders(table)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    _set_col_width(table, 0, box_w)
    _set_col_width(table, 1, USABLE_CM - box_w)

    row = table.rows[0]
    row.height = Cm(0.66)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    left, right = row.cells

    _cell_margins(left, top=0.0, bottom=0.0, left=0.18, right=0.18)
    _cell_margins(right, top=0.0, bottom=0.03, left=0.14, right=0.0)
    _shade(left, NAVY)
    _cell_border(right, "bottom", sz=8, color=NAVY)

    left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    right.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM

    p = left.paragraphs[0]
    _blank(p)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    _run(p, label, size=TITLE_PT, color="FFFFFF", bold=True)


def _entry(doc, entry):
    title = (entry.get("title") or "").strip()
    period = (entry.get("period") or "").strip()
    org = (entry.get("org") or "").strip()
    role = (entry.get("role") or "").strip()

    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(7)
    pf.space_after = Pt(2)
    pf.line_spacing = 1.1
    pf.tab_stops.add_tab_stop(Cm(USABLE_CM), WD_TAB_ALIGNMENT.RIGHT)

    if title:
        add_rich(p, title, bold=True)
        if role:
            _run(p, f" | {role}", bold=True)
    else:
        if period:
            _run(p, period, bold=True)
        right = " | ".join(x for x in (org, role) if x)
        if right:
            if period:
                p.add_run("\t")
            _run(p, right, bold=True)

    for text in entry.get("bullets") or []:
        _bullet(doc, str(text))


def _bullet(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.5)
    pf.first_line_indent = Cm(-0.5)
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    pf.line_spacing = 1.15
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.tab_stops.add_tab_stop(Cm(0.5))
    _run(p, "•\t", size=BODY_PT, color=NAVY, bold=True)
    add_rich(p, text)


# ---------- 对外接口 ----------
def build_resume(data: dict, photo_bytes: bytes = None) -> Document:
    doc = Document()
    _setup_page(doc)

    # 👇 核心修改：如果有照片，使用一行两列表格布局
    if photo_bytes:
        table = doc.add_table(rows=1, cols=2)
        _no_borders(table)
        _set_col_width(table, 0, 13.5) # 左侧文字宽
        _set_col_width(table, 1, 3.5)  # 右侧照片宽
        
        left_cell = table.cell(0, 0)
        right_cell = table.cell(0, 1)

        # 左上：姓名与意向
        p_name = left_cell.paragraphs[0]
        pf = p_name.paragraph_format
        pf.space_before = Pt(0); pf.space_after = Pt(4)
        _run(p_name, data.get("name") or "[姓名]", size=NAME_PT, color=NAVY, bold=True)
        if data.get("intent"):
            _run(p_name, f"  求职意向：{data.get('intent')}", size=11, color=NAVY, bold=True)

        # 左下：联系方式（嵌套小表格）
        contact = [c for c in (data.get("contact") or []) if str(c.get("value", "")).strip()]
        if contact:
            t_contact = left_cell.add_table(rows=(len(contact)+1)//2, cols=2)
            _no_borders(t_contact)
            t_contact.autofit = False
            _set_col_width(t_contact, 0, 6.5)
            _set_col_width(t_contact, 1, 6.5)
            for idx, row in enumerate(t_contact.rows):
                row.height = Cm(0.62)
                for cell in row.cells:
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    _cell_margins(cell, top=0.02, bottom=0.02, left=0.0, right=0.2)
            for i, item in enumerate(contact):
                cell = t_contact.cell(i // 2, i % 2)
                p_c = cell.paragraphs[0]
                _blank(p_c)
                _run(p_c, f"{item.get('label', '')}：", size=BODY_PT, color=MUTED)
                _run(p_c, str(item.get("value", "")), size=BODY_PT, color=INK)

        # 右侧：插入照片
        p_img = right_cell.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_img.add_run().add_picture(io.BytesIO(photo_bytes), width=Cm(2.5))

        # 移除外层表格最前面默认的空行
        first_p = doc.paragraphs[0]._element
        if first_p.tag.endswith('p') and len(first_p.getchildren()) == 0:
            first_p.getparent().remove(first_p)
    else:
        # 无照片时：保持原本的流式布局
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0); pf.space_after = Pt(4)
        _run(p, data.get("name") or "[姓名]", size=NAME_PT, color=NAVY, bold=True)
        if data.get("intent"):
            _run(p, f"  求职意向：{data.get('intent')}", size=11, color=NAVY, bold=True)
            
        contact = [c for c in (data.get("contact") or []) if str(c.get("value", "")).strip()]
        if contact:
            t_contact = doc.add_table(rows=(len(contact)+1)//2, cols=2)
            _no_borders(t_contact)
            t_contact.autofit = False
            _set_col_width(t_contact, 0, 8.5)
            _set_col_width(t_contact, 1, 8.5)
            for i, item in enumerate(contact):
                cell = t_contact.cell(i // 2, i % 2)
                p_c = cell.paragraphs[0]
                _blank(p_c)
                _run(p_c, f"{item.get('label', '')}：", size=BODY_PT, color=MUTED)
                _run(p_c, str(item.get("value", "")), size=BODY_PT, color=INK)

    _rule(doc, sz=18, color=NAVY, before=2, after=2)

    sections = data.get("sections") or []
    for section in sections:
        _spacer(doc, 3)
        _section_header(doc, section.get("title", ""), section.get("icon", ""))
        for entry in section.get("entries") or []:
            _entry(doc, entry)
        _rule(doc, sz=6, color=HAIRLINE, before=3, after=0)

    return doc


def build_resume_bytes(data: dict, photo_bytes: bytes = None) -> bytes:
    doc = build_resume(data, photo_bytes)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


MAX_INPUT_CHARS = 8000

def parse_json(text: str) -> dict:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", text, re.S)
        if not match:
            raise
        return json.loads(match.group(0))

def normalize(data: dict) -> dict:
    if not isinstance(data, dict):
        raise ValueError("模型返回的不是 JSON 对象")

    out = {
        "name": str(data.get("name") or "[姓名]").strip(),
        "intent": str(data.get("intent") or "").strip(),
        "contact": [],
        "sections": [],
    }

    for item in data.get("contact") or []:
        if not isinstance(item, dict):
            continue
        value = item.get("value")
        if value is None or not str(value).strip():
            continue
        out["contact"].append({"label": str(item.get("label") or "").strip(), "value": str(value).strip()})

    for section in data.get("sections") or []:
        if not isinstance(section, dict):
            continue
        entries: list[dict] = []
        for raw_entry in section.get("entries") or []:
            if not isinstance(raw_entry, dict):
                continue
            entry = {
                "title": str(raw_entry.get("title") or "").strip(),
                "period": str(raw_entry.get("period") or "").strip(),
                "org": str(raw_entry.get("org") or "").strip(),
                "role": str(raw_entry.get("role") or "").strip(),
                "bullets": [str(b).strip() for b in (raw_entry.get("bullets") or []) if b is not None and str(b).strip()],
            }
            if any(entry[k] for k in ("title", "period", "org", "role")):
                entries.append(entry)
        if entries:
            out["sections"].append({"title": str(section.get("title") or "经历").strip(), "icon": str(section.get("icon") or "").strip(), "entries": entries})
    return out

_PLACEHOLDER_RE = re.compile(r"\[[^\[\]]{1,40}\]")
def count_placeholders(data: dict) -> int:
    def walk(node):
        if isinstance(node, str):
            return len(_PLACEHOLDER_RE.findall(node))
        if isinstance(node, dict):
            return sum(walk(v) for v in node.values())
        if isinstance(node, (list, tuple)):
            return sum(walk(v) for v in node)
        return 0
    return walk(data)


# ---------------- HTML 预览 ----------------
def _rich(text: str) -> str:
    escaped = html.escape(str(text))
    return re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", escaped)

def preview_html(data: dict) -> str:
    parts = [
        '<div style="font-family:\'Microsoft YaHei\',sans-serif;color:#1a1a1a;font-size:13.5px;line-height:1.75">',
        f'<div style="font-size:24px;font-weight:700;color:#1F3864;letter-spacing:3px">{html.escape(data["name"])}</div>',
    ]
    if data["intent"]:
        parts.append(f'<div style="color:#1F3864;font-weight:600;margin-top:2px">求职意向：{html.escape(data["intent"])}</div>')

    if data["contact"]:
        rows = []
        cells = data["contact"]
        for i in range(0, len(cells), 2):
            tds = "".join(
                '<td style="width:50%;padding:3px 0">'
                f'<span style="color:#595959">{html.escape(c["label"])}：</span>{html.escape(c["value"])}</td>'
                for c in cells[i:i + 2]
            )
            rows.append(f"<tr>{tds}</tr>")
        parts.append('<table style="width:100%;border-collapse:collapse;margin-top:8px">' + "".join(rows) + "</table>")

    parts.append('<div style="border-top:2px solid #1F3864;margin:10px 0 12px 0"></div>')

    for section in data["sections"]:
        icon = f"{html.escape(section['icon'])}&nbsp;&nbsp;" if section["icon"] else ""
        parts.append(
            '<div style="display:flex;align-items:flex-end;margin:14px 0 6px 0">'
            f'<span style="background:#1F3864;color:#fff;font-weight:700;font-size:15px;padding:3px 10px;white-space:nowrap">{icon}{html.escape(section["title"])}</span>'
            '<span style="flex:1;border-bottom:1px solid #1F3864;margin-bottom:1px"></span></div>'
        )
        for entry in section["entries"]:
            if entry["title"]:
                head = f"<b>{_rich(entry['title'])}</b>"
                if entry["role"]: head += f" <b>| {_rich(entry['role'])}</b>"
            else:
                right = " | ".join(x for x in (entry["org"], entry["role"]) if x)
                head = f'<div style="display:flex;justify-content:space-between"><b>{_rich(entry["period"])}</b><b>{_rich(right)}</b></div>'
            parts.append(f'<div style="margin-top:8px">{head}</div>')
            for bullet in entry["bullets"]:
                parts.append(f'<div style="display:flex;gap:6px;margin-top:3px"><span style="color:#1F3864;font-weight:700">•</span><span style="text-align:justify">{_rich(bullet)}</span></div>')
        parts.append('<div style="border-top:1px solid #8EA9DB;margin:10px 0 0 0"></div>')

    parts.append("</div>")
    return "".join(parts)


# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="AI 简历生成器", page_icon="📄", layout="centered")

def get_api_key() -> str:
    try:
        key = st.secrets.get("DEEPSEEK_API_KEY", "")
    except Exception:
        key = ""
    return key or os.getenv("DEEPSEEK_API_KEY", "")

@st.cache_resource(show_spinner=False)
def get_client(api_key: str) -> OpenAI:
    return OpenAI(api_key=api_key, base_url="https://api.deepseek.com", timeout=90.0, max_retries=3)

def generate_resume_data(client: OpenAI, raw: str) -> dict:
    resp = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"这是求职者的原始素材：\n\n{raw}"},
        ],
        response_format={"type": "json_object"},
        temperature=0.3,
        max_tokens=3000,
    )
    return parse_json((resp.choices[0].message.content or "").strip())

if "resume" not in st.session_state:
    st.session_state.resume = None

st.title("📄 智能 AI 简历生成器")
st.caption("贴入大白话经历，生成与模板版式一致、可直接投递的 Word 简历。")

env_key = get_api_key()
api_key = env_key or st.sidebar.text_input("DeepSeek API Key", type="password")
if not env_key:
    st.sidebar.caption("建议把 Key 写入 `.streamlit/secrets.toml` 的 `DEEPSEEK_API_KEY`。")

with st.form("resume_form"):
    user_input = st.text_area(
        "你的真实经历（大白话即可）",
        height=220,
        placeholder="例如：我是计算机应届生，做过一个 Spring Boot + Vue 的电商管理系统……",
    )
    # 👇 新增：照片上传组件
    uploaded_photo = st.file_uploader("上传一寸证件照（可选，留空则无照片）", type=["jpg", "jpeg", "png"])
    
    submitted = st.form_submit_button("🚀 生成简历", type="primary")

if submitted:
    if not api_key:
        st.warning("⚠️ 请先配置 API Key")
    elif not user_input.strip():
        st.warning("⚠️ 请先输入你的经历")
    elif len(user_input) > MAX_INPUT_CHARS:
        st.warning(f"⚠️ 素材超过 {MAX_INPUT_CHARS} 字，请精简后再试")
    else:
        try:
            with st.spinner("AI 正在整理内容并排版…"):
                data = normalize(generate_resume_data(get_client(api_key), user_input))
                if not data["sections"]:
                    raise ValueError("模型没有返回任何经历板块，请把素材写得更具体一些")
                
                # 👇 提取照片数据并传入函数
                photo_bytes = uploaded_photo.getvalue() if uploaded_photo else None
                docx_bytes = build_resume_bytes(data, photo_bytes)
                
            st.session_state.resume = {"data": data, "docx": docx_bytes}
            st.success("✅ 简历已生成，可直接下载 Word 文件")
        except AuthenticationError:
            st.error("API Key 无效，请检查后重试。")
        except RateLimitError:
            st.error("请求过于频繁或余额不足，请稍后再试。")
        except (APIConnectionError, APITimeoutError):
            st.error("网络连接失败，请检查网络后重试。")
        except Exception as exc:
            st.error(f"生成失败：{exc}")

# 结果存在 session_state 里，任何 rerun 都不会丢
result = st.session_state.resume
if result:
    data = result["data"]
    missing = count_placeholders(data)
    if missing:
        st.info(f"提示：内容里还有 {missing} 处 `[占位符]`，请下载后补全再投递。")

    st.download_button(
        "📥 下载 Word 简历 (.docx)",
        data=result["docx"],
        file_name="我的简历.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )
    st.markdown("---")
    # 网页端暂时只预览文字，看照片效果请下载 Word
    st.markdown(preview_html(data), unsafe_allow_html=True)
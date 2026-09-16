# -*- coding: utf-8 -*-
"""
Word 工具集
依赖：pip install python-docx openpyxl
运行：python main.py

Tab1：标题处理（识别、重编号、格式统一）
Tab2：表格/正文提取到 Excel
"""

import os
import re
import traceback
import threading
import queue
import warnings
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

warnings.filterwarnings("ignore")

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


# =========================================================
# 常量
# =========================================================

TEMPLATE_PRESETS = {
    1: ["第{1}章", "第{1}篇", "第{1}部分", "第{1}编", "{1}.", "{1}、", "{1}"],
    2: ["{1}.{2}", "{1}.{2}.", "第{1}节", "（{2}）", "({2})", "{2}.", "{2}、", "{2}"],
    3: ["{1}.{2}.{3}", "{1}.{2}.{3}.", "（{3}）", "({3})", "{3}.", "{3}、", "{3}"],
    4: ["{1}.{2}.{3}.{4}", "（{4}）", "({4})", "{4}.", "{4}、", "{4}"],
    5: ["{1}.{2}.{3}.{4}.{5}", "{5}.", "{5}、", "{5}"],
}

TEMPLATE_DEFAULT = {1: "第{1}章", 2: "{1}.{2}", 3: "{1}.{2}.{3}", 4: "{4}", 5: "{5}"}

SEP_PRESETS = ["空格", "顿号（、）", "点（.）", "冒号（:）", "短横（-）", "无", "制表符"]
SEP_MAP = {
    "空格": " ", "顿号（、）": "、", "点（.）": ".", "冒号（:）": ":",
    "短横（-）": "-", "无": "", "制表符": "\t",
}

NUMFMT_PRESETS = ["阿拉伯数字", "中文数字"]
NUMFMT_MAP = {"阿拉伯数字": "arabic", "中文数字": "chinese"}

LINE_SPACING_PRESETS = ["单倍行距", "1.5 倍行距", "2 倍行距",
                        "固定值(磅)", "最小值(磅)"]
LINE_SPACING_MAP = {
    "单倍行距": "single",
    "1.5 倍行距": "1.5",
    "2 倍行距": "double",
    "固定值(磅)": "exact",
    "最小值(磅)": "atleast",
}

FONT_PRESETS = [
    "宋体", "黑体", "楷体", "仿宋", "微软雅黑", "等线",
    "华文中宋", "华文楷体", "华文仿宋", "华文行楷",
    "方正小标宋简体", "方正黑体简体",
    "Times New Roman", "Arial", "Calibri", "Cambria",
    "Georgia", "Verdana", "Tahoma", "Courier New",
]

FONT_SIZE_MAP = {
    "初号": 42.0, "小初": 36.0,
    "一号": 26.0, "小一": 24.0,
    "二号": 22.0, "小二": 18.0,
    "三号": 16.0, "小三": 15.0,
    "四号": 14.0, "小四": 12.0,
    "五号": 10.5, "小五": 9.0,
    "六号": 7.5, "小六": 6.5,
    "七号": 5.5, "八号": 5.0,
}
FONT_SIZE_PRESETS = list(FONT_SIZE_MAP.keys())


def parse_font_size(text):
    if text is None:
        return None
    text = str(text).strip()
    if not text:
        return None
    if text in FONT_SIZE_MAP:
        return FONT_SIZE_MAP[text]
    try:
        return float(text)
    except ValueError:
        pass
    m = re.search(r'(\d+(?:\.\d+)?)', text)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    return None


TEMPLATE_HELP_TEXT = """编号模板编写原则
============================

一、占位符
    {1}~{9}  表示第 1~9 级标题的当前序号

二、使用规则
    1. 每级模板中只能引用"本级"或"更高级"的占位符。
    2. 超出本级的占位符会被自动清除。
    3. 下级序号会随上级标题出现而自动重置为 1。
    4. 若上级标题未出现，上级序号按 1 显示。

三、数字格式
    可选"阿拉伯数字"或"中文数字"。

四、字体与字号
    字体下拉框提供常用中英文字体，也可手动输入。
    字号下拉框提供中文号数，也可直接输入磅值。

五、正文范围
    需在标题列表中指定"正文开始标题"和"正文结束标题"：
      - 右键点击标题行 → "设为正文开始" / "设为正文结束"
      - 如果"正文结束"是最后一个标题，它之后的段落一直
        到文档末尾都算正文。

六、导出清单
    "导出清单时去除编号"复选框：
      - 勾选：导出清单的"标题文字"列去掉原编号
      - 不勾选：保留原标题文字
"""


# =========================================================
# 中文数字
# =========================================================

_CN_DIGITS = '零一二三四五六七八九'


def num_to_chinese(n):
    try:
        n = int(n)
    except (TypeError, ValueError):
        return str(n)
    if n <= 0:
        return str(n)
    if n < 10:
        return _CN_DIGITS[n]
    if n < 20:
        return '十' + (_CN_DIGITS[n - 10] if n > 10 else '')
    if n < 100:
        tens = n // 10
        ones = n % 10
        return _CN_DIGITS[tens] + '十' + (_CN_DIGITS[ones] if ones else '')
    if n < 1000:
        h = n // 100
        r = n % 100
        s = _CN_DIGITS[h] + '百'
        if r == 0:
            return s
        if r < 10:
            return s + '零' + _CN_DIGITS[r]
        return s + num_to_chinese(r)
    if n < 10000:
        th = n // 1000
        r = n % 1000
        s = _CN_DIGITS[th] + '千'
        if r == 0:
            return s
        if r < 100:
            return s + '零' + num_to_chinese(r)
        return s + num_to_chinese(r)
    return str(n)


# =========================================================
# 文件名显示
# =========================================================

MAX_FILE_DISPLAY_CHARS = 46


def truncate_path(path, max_chars=MAX_FILE_DISPLAY_CHARS):
    if not path:
        return ''
    if len(path) <= max_chars:
        return path
    sep = '\\' if '\\' in path else '/'
    parts = re.split(r'[\\/]', path)
    filename = parts[-1] if parts else path
    prefix = "..." + sep
    if len(prefix) + len(filename) <= max_chars:
        return prefix + filename
    keep = max_chars - 3
    head = keep // 2
    tail = keep - head
    if head < 1:
        return "..." + filename[-keep:]
    return filename[:head] + "..." + filename[-tail:]


# =========================================================
# 段落遍历（含表格内段落）
# =========================================================

def _all_paragraph_objs(doc):
    from docx.text.paragraph import Paragraph
    result = []
    for p_el in doc.element.body.iter(qn('w:p')):
        result.append(Paragraph(p_el, doc))
    return result


def _is_in_table(p):
    parent = p._p.getparent()
    while parent is not None:
        if parent.tag == qn('w:tc'):
            return True
        if parent.tag == qn('w:body'):
            return False
        parent = parent.getparent()
    return False


# =========================================================
# 目录识别
# =========================================================

def is_toc_paragraph(paragraph):
    try:
        style_name = paragraph.style.name if paragraph.style else ''
    except Exception:
        style_name = ''
    if style_name:
        su = style_name.upper()
        if su.startswith('TOC') or style_name.startswith('目录'):
            return True
    try:
        for instr in paragraph._p.iter(qn('w:instrText')):
            t = instr.text or ''
            if re.search(r'\bTOC\b', t, re.I):
                return True
    except Exception:
        pass
    text = paragraph.text.strip()
    if re.search(r'[\.…·]{3,}\s*\d+\s*$', text):
        return True
    return False


# =========================================================
# 文档解析（标题处理）
# =========================================================

def get_outline_level(paragraph):
    pPr = paragraph._p.pPr
    if pPr is None:
        return None
    ol = pPr.find(W_NS + 'outlineLvl')
    if ol is None:
        return None
    val = ol.get(W_NS + 'val')
    if val is None:
        return None
    try:
        n = int(val) + 1
        if 1 <= n <= 9:
            return n
    except ValueError:
        pass
    return None


def get_style_level(style_name):
    if not style_name:
        return None
    m = re.match(r'^(?:Heading|标题)\s*(\d+)', style_name.strip(), re.I)
    if m:
        return int(m.group(1))
    return None


def detect_number_pattern(text):
    m = re.match(r'^(第[一二三四五六七八九十百千万零〇\d]+[章节篇部分编])\s*', text)
    if m:
        prefix = m.group(1)
        return (2 if '节' in prefix else 1), prefix
    m = re.match(r'^(\d+(?:[\.．\-]\d+)+)([\.．、\s:：\-])', text)
    if m:
        return m.group(1).replace('．', '.').count('.') + 1, m.group(0)
    m = re.match(r'^(\d+)([\.．、\s:：\)）\-])', text)
    if m:
        return 1, m.group(0)
    m = re.match(r'^([一二三四五六七八九十]+)[、.．]\s*', text)
    if m:
        return 1, m.group(0)
    m = re.match(r'^[\(（]\d+[\)）]\s*', text)
    if m:
        return 3, m.group(0)
    return None, None


def load_paragraphs(docx_path):
    doc = Document(docx_path)
    paras = []
    all_paras = _all_paragraph_objs(doc)
    for i, p in enumerate(all_paras):
        text = p.text.strip()
        if not text:
            continue
        try:
            style_name = p.style.name if p.style else ''
        except Exception:
            style_name = ''
        outline = get_outline_level(p)
        is_toc = is_toc_paragraph(p)
        in_table = _is_in_table(p)
        paras.append({
            'index': i,
            'text': text,
            'style': style_name,
            'outline': outline,
            'is_title': False,
            'level': None,
            'source': '',
            'is_toc': is_toc,
            'in_table': in_table,
        })
    return paras


def detect_heading(para, include_table=True):
    if para.get('is_toc'):
        return None, None
    if not include_table and para.get('in_table'):
        return None, None
    if para['outline']:
        return para['outline'], '大纲级别'
    lvl = get_style_level(para['style'])
    if lvl:
        return lvl, '内置样式'
    lvl, _ = detect_number_pattern(para['text'])
    if lvl:
        return lvl, '编号模式'
    return None, None


def remove_old_number(text):
    patterns = [
        r'^第[一二三四五六七八九十百千万零〇\d]+[章节篇部分编][\s:：、.．\-]*',
        r'^\d+(?:[\.．\-]\d+)+[\.．、\s:：\-]*',
        r'^[\(（\[\【]\d+[\)）\]\】][\s.．、:：\-]*',
        r'^\d+[\.．、\s:：\)）\-]+\s*',
        r'^\d+\s+',
        r'^[一二三四五六七八九十]+[、.．\s:：]+',
        r'^[①-⑳]\s*',
        r'^(?=[IVXLCivxlc]+[\.\s])[IVXLCivxlc]+[\.\s]+',
    ]
    for pat in patterns:
        new_text = re.sub(pat, '', text, count=1)
        if new_text != text:
            text = new_text
    return text.strip()


# =========================================================
# 编号生成
# =========================================================

def generate_numbers(titles, templates, separators,
                     number_formats=None, enabled_levels=None):
    if number_formats is None:
        number_formats = {}
    counters = [0] * 10
    result = []
    for t in titles:
        if not t.get('is_title'):
            result.append('')
            continue
        level = t.get('level')
        try:
            level = int(level)
        except (TypeError, ValueError):
            result.append('')
            continue
        if level < 1 or level > 9:
            result.append('')
            continue

        counters[level - 1] += 1
        for i in range(level, 10):
            counters[i] = 0

        if enabled_levels is not None and level not in enabled_levels:
            result.append('')
            continue

        tmpl = templates.get(level, '{' + str(level) + '}')
        s = tmpl
        for i in range(1, level + 1):
            cnt = counters[i - 1]
            if cnt == 0:
                cnt = 1
            fmt = number_formats.get(i, 'arabic')
            cnt_str = num_to_chinese(cnt) if fmt == 'chinese' else str(cnt)
            s = s.replace('{' + str(i) + '}', cnt_str)
        s = re.sub(r'\{\d+\}', '', s)
        result.append(s + separators.get(level, ' '))
    return result


# =========================================================
# 段落文字处理
# =========================================================

def _replace_paragraph_text_preserving_format(p, new_text):
    runs = list(p.runs)
    if not runs:
        p.add_run(new_text)
        return
    base_run = None
    for r in runs:
        if (r.text or '').strip():
            base_run = r
            break
    if base_run is None:
        base_run = runs[0]
    base_run.text = new_text
    for r in runs:
        if r is not base_run:
            r._element.getparent().remove(r._element)


def _remove_paragraph_numbering(p):
    pPr = p._p.pPr
    if pPr is None:
        return
    numPr = pPr.find(qn('w:numPr'))
    if numPr is not None:
        pPr.remove(numPr)


def _remove_style_numbering(doc, num_levels=None):
    for lvl in range(1, 10):
        if num_levels is not None and lvl not in num_levels:
            continue
        for name in (f"标题 {lvl}", f"Heading {lvl}"):
            try:
                st = doc.styles[name]
            except KeyError:
                continue
            el = st.element
            if el is None:
                continue
            pPr = el.find(qn('w:pPr'))
            if pPr is None:
                continue
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                pPr.remove(numPr)


def _apply_heading_style(p, doc, level):
    for name in (f"标题 {level}", f"Heading {level}"):
        try:
            p.style = doc.styles[name]
            return True
        except KeyError:
            continue
    return False


def _set_style_font_cn(style, font_name):
    style.font.name = font_name
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), font_name)


def _apply_format_to_styles(doc, format_settings):
    if not format_settings:
        return
    for level, fmt in format_settings.items():
        for name in (f"标题 {level}", f"Heading {level}"):
            try:
                st = doc.styles[name]
            except KeyError:
                continue
            try:
                if fmt.get('font_name'):
                    _set_style_font_cn(st, fmt['font_name'])
                if fmt.get('font_size'):
                    st.font.size = Pt(fmt['font_size'])
                st.font.bold = bool(fmt.get('bold', False))
                if fmt.get('color_rgb') is not None:
                    st.font.color.rgb = RGBColor(*fmt['color_rgb'])
                if fmt.get('alignment') is not None:
                    st.paragraph_format.alignment = fmt['alignment']
                if fmt.get('space_before') is not None:
                    st.paragraph_format.space_before = Pt(fmt['space_before'])
                if fmt.get('space_after') is not None:
                    st.paragraph_format.space_after = Pt(fmt['space_after'])
            except Exception:
                traceback.print_exc()
            break


# =========================================================
# 正文格式应用
# =========================================================

def _apply_run_font(run, fmt):
    if fmt.get('font_name'):
        fn = fmt['font_name']
        run.font.name = fn
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = OxmlElement('w:rFonts')
            rpr.append(rfonts)
        rfonts.set(qn('w:eastAsia'), fn)
        rfonts.set(qn('w:ascii'), fn)
        rfonts.set(qn('w:hAnsi'), fn)
    if fmt.get('font_size'):
        run.font.size = Pt(fmt['font_size'])
    run.font.bold = bool(fmt.get('bold', False))
    if fmt.get('color_rgb') is not None:
        run.font.color.rgb = RGBColor(*fmt['color_rgb'])


def _apply_body_format(doc, all_paras, title_indexes, body_fmt,
                       body_range=None):
    if not body_fmt:
        return
    start_idx, end_idx = (body_range if body_range else (None, None))
    for i, p in enumerate(all_paras):
        if i in title_indexes:
            continue
        if not p.text.strip():
            continue
        if start_idx is not None and i <= start_idx:
            continue
        if end_idx is not None and i >= end_idx:
            continue

        for r in p.runs:
            try:
                _apply_run_font(r, body_fmt)
            except Exception:
                traceback.print_exc()

        pf = p.paragraph_format
        if body_fmt.get('alignment') is not None:
            pf.alignment = body_fmt['alignment']

        ls_type = body_fmt.get('line_spacing_type')
        ls_val = body_fmt.get('line_spacing_value', 18)
        try:
            if ls_type == 'single':
                pf.line_spacing = 1.0
                pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            elif ls_type == '1.5':
                pf.line_spacing = 1.5
                pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            elif ls_type == 'double':
                pf.line_spacing = 2.0
                pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
            elif ls_type == 'exact':
                pf.line_spacing = Pt(ls_val)
                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            elif ls_type == 'atleast':
                pf.line_spacing = Pt(ls_val)
                pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        except Exception:
            traceback.print_exc()

        if body_fmt.get('space_before') is not None:
            pf.space_before = Pt(body_fmt['space_before'])
        if body_fmt.get('space_after') is not None:
            pf.space_after = Pt(body_fmt['space_after'])


# =========================================================
# 核心导出（标题处理）
# =========================================================

def export_document(src_path, out_path, titles, format_settings,
                    templates, separators, number_formats,
                    fmt_levels, num_levels,
                    apply_format=True, body_format=None,
                    body_range=None,
                    progress_cb=None):
    def report(pct, msg=None):
        if progress_cb:
            try:
                progress_cb(pct, msg)
            except Exception:
                pass

    report(3, "准备标题列表...")
    title_list = [
        t for t in titles
        if t.get('is_title') and isinstance(t.get('level'), int)
        and 1 <= t['level'] <= 9
    ]
    title_list.sort(key=lambda x: x['index'])
    if not title_list:
        raise RuntimeError("没有可导出的有效标题")

    report(10, "读取原文档...")
    doc = Document(src_path)
    all_paras = _all_paragraph_objs(doc)

    report(20, "删除旧编号...")
    _remove_style_numbering(doc, num_levels)
    for t in title_list:
        if t['level'] not in num_levels:
            continue
        idx = t['index']
        if idx >= len(all_paras):
            continue
        p = all_paras[idx]
        clean = remove_old_number(p.text)
        if clean != p.text:
            _replace_paragraph_text_preserving_format(p, clean)
        _remove_paragraph_numbering(p)

    report(40, "生成新编号...")
    nums = generate_numbers(title_list, templates, separators,
                            number_formats, num_levels)

    report(50, "写入新编号...")
    written = 0
    total_num = sum(1 for x in nums if x)
    for t, num in zip(title_list, nums):
        if not num:
            continue
        idx = t['index']
        if idx >= len(all_paras):
            continue
        p = all_paras[idx]
        new_text = num + p.text
        _replace_paragraph_text_preserving_format(p, new_text)
        written += 1
        if total_num and (written % 5 == 0 or written == total_num):
            report(50 + int(15 * written / total_num),
                   f"写入编号 {written}/{total_num}")

    if apply_format:
        report(70, "应用标题样式与格式...")
        for t in title_list:
            lvl = t['level']
            if lvl not in fmt_levels:
                continue
            idx = t['index']
            if idx >= len(all_paras):
                continue
            p = all_paras[idx]
            _apply_heading_style(p, doc, lvl)

        report(80, "应用标题样式属性...")
        _apply_format_to_styles(doc, format_settings)

    if body_format:
        report(85, "应用正文格式...")
        title_indexes = {t['index'] for t in title_list}
        _apply_body_format(doc, all_paras, title_indexes, body_format,
                           body_range=body_range)

    report(95, "保存文档...")
    doc.save(out_path)
    report(100, "完成")


# =========================================================
# 导出标题清单
# =========================================================

def _export_titles_excel(titles, nums, file_path):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment
    except ImportError:
        raise RuntimeError(
            "导出 Excel 需要安装 openpyxl：\n    pip install openpyxl\n"
            "或改用 .txt 保存。"
        )

    wb = Workbook()
    ws = wb.active
    ws.title = "标题清单"
    ws.append(["序号", "级别", "编号", "标题文字", "识别来源", "段落索引"])
    for i, (t, num) in enumerate(zip(titles, nums), 1):
        ws.append([i, t.get('level'), num.strip(), t.get('text', ''),
                   t.get('source', ''), t.get('index')])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center')
    for col, w in zip("ABCDEF", [6, 6, 18, 60, 12, 10]):
        ws.column_dimensions[col].width = w

    ws2 = wb.create_sheet("横向排列")
    ws2.append(["级别"] + [str(i + 1) for i in range(len(titles))])
    ws2.append(["编号"] + [num.strip() for num in nums])
    ws2.append(["标题文字"] + [t.get('text', '') for t in titles])
    for c in ws2["A"]:
        c.font = Font(bold=True)
    ws2.column_dimensions["A"].width = 12

    wb.save(file_path)


def _export_titles_text(titles, nums, file_path):
    lines = ["序号\t级别\t编号\t标题文字\t识别来源"]
    for i, (t, num) in enumerate(zip(titles, nums), 1):
        lines.append("\t".join([
            str(i), str(t.get('level', '')), num.strip(),
            t.get('text', ''), t.get('source', ''),
        ]))
    with open(file_path, 'w', encoding='utf-8-sig') as f:
        f.write("\n".join(lines))


def export_titles_list(titles, file_path, templates, separators,
                       number_formats, num_levels=None,
                       strip_number=False):
    """
    strip_number=True 时，导出的"标题文字"列会去掉标题原有的编号。
    """
    sorted_titles = sorted(titles, key=lambda x: x['index'])
    if strip_number:
        sorted_titles = [
            {**t, 'text': remove_old_number(t.get('text', ''))}
            for t in sorted_titles
        ]
    nums = generate_numbers(sorted_titles, templates, separators,
                            number_formats, num_levels)
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ('.xlsx', '.xls'):
        _export_titles_excel(sorted_titles, nums, file_path)
    else:
        _export_titles_text(sorted_titles, nums, file_path)


# =========================================================
# ★ 表格 / 正文提取到 Excel
# =========================================================

def _find_first_colon(line):
    """查找中/英文冒号，返回最靠前的索引，无则 -1。"""
    idx_cn = line.find('：')
    idx_en = line.find(':')
    if idx_cn >= 0 and idx_en >= 0:
        return min(idx_cn, idx_en)
    if idx_cn >= 0:
        return idx_cn
    if idx_en >= 0:
        return idx_en
    return -1


def _is_heading_paragraph(p):
    """
    判断段落是否为标题段落：
      1. Word 标题样式（Heading N / 标题 N）
      2. 段落 XML 中带大纲级别（outlineLvl）
      3. 编号模式：
         - 多级数字编号："7.8"、"7.8.1"、"7.8.1.2" 等
         - "第X章/节/篇/部分/编"
    """
    try:
        style_name = p.style.name if p.style else ''
    except Exception:
        style_name = ''
    if style_name:
        if re.match(r'^(?:Heading|标题)\s*\d+', style_name.strip(), re.I):
            return True

    try:
        pPr = p._p.pPr
        if pPr is not None:
            ol = pPr.find(W_NS + 'outlineLvl')
            if ol is not None:
                val = ol.get(W_NS + 'val')
                if val is not None:
                    try:
                        n = int(val)
                        if 0 <= n <= 8:
                            return True
                    except ValueError:
                        pass
    except Exception:
        pass

    try:
        text = p.text.strip()
    except Exception:
        text = ''
    if not text:
        return False
    if re.match(r'^\d+(?:[\.．]\d+)+(?:\s|$|[、：:\.．])', text):
        return True
    if re.match(r'^第[一二三四五六七八九十百千万零〇\d]+[章节篇部分编]', text):
        return True

    return False


def _parse_paragraph_block(paragraphs):
    """
    解析一段段落列表，返回 {列名: 内容}。
    规则：
      - 每段 "列名: 内容" 开始新字段
      - 后续段落作为内容，直到遇到下一个 "列名:" 或标题段落
      - 空段落保留为空行（内容里）
    """
    result = {}
    current_key = None
    current_value_lines = []

    def flush():
        if current_key is not None:
            result[current_key] = '\n'.join(current_value_lines).strip()

    for p in paragraphs:
        try:
            text = p.text.strip()
        except Exception:
            text = ''

        if not text:
            if current_key is not None:
                current_value_lines.append('')
            continue

        if _is_heading_paragraph(p):
            flush()
            current_key = None
            current_value_lines = []
            continue

        idx = _find_first_colon(text)
        if idx >= 0:
            key = text[:idx].strip()
            value = text[idx + 1:].strip()
            if key:
                flush()
                current_key = key
                current_value_lines = [value] if value else []
                continue

        if current_key is not None:
            current_value_lines.append(text)

    flush()
    return result


def _parse_cell_blocks(cell):
    """
    解析单个单元格，按标题段落切分成多个数据块，
    返回 list[dict]。
    """
    try:
        paragraphs = list(cell.paragraphs)
    except Exception:
        return []

    rows = []
    current_block = []

    def flush_block():
        if current_block:
            row_data = _parse_paragraph_block(current_block)
            if row_data:
                rows.append(row_data)

    for p in paragraphs:
        if _is_heading_paragraph(p):
            flush_block()
            current_block = []
            continue
        current_block.append(p)

    flush_block()
    return rows


def _unique_cells(row):
    """取一行的唯一单元格（去掉合并单元格的重复项）。"""
    seen = set()
    result = []
    try:
        cells = row.cells
    except Exception:
        return result
    for c in cells:
        try:
            tc = c._tc
        except Exception:
            tc = None
        if tc is None:
            result.append(c)
            continue
        key = id(tc)
        if key in seen:
            continue
        seen.add(key)
        result.append(c)
    return result


def _parse_one_table(table):
    """
    解析一个表格，返回 list[dict]（每 dict 对应 Excel 一行）。
    """
    rows = []
    try:
        nrows = len(table.rows)
        ncols = len(table.columns)
    except Exception:
        nrows = ncols = 0
    if nrows == 0 or ncols == 0:
        return rows

    # 1x1 单元格
    if nrows == 1 and ncols == 1:
        try:
            cell = table.cell(0, 0)
        except Exception:
            return rows
        return _parse_cell_blocks(cell)

    # 一般多行表格
    for r in table.rows:
        cells = _unique_cells(r)
        if not cells:
            continue
        texts = [c.text.strip() for c in cells]

        is_kv_row = False
        if len(cells) >= 2:
            all_short = all(len(t) < 30 for t in texts)
            all_no_colon = all(_find_first_colon(t) < 0 for t in texts)
            all_no_newline = all('\n' not in t for t in texts)
            if all_short and all_no_colon and all_no_newline:
                is_kv_row = True

        if is_kv_row:
            row_data = {}
            for i in range(0, len(texts) - 1, 2):
                k = texts[i]
                v = texts[i + 1]
                if k:
                    row_data[k] = v
            if len(texts) % 2 == 1:
                k = texts[-1]
                if k:
                    row_data[k] = ''
            if row_data:
                rows.append(row_data)
            continue

        for c in cells:
            blocks = _parse_cell_blocks(c)
            if blocks:
                rows.extend(blocks)

    return rows


def extract_from_body_paragraphs(doc):
    """从正文段落（排除表格内）中提取 "列名: 内容" 数据。"""
    from docx.text.paragraph import Paragraph

    body_paragraphs = []
    for p_el in doc.element.body.iter(qn('w:p')):
        p = Paragraph(p_el, doc)
        if _is_in_table(p):
            continue
        body_paragraphs.append(p)

    rows = []
    current_block = []

    def flush_block():
        if current_block:
            row_data = _parse_paragraph_block(current_block)
            if row_data:
                rows.append(row_data)

    for p in body_paragraphs:
        if _is_heading_paragraph(p):
            flush_block()
            current_block = []
            continue
        current_block.append(p)

    flush_block()
    return rows


def extract_tables_from_word(docx_path):
    """
    从 Word 文档提取所有表格 + 正文中的数据块，
    返回 (columns, rows)。
    """
    doc = Document(docx_path)
    rows = []

    for table in doc.tables:
        try:
            table_rows = _parse_one_table(table)
        except Exception:
            traceback.print_exc()
            table_rows = []
        rows.extend(table_rows)

    try:
        body_rows = extract_from_body_paragraphs(doc)
        rows.extend(body_rows)
    except Exception:
        traceback.print_exc()

    columns = []
    seen = set()
    for r in rows:
        for k in r.keys():
            if k and k not in seen:
                columns.append(k)
                seen.add(k)

    return columns, rows


def export_tables_to_excel(columns, rows, file_path):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment
    except ImportError:
        raise RuntimeError(
            "导出 Excel 需要安装 openpyxl：\n    pip install openpyxl"
        )

    wb = Workbook()
    ws = wb.active
    ws.title = "提取数据"
    ws.append(columns)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center', vertical='center')

    for r in rows:
        ws.append([r.get(c, '') for c in columns])

    for i, col in enumerate(columns, 1):
        max_len = len(str(col))
        for r in rows:
            v = str(r.get(col, ''))
            for line in v.split('\n'):
                est = sum(2 if ord(ch) > 127 else 1 for ch in line)
                max_len = max(max_len, est)
        letter = ws.cell(row=1, column=i).column_letter
        ws.column_dimensions[letter].width = min(60, max(8, max_len + 2))

    wb.save(file_path)


# =========================================================
# GUI 部件（标题处理）
# =========================================================

ALIGN_MAP = {
    '左对齐': WD_ALIGN_PARAGRAPH.LEFT, '居中': WD_ALIGN_PARAGRAPH.CENTER,
    '右对齐': WD_ALIGN_PARAGRAPH.RIGHT, '两端对齐': WD_ALIGN_PARAGRAPH.JUSTIFY,
}
ALIGN_NAMES = list(ALIGN_MAP.keys())


def show_template_help(parent):
    win = tk.Toplevel(parent)
    win.title("编号模板编写原则")
    win.geometry("680x700")
    txt = tk.Text(win, wrap='word', font=('Consolas', 10))
    txt.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    txt.insert('1.0', TEMPLATE_HELP_TEXT)
    txt.config(state='disabled')
    ttk.Button(win, text="知道了", command=win.destroy).pack(pady=(0, 10))
    win.transient(parent)
    win.grab_set()


class ProgressDialog:
    def __init__(self, parent, title="处理中..."):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.transient(parent)
        self.top.resizable(False, False)

        w, h = 420, 120
        self.top.update_idletasks()
        try:
            px = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
            py = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        except Exception:
            px = py = 200
        self.top.geometry(f"{w}x{h}+{max(0, px)}+{max(0, py)}")

        self.top.protocol("WM_DELETE_WINDOW", lambda: None)
        try:
            self.top.grab_set()
        except Exception:
            pass

        self.msg_var = tk.StringVar(value="准备中...")
        ttk.Label(self.top, textvariable=self.msg_var).pack(
            padx=16, pady=(16, 6), anchor='w')
        self.pb = ttk.Progressbar(self.top, mode='determinate',
                                  maximum=100, length=380)
        self.pb.pack(padx=16, pady=(0, 12))
        self.pb['value'] = 0
        self._closed = False

    def update(self, pct, msg=None):
        if self._closed:
            return
        try:
            self.pb['value'] = max(0, min(100, int(pct)))
            if msg:
                self.msg_var.set(msg)
            self.top.update_idletasks()
        except Exception:
            pass

    def close(self):
        if self._closed:
            return
        self._closed = True
        try:
            self.top.grab_release()
        except Exception:
            pass
        try:
            self.top.destroy()
        except Exception:
            pass


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient='vertical',
                                 command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner_id = self.canvas.create_window(
            (0, 0), window=self.inner, anchor='nw')
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.inner.bind('<Configure>', self._on_inner_config)
        self.canvas.bind('<Configure>', self._on_canvas_config)
        self.canvas.bind('<Enter>', self._bind_wheel)
        self.canvas.bind('<Leave>', self._unbind_wheel)
        self.inner.bind('<Enter>', self._bind_wheel)
        self.inner.bind('<Leave>', self._unbind_wheel)
        self._wheel_bound = False

    def _bind_wheel(self, event=None):
        if not self._wheel_bound:
            self.canvas.bind_all('<MouseWheel>', self._on_wheel)
            self._wheel_bound = True

    def _unbind_wheel(self, event=None):
        try:
            x, y = self.canvas.winfo_pointerxy()
            widget = self.canvas.winfo_containing(x, y)
            if widget is not None and (
                    widget is self.canvas or widget is self.inner
                    or str(widget).startswith(str(self.inner))):
                return
        except Exception:
            pass
        try:
            self.canvas.unbind_all('<MouseWheel>')
        except Exception:
            pass
        self._wheel_bound = False

    def _on_inner_config(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def _on_canvas_config(self, event):
        self.canvas.itemconfig(self.inner_id, width=event.width)

    def _on_wheel(self, event):
        if not event.delta:
            return
        if abs(event.delta) >= 120:
            step = -1 * (event.delta // 120)
        else:
            step = -1 if event.delta > 0 else 1
        self.canvas.yview_scroll(step, 'units')


class LevelFormatPanel:
    def __init__(self, parent, level, defaults):
        self.level = level
        self.frame = ttk.LabelFrame(parent, text=f"{level} 级标题格式")
        self.frame.pack(fill=tk.X, padx=4, pady=2)

        header = ttk.Frame(self.frame)
        header.pack(fill=tk.X, padx=4, pady=(2, 0))
        self.enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(header, text="启用该级格式修改",
                        variable=self.enabled_var,
                        command=self._on_toggle).pack(side=tk.LEFT)

        self.body = ttk.Frame(self.frame)
        self.body.pack(fill=tk.X, padx=4, pady=2)

        row1 = ttk.Frame(self.body)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="字体").pack(side=tk.LEFT)
        self.font_var = tk.StringVar(
            value=defaults.get('font_name', '宋体'))
        ttk.Combobox(row1, textvariable=self.font_var,
                     values=FONT_PRESETS, width=14).pack(
            side=tk.LEFT, padx=(2, 6))

        ttk.Label(row1, text="字号").pack(side=tk.LEFT)
        default_size = defaults.get('font_size_name', '小四')
        self.size_var = tk.StringVar(value=default_size)
        ttk.Combobox(row1, textvariable=self.size_var,
                     values=FONT_SIZE_PRESETS, width=8).pack(
            side=tk.LEFT, padx=(2, 6))

        self.bold_var = tk.BooleanVar(value=defaults.get('bold', True))
        ttk.Checkbutton(row1, text="加粗",
                        variable=self.bold_var).pack(
            side=tk.LEFT, padx=(0, 6))

        ttk.Label(row1, text="对齐").pack(side=tk.LEFT)
        self.align_var = tk.StringVar(value=defaults.get('align', '左对齐'))
        ttk.Combobox(row1, textvariable=self.align_var, values=ALIGN_NAMES,
                     width=7, state='readonly').pack(
            side=tk.LEFT, padx=(2, 6))

        ttk.Label(row1, text="颜色").pack(side=tk.LEFT)
        self.color_rgb = defaults.get('color_rgb', (0, 0, 0))
        self.color_btn = tk.Button(row1, text="  ",
                                   bg=self._rgb_to_hex(self.color_rgb),
                                   width=2, command=self.pick_color)
        self.color_btn.pack(side=tk.LEFT, padx=(2, 6))

        row2 = ttk.Frame(self.body)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="段前(磅)").pack(side=tk.LEFT)
        self.before_var = tk.IntVar(value=defaults.get('space_before', 6))
        ttk.Spinbox(row2, from_=0, to=100, textvariable=self.before_var,
                    width=4).pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(row2, text="段后(磅)").pack(side=tk.LEFT)
        self.after_var = tk.IntVar(value=defaults.get('space_after', 6))
        ttk.Spinbox(row2, from_=0, to=100, textvariable=self.after_var,
                    width=4).pack(side=tk.LEFT, padx=(2, 6))

    def _on_toggle(self):
        self.set_enabled(self.enabled_var.get())

    def set_enabled(self, enabled):
        state = 'normal' if enabled else 'disabled'
        for w in self.body.winfo_children():
            try:
                w.configure(state=state)
            except Exception:
                pass
            for sub in w.winfo_children():
                try:
                    if isinstance(sub, ttk.Combobox):
                        sub.configure(state='normal' if enabled
                                      else 'disabled')
                    else:
                        sub.configure(state=state)
                except Exception:
                    pass

    def is_enabled(self):
        return bool(self.enabled_var.get())

    @staticmethod
    def _rgb_to_hex(rgb):
        return '#%02x%02x%02x' % rgb

    def pick_color(self):
        rgb, hx = colorchooser.askcolor(color=self._rgb_to_hex(self.color_rgb))
        if rgb:
            self.color_rgb = tuple(int(x) for x in rgb)
            self.color_btn.config(bg=hx)

    def get_settings(self):
        size_val = parse_font_size(self.size_var.get())
        if size_val is None:
            size_val = 12.0
        return {
            'font_name': self.font_var.get().strip(),
            'font_size': size_val,
            'bold': bool(self.bold_var.get()),
            'color_rgb': self.color_rgb,
            'alignment': ALIGN_MAP.get(self.align_var.get(),
                                       WD_ALIGN_PARAGRAPH.LEFT),
            'space_before': int(self.before_var.get()),
            'space_after': int(self.after_var.get()),
        }


class BodyFormatPanel:
    def __init__(self, parent):
        self.frame = ttk.LabelFrame(
            parent, text="正文格式（只应用到指定的正文范围内）")
        self.frame.pack(fill=tk.X, padx=4, pady=4)

        header = ttk.Frame(self.frame)
        header.pack(fill=tk.X, padx=4, pady=(2, 0))
        self.enabled_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(header, text="启用正文格式修改",
                        variable=self.enabled_var,
                        command=self._on_toggle).pack(side=tk.LEFT)
        ttk.Label(header,
                  text="（需先右键标题行设置“正文开始”和“正文结束”；"
                       "若“正文结束”是最后一个标题，之后到文档末尾都算正文）",
                  foreground='#666').pack(side=tk.LEFT, padx=8)

        self.body = ttk.Frame(self.frame)
        self.body.pack(fill=tk.X, padx=4, pady=2)

        row1 = ttk.Frame(self.body)
        row1.pack(fill=tk.X, pady=2)

        ttk.Label(row1, text="字体").pack(side=tk.LEFT)
        self.font_var = tk.StringVar(value="宋体")
        ttk.Combobox(row1, textvariable=self.font_var,
                     values=FONT_PRESETS, width=14).pack(
            side=tk.LEFT, padx=(2, 6))

        ttk.Label(row1, text="字号").pack(side=tk.LEFT)
        self.size_var = tk.StringVar(value="小四")
        ttk.Combobox(row1, textvariable=self.size_var,
                     values=FONT_SIZE_PRESETS, width=8).pack(
            side=tk.LEFT, padx=(2, 6))

        self.bold_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row1, text="加粗",
                        variable=self.bold_var).pack(
            side=tk.LEFT, padx=(0, 6))

        ttk.Label(row1, text="对齐").pack(side=tk.LEFT)
        self.align_var = tk.StringVar(value="两端对齐")
        ttk.Combobox(row1, textvariable=self.align_var, values=ALIGN_NAMES,
                     width=7, state='readonly').pack(
            side=tk.LEFT, padx=(2, 6))

        ttk.Label(row1, text="颜色").pack(side=tk.LEFT)
        self.color_rgb = (0, 0, 0)
        self.color_btn = tk.Button(row1, text="  ",
                                   bg=self._rgb_to_hex(self.color_rgb),
                                   width=2, command=self.pick_color)
        self.color_btn.pack(side=tk.LEFT, padx=(2, 6))

        row2 = ttk.Frame(self.body)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="行间距").pack(side=tk.LEFT)
        self.ls_var = tk.StringVar(value="单倍行距")
        self.ls_cb = ttk.Combobox(row2, textvariable=self.ls_var,
                                  values=LINE_SPACING_PRESETS,
                                  width=12, state='readonly')
        self.ls_cb.pack(side=tk.LEFT, padx=(2, 4))
        self.ls_cb.bind('<<ComboboxSelected>>', self._on_ls_change)

        ttk.Label(row2, text="数值(磅)").pack(side=tk.LEFT)
        self.ls_val_var = tk.IntVar(value=18)
        self.ls_val_sb = ttk.Spinbox(row2, from_=6, to=200,
                                     textvariable=self.ls_val_var,
                                     width=5, state='disabled')
        self.ls_val_sb.pack(side=tk.LEFT, padx=(2, 12))

        ttk.Label(row2, text="段前(磅)").pack(side=tk.LEFT)
        self.before_var = tk.IntVar(value=0)
        ttk.Spinbox(row2, from_=0, to=100, textvariable=self.before_var,
                    width=4).pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(row2, text="段后(磅)").pack(side=tk.LEFT)
        self.after_var = tk.IntVar(value=0)
        ttk.Spinbox(row2, from_=0, to=100, textvariable=self.after_var,
                    width=4).pack(side=tk.LEFT, padx=(2, 6))

        self._on_toggle()

    def _on_ls_change(self, event=None):
        val = self.ls_var.get()
        if val in ("固定值(磅)", "最小值(磅)"):
            try:
                self.ls_val_sb.configure(state='normal')
            except Exception:
                pass
        else:
            try:
                self.ls_val_sb.configure(state='disabled')
            except Exception:
                pass

    def _on_toggle(self):
        self.set_enabled(self.enabled_var.get())
        if self.enabled_var.get():
            self._on_ls_change()

    def set_enabled(self, enabled):
        state = 'normal' if enabled else 'disabled'
        for w in self.body.winfo_children():
            try:
                w.configure(state=state)
            except Exception:
                pass
            for sub in w.winfo_children():
                try:
                    if isinstance(sub, ttk.Combobox):
                        sub.configure(state='normal' if enabled
                                      else 'disabled')
                    else:
                        sub.configure(state=state)
                except Exception:
                    pass
        if enabled:
            self._on_ls_change()

    def is_enabled(self):
        return bool(self.enabled_var.get())

    @staticmethod
    def _rgb_to_hex(rgb):
        return '#%02x%02x%02x' % rgb

    def pick_color(self):
        rgb, hx = colorchooser.askcolor(color=self._rgb_to_hex(self.color_rgb))
        if rgb:
            self.color_rgb = tuple(int(x) for x in rgb)
            self.color_btn.config(bg=hx)

    def get_settings(self):
        ls_name = self.ls_var.get()
        ls_type = LINE_SPACING_MAP.get(ls_name, 'single')
        try:
            ls_val = int(self.ls_val_var.get())
        except (TypeError, ValueError):
            ls_val = 18
        size_val = parse_font_size(self.size_var.get())
        if size_val is None:
            size_val = 12.0
        return {
            'font_name': self.font_var.get().strip(),
            'font_size': size_val,
            'bold': bool(self.bold_var.get()),
            'color_rgb': self.color_rgb,
            'alignment': ALIGN_MAP.get(self.align_var.get(),
                                       WD_ALIGN_PARAGRAPH.JUSTIFY),
            'line_spacing_type': ls_type,
            'line_spacing_value': ls_val,
            'space_before': int(self.before_var.get()),
            'space_after': int(self.after_var.get()),
        }


class AddParagraphDialog:
    def __init__(self, parent, all_paras, existing_indexes):
        self.top = tk.Toplevel(parent)
        self.top.title("从所有段落中添加标题")
        self.top.geometry("820x480")
        self.result = []
        self.all_paras = all_paras
        self.existing = set(existing_indexes)

        ttk.Label(self.top,
                  text="勾选要添加为标题的段落（可多选）").pack(
            anchor='w', padx=8, pady=6)

        cols = ("选择", "序号", "样式", "表格", "内容")
        self.tree = ttk.Treeview(self.top, columns=cols, show='headings',
                                 height=16)
        for c in cols:
            self.tree.heading(c, text=c)
        self.tree.column("选择", width=50, anchor='center')
        self.tree.column("序号", width=60, anchor='center')
        self.tree.column("样式", width=110, anchor='w')
        self.tree.column("表格", width=60, anchor='center')
        self.tree.column("内容", width=520, anchor='w')
        self.tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        self.tree.bind('<Button-1>', self.on_click)
        self.tree.bind('<MouseWheel>', self._on_tree_wheel)

        self.check_state = {}
        for p in all_paras:
            iid = str(p['index'])
            mark = '✓' if p['index'] in self.existing else ''
            in_tbl = '是' if p.get('in_table') else ''
            self.tree.insert('', tk.END, iid=iid, values=(
                mark, p['index'] + 1, p['style'], in_tbl, p['text'][:80]
            ))
            self.check_state[iid] = (p['index'] in self.existing)

        btns = ttk.Frame(self.top)
        btns.pack(fill=tk.X, padx=8, pady=6)
        ttk.Button(btns, text="取消", command=self.top.destroy).pack(
            side=tk.RIGHT, padx=4)
        ttk.Button(btns, text="确定", command=self.confirm).pack(
            side=tk.RIGHT, padx=4)

        self.top.transient(parent)
        self.top.grab_set()
        parent.wait_window(self.top)

    def _on_tree_wheel(self, event):
        if not event.delta:
            return 'break'
        if abs(event.delta) >= 120:
            step = -1 * (event.delta // 120)
        else:
            step = -1 if event.delta > 0 else 1
        self.tree.yview_scroll(step, 'units')
        return 'break'

    def on_click(self, event):
        if self.tree.identify_region(event.x, event.y) != 'cell':
            return
        if self.tree.identify_column(event.x) != '#1':
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        self.check_state[iid] = not self.check_state.get(iid, False)
        vals = list(self.tree.item(iid, 'values'))
        vals[0] = '✓' if self.check_state[iid] else ''
        self.tree.item(iid, values=vals)

    def confirm(self):
        self.result = [int(iid) for iid, chk
                       in self.check_state.items() if chk]
        self.top.destroy()


# =========================================================
# Tab 1：标题处理
# =========================================================

class App(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.root = self.winfo_toplevel()

        self.src_path = None
        self.all_paras = []
        self.items = []

        self.format_panels = {}
        self.template_vars = {}
        self.sep_vars = {}
        self.numfmt_vars = {}
        self.num_enabled_vars = {}

        self.include_table_var = tk.BooleanVar(value=True)
        # ★ 导出清单是否去除编号
        self.strip_number_var = tk.BooleanVar(value=True)

        self.body_start_idx = None
        self.body_end_idx = None

        self.body_panel = None

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=(6, 2))

        ttk.Button(top, text="选择 Word 文档",
                   command=self.select_file).pack(side=tk.LEFT)
        self.file_label = ttk.Label(
            top, text="未选择文件", foreground='#555',
            width=MAX_FILE_DISPLAY_CHARS + 2, anchor='w')
        self.file_label.pack(side=tk.LEFT, padx=6)

        ttk.Button(top, text="识别标题",
                   command=self.detect_titles).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="从所有段落添加",
                   command=self.open_add_dialog).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="移除选中",
                   command=self.remove_selected).pack(side=tk.LEFT, padx=4)

        bottom = ttk.Frame(self)
        bottom.pack(fill=tk.X, padx=8, pady=(2, 8), side=tk.BOTTOM)

        ttk.Button(bottom, text="导出为新 Word 文件",
                   command=self.export).pack(side=tk.RIGHT, padx=4)
        ttk.Button(bottom, text="导出标题清单（Excel/文本）",
                   command=self.export_titles).pack(side=tk.RIGHT, padx=4)

        scroll = ScrollableFrame(self)
        scroll.pack(fill=tk.BOTH, expand=True, padx=6, pady=2)
        body = scroll.inner

        mid = ttk.LabelFrame(
            body,
            text="标题列表（点击第一列勾选/取消；双击级别或文字列修改；右键设置正文范围）")
        mid.pack(fill=tk.X, padx=2, pady=4)

        quick = ttk.Frame(mid)
        quick.pack(fill=tk.X, padx=4, pady=(4, 2))

        ttk.Label(quick, text="批量操作：").pack(side=tk.LEFT)
        ttk.Button(quick, text="全选", width=6,
                   command=lambda: self._set_checked(
                       self._get_all_iids(), True)).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick, text="全不选", width=6,
                   command=lambda: self._set_checked(
                       self._get_all_iids(), False)).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick, text="反选", width=6,
                   command=self._invert_checked).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick, text="切换选中 (空格)", width=14,
                   command=self._toggle_selected).pack(side=tk.LEFT, padx=2)

        # ★ 新增：导出清单去除编号
        ttk.Checkbutton(quick, text="导出清单时去除编号",
                        variable=self.strip_number_var).pack(
            side=tk.LEFT, padx=10)

        self.body_range_var = tk.StringVar(
            value="正文范围：未设置（右键标题行可设置）")
        ttk.Label(quick, textvariable=self.body_range_var,
                  foreground='#0066cc').pack(side=tk.LEFT, padx=10)

        cols = ("包含", "序号", "级别", "标题文字", "来源",
                "段索", "正文起", "正文止")
        self.tree = ttk.Treeview(mid, columns=cols, show='headings',
                                 height=8, selectmode='extended')
        for c in cols:
            self.tree.heading(c, text=c)
        self.tree.column("包含", width=48, anchor='center')
        self.tree.column("序号", width=48, anchor='center')
        self.tree.column("级别", width=56, anchor='center')
        self.tree.column("标题文字", width=440, anchor='w')
        self.tree.column("来源", width=110, anchor='center')
        self.tree.column("段索", width=60, anchor='center')
        self.tree.column("正文起", width=56, anchor='center')
        self.tree.column("正文止", width=56, anchor='center')
        self.tree.pack(fill=tk.X, side=tk.LEFT, padx=4, pady=4)
        sb = ttk.Scrollbar(mid, orient='vertical', command=self.tree.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=sb.set)

        self.tree.bind('<Button-1>', self.on_tree_click)
        self.tree.bind('<Double-1>', self.on_tree_double)
        self.tree.bind('<MouseWheel>', self.on_tree_wheel)
        self.tree.bind('<space>', self._on_space_key)
        self.tree.bind('<Button-3>', self._on_right_click)

        opt_wrap = ttk.Frame(body)
        opt_wrap.pack(fill=tk.X, padx=2, pady=4)

        mode_wrap = ttk.LabelFrame(opt_wrap, text="处理模式")
        mode_wrap.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.mode_var = tk.StringVar(value='reformat')
        ttk.Radiobutton(mode_wrap, text="重新编号 + 修改标题格式",
                        variable=self.mode_var, value='reformat',
                        command=self.on_mode_change).pack(
            side=tk.LEFT, padx=10, pady=2)
        ttk.Radiobutton(mode_wrap, text="只重新编号（保持原标题格式）",
                        variable=self.mode_var, value='renumber_only',
                        command=self.on_mode_change).pack(
            side=tk.LEFT, padx=10, pady=2)

        recog_wrap = ttk.LabelFrame(opt_wrap, text="识别选项")
        recog_wrap.pack(side=tk.LEFT, fill=tk.X, padx=(4, 0))

        ttk.Checkbutton(
            recog_wrap,
            text="识别表格里的标题",
            variable=self.include_table_var).pack(
            side=tk.LEFT, padx=10, pady=2)

        paned = ttk.Frame(body)
        paned.pack(fill=tk.X, padx=2, pady=2)

        left = ttk.Frame(paned)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        right = ttk.Frame(paned)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0))

        fmt_wrap = ttk.LabelFrame(
            left, text="按级别统一设置标题格式（每级可单独启用/禁用）")
        fmt_wrap.pack(fill=tk.X)
        default_fonts = {
            1: ('黑体', '三号', True),
            2: ('楷体', '四号', True),
            3: ('宋体', '小四', False),
        }
        for lvl in (1, 2, 3):
            fn, fs_name, bd = default_fonts[lvl]
            self.format_panels[lvl] = LevelFormatPanel(fmt_wrap, lvl, {
                'font_name': fn,
                'font_size_name': fs_name,
                'bold': bd,
                'align': '左对齐', 'color_rgb': (0, 0, 0),
                'space_before': 6 if lvl == 1 else 3,
                'space_after': 6 if lvl == 1 else 3,
            })

        num_wrap = ttk.LabelFrame(
            right, text="编号模板（每级可单独启用/禁用）")
        num_wrap.pack(fill=tk.X)

        header = ttk.Frame(num_wrap)
        header.pack(fill=tk.X, padx=4, pady=(4, 2))
        ttk.Label(header, text="启用", width=4).pack(side=tk.LEFT)
        ttk.Label(header, text="级别", width=5).pack(side=tk.LEFT)
        ttk.Label(header, text="编号模板").pack(side=tk.LEFT, padx=(2, 0))
        ttk.Label(header, text="分隔符").pack(side=tk.LEFT, padx=(28, 0))
        ttk.Label(header, text="数字格式").pack(side=tk.LEFT, padx=(14, 0))
        ttk.Button(header, text="📖 编写原则",
                   command=lambda: show_template_help(self.root)).pack(
            side=tk.RIGHT, padx=4)

        for lvl in (1, 2, 3):
            row = ttk.Frame(num_wrap)
            row.pack(fill=tk.X, padx=4, pady=3)

            ev = tk.BooleanVar(value=True)
            ttk.Checkbutton(row, variable=ev,
                            width=2).pack(side=tk.LEFT, padx=(2, 8))
            self.num_enabled_vars[lvl] = ev

            ttk.Label(row, text=f"{lvl} 级", width=5).pack(side=tk.LEFT)

            tv = tk.StringVar(value=TEMPLATE_DEFAULT[lvl])
            ttk.Combobox(row, textvariable=tv,
                         values=TEMPLATE_PRESETS.get(lvl, []),
                         width=18).pack(side=tk.LEFT, padx=(2, 8))
            self.template_vars[lvl] = tv

            ttk.Label(row, text="分隔符").pack(side=tk.LEFT)
            sv = tk.StringVar(value="空格")
            ttk.Combobox(row, textvariable=sv, values=SEP_PRESETS,
                         width=8).pack(side=tk.LEFT, padx=2)
            self.sep_vars[lvl] = sv

            ttk.Label(row, text="数字格式").pack(side=tk.LEFT, padx=(6, 0))
            nv = tk.StringVar(value="阿拉伯数字")
            ttk.Combobox(row, textvariable=nv, values=NUMFMT_PRESETS,
                         width=9, state='readonly').pack(side=tk.LEFT, padx=2)
            self.numfmt_vars[lvl] = nv

        ttk.Label(num_wrap,
                  text="提示：编号为纯文本，直接写入标题前。",
                  foreground='#666').pack(anchor='w', padx=8, pady=(2, 4))

        self.body_panel = BodyFormatPanel(body)

    # ---------- 正文范围 ----------
    def _update_body_range_label(self):
        start_text = '未设置'
        end_text = '未设置'
        if self.body_start_idx is not None:
            for it in self.items:
                if it.get('index') == self.body_start_idx:
                    start_text = f"【{it.get('text', '')[:18]}】"
                    break
        if self.body_end_idx is not None:
            for it in self.items:
                if it.get('index') == self.body_end_idx:
                    end_text = f"【{it.get('text', '')[:18]}】"
                    break

        end_extra = ""
        if self.body_end_idx is not None and self.items:
            title_idxs = sorted([it['index'] for it in self.items
                                 if it.get('is_title')])
            if title_idxs and self.body_end_idx == title_idxs[-1]:
                end_extra = "（延伸到文档末尾）"

        self.body_range_var.set(
            f"正文范围：{start_text} → {end_text}{end_extra}")

    def _set_body_start(self, iid):
        vals = self.tree.item(iid, 'values')
        try:
            idx = int(vals[5])
        except (TypeError, ValueError, IndexError):
            return
        self.body_start_idx = idx
        self._update_body_range_label()
        self.refresh_tree()

    def _set_body_end(self, iid):
        vals = self.tree.item(iid, 'values')
        try:
            idx = int(vals[5])
        except (TypeError, ValueError, IndexError):
            return
        self.body_end_idx = idx
        self._update_body_range_label()
        self.refresh_tree()

    def _clear_body_range(self):
        self.body_start_idx = None
        self.body_end_idx = None
        self._update_body_range_label()
        self.refresh_tree()

    def _clear_body_start(self):
        self.body_start_idx = None
        self._update_body_range_label()
        self.refresh_tree()

    def _clear_body_end(self):
        self.body_end_idx = None
        self._update_body_range_label()
        self.refresh_tree()

    def _run_with_progress(self, work, on_done=None, title="处理中..."):
        dlg = ProgressDialog(self.root, title=title)
        q = queue.Queue()

        def progress_cb(pct, msg=None):
            q.put(('progress', pct, msg))

        def worker():
            try:
                result = work(progress_cb)
                q.put(('done', result, None))
            except Exception as e:
                q.put(('error', e, None))

        threading.Thread(target=worker, daemon=True).start()

        def poll():
            try:
                while True:
                    kind, a, b = q.get_nowait()
                    if kind == 'progress':
                        dlg.update(a, b)
                    elif kind == 'done':
                        dlg.close()
                        if on_done:
                            try:
                                on_done(a)
                            except Exception:
                                traceback.print_exc()
                        return
                    elif kind == 'error':
                        dlg.close()
                        messagebox.showerror("错误", f"处理失败：{a}")
                        return
            except queue.Empty:
                pass
            self.root.after(80, poll)

        poll()

    def _get_all_iids(self):
        return list(self.tree.get_children())

    def _get_selected_iids(self):
        return list(self.tree.selection())

    def _set_checked(self, iids, checked):
        for iid in iids:
            vals = list(self.tree.item(iid, 'values'))
            if not vals:
                continue
            vals[0] = '✓' if checked else ''
            self.tree.item(iid, values=vals)

    def _invert_checked(self):
        for iid in self._get_all_iids():
            vals = list(self.tree.item(iid, 'values'))
            if not vals:
                continue
            vals[0] = '' if vals[0] == '✓' else '✓'
            self.tree.item(iid, values=vals)

    def _toggle_selected(self):
        iids = self._get_selected_iids()
        if not iids:
            return
        all_checked = True
        for iid in iids:
            vals = self.tree.item(iid, 'values')
            if not vals or vals[0] != '✓':
                all_checked = False
                break
        self._set_checked(iids, not all_checked)

    def _on_space_key(self, event):
        self._toggle_selected()
        return 'break'

    def _on_right_click(self, event):
        iid = self.tree.identify_row(event.y)
        if iid and iid not in self.tree.selection():
            self.tree.selection_set(iid)

        is_start = False
        is_end = False
        if iid:
            v = self.tree.item(iid, 'values')
            try:
                is_start = (v[6] == '起')
                is_end = (v[7] == '止')
            except Exception:
                pass

        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="切换选中 (Space)",
                         command=self._toggle_selected)
        menu.add_separator()

        if is_start:
            menu.add_command(label="取消「正文开始」标记",
                             command=self._clear_body_start)
        else:
            menu.add_command(label="将选中行设为「正文开始」",
                             command=lambda: self._set_body_start(iid))

        if is_end:
            menu.add_command(label="取消「正文结束」标记",
                             command=self._clear_body_end)
        else:
            menu.add_command(label="将选中行设为「正文结束」",
                             command=lambda: self._set_body_end(iid))

        menu.add_separator()
        menu.add_command(label="清除正文范围",
                         command=self._clear_body_range)
        menu.add_separator()
        menu.add_command(label="全部勾选",
                         command=lambda: self._set_checked(
                             self._get_all_iids(), True))
        menu.add_command(label="全部取消",
                         command=lambda: self._set_checked(
                             self._get_all_iids(), False))
        menu.add_command(label="反选", command=self._invert_checked)
        menu.add_separator()
        menu.add_command(label="将选中行设为 1 级",
                         command=lambda: self._set_level_for_selected(1))
        menu.add_command(label="将选中行设为 2 级",
                         command=lambda: self._set_level_for_selected(2))
        menu.add_command(label="将选中行设为 3 级",
                         command=lambda: self._set_level_for_selected(3))
        menu.add_separator()
        menu.add_command(label="从标题列表移除选中行",
                         command=self.remove_selected)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _set_level_for_selected(self, level):
        for iid in self._get_selected_iids():
            vals = list(self.tree.item(iid, 'values'))
            if not vals:
                continue
            vals[2] = level
            self.tree.item(iid, values=vals)

    def on_tree_wheel(self, event):
        if not event.delta:
            return 'break'
        if abs(event.delta) >= 120:
            step = -1 * (event.delta // 120)
        else:
            step = -1 if event.delta > 0 else 1
        self.tree.yview_scroll(step, 'units')
        return 'break'

    def _bind_tooltip(self, widget, text):
        for seq in ('<Enter>', '<Leave>'):
            try:
                widget.unbind(seq)
            except Exception:
                pass
        tip = {'win': None}

        def show(_e=None):
            if tip['win'] is not None:
                return
            x = widget.winfo_rootx() + 10
            y = widget.winfo_rooty() + widget.winfo_height() + 4
            tw = tk.Toplevel(widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            tk.Label(tw, text=text, justify='left',
                     background="#ffffe1", relief='solid', borderwidth=1,
                     font=('Microsoft YaHei', 9), wraplength=520
                     ).pack(ipadx=4, ipady=2)
            tip['win'] = tw

        def hide(_e=None):
            if tip['win'] is not None:
                try:
                    tip['win'].destroy()
                except Exception:
                    pass
                tip['win'] = None

        widget.bind('<Enter>', show)
        widget.bind('<Leave>', hide)

    def on_mode_change(self):
        enabled = (self.mode_var.get() == 'reformat')
        for panel in self.format_panels.values():
            if not enabled:
                panel.set_enabled(False)
            else:
                panel.set_enabled(panel.enabled_var.get())

    def select_file(self):
        path = filedialog.askopenfilename(
            title="选择 Word 文档",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if path:
            self.src_path = path
            self.file_label.config(text=truncate_path(path))
            self._bind_tooltip(self.file_label, path)
            self.all_paras = []
            self.items = []
            self.body_start_idx = None
            self.body_end_idx = None
            self._update_body_range_label()
            self.refresh_tree()

    def detect_titles(self):
        if not self.src_path:
            messagebox.showwarning("提示", "请先选择 Word 文档")
            return
        try:
            self.all_paras = load_paragraphs(self.src_path)
        except Exception as e:
            messagebox.showerror("错误", f"读取文档失败：{e}")
            return

        include_table = bool(self.include_table_var.get())
        self.items = []
        for p in self.all_paras:
            level, source = detect_heading(p, include_table=include_table)
            item = dict(p)
            if level:
                item['is_title'] = True
                item['level'] = level
                src = source or ''
                if p.get('in_table'):
                    src = (src + ' [表格]').strip()
                item['source'] = src
            self.items.append(item)

        self.body_start_idx = None
        self.body_end_idx = None
        self._update_body_range_label()
        self.refresh_tree()

        n = sum(1 for x in self.items if x['is_title'])
        tip = "" if include_table else "（未包含表格内段落）"
        messagebox.showinfo("识别完成",
                            f"共识别到 {n} 个标题{tip}（目录段落已排除）\n"
                            "右键标题行可标记「正文开始」/「正文结束」。")

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        shown = [it for it in self.items if it['is_title']]
        shown.sort(key=lambda x: x['index'])
        for i, it in enumerate(shown):
            is_start = (it['index'] == self.body_start_idx)
            is_end = (it['index'] == self.body_end_idx)
            self.tree.insert('', tk.END, iid=str(i), values=(
                '✓', i + 1, it['level'], it['text'],
                it['source'], it['index'],
                '起' if is_start else '',
                '止' if is_end else ''))

    def on_tree_click(self, event):
        if self.tree.identify_region(event.x, event.y) != 'cell':
            return
        if self.tree.identify_column(event.x) != '#1':
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        vals = list(self.tree.item(iid, 'values'))
        vals[0] = '' if vals[0] == '✓' else '✓'
        self.tree.item(iid, values=vals)

    def on_tree_double(self, event):
        if self.tree.identify_region(event.x, event.y) != 'cell':
            return
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if col == '#3':
            self.edit_level(iid)
        elif col == '#4':
            self.edit_text(iid)

    def edit_level(self, iid):
        vals = list(self.tree.item(iid, 'values'))
        win = tk.Toplevel(self.root)
        win.title("修改级别")
        ttk.Label(win, text="请输入 1~9 的级别：").pack(padx=10, pady=6)
        v = tk.StringVar(value=str(vals[2]))
        ent = ttk.Entry(win, textvariable=v, width=8)
        ent.pack(padx=10, pady=4)
        ent.focus_set()

        def ok():
            try:
                lv = int(v.get())
                if not 1 <= lv <= 9:
                    raise ValueError
            except ValueError:
                messagebox.showerror("错误", "级别必须是 1~9 的整数",
                                     parent=win)
                return
            vals[2] = lv
            self.tree.item(iid, values=vals)
            win.destroy()

        ttk.Button(win, text="确定", command=ok).pack(pady=6)
        win.transient(self.root)
        win.grab_set()

    def edit_text(self, iid):
        vals = list(self.tree.item(iid, 'values'))
        win = tk.Toplevel(self.root)
        win.title("修改标题文字")
        ttk.Label(win, text="编辑标题文字（可手动去掉旧编号）：").pack(
            padx=10, pady=6)
        txt = tk.Text(win, width=60, height=4)
        txt.pack(padx=10, pady=4)
        txt.insert('1.0', str(vals[3]))
        txt.focus_set()

        def ok():
            new_text = txt.get('1.0', 'end').strip().replace('\n', ' ')
            vals[3] = new_text
            self.tree.item(iid, values=vals)
            win.destroy()

        ttk.Button(win, text="确定", command=ok).pack(pady=6)
        win.transient(self.root)
        win.grab_set()

    def remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在列表中选择要移除的项")
            return
        remove_idx = set()
        for iid in sel:
            vals = self.tree.item(iid, 'values')
            try:
                remove_idx.add(int(vals[5]))
            except (TypeError, ValueError, IndexError):
                continue
        for it in self.items:
            if it['index'] in remove_idx:
                it['is_title'] = False
                it['level'] = None
                it['source'] = ''
        if self.body_start_idx in remove_idx:
            self.body_start_idx = None
        if self.body_end_idx in remove_idx:
            self.body_end_idx = None
        self._update_body_range_label()
        self.refresh_tree()

    def open_add_dialog(self):
        if not self.all_paras:
            if not self.src_path:
                messagebox.showwarning("提示", "请先选择 Word 文档")
                return
            try:
                self.all_paras = load_paragraphs(self.src_path)
            except Exception as e:
                messagebox.showerror("错误", f"读取文档失败：{e}")
                return
        if not self.items:
            self.items = [dict(p, is_title=False, level=None, source='')
                          for p in self.all_paras]

        existing = [it['index'] for it in self.items if it['is_title']]
        dlg = AddParagraphDialog(self.root, self.all_paras, existing)
        if not dlg.result:
            return

        selected = set(dlg.result)
        for it in self.items:
            it['is_title'] = it['index'] in selected
            if it['is_title'] and not it.get('level'):
                lv, src = detect_heading(it)
                it['level'] = lv or 1
                it['source'] = src or '手动标记'
            if it['is_title'] and not it.get('source'):
                it['source'] = '手动标记'
            if not it['is_title']:
                it['level'] = None
                it['source'] = ''
        self.refresh_tree()

    def _collect_templates(self):
        return {lvl: var.get().strip()
                for lvl, var in self.template_vars.items()
                if var.get().strip()}

    def _collect_separators(self):
        return {lvl: SEP_MAP.get(var.get(), var.get())
                for lvl, var in self.sep_vars.items()}

    def _collect_number_formats(self):
        return {lvl: NUMFMT_MAP.get(var.get(), 'arabic')
                for lvl, var in self.numfmt_vars.items()}

    def _collect_enabled_num_levels(self):
        return {lvl for lvl, var in self.num_enabled_vars.items() if var.get()}

    def _collect_format_settings(self):
        result = {}
        for lvl, panel in self.format_panels.items():
            if panel.is_enabled():
                result[lvl] = panel.get_settings()
        return result

    def _collect_checked_titles(self):
        checked = []
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, 'values')
            if vals[0] != '✓':
                continue
            try:
                idx = int(vals[5])
            except (TypeError, ValueError, IndexError):
                continue
            try:
                lv = int(vals[2])
                if not 1 <= lv <= 9:
                    raise ValueError
            except (TypeError, ValueError):
                continue
            src = ''
            for it in self.items:
                if it['index'] == idx:
                    src = it.get('source', '')
                    break
            checked.append({
                'index': idx, 'level': lv,
                'text': str(vals[3]) if len(vals) > 3 else '',
                'is_title': True, 'source': src,
            })
        return checked

    def export(self):
        if not self.src_path:
            messagebox.showwarning("提示", "请先选择 Word 文档")
            return
        if not self.items:
            messagebox.showwarning("提示", "请先识别标题")
            return

        title_list = self._collect_checked_titles()
        if not title_list:
            messagebox.showwarning("提示",
                                   "列表中没有任何被勾选、且级别有效的标题")
            return

        out_path = filedialog.asksaveasfilename(
            title="另存为新文件", defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")])
        if not out_path:
            return

        apply_format = (self.mode_var.get() == 'reformat')
        format_settings = (self._collect_format_settings()
                           if apply_format else {})
        fmt_levels = set(format_settings.keys()) if apply_format else set()

        templates = self._collect_templates()
        separators = self._collect_separators()
        number_formats = self._collect_number_formats()
        num_levels = self._collect_enabled_num_levels()

        body_format = None
        if self.body_panel is not None and self.body_panel.is_enabled():
            body_format = self.body_panel.get_settings()

        if body_format:
            if self.body_start_idx is None or self.body_end_idx is None:
                messagebox.showwarning(
                    "提示",
                    "启用正文格式前，需要先在标题列表中设置：\n"
                    "  · 正文开始标题\n"
                    "  · 正文结束标题\n\n"
                    "请右键标题行，选择「将选中行设为『正文开始』」和\n"
                    "「将选中行设为『正文结束』」。")
                return

            title_idxs = sorted([t['index'] for t in title_list])
            is_last_title = (title_idxs and
                             self.body_end_idx == title_idxs[-1])

            if not is_last_title:
                if self.body_start_idx >= self.body_end_idx:
                    messagebox.showwarning(
                        "提示",
                        "「正文开始」标题必须位于「正文结束」标题之前。\n"
                        "请重新设置。")
                    return

        if not num_levels and not fmt_levels and not body_format:
            messagebox.showwarning(
                "提示",
                "编号、标题格式、正文格式都没有启用，无需导出。")
            return

        used_levels = {t['level'] for t in title_list}
        missing = [lv for lv in sorted(used_levels)
                   if lv in num_levels and lv not in templates]
        if missing:
            if not messagebox.askyesno(
                    "提示",
                    f"以下启用编号的级别没有设置编号模板：{missing}\n"
                    "未设置的级别将使用默认模板 '{级别}'。\n是否继续？"):
                return

        fmt_str = ','.join(str(x) for x in sorted(fmt_levels)) or '无'
        num_str = ','.join(str(x) for x in sorted(num_levels)) or '无'
        body_str = '启用' if body_format else '未启用'

        body_range = None
        if body_format:
            start_idx = self.body_start_idx
            end_idx = self.body_end_idx
            title_idxs = sorted([t['index'] for t in title_list])
            if title_idxs and end_idx == title_idxs[-1]:
                body_range = (start_idx, None)
            else:
                body_range = (start_idx, end_idx)

        def work(progress_cb):
            export_document(
                self.src_path, out_path, title_list, format_settings,
                templates, separators, number_formats,
                fmt_levels, num_levels,
                apply_format=apply_format,
                body_format=body_format,
                body_range=body_range,
                progress_cb=progress_cb)
            return True

        def on_done(_):
            mode_txt = ("修改格式 + 重新编号" if apply_format
                        else "只重新编号")
            range_txt = ""
            if body_range:
                if body_range[1] is None:
                    range_txt = (f"\n  正文范围：段索引 "
                                 f"{body_range[0]} 之后（延伸到文档末尾）")
                else:
                    range_txt = (f"\n  正文范围：段索引 "
                                 f"{body_range[0]} ~ {body_range[1]}")
            messagebox.showinfo(
                "完成",
                f"[{mode_txt}]\n"
                f"  修改标题格式的级别：{fmt_str}\n"
                f"  重新编号的级别：{num_str}\n"
                f"  正文格式：{body_str}{range_txt}\n"
                f"  编号方式：纯文本编号\n"
                f"  输出文件：\n{out_path}")

        self._run_with_progress(work, on_done=on_done,
                                title="导出 Word 中...")

    def export_titles(self):
        if not self.items:
            messagebox.showwarning("提示", "请先识别标题")
            return
        title_list = self._collect_checked_titles()
        if not title_list:
            messagebox.showwarning("提示",
                                   "列表中没有任何被勾选、且级别有效的标题")
            return

        path = filedialog.asksaveasfilename(
            title="导出标题清单", defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("文本文件", "*.txt")])
        if not path:
            return

        templates = self._collect_templates()
        separators = self._collect_separators()
        number_formats = self._collect_number_formats()
        strip = bool(self.strip_number_var.get())
        try:
            export_titles_list(title_list, path, templates, separators,
                               number_formats, None,
                               strip_number=strip)
            messagebox.showinfo("完成", f"已导出标题清单：\n{path}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("错误", f"导出失败：{e}")


# =========================================================
# Tab 2：表格提取到 Excel
# =========================================================

class TableExtractApp(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.root = self.winfo_toplevel()

        self.word_path = None
        self.excel_path = None
        self.columns = []
        self.rows = []

        self._build_ui()

    def _build_ui(self):
        wrap = ttk.LabelFrame(self, text="文件选择")
        wrap.pack(fill=tk.X, padx=8, pady=8)

        row1 = ttk.Frame(wrap)
        row1.pack(fill=tk.X, padx=6, pady=4)
        ttk.Button(row1, text="选择 Word 文档",
                   command=self.select_word).pack(side=tk.LEFT)
        self.word_label = ttk.Label(
            row1, text="未选择", foreground='#555',
            width=MAX_FILE_DISPLAY_CHARS + 2, anchor='w')
        self.word_label.pack(side=tk.LEFT, padx=6)

        row2 = ttk.Frame(wrap)
        row2.pack(fill=tk.X, padx=6, pady=4)
        ttk.Button(row2, text="选择输出 Excel",
                   command=self.select_excel).pack(side=tk.LEFT)
        self.excel_label = ttk.Label(
            row2, text="未选择", foreground='#555',
            width=MAX_FILE_DISPLAY_CHARS + 2, anchor='w')
        self.excel_label.pack(side=tk.LEFT, padx=6)

        ops = ttk.Frame(self)
        ops.pack(fill=tk.X, padx=8, pady=4)

        ttk.Button(ops, text="提取预览",
                   command=self.extract).pack(side=tk.LEFT)
        ttk.Button(ops, text="提取并导出 Excel",
                   command=self.extract_and_export).pack(side=tk.LEFT, padx=6)
        ttk.Button(ops, text="清空",
                   command=self.clear).pack(side=tk.LEFT, padx=6)

        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(ops, textvariable=self.status_var,
                  foreground='#0066cc').pack(side=tk.LEFT, padx=12)

        info = ttk.LabelFrame(self, text="提取规则")
        info.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(
            info, justify='left',
            text=(
                "· 提取范围：所有表格 + 正文段落（正文排除表格内的）\n"
                "· 多行表格：短且无冒号的 key-value 行识别为一行数据\n"
                "· 合并大单元格：按标题段落切分成多个数据块，每块 Excel 一行\n"
                "· 标题判定：Word 标题样式 / 大纲级别 / 编号模式\n"
                "  （如 \"7.8 上下电测试\"、\"7.8.1 xxx\"、\"第X章 xxx\"）\n"
                "· 单元格内：每段 \"列名: 内容\" 开始新字段；内容跨多行直到\n"
                "  遇到下一个 \"列名:\" 或标题段落"),
            foreground='#444'
        ).pack(anchor='w', padx=8, pady=6)

        mid = ttk.LabelFrame(self, text="提取结果预览")
        mid.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        self.preview_tree = ttk.Treeview(mid, show='headings')
        self.preview_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT,
                               padx=4, pady=4)
        sb_y = ttk.Scrollbar(mid, orient='vertical',
                             command=self.preview_tree.yview)
        sb_y.pack(side=tk.RIGHT, fill=tk.Y)
        sb_x = ttk.Scrollbar(mid, orient='horizontal',
                             command=self.preview_tree.xview)
        sb_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.preview_tree.configure(yscrollcommand=sb_y.set,
                                    xscrollcommand=sb_x.set)

    def select_word(self):
        path = filedialog.askopenfilename(
            title="选择 Word 文档",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if path:
            self.word_path = path
            self.word_label.config(text=truncate_path(path))

    def select_excel(self):
        path = filedialog.asksaveasfilename(
            title="选择输出 Excel 文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.excel_path = path
            self.excel_label.config(text=truncate_path(path))

    def extract(self):
        if not self.word_path:
            messagebox.showwarning("提示", "请先选择 Word 文档")
            return
        try:
            cols, rows = extract_tables_from_word(self.word_path)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("错误", f"提取失败：{e}")
            return

        self.columns = cols
        self.rows = rows
        self._refresh_preview()
        self.status_var.set(
            f"提取完成：{len(self.rows)} 行，{len(self.columns)} 列")

    def extract_and_export(self):
        if not self.word_path:
            messagebox.showwarning("提示", "请先选择 Word 文档")
            return

        if not self.excel_path:
            path = filedialog.asksaveasfilename(
                title="另存为 Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel 文件", "*.xlsx")])
            if not path:
                return
            self.excel_path = path
            self.excel_label.config(text=truncate_path(path))

        try:
            cols, rows = extract_tables_from_word(self.word_path)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("错误", f"提取失败：{e}")
            return

        self.columns = cols
        self.rows = rows
        self._refresh_preview()

        if not self.columns:
            messagebox.showwarning("提示", "未提取到任何内容")
            self.status_var.set("未提取到任何内容")
            return

        try:
            export_tables_to_excel(self.columns, self.rows, self.excel_path)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("错误", f"导出失败：{e}")
            return

        self.status_var.set(
            f"已导出：{len(self.rows)} 行，{len(self.columns)} 列")
        messagebox.showinfo(
            "完成",
            f"已提取 {len(self.rows)} 行、{len(self.columns)} 列。\n"
            f"输出文件：\n{self.excel_path}")

    def _refresh_preview(self):
        self.preview_tree.delete(*self.preview_tree.get_children())
        if not self.columns:
            self.preview_tree['columns'] = ()
            return
        self.preview_tree['columns'] = self.columns
        for col in self.columns:
            self.preview_tree.heading(col, text=str(col))
            self.preview_tree.column(col, width=140, anchor='w')
        for r in self.rows:
            values = [str(r.get(c, '')) for c in self.columns]
            self.preview_tree.insert('', tk.END, values=values)

    def clear(self):
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_tree['columns'] = ()
        self.columns = []
        self.rows = []
        self.status_var.set("已清空")


# =========================================================
# 入口
# =========================================================

def main():
    root = tk.Tk()
    root.title("Word 工具集")

    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    w = min(1150, sw - 60)
    h = min(820, sh - 90)
    x = max(0, (sw - w) // 2)
    y = max(0, (sh - h) // 2)
    root.geometry(f"{w}x{h}+{x}+{y}")
    root.minsize(900, 500)

    try:
        style = ttk.Style()
        style.theme_use('clam')
    except Exception:
        pass

    nb = ttk.Notebook(root)
    nb.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

    tab1 = App(nb)
    nb.add(tab1, text="  标题处理  ")

    tab2 = TableExtractApp(nb)
    nb.add(tab2, text="  表格/正文提取到 Excel  ")

    root.mainloop()


if __name__ == "__main__":
    main()

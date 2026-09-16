# -*- coding: utf-8 -*-
"""
Word 标题识别、格式统一与自动编号工具
依赖：
    pip install python-docx pywin32

运行：
    python main.py
"""

import os
import re
import sys
import shutil
import tempfile
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------- COM 支持 ----------
try:
    import win32com.client as win32
    import pythoncom
    HAS_COM = True
except ImportError:
    HAS_COM = False


W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


# =========================================================
# 1. 文档解析
# =========================================================

def get_outline_level(paragraph):
    """从段落 XML 获取大纲级别 (1-9)，无则返回 None"""
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
    """从样式名解析标题级别，如 'Heading 1'、'标题 2'"""
    if not style_name:
        return None
    m = re.match(r'^(?:Heading|标题)\s*(\d+)', style_name.strip(), re.I)
    if m:
        return int(m.group(1))
    return None


def detect_number_pattern(text):
    """从文本前缀判断编号级别，返回 (级别, 前缀文本)，无则 (None, None)"""
    # 第一章 / 第一节 / 第一篇 / 第一部分
    m = re.match(r'^(第[一二三四五六七八九十百千万零〇\d]+[章节篇部分])\s*', text)
    if m:
        prefix = m.group(1)
        if '节' in prefix:
            return 2, prefix
        return 1, prefix

    # 多级数字 1.1.1 / 1.1
    m = re.match(r'^(\d+(?:[\.．]\d+)+)([\.．、\s])', text)
    if m:
        num = m.group(1).replace('．', '.')
        return num.count('.') + 1, m.group(0)

    # 单级数字 1. / 1、
    m = re.match(r'^(\d+)([\.．、\s])', text)
    if m:
        return 1, m.group(0)

    # 一、二、
    m = re.match(r'^([一二三四五六七八九十]+)[、.．]\s*', text)
    if m:
        return 1, m.group(0)

    # (1) （1）
    m = re.match(r'^[\(（]\d+[\)）]\s*', text)
    if m:
        return 3, m.group(0)

    return None, None


def load_paragraphs(docx_path):
    """加载文档所有非空段落，返回列表"""
    doc = Document(docx_path)
    paras = []
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            continue
        style_name = p.style.name if p.style else ''
        outline = get_outline_level(p)
        paras.append({
            'index': i,
            'text': text,
            'style': style_name,
            'outline': outline,
            'is_title': False,
            'level': None,
            'source': '',
        })
    return paras


def detect_heading(para):
    """识别段落是否为标题，返回 (级别, 来源)，优先级：大纲级别 > 内置样式 > 编号模式"""
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
    """删除标题开头的旧编号"""
    patterns = [
        r'^第[一二三四五六七八九十百千万零〇\d]+[章节篇部分][\s:：、.．]*',
        r'^\d+(?:[\.．]\d+)+[\.．、\s:：]*',
        r'^\d+[\.．、\s:：]+',
        r'^[一二三四五六七八九十]+[、.．\s:：]+',
        r'^[\(（]\d+[\)）][\s.．、:：]*',
        r'^[①-⑳]\s*',
    ]
    for pat in patterns:
        new_text = re.sub(pat, '', text, count=1)
        if new_text != text:
            text = new_text
    return text.strip()


# =========================================================
# 2. 编号生成
# =========================================================

def generate_numbers(titles, templates, separators):
    """
    titles: 按文档顺序排列的标题列表，每项含 'level' 和 'is_title'
    templates: {1: '第{1}章', 2: '{1}.{2}', ...}
    separators: {1: ' ', ...}
    返回与 titles 等长的前缀字符串列表（非标题返回 ''）
    """
    counters = [0] * 10
    result = []
    for t in titles:
        if not t.get('is_title'):
            result.append('')
            continue
        level = t.get('level') or 1
        if level < 1 or level > 9:
            level = 1
        counters[level - 1] += 1
        for i in range(level, 10):
            counters[i] = 0
        tmpl = templates.get(level, '{' + str(level) + '}')
        s = tmpl
        for i in range(1, level + 1):
            s = s.replace('{' + str(i) + '}', str(counters[i - 1]))
        sep = separators.get(level, ' ')
        result.append(s + sep)
    return result


# =========================================================
# 3. python-docx 处理
# =========================================================

def _clear_paragraph_runs(p):
    """清空段落所有 run"""
    for r in list(p.runs):
        r._element.getparent().remove(r._element)


def _set_paragraph_text(p, text):
    """重写段落文字，保留段落本身"""
    _clear_paragraph_runs(p)
    p.add_run(text)


def _apply_heading_style(p, doc, level):
    """应用内置标题样式"""
    for name in (f"标题 {level}", f"Heading {level}"):
        try:
            p.style = doc.styles[name]
            return True
        except KeyError:
            continue
    return False


def _set_style_font_cn(style, font_name):
    """设置样式的中文字体"""
    style.font.name = font_name
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), font_name)


def _apply_format_to_styles(doc, format_settings):
    """将格式设置应用到内置标题样式"""
    for level, fmt in format_settings.items():
        applied = False
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
                applied = True
            except Exception:
                traceback.print_exc()
            break


def _clean_and_style(doc, title_list):
    """去掉旧编号、应用标题样式"""
    for t in title_list:
        idx = t['index']
        if idx >= len(doc.paragraphs):
            continue
        p = doc.paragraphs[idx]
        clean = remove_old_number(p.text)
        _set_paragraph_text(p, clean)
        _apply_heading_style(p, doc, t['level'])


# =========================================================
# 4. COM 自动编号
# =========================================================

def get_word_app():
    """获取 Word 或 WPS 的 COM 应用对象"""
    if not HAS_COM:
        return None
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    for prog_id in ("Word.Application", "Kwps.Application", "kwps.Application"):
        try:
            app = win32.DispatchEx(prog_id)
            app.Visible = False
            app.DisplayAlerts = False
            return app
        except Exception:
            continue
    return None


def apply_auto_numbering(doc, title_list, templates, separators):
    """
    在 Word COM 文档中为标题应用多级自动编号
    doc: Word COM Document
    title_list: [{'index': 0-based 段落序号, 'level': 级别}, ...] 按文档顺序
    返回 True 成功 / False 失败
    """
    try:
        list_gallery = doc.Application.ListGalleries(2)  # wdOutlineNumberGallery
        try:
            list_template = list_gallery.ListTemplates(1)
        except Exception:
            list_template = list_gallery.ListTemplates.Add()

        # 配置每一级
        for lvl_num in range(1, 10):
            tmpl = templates.get(lvl_num)
            if not tmpl:
                continue
            try:
                level_obj = list_template.ListLevels(lvl_num)
                w_tmpl = re.sub(r'\{(\d+)\}', lambda m: '%' + m.group(1), tmpl)
                sep = separators.get(lvl_num, ' ')
                level_obj.NumberFormat = w_tmpl + sep
                level_obj.NumberStyle = 0       # wdListNumberStyleArabic
                level_obj.StartAt = 1
                level_obj.NumberPosition = 0
                level_obj.TextPosition = 0
                level_obj.TrailingCharacter = 0 # wdTrailingTab
                try:
                    level_obj.LinkedStyle = f"标题 {lvl_num}"
                except Exception:
                    try:
                        level_obj.LinkedStyle = f"Heading {lvl_num}"
                    except Exception:
                        pass
            except Exception:
                continue

        # 逐个应用标题样式并应用列表
        for t in title_list:
            try:
                para = doc.Paragraphs(t['index'] + 1)
                for sname in (f"标题 {t['level']}", f"Heading {t['level']}"):
                    try:
                        para.Style = doc.Styles(sname)
                        break
                    except Exception:
                        continue
                try:
                    para.Range.ListFormat.ApplyListTemplateWithLevel(
                        ListTemplate=list_template,
                        ContinuePreviousList=True,
                        ApplyTo=0,                # wdListApplyToWholeList
                        DefaultListBehavior=2,    # wdWord10ListBehavior
                        ApplyLevel=t['level']
                    )
                except Exception:
                    # 部分版本没有 ApplyListTemplateWithLevel
                    try:
                        para.Range.ListFormat.ApplyListTemplate(
                            ListTemplate=list_template,
                            ContinuePreviousList=True,
                            ApplyTo=0,
                            DefaultListBehavior=2
                        )
                    except Exception:
                        pass
            except Exception:
                continue
        return True
    except Exception:
        traceback.print_exc()
        return False


def try_com_export(src_path, out_path, title_list, format_settings,
                   templates, separators):
    """尝试用 COM 自动编号导出"""
    if not HAS_COM:
        return False
    tmp_dir = tempfile.mkdtemp(prefix='word_title_')
    tmp_path = os.path.join(tmp_dir, 'stage1.docx')
    app = None
    try:
        # 1. 用 python-docx 预处理：去旧编号 + 应用样式 + 格式
        doc = Document(src_path)
        _clean_and_style(doc, title_list)
        _apply_format_to_styles(doc, format_settings)
        doc.save(tmp_path)

        # 2. 用 COM 打开并应用自动编号
        app = get_word_app()
        if app is None:
            return False
        wdoc = app.Documents.Open(os.path.abspath(tmp_path), ReadOnly=False)
        try:
            if not apply_auto_numbering(wdoc, title_list, templates, separators):
                return False
            try:
                wdoc.SaveAs2(os.path.abspath(out_path), FileFormat=16)  # docx
            except Exception:
                wdoc.SaveAs(os.path.abspath(out_path))
            return True
        finally:
            try:
                wdoc.Close(SaveChanges=False)
            except Exception:
                pass
    except Exception:
        traceback.print_exc()
        return False
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        shutil.rmtree(tmp_dir, ignore_errors=True)


# =========================================================
# 5. 纯文本回退导出
# =========================================================

def text_only_export(src_path, out_path, all_titles, title_list,
                     format_settings, templates, separators):
    """纯文本编号回退"""
    doc = Document(src_path)
    _clean_and_style(doc, title_list)
    _apply_format_to_styles(doc, format_settings)

    # 生成编号（按文档顺序）
    all_sorted = sorted(all_titles, key=lambda x: x['index'])
    nums = generate_numbers(all_sorted, templates, separators)

    for t, num in zip(all_sorted, nums):
        if not t.get('is_title'):
            continue
        idx = t['index']
        if idx >= len(doc.paragraphs):
            continue
        p = doc.paragraphs[idx]
        clean = p.text  # 已经去掉了旧编号
        _set_paragraph_text(p, num + clean)

    doc.save(out_path)


# =========================================================
# 6. 导出主入口
# =========================================================

def export_document(src_path, out_path, titles, format_settings,
                    templates, separators, use_auto_number=True):
    """
    titles: 完整列表（含 is_title 标记），按文档顺序
    format_settings: {level: {'font_name','font_size','bold','color_rgb','alignment','space_before','space_after'}}
    templates: {level: '第{1}章'}
    separators: {level: ' '}
    """
    title_list = [t for t in titles if t.get('is_title') and t.get('level')]
    title_list.sort(key=lambda x: x['index'])

    if use_auto_number and HAS_COM:
        if try_com_export(src_path, out_path, title_list,
                          format_settings, templates, separators):
            return True
        # 自动编号失败，回退纯文本
    text_only_export(src_path, out_path, titles, title_list,
                     format_settings, templates, separators)
    return False


# =========================================================
# 7. GUI
# =========================================================

ALIGN_MAP = {
    '左对齐': WD_ALIGN_PARAGRAPH.LEFT,
    '居中': WD_ALIGN_PARAGRAPH.CENTER,
    '右对齐': WD_ALIGN_PARAGRAPH.RIGHT,
    '两端对齐': WD_ALIGN_PARAGRAPH.JUSTIFY,
}
ALIGN_NAMES = list(ALIGN_MAP.keys())


class LevelFormatPanel:
    """单个级别格式设置面板"""
    def __init__(self, parent, level, defaults):
        self.level = level
        self.frame = ttk.LabelFrame(parent, text=f"{level} 级标题格式")
        self.frame.pack(fill=tk.X, padx=6, pady=3)

        row1 = ttk.Frame(self.frame)
        row1.pack(fill=tk.X, padx=4, pady=2)

        ttk.Label(row1, text="字体").pack(side=tk.LEFT)
        self.font_var = tk.StringVar(value=defaults.get('font_name', '宋体'))
        ttk.Entry(row1, textvariable=self.font_var, width=10).pack(side=tk.LEFT, padx=(2, 8))

        ttk.Label(row1, text="字号").pack(side=tk.LEFT)
        self.size_var = tk.IntVar(value=defaults.get('font_size', 14))
        ttk.Spinbox(row1, from_=8, to=72, textvariable=self.size_var, width=5).pack(side=tk.LEFT, padx=(2, 8))

        self.bold_var = tk.BooleanVar(value=defaults.get('bold', True))
        ttk.Checkbutton(row1, text="加粗", variable=self.bold_var).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(row1, text="对齐").pack(side=tk.LEFT)
        self.align_var = tk.StringVar(value=defaults.get('align', '左对齐'))
        ttk.Combobox(row1, textvariable=self.align_var, values=ALIGN_NAMES,
                     width=8, state='readonly').pack(side=tk.LEFT, padx=(2, 8))

        ttk.Label(row1, text="颜色").pack(side=tk.LEFT)
        self.color_rgb = defaults.get('color_rgb', (0, 0, 0))
        self.color_btn = tk.Button(row1, text="    ", bg=self._rgb_to_hex(self.color_rgb),
                                   width=3, command=self.pick_color)
        self.color_btn.pack(side=tk.LEFT, padx=(2, 8))

        row2 = ttk.Frame(self.frame)
        row2.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(row2, text="段前(磅)").pack(side=tk.LEFT)
        self.before_var = tk.IntVar(value=defaults.get('space_before', 6))
        ttk.Spinbox(row2, from_=0, to=100, textvariable=self.before_var, width=5).pack(side=tk.LEFT, padx=(2, 8))
        ttk.Label(row2, text="段后(磅)").pack(side=tk.LEFT)
        self.after_var = tk.IntVar(value=defaults.get('space_after', 6))
        ttk.Spinbox(row2, from_=0, to=100, textvariable=self.after_var, width=5).pack(side=tk.LEFT, padx=(2, 8))

    @staticmethod
    def _rgb_to_hex(rgb):
        return '#%02x%02x%02x' % rgb

    def pick_color(self):
        rgb, hx = colorchooser.askcolor(color=self._rgb_to_hex(self.color_rgb))
        if rgb:
            self.color_rgb = tuple(int(x) for x in rgb)
            self.color_btn.config(bg=hx)

    def get_settings(self):
        return {
            'font_name': self.font_var.get().strip(),
            'font_size': int(self.size_var.get()),
            'bold': bool(self.bold_var.get()),
            'color_rgb': self.color_rgb,
            'alignment': ALIGN_MAP.get(self.align_var.get(), WD_ALIGN_PARAGRAPH.LEFT),
            'space_before': int(self.before_var.get()),
            'space_after': int(self.after_var.get()),
        }


class AddParagraphDialog:
    """从所有段落中选择段落标记为标题"""
    def __init__(self, parent, all_paras, existing_indexes):
        self.top = tk.Toplevel(parent)
        self.top.title("从所有段落中添加标题")
        self.top.geometry("820x520")
        self.result = []          # 返回选中的段落列表
        self.all_paras = all_paras
        self.existing = set(existing_indexes)

        ttk.Label(self.top, text="勾选要添加为标题的段落（可多选）").pack(anchor='w', padx=8, pady=6)

        cols = ("选择", "序号", "样式", "大纲", "内容")
        self.tree = ttk.Treeview(self.top, columns=cols, show='headings', height=18)
        for c in cols:
            self.tree.heading(c, text=c)
        self.tree.column("选择", width=50, anchor='center')
        self.tree.column("序号", width=60, anchor='center')
        self.tree.column("样式", width=110, anchor='w')
        self.tree.column("大纲", width=60, anchor='center')
        self.tree.column("内容", width=520, anchor='w')
        self.tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        self.tree.bind('<Button-1>', self.on_click)

        self.check_state = {}   # iid -> bool
        for p in all_paras:
            iid = str(p['index'])
            mark = '✓' if p['index'] in self.existing else ''
            outline = p.get('outline') or ''
            self.tree.insert('', tk.END, iid=iid, values=(
                mark, p['index'] + 1, p['style'], outline, p['text'][:80]
            ))
            self.check_state[iid] = (p['index'] in self.existing)

        btns = ttk.Frame(self.top)
        btns.pack(fill=tk.X, padx=8, pady=6)
        ttk.Button(btns, text="取消", command=self.top.destroy).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btns, text="确定", command=self.confirm).pack(side=tk.RIGHT, padx=4)

        self.top.transient(parent)
        self.top.grab_set()
        parent.wait_window(self.top)

    def on_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != 'cell':
            return
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if col == '#1':
            self.check_state[iid] = not self.check_state.get(iid, False)
            vals = list(self.tree.item(iid, 'values'))
            vals[0] = '✓' if self.check_state[iid] else ''
            self.tree.item(iid, values=vals)

    def confirm(self):
        self.result = [int(iid) for iid, chk in self.check_state.items() if chk]
        self.top.destroy()


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Word 标题识别、格式统一与自动编号工具")
        self.root.geometry("1060x780")

        self.src_path = None
        self.all_paras = []      # 加载文档后的所有段落
        self.items = []          # 当前标题列表（可增删、改级别）

        self.format_panels = {}
        self.template_vars = {}  # {level: StringVar}
        self.sep_vars = {}       # {level: StringVar}

        self._build_ui()

    # ---------- UI ----------
    def _build_ui(self):
        # 顶部
        top = ttk.Frame(self.root)
        top.pack(fill=tk.X, padx=8, pady=6)

        ttk.Button(top, text="选择 Word 文档", command=self.select_file).pack(side=tk.LEFT)
        self.file_label = ttk.Label(top, text="未选择文件", foreground='#555')
        self.file_label.pack(side=tk.LEFT, padx=10)

        ttk.Button(top, text="识别标题", command=self.detect_titles).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="从所有段落添加", command=self.open_add_dialog).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="移除选中", command=self.remove_selected).pack(side=tk.LEFT, padx=4)

        # 中部：标题列表
        mid = ttk.LabelFrame(self.root, text="标题列表（点击第一列勾选/取消；双击“级别”列可修改）")
        mid.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        cols = ("包含", "序号", "级别", "标题文字", "来源", "段落索引")
        self.tree = ttk.Treeview(mid, columns=cols, show='headings', height=14)
        for c in cols:
            self.tree.heading(c, text=c)
        self.tree.column("包含", width=50, anchor='center')
        self.tree.column("序号", width=50, anchor='center')
        self.tree.column("级别", width=60, anchor='center')
        self.tree.column("标题文字", width=560, anchor='w')
        self.tree.column("来源", width=90, anchor='center')
        self.tree.column("段落索引", width=80, anchor='center')
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=4, pady=4)
        sb = ttk.Scrollbar(mid, orient='vertical', command=self.tree.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=sb.set)

        self.tree.bind('<Button-1>', self.on_tree_click)
        self.tree.bind('<Double-1>', self.on_tree_double)

        # 格式 + 编号 面板
        paned = ttk.Frame(self.root)
        paned.pack(fill=tk.X, padx=8, pady=4)

        left = ttk.Frame(paned)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        right = ttk.Frame(paned)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))

        # 格式
        fmt_wrap = ttk.LabelFrame(left, text="按级别统一设置格式")
        fmt_wrap.pack(fill=tk.BOTH, expand=True)
        default_fonts = {1: ('黑体', 16, True), 2: ('楷体', 14, True), 3: ('宋体', 12, False)}
        for lvl in (1, 2, 3):
            fn, fs, bd = default_fonts[lvl]
            self.format_panels[lvl] = LevelFormatPanel(fmt_wrap, lvl, {
                'font_name': fn, 'font_size': fs, 'bold': bd,
                'align': '左对齐', 'color_rgb': (0, 0, 0),
                'space_before': 6 if lvl == 1 else 3,
                'space_after': 6 if lvl == 1 else 3,
            })

        # 编号模板
        num_wrap = ttk.LabelFrame(right, text="编号模板与分隔符（占位符 {1},{2},{3} 表示各级序号）")
        num_wrap.pack(fill=tk.BOTH, expand=True)
        defaults = {
            1: ("第{1}章", " "),
            2: ("{1}.{2}", " "),
            3: ("{1}.{2}.{3}", " "),
        }
        for lvl in (1, 2, 3):
            row = ttk.Frame(num_wrap)
            row.pack(fill=tk.X, padx=4, pady=4)
            ttk.Label(row, text=f"{lvl} 级模板").pack(side=tk.LEFT)
            tv = tk.StringVar(value=defaults[lvl][0])
            ttk.Entry(row, textvariable=tv, width=22).pack(side=tk.LEFT, padx=(4, 10))
            ttk.Label(row, text="分隔符").pack(side=tk.LEFT)
            sv = tk.StringVar(value=defaults[lvl][1])
            ttk.Entry(row, textvariable=sv, width=6).pack(side=tk.LEFT, padx=4)
            self.template_vars[lvl] = tv
            self.sep_vars[lvl] = sv

        self.auto_number_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(num_wrap,
                        text="优先使用 Word 自动编号（失败自动回退为纯文本编号）",
                        variable=self.auto_number_var).pack(anchor='w', padx=6, pady=6)

        # 底部按钮
        bottom = ttk.Frame(self.root)
        bottom.pack(fill=tk.X, padx=8, pady=8)
        ttk.Button(bottom, text="退出", command=self.root.quit).pack(side=tk.RIGHT, padx=4)
        ttk.Button(bottom, text="导出为新 Word 文件",
                   command=self.export).pack(side=tk.RIGHT, padx=4)

    # ---------- 事件 ----------
    def select_file(self):
        path = filedialog.askopenfilename(
            title="选择 Word 文档",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
        )
        if path:
            self.src_path = path
            self.file_label.config(text=path)
            # 重置
            self.all_paras = []
            self.items = []
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

        self.items = []
        for p in self.all_paras:
            level, source = detect_heading(p)
            item = dict(p)
            if level:
                item['is_title'] = True
                item['level'] = level
                item['source'] = source
            else:
                item['is_title'] = False
                item['level'] = None
                item['source'] = ''
            self.items.append(item)

        # 只显示识别为标题的 + 用户已添加的
        self.refresh_tree()
        n = sum(1 for x in self.items if x['is_title'])
        messagebox.showinfo("识别完成", f"共识别到 {n} 个标题（可手动增删、调整级别）")

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        # 只显示 is_title = True 的
        shown = [it for it in self.items if it['is_title']]
        shown.sort(key=lambda x: x['index'])
        for i, it in enumerate(shown):
            self.tree.insert('', tk.END, iid=str(i), values=(
                '✓',
                i + 1,
                it['level'],
                it['text'],
                it['source'],
                it['index'],
            ))

    def _visible_items(self):
        shown = [it for it in self.items if it['is_title']]
        shown.sort(key=lambda x: x['index'])
        return shown

    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != 'cell':
            return
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if col == '#1':  # 第一列“包含”
            vals = list(self.tree.item(iid, 'values'))
            if vals[0] == '✓':
                vals[0] = ''
            else:
                vals[0] = '✓'
            self.tree.item(iid, values=vals)

    def on_tree_double(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != 'cell':
            return
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if not iid:
            return

        if col == '#3':  # 级别列
            self.edit_level(iid)
        elif col == '#4':  # 标题文字列
            self.edit_text(iid)

    def edit_level(self, iid):
        vals = list(self.tree.item(iid, 'values'))
        cur = str(vals[2])
        win = tk.Toplevel(self.root)
        win.title("修改级别")
        ttk.Label(win, text="请输入 1~9 的级别：").pack(padx=10, pady=6)
        v = tk.StringVar(value=cur)
        ent = ttk.Entry(win, textvariable=v, width=8)
        ent.pack(padx=10, pady=4)
        ent.focus_set()

        def ok():
            try:
                lv = int(v.get())
                if not 1 <= lv <= 9:
                    raise ValueError
            except ValueError:
                messagebox.showerror("错误", "级别必须是 1~9 的整数", parent=win)
                return
            vals[2] = lv
            self.tree.item(iid, values=vals)
            # 同步到 self.items
            idx = int(vals[5])
            for it in self.items:
                if it['index'] == idx:
                    it['level'] = lv
                    break
            win.destroy()

        ttk.Button(win, text="确定", command=ok).pack(pady=6)
        win.transient(self.root)
        win.grab_set()

    def edit_text(self, iid):
        vals = list(self.tree.item(iid, 'values'))
        cur = str(vals[3])
        win = tk.Toplevel(self.root)
        win.title("修改标题文字")
        ttk.Label(win, text="编辑标题文字（可手动去掉旧编号）：").pack(padx=10, pady=6)
        txt = tk.Text(win, width=60, height=4)
        txt.pack(padx=10, pady=4)
        txt.insert('1.0', cur)
        txt.focus_set()

        def ok():
            new_text = txt.get('1.0', 'end').strip().replace('\n', ' ')
            vals[3] = new_text
            self.tree.item(iid, values=vals)
            idx = int(vals[5])
            for it in self.items:
                if it['index'] == idx:
                    it['text'] = new_text
                    break
            win.destroy()

        ttk.Button(win, text="确定", command=ok).pack(pady=6)
        win.transient(self.root)
        win.grab_set()

    def remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在列表中选择要移除的项")
            return
        for iid in sel:
            vals = self.tree.item(iid, 'values')
            idx = int(vals[5])
            for it in self.items:
                if it['index'] == idx:
                    it['is_title'] = False
                    it['level'] = None
                    it['source'] = ''
                    break
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

        # 保证 items 与 all_paras 同步
        if not self.items:
            self.items = [dict(p, is_title=False, level=None, source='') for p in self.all_paras]

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

    # ---------- 导出 ----------
    def export(self):
        if not self.src_path:
            messagebox.showwarning("提示", "请先选择 Word 文档")
            return
        if not self.items:
            messagebox.showwarning("提示", "请先识别标题")
            return

        # 只导出被勾选的（第一列是 ✓）
        checked_indexes = set()
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, 'values')
            if vals[0] == '✓':
                checked_indexes.add(int(vals[5]))
        if not checked_indexes:
            messagebox.showwarning("提示", "列表中没有任何被勾选的标题")
            return

        # 应用列表中的修改（级别、文字）到 items
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, 'values')
            idx = int(vals[5])
            for it in self.items:
                if it['index'] == idx:
                    try:
                        it['level'] = int(vals[2])
                    except Exception:
                        pass
                    it['text'] = str(vals[3])
                    break

        # 未勾选的取消 is_title
        for it in self.items:
            if it.get('is_title') and it['index'] not in checked_indexes:
                it['is_title'] = False

        # 输出文件
        out_path = filedialog.asksaveasfilename(
            title="另存为新文件",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")]
        )
        if not out_path:
            return

        # 收集格式
        format_settings = {lvl: panel.get_settings()
                           for lvl, panel in self.format_panels.items()}
        templates = {lvl: v.get() for lvl, v in self.template_vars.items()}
        separators = {lvl: v.get() for lvl, v in self.sep_vars.items()}

        try:
            used_auto = export_document(
                self.src_path, out_path, self.items, format_settings,
                templates, separators,
                use_auto_number=self.auto_number_var.get()
            )
            if used_auto:
                messagebox.showinfo("完成", f"已使用 Word 自动编号导出：\n{out_path}")
            else:
                messagebox.showinfo("完成",
                                    f"已使用纯文本编号导出（自动编号不可用或失败）：\n{out_path}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("错误", f"导出失败：{e}")


# =========================================================
# 8. 入口
# =========================================================

def main():
    root = tk.Tk()
    try:
        # 稍微美化一下
        style = ttk.Style()
        style.theme_use('clam')
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

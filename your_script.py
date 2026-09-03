import sys
import re
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook, load_workbook
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QListWidget, QListWidgetItem, QLabel,
    QComboBox, QSpinBox, QDoubleSpinBox, QCheckBox, QTableWidget, QTableWidgetItem,
    QTabWidget, QLineEdit, QMessageBox, QProgressBar, QTextEdit,
    QSplitter, QGroupBox, QFormLayout, QDialog, QDialogButtonBox
)
from PyQt5.QtCore import Qt

# ================== 辅助函数 ==================

def is_heading_paragraph(para):
    """
    判断段落是否为标题，返回 (是否, 层级)
    完全依据编号样式识别，不依赖字体格式。
    支持：
    1. 数字编号：1, 1.1, 1.1.1
    2. 中文序号：一、 二、 （一） 1） 等
    """
    text = para.text.strip()
    if not text:
        return False, 0

    # 1. 数字编号：1, 1.1, 1.1.1
    if re.match(r'^(\d+(\.\d+)*)\s+', text):
        parts = text.split()[0].split('.')
        level = min(len(parts), 3)
        return True, level

    # 2. 中文序号
    # 一级：一、 二、 三、
    if re.match(r'^[一二三四五六七八九十百千]+、', text):
        return True, 1
    # 二级：（一） （二）
    if re.match(r'^（[一二三四五六七八九十百千]+）', text):
        return True, 2
    # 二级：1） 2）
    if re.match(r'^\d+）', text):
        return True, 2
    # 三级：① ② 等
    if re.match(r'^[①-⑩]', text):
        return True, 3

    return False, 0

def set_heading_style(paragraph, level):
    """将段落设置为内置标题样式"""
    style_name = f'Heading {level}'
    paragraph.style = paragraph.part.document.styles[style_name]

def add_toc(paragraph, levels=3):
    """在指定段落插入目录域"""
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = f'TOC \\o "1-{levels}" \\h \\z \\u'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t')
    t.text = "目录：请在此处右键选择“更新域”以生成目录内容。"
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')

    run._r.append(fldChar)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(t)
    run._r.append(fldChar3)

def get_table_name(table, doc, index):
    """尝试获取表格名称：简化返回表格序号"""
    return f"表格 {index+1}"

# ================== 主窗口 ==================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word 智能标题与表格处理工具")
        self.setGeometry(100, 100, 1000, 800)

        self.current_doc_path = None
        self.doc = None
        self.headings = []  # 存储 (paragraph, level, text, index)

        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 文件选择区
        file_layout = QHBoxLayout()
        self.btn_open = QPushButton("打开 Word 文档")
        self.btn_open.clicked.connect(self.open_document)
        self.lbl_file = QLabel("未选择文件")
        file_layout.addWidget(self.btn_open)
        file_layout.addWidget(self.lbl_file, 1)
        main_layout.addLayout(file_layout)

        # 选项卡
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # 选项卡1：标题识别与目录
        self.tab1 = QWidget()
        self.tabs.addTab(self.tab1, "标题识别与目录")
        self.init_tab1()

        # 选项卡2：Word 表格填充
        self.tab2 = QWidget()
        self.tabs.addTab(self.tab2, "表格填充")
        self.init_tab2()

        # 选项卡3：Word ↔ Excel
        self.tab3 = QWidget()
        self.tabs.addTab(self.tab3, "Word ↔ Excel")
        self.init_tab3()

        # 状态栏
        self.statusBar().showMessage("就绪")

    def init_tab1(self):
        layout = QVBoxLayout(self.tab1)

        # 扫描按钮
        self.btn_scan = QPushButton("扫描标题")
        self.btn_scan.clicked.connect(self.scan_headings)
        layout.addWidget(self.btn_scan)

        # 标题列表
        self.list_headings = QListWidget()
        self.list_headings.setSelectionMode(QListWidget.MultiSelection)
        layout.addWidget(QLabel("识别到的标题（可多选，用于应用样式或选择格式来源）："))
        layout.addWidget(self.list_headings)

        # 应用样式按钮
        self.btn_apply_style = QPushButton("将选中标题应用为 Word 标题样式")
        self.btn_apply_style.clicked.connect(self.apply_heading_styles)
        layout.addWidget(self.btn_apply_style)

        # 格式统一区域
        group_format = QGroupBox("标题格式统一与自动编码")
        format_layout = QVBoxLayout()

        # 编号样式选择
        format_row1 = QHBoxLayout()
        format_row1.addWidget(QLabel("编号样式："))
        self.combo_number_style = QComboBox()
        self.combo_number_style.addItems(["数字编号：1, 1.1, 1.1.1", "中文序号：一、（一）、1）", "保持原编号"])
        format_row1.addWidget(self.combo_number_style)
        format_layout.addLayout(format_row1)

        # 格式来源选择
        format_row2 = QHBoxLayout()
        format_row2.addWidget(QLabel("格式来源："))
        self.radio_format_source = QComboBox()
        self.radio_format_source.addItems(["从选定标题复制格式", "手动设置格式"])
        format_row2.addWidget(self.radio_format_source)
        format_layout.addLayout(format_row2)

        # 手动设置格式控件（初始隐藏）
        self.manual_format_widget = QWidget()
        manual_layout = QHBoxLayout()
        manual_layout.addWidget(QLabel("字体："))
        self.combo_font = QComboBox()
        self.combo_font.addItems(["宋体", "黑体", "微软雅黑", "仿宋", "楷体", "Arial", "Times New Roman"])
        manual_layout.addWidget(self.combo_font)
        manual_layout.addWidget(QLabel("字号："))
        self.combo_font_size = QComboBox()
        self.combo_font_size.addItems(["12", "14", "16", "18", "20", "22", "24", "26", "28", "36"])
        self.combo_font_size.setCurrentText("16")
        manual_layout.addWidget(self.combo_font_size)
        self.check_bold = QCheckBox("加粗")
        self.check_bold.setChecked(True)
        manual_layout.addWidget(self.check_bold)
        self.manual_format_widget.setLayout(manual_layout)
        self.manual_format_widget.setVisible(False)
        format_layout.addWidget(self.manual_format_widget)

        # 连接信号：切换格式来源时显示/隐藏手动设置
        self.radio_format_source.currentIndexChanged.connect(
            lambda idx: self.manual_format_widget.setVisible(idx == 1)
        )

        # 执行按钮
        self.btn_unify_format = QPushButton("统一格式并自动编码")
        self.btn_unify_format.clicked.connect(self.unify_format_and_number)
        format_layout.addWidget(self.btn_unify_format)

        group_format.setLayout(format_layout)
        layout.addWidget(group_format)

        # 目录设置
        group_dir = QGroupBox("生成目录")
        dir_layout = QFormLayout()
        self.spin_levels = QSpinBox()
        self.spin_levels.setRange(1, 3)
        self.spin_levels.setValue(3)
        dir_layout.addRow("显示级别：", self.spin_levels)
        self.btn_insert_toc = QPushButton("插入目录（含目录标题）")
        self.btn_insert_toc.clicked.connect(self.insert_toc)
        dir_layout.addRow(self.btn_insert_toc)
        group_dir.setLayout(dir_layout)
        layout.addWidget(group_dir)

    def init_tab2(self):
        layout = QVBoxLayout(self.tab2)

        # 表格选择
        self.btn_list_tables = QPushButton("列出所有表格")
        self.btn_list_tables.clicked.connect(self.list_tables)
        layout.addWidget(self.btn_list_tables)

        self.combo_tables = QComboBox()
        self.combo_tables.currentIndexChanged.connect(self.on_table_selected)
        layout.addWidget(QLabel("选择要填充的表格："))
        layout.addWidget(self.combo_tables)

        # 表格列选择
        self.combo_seq_col = QComboBox()
        self.combo_title_col = QComboBox()
        layout.addWidget(QLabel("序号列（表头名称或列号）："))
        layout.addWidget(self.combo_seq_col)
        layout.addWidget(QLabel("标题列："))
        layout.addWidget(self.combo_title_col)

        # 层级选择
        self.check_level1 = QCheckBox("一级标题")
        self.check_level1.setChecked(True)
        self.check_level2 = QCheckBox("二级标题")
        self.check_level2.setChecked(True)
        self.check_level3 = QCheckBox("三级标题")
        self.check_level3.setChecked(True)
        level_layout = QHBoxLayout()
        level_layout.addWidget(self.check_level1)
        level_layout.addWidget(self.check_level2)
        level_layout.addWidget(self.check_level3)
        layout.addWidget(QLabel("填充层级："))
        layout.addLayout(level_layout)

        # 填充按钮
        self.btn_fill_table = QPushButton("填充表格")
        self.btn_fill_table.clicked.connect(self.fill_table)
        layout.addWidget(self.btn_fill_table)

    def init_tab3(self):
        layout = QVBoxLayout(self.tab3)

        # 导出 Word -> Excel
        group_export = QGroupBox("Word 表格导出到 Excel")
        export_layout = QVBoxLayout()
        self.btn_export = QPushButton("选择表格并导出")
        self.btn_export.clicked.connect(self.export_to_excel)
        export_layout.addWidget(self.btn_export)
        group_export.setLayout(export_layout)
        layout.addWidget(group_export)

        # 导入 Excel -> Word
        group_import = QGroupBox("Excel 数据导入 Word 表格")
        import_layout = QVBoxLayout()
        self.btn_import = QPushButton("选择 Excel 文件并导入")
        self.btn_import.clicked.connect(self.import_from_excel)
        import_layout.addWidget(self.btn_import)
        group_import.setLayout(import_layout)
        layout.addWidget(group_import)

    # ================== 文件操作 ==================
    def open_document(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 Word 文档", "", "Word 文档 (*.docx)")
        if not file_path:
            return
        self.current_doc_path = file_path
        self.lbl_file.setText(file_path)
        self.doc = Document(file_path)
        self.statusBar().showMessage(f"已打开：{file_path}")
        # 清空之前的数据
        self.headings = []
        self.list_headings.clear()
        self.combo_tables.clear()

    def get_new_save_path(self):
        """生成新文件保存路径（不覆盖原文件）"""
        if not self.current_doc_path:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return None
        dir_path = os.path.dirname(self.current_doc_path)
        base_name = os.path.splitext(os.path.basename(self.current_doc_path))[0]
        new_name = f"{base_name}_已处理.docx"
        new_path = os.path.join(dir_path, new_name)
        # 如果文件已存在，弹窗让用户选择其他位置或覆盖
        if os.path.exists(new_path):
            new_path, _ = QFileDialog.getSaveFileName(
                self, "保存新文件", new_path, "Word 文档 (*.docx)"
            )
            if not new_path:
                return None
        return new_path

    # ================== 标题识别 ==================
    def scan_headings(self):
        if not self.doc:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return
        self.headings = []
        self.list_headings.clear()

        for i, para in enumerate(self.doc.paragraphs):
            is_heading, level = is_heading_paragraph(para)
            if is_heading:
                self.headings.append((para, level, para.text.strip(), i))
                item = QListWidgetItem(f"[{level}] {para.text.strip()}")
                item.setData(Qt.UserRole, len(self.headings)-1)
                self.list_headings.addItem(item)

        if not self.headings:
            QMessageBox.information(self, "提示", "未识别到标题")
        else:
            self.statusBar().showMessage(f"识别到 {len(self.headings)} 个标题")

    def apply_heading_styles(self):
        selected_items = self.list_headings.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选择要应用样式的标题")
            return
        for item in selected_items:
            idx = item.data(Qt.UserRole)
            para, level, text, _ = self.headings[idx]
            set_heading_style(para, level)
        # 另存为新文件
        new_path = self.get_new_save_path()
        if new_path:
            self.doc.save(new_path)
            QMessageBox.information(self, "完成", f"已应用标题样式，新文件已保存为：{new_path}")

    # ================== 格式统一与自动编码 ==================
    def unify_format_and_number(self):
        """统一标题格式并自动编码"""
        if not self.doc or not self.headings:
            QMessageBox.warning(self, "警告", "请先打开文档并扫描标题")
            return

        # 确定格式来源
        format_source = self.radio_format_source.currentIndex()
        if format_source == 0:  # 从选定标题复制格式
            selected_items = self.list_headings.selectedItems()
            if len(selected_items) != 1:
                QMessageBox.warning(self, "警告", "请选择一个“完美格式”标题作为格式来源")
                return
            source_idx = selected_items[0].data(Qt.UserRole)
            source_para, source_level, _, _ = self.headings[source_idx]
            # 获取源标题的格式
            source_format = self.get_paragraph_format(source_para)
            # 应用到所有同级标题
            for i, (para, level, text, _) in enumerate(self.headings):
                if level == source_level:
                    self.apply_paragraph_format(para, source_format)
        else:  # 手动设置格式
            font_name = self.combo_font.currentText()
            font_size = float(self.combo_font_size.currentText())
            bold = self.check_bold.isChecked()
            # 应用到所有标题
            for para, level, text, _ in self.headings:
                self.apply_manual_format(para, font_name, font_size, bold)

        # 自动编码
        number_style = self.combo_number_style.currentIndex()
        self.auto_number_headings(number_style)

        # 另存为新文件
        new_path = self.get_new_save_path()
        if new_path:
            self.doc.save(new_path)
            QMessageBox.information(self, "完成", f"格式统一与自动编码完成，新文件已保存为：{new_path}")

    def get_paragraph_format(self, para):
        """获取段落的格式信息（字体、字号、加粗、颜色等）"""
        format_info = {}
        if para.runs:
            run = para.runs[0]
            format_info['font_name'] = run.font.name
            format_info['font_size'] = run.font.size.pt if run.font.size else None
            format_info['bold'] = run.font.bold
            format_info['color'] = run.font.color.rgb if run.font.color and run.font.color.rgb else None
        else:
            # 使用段落样式
            style = para.style
            format_info['font_name'] = style.font.name
            format_info['font_size'] = style.font.size.pt if style.font.size else None
            format_info['bold'] = style.font.bold
            format_info['color'] = style.font.color.rgb if style.font.color and style.font.color.rgb else None
        return format_info

    def apply_paragraph_format(self, para, format_info):
        """将格式应用到段落的所有 run"""
        for run in para.runs:
            if format_info.get('font_name'):
                run.font.name = format_info['font_name']
            if format_info.get('font_size'):
                run.font.size = Pt(format_info['font_size'])
            if format_info.get('bold') is not None:
                run.font.bold = format_info['bold']
            if format_info.get('color'):
                run.font.color.rgb = format_info['color']

    def apply_manual_format(self, para, font_name, font_size, bold):
        """手动设置格式应用到段落"""
        for run in para.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = bold

    def auto_number_headings(self, number_style):
        """自动为标题添加编号"""
        if number_style == 0:  # 数字编号：1, 1.1, 1.1.1
            self.number_by_digit()
        elif number_style == 1:  # 中文序号：一、（一）、1）
            self.number_by_chinese()
        # number_style == 2: 保持原编号，不处理

    def number_by_digit(self):
        """数字编号：一级 1, 二级 1.1, 三级 1.1.1"""
        counters = [0, 0, 0]
        for para, level, text, _ in self.headings:
            counters[level-1] += 1
            # 重置下级计数器
            for i in range(level, 3):
                counters[i] = 0
            # 生成编号
            if level == 1:
                new_num = str(counters[0])
            elif level == 2:
                new_num = f"{counters[0]}.{counters[1]}"
            else:
                new_num = f"{counters[0]}.{counters[1]}.{counters[2]}"
            # 替换标题文本中的编号部分（保留原文字）
            new_text = f"{new_num} {self.remove_old_number(text)}"
            self.set_paragraph_text(para, new_text)

    def number_by_chinese(self):
        """中文序号：一级 一、 二级 （一） 三级 1）"""
        level1_count = 0
        level2_count = 0
        level3_count = 0
        for para, level, text, _ in self.headings:
            if level == 1:
                level1_count += 1
                level2_count = 0
                level3_count = 0
                new_num = self.to_chinese_number(level1_count) + "、"
            elif level == 2:
                level2_count += 1
                level3_count = 0
                new_num = "（" + self.to_chinese_number(level2_count) + "）"
            else:
                level3_count += 1
                new_num = str(level3_count) + "）"
            new_text = f"{new_num}{self.remove_old_number(text)}"
            self.set_paragraph_text(para, new_text)

    def to_chinese_number(self, num):
        """将数字转换为中文数字（1-99）"""
        chinese_digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
        if num <= 10:
            return chinese_digits[num] if num < 10 else '十'
        elif num < 20:
            return '十' + chinese_digits[num % 10] if num % 10 != 0 else '十'
        elif num < 100:
            tens = num // 10
            ones = num % 10
            result = chinese_digits[tens] + '十'
            if ones != 0:
                result += chinese_digits[ones]
            return result
        else:
            return str(num)

    def remove_old_number(self, text):
        """移除标题文本开头的旧编号"""
        # 移除数字编号：1, 1.1, 1.1.1
        text = re.sub(r'^(\d+(\.\d+)*)\s+', '', text)
        # 移除中文序号：一、
        text = re.sub(r'^[一二三四五六七八九十百千]+、', '', text)
        # 移除中文括号序号：（一）
        text = re.sub(r'^（[一二三四五六七八九十百千]+）', '', text)
        # 移除数字括号序号：1）
        text = re.sub(r'^\d+）', '', text)
        # 移除圈号序号：①
        text = re.sub(r'^[①-⑩]', '', text)
        return text.strip()

    def set_paragraph_text(self, para, new_text):
        """设置段落文本（保留第一个run的格式）"""
        if para.runs:
            para.runs[0].text = new_text
            # 删除多余的run
            for run in para.runs[1:]:
                run.text = ""
        else:
            para.add_run(new_text)

    # ================== 目录生成 ==================
    def insert_toc(self):
        if not self.doc:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return
        levels = self.spin_levels.value()

        first_para = self.doc.paragraphs[0]

        # 先插入目录域
        toc_para = first_para.insert_paragraph_before()
        add_toc(toc_para, levels)

        # 再在目录域上方插入“目录”标题
        title_para = toc_para.insert_paragraph_before()
        title_run = title_para.add_run("目录")
        title_run.bold = True
        title_run.font.size = Pt(16)
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 另存为新文件
        new_path = self.get_new_save_path()
        if new_path:
            self.doc.save(new_path)
            QMessageBox.information(self, "完成", f"目录已插入，新文件已保存为：{new_path}\n请在 Word 中更新域以显示页码")

    # ================== 表格填充 ==================
    def list_tables(self):
        if not self.doc:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return
        self.combo_tables.clear()
        for i, table in enumerate(self.doc.tables):
            name = get_table_name(table, self.doc, i)
            self.combo_tables.addItem(f"{i+1}: {name}", i)
        if not self.doc.tables:
            QMessageBox.information(self, "提示", "文档中没有表格")
        else:
            self.statusBar().showMessage(f"找到 {len(self.doc.tables)} 个表格")

    def on_table_selected(self, index):
        if index < 0 or not self.doc:
            return
        table_idx = self.combo_tables.itemData(index)
        if table_idx is None:
            return
        table = self.doc.tables[table_idx]
        self.combo_seq_col.clear()
        self.combo_title_col.clear()
        headers = []
        if len(table.rows) > 0:
            for cell in table.rows[0].cells:
                headers.append(cell.text.strip())
        for col_idx in range(len(table.columns)):
            if headers and col_idx < len(headers):
                label = f"{col_idx+1} - {headers[col_idx]}" if headers[col_idx] else f"列 {col_idx+1}"
            else:
                label = f"列 {col_idx+1}"
            self.combo_seq_col.addItem(label, col_idx)
            self.combo_title_col.addItem(label, col_idx)

    def fill_table(self):
        if not self.doc or not self.headings:
            QMessageBox.warning(self, "警告", "请先打开文档并扫描标题")
            return
        table_idx = self.combo_tables.currentData()
        if table_idx is None:
            QMessageBox.warning(self, "警告", "请先选择表格")
            return
        seq_col = self.combo_seq_col.currentData()
        title_col = self.combo_title_col.currentData()
        if seq_col is None or title_col is None:
            QMessageBox.warning(self, "警告", "请指定列")
            return

        levels_wanted = []
        if self.check_level1.isChecked():
            levels_wanted.append(1)
        if self.check_level2.isChecked():
            levels_wanted.append(2)
        if self.check_level3.isChecked():
            levels_wanted.append(3)

        titles_to_fill = []
        for para, level, text, _ in self.headings:
            if level in levels_wanted:
                titles_to_fill.append((level, text))

        if not titles_to_fill:
            QMessageBox.warning(self, "警告", "没有符合层级的标题")
            return

        table = self.doc.tables[table_idx]
        data_rows = len(table.rows) - 1
        needed_rows = len(titles_to_fill)
        if data_rows < needed_rows:
            for _ in range(needed_rows - data_rows):
                table.add_row()

        for i, (level, text) in enumerate(titles_to_fill):
            row = table.rows[i + 1]
            seq_text = self.extract_seq_text(text, i+1)
            row.cells[seq_col].text = seq_text
            row.cells[title_col].text = text

        # 另存为新文件
        new_path = self.get_new_save_path()
        if new_path:
            self.doc.save(new_path)
            QMessageBox.information(self, "完成", f"表格填充完成，新文件已保存为：{new_path}")

    def extract_seq_text(self, text, fallback_idx):
        """从标题文本中提取序号部分"""
        # 数字编号：1, 1.1
        match = re.match(r'^(\d+(\.\d+)*)\s+', text)
        if match:
            return match.group(1)
        # 中文序号：一、
        match = re.match(r'^([一二三四五六七八九十百千]+、)', text)
        if match:
            return match.group(1)
        # 中文括号序号：（一）
        match = re.match(r'^(（[一二三四五六七八九十百千]+）)', text)
        if match:
            return match.group(1)
        # 数字括号序号：1）
        match = re.match(r'^(\d+）)', text)
        if match:
            return match.group(1)
        # 圈号序号：①
        match = re.match(r'^([①-⑩])', text)
        if match:
            return match.group(1)
        return str(fallback_idx)

    # ================== Excel 导出导入 ==================
    def export_to_excel(self):
        if not self.doc:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return
        if not self.doc.tables:
            QMessageBox.warning(self, "警告", "文档中没有表格")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("选择要导出的表格")
        layout = QVBoxLayout(dialog)
        list_widget = QListWidget()
        list_widget.setSelectionMode(QListWidget.MultiSelection)
        for i, table in enumerate(self.doc.tables):
            name = get_table_name(table, self.doc, i)
            item = QListWidgetItem(f"{i+1}: {name}")
            item.setData(Qt.UserRole, i)
            list_widget.addItem(item)
        layout.addWidget(list_widget)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec_() != QDialog.Accepted:
            return
        selected = [item.data(Qt.UserRole) for item in list_widget.selectedItems()]
        if not selected:
            return
        save_path, _ = QFileDialog.getSaveFileName(self, "保存 Excel 文件", "", "Excel 文件 (*.xlsx)")
        if not save_path:
            return
        wb = Workbook()
        wb.remove(wb.active)
        for idx in selected:
            table = self.doc.tables[idx]
            ws = wb.create_sheet(title=f"表格{idx+1}")
            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    ws.cell(row=row_idx+1, column=col_idx+1, value=cell.text)
        wb.save(save_path)
        QMessageBox.information(self, "完成", f"已导出 {len(selected)} 个表格到 {save_path}")

    def import_from_excel(self):
        if not self.doc:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return
        if not self.doc.tables:
            QMessageBox.warning(self, "警告", "文档中没有表格")
            return
        excel_path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx)")
        if not excel_path:
            return
        wb = load_workbook(excel_path, data_only=True)
        sheet_names = wb.sheetnames
        if not sheet_names:
            QMessageBox.warning(self, "警告", "Excel 文件为空")
            return
        ws = wb[sheet_names[0]]
        headers = [cell.value for cell in ws[1]]
        data_rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            data_rows.append(row)
        if not data_rows:
            QMessageBox.warning(self, "警告", "Excel 没有数据")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("选择目标表格和列映射")
        layout = QFormLayout(dialog)
        combo_table = QComboBox()
        for i, table in enumerate(self.doc.tables):
            combo_table.addItem(f"{i+1}: {get_table_name(table, self.doc, i)}", i)
        layout.addRow("目标表格：", combo_table)

        combo_word_col = QComboBox()
        table_idx = combo_table.currentData()
        if table_idx is not None:
            table = self.doc.tables[table_idx]
            for col_idx in range(len(table.columns)):
                header = table.rows[0].cells[col_idx].text.strip() if table.rows else f"列{col_idx+1}"
                combo_word_col.addItem(f"{col_idx+1} - {header}", col_idx)
        combo_table.currentIndexChanged.connect(
            lambda: self.update_import_cols(combo_table, combo_word_col)
        )
        layout.addRow("Word 列：", combo_word_col)

        combo_excel_col = QComboBox()
        for i, h in enumerate(headers):
            label = f"{i+1} - {h}" if h else f"列 {i+1}"
            combo_excel_col.addItem(label, i)
        layout.addRow("Excel 列：", combo_excel_col)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addRow(buttons)

        if dialog.exec_() != QDialog.Accepted:
            return

        target_table_idx = combo_table.currentData()
        word_col = combo_word_col.currentData()
        excel_col = combo_excel_col.currentData()
        if target_table_idx is None or word_col is None or excel_col is None:
            return

        table = self.doc.tables[target_table_idx]
        needed = len(data_rows)
        current_data_rows = len(table.rows) - 1
        if current_data_rows < needed:
            for _ in range(needed - current_data_rows):
                table.add_row()
        for i, row_data in enumerate(data_rows):
            value = row_data[excel_col]
            table.rows[i+1].cells[word_col].text = str(value) if value is not None else ""

        # 另存为新文件
        new_path = self.get_new_save_path()
        if new_path:
            self.doc.save(new_path)
            QMessageBox.information(self, "完成", f"导入完成，新文件已保存为：{new_path}")

    def update_import_cols(self, combo_table, combo_word_col):
        combo_word_col.clear()
        table_idx = combo_table.currentData()
        if table_idx is None:
            return
        table = self.doc.tables[table_idx]
        for col_idx in range(len(table.columns)):
            header = table.rows[0].cells[col_idx].text.strip() if table.rows else f"列{col_idx+1}"
            combo_word_col.addItem(f"{col_idx+1} - {header}", col_idx)

# ================== 程序入口 ==================
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

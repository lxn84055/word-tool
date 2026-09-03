import sys
import re
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from openpyxl import Workbook, load_workbook
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QListWidget, QListWidgetItem, QLabel,
    QComboBox, QSpinBox, QDoubleSpinBox, QCheckBox, QTableWidget, QTableWidgetItem,
    QTabWidget, QLineEdit, QMessageBox, QTextEdit, QDialog, QDialogButtonBox,
    QGroupBox, QFormLayout, QMenu, QAction, QColorDialog, QFontComboBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

# ================== 辅助函数 ==================

def is_heading_paragraph(para):
    """判断段落是否为标题，返回 (是否, 层级)"""
    text = para.text.strip()
    if not text:
        return False, 0

    # 数字编号：1, 1.1, 1.1.1
    if re.match(r'^(\d+(\.\d+)*)\s+', text):
        parts = text.split()[0].split('.')
        level = min(len(parts), 3)
        return True, level

    # 中文序号一级：一、 二、
    if re.match(r'^[一二三四五六七八九十百千]+、', text):
        return True, 1
    # 中文序号二级：（一）
    if re.match(r'^（[一二三四五六七八九十百千]+）', text):
        return True, 2
    # 数字括号二级：1）
    if re.match(r'^\d+）', text):
        return True, 2
    # 圈号三级：①
    if re.match(r'^[①-⑩]', text):
        return True, 3

    return False, 0

def set_heading_style(paragraph, level):
    """设置为内置标题样式"""
    try:
        paragraph.style = paragraph.part.document.styles[f'Heading {level}']
    except:
        pass

def add_toc(paragraph, levels=3):
    """插入目录域"""
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
    return f"表格 {index+1}"

def is_empty_paragraph(para):
    """判断是否为空段落（无文本或只有空白）"""
    return not para.text.strip()

def count_empty_paragraphs_before(doc, para_index):
    """计算段落前的连续空段落数"""
    count = 0
    idx = para_index - 1
    while idx >= 0:
        para = doc.paragraphs[idx]
        if is_empty_paragraph(para):
            count += 1
            idx -= 1
        else:
            break
    return count

def count_empty_paragraphs_after(doc, para_index):
    """计算段落后的连续空段落数"""
    count = 0
    idx = para_index + 1
    while idx < len(doc.paragraphs):
        para = doc.paragraphs[idx]
        if is_empty_paragraph(para):
            count += 1
            idx += 1
        else:
            break
    return count

# ================== 主窗口 ==================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word 智能标题与表格处理工具")
        self.setGeometry(100, 100, 1200, 900)

        self.current_doc_path = None
        self.doc = None
        self.headings = []  # 存储 (paragraph_index, level, text)

        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 文件选择
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

        self.tab1 = QWidget()
        self.tabs.addTab(self.tab1, "标题识别与目录")
        self.init_tab1()

        self.tab2 = QWidget()
        self.tabs.addTab(self.tab2, "表格填充")
        self.init_tab2()

        self.tab3 = QWidget()
        self.tabs.addTab(self.tab3, "Word ↔ Excel")
        self.init_tab3()

        self.statusBar().showMessage("就绪")

    def init_tab1(self):
        layout = QVBoxLayout(self.tab1)

        # 扫描按钮
        self.btn_scan = QPushButton("扫描标题")
        self.btn_scan.clicked.connect(self.scan_headings)
        layout.addWidget(self.btn_scan)

        # 标题列表（带右键菜单）
        self.list_headings = QListWidget()
        self.list_headings.setSelectionMode(QListWidget.MultiSelection)
        self.list_headings.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list_headings.customContextMenuRequested.connect(self.show_heading_context_menu)
        layout.addWidget(QLabel("识别到的标题（右键可管理）："))
        layout.addWidget(self.list_headings)

        # 按钮行
        btn_row = QHBoxLayout()
        self.btn_add_heading = QPushButton("从段落中选取添加标题")
        self.btn_add_heading.clicked.connect(self.show_paragraph_picker)
        self.btn_delete_heading = QPushButton("删除选中标题")
        self.btn_delete_heading.clicked.connect(self.delete_selected_headings)
        self.btn_mark_non_heading = QPushButton("标记为非标题")
        self.btn_mark_non_heading.clicked.connect(self.mark_selected_as_non_heading)
        btn_row.addWidget(self.btn_add_heading)
        btn_row.addWidget(self.btn_delete_heading)
        btn_row.addWidget(self.btn_mark_non_heading)
        layout.addLayout(btn_row)

        # 调整层级
        level_row = QHBoxLayout()
        level_row.addWidget(QLabel("调整选中标题层级为："))
        self.combo_set_level = QComboBox()
        self.combo_set_level.addItems(["一级", "二级", "三级"])
        level_row.addWidget(self.combo_set_level)
        self.btn_set_level = QPushButton("应用层级")
        self.btn_set_level.clicked.connect(self.set_selected_level)
        level_row.addWidget(self.btn_set_level)
        layout.addLayout(level_row)

        # 应用样式
        self.btn_apply_style = QPushButton("将选中标题应用为 Word 标题样式")
        self.btn_apply_style.clicked.connect(self.apply_heading_styles)
        layout.addWidget(self.btn_apply_style)

        # 格式统一区域
        group_format = QGroupBox("标题格式统一与自动编号")
        format_layout = QVBoxLayout()

        # 作用层级
        level_select_row = QHBoxLayout()
        level_select_row.addWidget(QLabel("作用层级："))
        self.check_unify_level1 = QCheckBox("一级")
        self.check_unify_level1.setChecked(True)
        self.check_unify_level2 = QCheckBox("二级")
        self.check_unify_level2.setChecked(True)
        self.check_unify_level3 = QCheckBox("三级")
        self.check_unify_level3.setChecked(True)
        level_select_row.addWidget(self.check_unify_level1)
        level_select_row.addWidget(self.check_unify_level2)
        level_select_row.addWidget(self.check_unify_level3)
        format_layout.addLayout(level_select_row)

        # 编号样式
        num_row = QHBoxLayout()
        num_row.addWidget(QLabel("编号样式："))
        self.combo_number_style = QComboBox()
        self.combo_number_style.addItems(["数字编号：1, 1.1, 1.1.1", "中文序号：一、（一）、1）", "保持原编号"])
        num_row.addWidget(self.combo_number_style)
        format_layout.addLayout(num_row)

        # 格式来源
        source_row = QHBoxLayout()
        source_row.addWidget(QLabel("格式来源："))
        self.radio_format_source = QComboBox()
        self.radio_format_source.addItems(["从选定标题复制格式", "手动设置格式"])
        source_row.addWidget(self.radio_format_source)
        format_layout.addLayout(source_row)

        # 手动设置控件
        self.manual_widget = QWidget()
        manual_layout = QVBoxLayout()
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("字体："))
        self.combo_font = QFontComboBox()
        row1.addWidget(self.combo_font)
        row1.addWidget(QLabel("字号："))
        self.combo_font_size = QComboBox()
        self.combo_font_size.addItems(["9","10","10.5","11","12","14","16","18","20","22","24","26","28","36"])
        self.combo_font_size.setCurrentText("16")
        row1.addWidget(self.combo_font_size)
        manual_layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.check_bold = QCheckBox("加粗")
        self.check_bold.setChecked(True)
        self.check_italic = QCheckBox("斜体")
        self.check_color = QCheckBox("字体颜色")
        self.btn_color = QPushButton("选择颜色")
        self.btn_color.clicked.connect(self.pick_color)
        self.color_value = None
        row2.addWidget(self.check_bold)
        row2.addWidget(self.check_italic)
        row2.addWidget(self.check_color)
        row2.addWidget(self.btn_color)
        manual_layout.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("对齐："))
        self.combo_alignment = QComboBox()
        self.combo_alignment.addItems(["左对齐", "居中", "右对齐", "两端对齐"])
        row3.addWidget(self.combo_alignment)
        manual_layout.addLayout(row3)

        row4 = QHBoxLayout()
        row4.addWidget(QLabel("段前(磅)："))
        self.spin_space_before = QDoubleSpinBox()
        self.spin_space_before.setRange(0, 100)
        self.spin_space_before.setValue(0)
        row4.addWidget(self.spin_space_before)
        row4.addWidget(QLabel("段后(磅)："))
        self.spin_space_after = QDoubleSpinBox()
        self.spin_space_after.setRange(0, 100)
        self.spin_space_after.setValue(0)
        row4.addWidget(self.spin_space_after)
        manual_layout.addLayout(row4)

        row5 = QHBoxLayout()
        row5.addWidget(QLabel("行距："))
        self.combo_line_spacing = QComboBox()
        self.combo_line_spacing.addItems(["单倍行距", "1.5倍行距", "2倍行距", "固定值"])
        row5.addWidget(self.combo_line_spacing)
        row5.addWidget(QLabel("首行缩进(磅)："))
        self.spin_first_indent = QDoubleSpinBox()
        self.spin_first_indent.setRange(0, 100)
        self.spin_first_indent.setValue(0)
        row5.addWidget(self.spin_first_indent)
        manual_layout.addLayout(row5)

        row6 = QHBoxLayout()
        row6.addWidget(QLabel("左缩进(磅)："))
        self.spin_left_indent = QDoubleSpinBox()
        self.spin_left_indent.setRange(0, 100)
        self.spin_left_indent.setValue(0)
        row6.addWidget(self.spin_left_indent)
        row6.addWidget(QLabel("前空行数："))
        self.spin_empty_before = QSpinBox()
        self.spin_empty_before.setRange(0, 10)
        self.spin_empty_before.setValue(0)
        row6.addWidget(self.spin_empty_before)
        row6.addWidget(QLabel("后空行数："))
        self.spin_empty_after = QSpinBox()
        self.spin_empty_after.setRange(0, 10)
        self.spin_empty_after.setValue(0)
        row6.addWidget(self.spin_empty_after)
        manual_layout.addLayout(row6)

        self.manual_widget.setLayout(manual_layout)
        self.manual_widget.setVisible(False)
        format_layout.addWidget(self.manual_widget)

        self.radio_format_source.currentIndexChanged.connect(
            lambda idx: self.manual_widget.setVisible(idx == 1)
        )

        # 执行按钮
        self.btn_unify = QPushButton("统一格式并自动编号")
        self.btn_unify.clicked.connect(self.unify_format_and_number)
        format_layout.addWidget(self.btn_unify)

        group_format.setLayout(format_layout)
        layout.addWidget(group_format)

        # 目录设置
        group_dir = QGroupBox("生成目录")
        dir_layout = QFormLayout()
        self.spin_toc_levels = QSpinBox()
        self.spin_toc_levels.setRange(1, 3)
        self.spin_toc_levels.setValue(3)
        dir_layout.addRow("显示级别：", self.spin_toc_levels)
        self.btn_insert_toc = QPushButton("插入目录（含目录标题）")
        self.btn_insert_toc.clicked.connect(self.insert_toc)
        dir_layout.addRow(self.btn_insert_toc)
        group_dir.setLayout(dir_layout)
        layout.addWidget(group_dir)

    def init_tab2(self):
        layout = QVBoxLayout(self.tab2)
        self.btn_list_tables = QPushButton("列出所有表格")
        self.btn_list_tables.clicked.connect(self.list_tables)
        layout.addWidget(self.btn_list_tables)

        self.combo_tables = QComboBox()
        self.combo_tables.currentIndexChanged.connect(self.on_table_selected)
        layout.addWidget(QLabel("选择要填充的表格："))
        layout.addWidget(self.combo_tables)

        self.combo_seq_col = QComboBox()
        self.combo_title_col = QComboBox()
        layout.addWidget(QLabel("序号列："))
        layout.addWidget(self.combo_seq_col)
        layout.addWidget(QLabel("标题列："))
        layout.addWidget(self.combo_title_col)

        level_row = QHBoxLayout()
        self.check_fill_l1 = QCheckBox("一级")
        self.check_fill_l1.setChecked(True)
        self.check_fill_l2 = QCheckBox("二级")
        self.check_fill_l2.setChecked(True)
        self.check_fill_l3 = QCheckBox("三级")
        self.check_fill_l3.setChecked(True)
        level_row.addWidget(QLabel("填充层级："))
        level_row.addWidget(self.check_fill_l1)
        level_row.addWidget(self.check_fill_l2)
        level_row.addWidget(self.check_fill_l3)
        layout.addLayout(level_row)

        self.btn_fill_table = QPushButton("填充表格")
        self.btn_fill_table.clicked.connect(self.fill_table)
        layout.addWidget(self.btn_fill_table)

    def init_tab3(self):
        layout = QVBoxLayout(self.tab3)
        group_export = QGroupBox("Word 表格导出到 Excel")
        export_layout = QVBoxLayout()
        self.btn_export = QPushButton("选择表格并导出")
        self.btn_export.clicked.connect(self.export_to_excel)
        export_layout.addWidget(self.btn_export)
        group_export.setLayout(export_layout)
        layout.addWidget(group_export)

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
        self.headings = []
        self.list_headings.clear()
        self.combo_tables.clear()
        self.statusBar().showMessage(f"已打开：{file_path}")

    def get_new_save_path(self):
        if not self.current_doc_path:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return None
        dir_path = os.path.dirname(self.current_doc_path)
        base_name = os.path.splitext(os.path.basename(self.current_doc_path))[0]
        new_name = f"{base_name}_已处理.docx"
        new_path = os.path.join(dir_path, new_name)
        if os.path.exists(new_path):
            new_path, _ = QFileDialog.getSaveFileName(self, "保存新文件", new_path, "Word 文档 (*.docx)")
            if not new_path:
                return None
        return new_path

    # ================== 标题扫描与管理 ==================
    def scan_headings(self):
        if not self.doc:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return
        self.headings = []
        self.list_headings.clear()
        for i, para in enumerate(self.doc.paragraphs):
            is_heading, level = is_heading_paragraph(para)
            if is_heading:
                self.headings.append((i, level, para.text.strip()))
                item = QListWidgetItem(f"[{level}] {para.text.strip()}")
                item.setData(Qt.UserRole, len(self.headings)-1)
                self.list_headings.addItem(item)
        if not self.headings:
            QMessageBox.information(self, "提示", "未识别到标题")
        else:
            self.statusBar().showMessage(f"识别到 {len(self.headings)} 个标题")

    def show_heading_context_menu(self, pos):
        menu = QMenu()
        act_add = menu.addAction("从段落中选取添加标题")
        act_delete = menu.addAction("删除选中标题")
        act_mark = menu.addAction("标记为非标题")
        act_level1 = menu.addAction("设为一级标题")
        act_level2 = menu.addAction("设为二级标题")
        act_level3 = menu.addAction("设为三级标题")
        action = menu.exec_(self.list_headings.mapToGlobal(pos))
        if action == act_add:
            self.show_paragraph_picker()
        elif action == act_delete:
            self.delete_selected_headings()
        elif action == act_mark:
            self.mark_selected_as_non_heading()
        elif action == act_level1:
            self.set_selected_level_value(1)
        elif action == act_level2:
            self.set_selected_level_value(2)
        elif action == act_level3:
            self.set_selected_level_value(3)

    def show_paragraph_picker(self):
        if not self.doc:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("选择要标记为标题的段落")
        dialog.setGeometry(200, 200, 600, 500)
        layout = QVBoxLayout(dialog)
        label = QLabel("勾选要标记为标题的段落（已识别的标题置灰）：")
        layout.addWidget(label)
        list_widget = QListWidget()
        list_widget.setSelectionMode(QListWidget.MultiSelection)
        heading_indices = set(h[0] for h in self.headings)
        for i, para in enumerate(self.doc.paragraphs):
            if not para.text.strip():
                continue
            item = QListWidgetItem(f"{i+1}: {para.text.strip()[:60]}")
            item.setData(Qt.UserRole, i)
            if i in heading_indices:
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled)  # 置灰
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
        # 弹窗选择层级
        level_dialog = QDialog(self)
        level_dialog.setWindowTitle("选择标题层级")
        level_layout = QFormLayout(level_dialog)
        combo_level = QComboBox()
        combo_level.addItems(["一级", "二级", "三级"])
        level_layout.addRow("层级：", combo_level)
        level_buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        level_buttons.accepted.connect(level_dialog.accept)
        level_buttons.rejected.connect(level_dialog.reject)
        level_layout.addRow(level_buttons)
        if level_dialog.exec_() != QDialog.Accepted:
            return
        level = combo_level.currentIndex() + 1
        for idx in selected:
            para = self.doc.paragraphs[idx]
            self.headings.append((idx, level, para.text.strip()))
            item = QListWidgetItem(f"[{level}] {para.text.strip()}")
            item.setData(Qt.UserRole, len(self.headings)-1)
            self.list_headings.addItem(item)
        self.statusBar().showMessage(f"已添加 {len(selected)} 个标题")

    def delete_selected_headings(self):
        selected_items = self.list_headings.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选择要删除的标题")
            return
        # 询问是否修改文档格式
        reply = QMessageBox.question(self, "确认", "是否同时清除文档中这些段落的标题格式？",
                                     QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            return
        clear_format = (reply == QMessageBox.Yes)
        indices_to_remove = []
        for item in selected_items:
            idx = item.data(Qt.UserRole)
            indices_to_remove.append(idx)
            if clear_format:
                para_idx = self.headings[idx][0]
                para = self.doc.paragraphs[para_idx]
                para.style = self.doc.styles['Normal']
        # 从列表移除
        indices_to_remove.sort(reverse=True)
        for idx in indices_to_remove:
            self.headings.pop(idx)
        # 刷新列表
        self.refresh_heading_list()

    def mark_selected_as_non_heading(self):
        selected_items = self.list_headings.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选择要标记的标题")
            return
        reply = QMessageBox.question(self, "确认", "是否同时清除文档中这些段落的标题格式？",
                                     QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            return
        clear_format = (reply == QMessageBox.Yes)
        indices_to_remove = []
        for item in selected_items:
            idx = item.data(Qt.UserRole)
            indices_to_remove.append(idx)
            if clear_format:
                para_idx = self.headings[idx][0]
                para = self.doc.paragraphs[para_idx]
                para.style = self.doc.styles['Normal']
        indices_to_remove.sort(reverse=True)
        for idx in indices_to_remove:
            self.headings.pop(idx)
        self.refresh_heading_list()

    def set_selected_level(self):
        level = self.combo_set_level.currentIndex() + 1
        self.set_selected_level_value(level)

    def set_selected_level_value(self, level):
        selected_items = self.list_headings.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选择标题")
            return
        for item in selected_items:
            idx = item.data(Qt.UserRole)
            para_idx, _, text = self.headings[idx]
            self.headings[idx] = (para_idx, level, text)
        self.refresh_heading_list()

    def refresh_heading_list(self):
        self.list_headings.clear()
        for i, (para_idx, level, text) in enumerate(self.headings):
            item = QListWidgetItem(f"[{level}] {text}")
            item.setData(Qt.UserRole, i)
            self.list_headings.addItem(item)

    def apply_heading_styles(self):
        selected_items = self.list_headings.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选择要应用样式的标题")
            return
        for item in selected_items:
            idx = item.data(Qt.UserRole)
            para_idx, level, _ = self.headings[idx]
            para = self.doc.paragraphs[para_idx]
            set_heading_style(para, level)
        new_path = self.get_new_save_path()
        if new_path:
            self.doc.save(new_path)
            QMessageBox.information(self, "完成", f"已应用标题样式，新文件保存为：{new_path}")

    # ================== 格式统一与自动编号 ==================
    def pick_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.color_value = color
            self.btn_color.setStyleSheet(f"background-color: {color.name()}")
            self.check_color.setChecked(True)

    def get_selected_levels(self):
        levels = []
        if self.check_unify_level1.isChecked():
            levels.append(1)
        if self.check_unify_level2.isChecked():
            levels.append(2)
        if self.check_unify_level3.isChecked():
            levels.append(3)
        return levels

    def get_paragraph_full_format(self, para, para_index):
        """获取段落完整格式，包括前后空行数"""
        format_info = {}
        if para.runs:
            run = para.runs[0]
            format_info['font_name'] = run.font.name
            format_info['font_size'] = run.font.size.pt if run.font.size else None
            format_info['bold'] = run.font.bold
            format_info['italic'] = run.font.italic
            format_info['color'] = run.font.color.rgb if run.font.color and run.font.color.rgb else None
        else:
            style = para.style
            format_info['font_name'] = style.font.name
            format_info['font_size'] = style.font.size.pt if style.font.size else None
            format_info['bold'] = style.font.bold
            format_info['italic'] = style.font.italic
            format_info['color'] = style.font.color.rgb if style.font.color and style.font.color.rgb else None

        pf = para.paragraph_format
        format_info['alignment'] = pf.alignment
        format_info['space_before'] = pf.space_before.pt if pf.space_before else 0
        format_info['space_after'] = pf.space_after.pt if pf.space_after else 0
        format_info['line_spacing'] = pf.line_spacing
        format_info['line_spacing_rule'] = pf.line_spacing_rule
        format_info['first_line_indent'] = pf.first_line_indent.pt if pf.first_line_indent else 0
        format_info['left_indent'] = pf.left_indent.pt if pf.left_indent else 0
        format_info['right_indent'] = pf.right_indent.pt if pf.right_indent else 0
        format_info['keep_with_next'] = pf.keep_with_next
        format_info['keep_together'] = pf.keep_together
        format_info['page_break_before'] = pf.page_break_before
        format_info['widow_control'] = pf.widow_control

        # 前后空行数
        format_info['empty_before'] = count_empty_paragraphs_before(self.doc, para_index)
        format_info['empty_after'] = count_empty_paragraphs_after(self.doc, para_index)

        return format_info

    def apply_paragraph_full_format(self, para, para_index, format_info):
        """应用完整格式（包括前后空行处理）"""
        # 应用字体格式
        for run in para.runs:
            if format_info.get('font_name'):
                run.font.name = format_info['font_name']
                if run._element.rPr is not None:
                    rFonts = run._element.rPr.find(qn('w:rFonts'))
                    if rFonts is None:
                        rFonts = OxmlElement('w:rFonts')
                        run._element.rPr.append(rFonts)
                    rFonts.set(qn('w:eastAsia'), format_info['font_name'])
            if format_info.get('font_size'):
                run.font.size = Pt(format_info['font_size'])
            if format_info.get('bold') is not None:
                run.font.bold = format_info['bold']
            if format_info.get('italic') is not None:
                run.font.italic = format_info['italic']
            if format_info.get('color'):
                run.font.color.rgb = format_info['color']

        # 应用段落格式
        pf = para.paragraph_format
        if format_info.get('alignment') is not None:
            pf.alignment = format_info['alignment']
        pf.space_before = Pt(format_info.get('space_before', 0))
        pf.space_after = Pt(format_info.get('space_after', 0))
        if format_info.get('line_spacing') is not None:
            pf.line_spacing = format_info['line_spacing']
        if format_info.get('line_spacing_rule') is not None:
            pf.line_spacing_rule = format_info['line_spacing_rule']
        pf.first_line_indent = Pt(format_info.get('first_line_indent', 0))
        pf.left_indent = Pt(format_info.get('left_indent', 0))
        if format_info.get('right_indent') is not None:
            pf.right_indent = Pt(format_info['right_indent'])
        if format_info.get('keep_with_next') is not None:
            pf.keep_with_next = format_info['keep_with_next']
        if format_info.get('keep_together') is not None:
            pf.keep_together = format_info['keep_together']
        if format_info.get('page_break_before') is not None:
            pf.page_break_before = format_info['page_break_before']
        if format_info.get('widow_control') is not None:
            pf.widow_control = format_info['widow_control']

        # 处理前后空行
        target_before = format_info.get('empty_before', 0)
        target_after = format_info.get('empty_after', 0)
        self.adjust_empty_paragraphs(para_index, target_before, target_after)

    def adjust_empty_paragraphs(self, para_index, target_before, target_after):
        """调整标题前后的空段落数"""
        # 处理前方空行
        current_before = count_empty_paragraphs_before(self.doc, para_index)
        if current_before > target_before:
            # 删除多余空行
            for _ in range(current_before - target_before):
                idx = para_index - 1
                if idx >= 0 and is_empty_paragraph(self.doc.paragraphs[idx]):
                    self.doc.paragraphs[idx]._element.getparent().remove(self.doc.paragraphs[idx]._element)
                    para_index -= 1
        elif current_before < target_before:
            # 插入空行
            for _ in range(target_before - current_before):
                new_para = self.doc.paragraphs[para_index].insert_paragraph_before()
                new_para.style = self.doc.styles['Normal']

        # 处理后方空行（需重新定位段落索引）
        current_after = count_empty_paragraphs_after(self.doc, para_index)
        if current_after > target_after:
            for _ in range(current_after - target_after):
                idx = para_index + 1
                if idx < len(self.doc.paragraphs) and is_empty_paragraph(self.doc.paragraphs[idx]):
                    self.doc.paragraphs[idx]._element.getparent().remove(self.doc.paragraphs[idx]._element)
        elif current_after < target_after:
            for _ in range(target_after - current_after):
                new_para = self.doc.paragraphs[para_index].insert_paragraph_before()
                new_para.style = self.doc.styles['Normal']
                # 移动到标题后面
                new_para._element.addnext(self.doc.paragraphs[para_index]._element)

    def apply_manual_format_to_para(self, para, para_index):
        """手动格式应用（只修改用户填写的项目）"""
        # 字体设置
        if self.check_bold.isChecked() or self.check_italic.isChecked() or self.check_color.isChecked() or True:
            font_name = self.combo_font.currentFont().family()
            font_size = float(self.combo_font_size.currentText())
            for run in para.runs:
                run.font.name = font_name
                if run._element.rPr is not None:
                    rFonts = run._element.rPr.find(qn('w:rFonts'))
                    if rFonts is None:
                        rFonts = OxmlElement('w:rFonts')
                        run._element.rPr.append(rFonts)
                    rFonts.set(qn('w:eastAsia'), font_name)
                run.font.size = Pt(font_size)
                if self.check_bold.isChecked():
                    run.font.bold = True
                if self.check_italic.isChecked():
                    run.font.italic = True
                if self.check_color.isChecked() and self.color_value:
                    run.font.color.rgb = RGBColor(self.color_value.red(), self.color_value.green(), self.color_value.blue())

        pf = para.paragraph_format
        # 对齐
        align_map = {"左对齐": WD_ALIGN_PARAGRAPH.LEFT, "居中": WD_ALIGN_PARAGRAPH.CENTER,
                     "右对齐": WD_ALIGN_PARAGRAPH.RIGHT, "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY}
        pf.alignment = align_map.get(self.combo_alignment.currentText(), WD_ALIGN_PARAGRAPH.LEFT)
        # 间距
        pf.space_before = Pt(self.spin_space_before.value())
        pf.space_after = Pt(self.spin_space_after.value())
        # 行距
        ls_map = {"单倍行距": 1.0, "1.5倍行距": 1.5, "2倍行距": 2.0}
        ls_text = self.combo_line_spacing.currentText()
        if ls_text in ls_map:
            pf.line_spacing = ls_map[ls_text]
        # 缩进
        pf.first_line_indent = Pt(self.spin_first_indent.value())
        pf.left_indent = Pt(self.spin_left_indent.value())
        # 前后空行
        self.adjust_empty_paragraphs(para_index, self.spin_empty_before.value(), self.spin_empty_after.value())

    def unify_format_and_number(self):
        if not self.doc or not self.headings:
            QMessageBox.warning(self, "警告", "请先打开文档并扫描标题")
            return

        levels_to_unify = self.get_selected_levels()
        if not levels_to_unify:
            QMessageBox.warning(self, "警告", "请选择至少一个层级")
            return

        format_source = self.radio_format_source.currentIndex()

        if format_source == 0:  # 从选定标题复制
            selected_items = self.list_headings.selectedItems()
            if len(selected_items) != 1:
                QMessageBox.warning(self, "警告", "请选择一个“完美格式”标题")
                return
            source_idx = selected_items[0].data(Qt.UserRole)
            source_para_idx, source_level, _ = self.headings[source_idx]
            source_para = self.doc.paragraphs[source_para_idx]
            source_format = self.get_paragraph_full_format(source_para, source_para_idx)
            # 应用到同层级的其他标题
            for i, (para_idx, level, text) in enumerate(self.headings):
                if level in levels_to_unify:
                    para = self.doc.paragraphs[para_idx]
                    self.apply_paragraph_full_format(para, para_idx, source_format)
        else:  # 手动设置
            for i, (para_idx, level, text) in enumerate(self.headings):
                if level in levels_to_unify:
                    para = self.doc.paragraphs[para_idx]
                    self.apply_manual_format_to_para(para, para_idx)

        # 自动编号
        number_style = self.combo_number_style.currentIndex()
        self.auto_number_headings(number_style)

        # 保存
        new_path = self.get_new_save_path()
        if new_path:
            self.doc.save(new_path)
            QMessageBox.information(self, "完成", f"格式统一与自动编号完成，新文件保存为：{new_path}")
            # 重新加载文档以更新索引
            self.doc = Document(new_path)
            self.scan_headings()

    def auto_number_headings(self, number_style):
        if number_style == 0:
            self.number_by_digit()
        elif number_style == 1:
            self.number_by_chinese()

    def number_by_digit(self):
        counters = [0, 0, 0]
        for i, (para_idx, level, text) in enumerate(self.headings):
            counters[level-1] += 1
            for j in range(level, 3):
                counters[j] = 0
            if level == 1:
                new_num = str(counters[0])
            elif level == 2:
                new_num = f"{counters[0]}.{counters[1]}"
            else:
                new_num = f"{counters[0]}.{counters[1]}.{counters[2]}"
            new_text = f"{new_num} {self.remove_old_number(text)}"
            para = self.doc.paragraphs[para_idx]
            self.set_paragraph_text(para, new_text)
            self.headings[i] = (para_idx, level, new_text)

    def number_by_chinese(self):
        l1 = l2 = l3 = 0
        for i, (para_idx, level, text) in enumerate(self.headings):
            if level == 1:
                l1 += 1
                l2 = l3 = 0
                new_num = self.to_chinese_number(l1) + "、"
            elif level == 2:
                l2 += 1
                l3 = 0
                new_num = "（" + self.to_chinese_number(l2) + "）"
            else:
                l3 += 1
                new_num = str(l3) + "）"
            new_text = f"{new_num}{self.remove_old_number(text)}"
            para = self.doc.paragraphs[para_idx]
            self.set_paragraph_text(para, new_text)
            self.headings[i] = (para_idx, level, new_text)

    def to_chinese_number(self, num):
        digits = ['零','一','二','三','四','五','六','七','八','九']
        if num <= 10:
            return digits[num] if num < 10 else '十'
        elif num < 20:
            return '十' + digits[num%10] if num%10 else '十'
        elif num < 100:
            return digits[num//10] + '十' + (digits[num%10] if num%10 else '')
        return str(num)

    def remove_old_number(self, text):
        text = re.sub(r'^(\d+(\.\d+)*)\s+', '', text)
        text = re.sub(r'^[一二三四五六七八九十百千]+、', '', text)
        text = re.sub(r'^（[一二三四五六七八九十百千]+）', '', text)
        text = re.sub(r'^\d+）', '', text)
        text = re.sub(r'^[①-⑩]', '', text)
        return text.strip()

    def set_paragraph_text(self, para, new_text):
        if para.runs:
            para.runs[0].text = new_text
            for run in para.runs[1:]:
                run.text = ""
        else:
            para.add_run(new_text)

    # ================== 目录生成 ==================
    def insert_toc(self):
        if not self.doc:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return
        levels = self.spin_toc_levels.value()
        first_para = self.doc.paragraphs[0]
        toc_para = first_para.insert_paragraph_before()
        add_toc(toc_para, levels)
        title_para = toc_para.insert_paragraph_before()
        title_run = title_para.add_run("目录")
        title_run.bold = True
        title_run.font.size = Pt(16)
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        new_path = self.get_new_save_path()
        if new_path:
            self.doc.save(new_path)
            QMessageBox.information(self, "完成", f"目录已插入，新文件保存为：{new_path}\n请在 Word 中更新域显示页码")

    # ================== 表格填充 ==================
    def list_tables(self):
        if not self.doc:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return
        self.combo_tables.clear()
        for i, table in enumerate(self.doc.tables):
            self.combo_tables.addItem(f"{i+1}: {get_table_name(table, self.doc, i)}", i)
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
            label = f"{col_idx+1} - {headers[col_idx]}" if col_idx < len(headers) and headers[col_idx] else f"列 {col_idx+1}"
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
        levels = []
        if self.check_fill_l1.isChecked():
            levels.append(1)
        if self.check_fill_l2.isChecked():
            levels.append(2)
        if self.check_fill_l3.isChecked():
            levels.append(3)
        titles = [(level, text) for _, level, text in self.headings if level in levels]
        if not titles:
            QMessageBox.warning(self, "警告", "没有符合层级的标题")
            return
        table = self.doc.tables[table_idx]
        needed = len(titles)
        current = len(table.rows) - 1
        if current < needed:
            for _ in range(needed - current):
                table.add_row()
        for i, (level, text) in enumerate(titles):
            row = table.rows[i+1]
            row.cells[seq_col].text = self.extract_seq_text(text, i+1)
            row.cells[title_col].text = text
        new_path = self.get_new_save_path()
        if new_path:
            self.doc.save(new_path)
            QMessageBox.information(self, "完成", f"表格填充完成，新文件保存为：{new_path}")

    def extract_seq_text(self, text, fallback):
        match = re.match(r'^(\d+(\.\d+)*)\s+', text)
        if match:
            return match.group(1)
        match = re.match(r'^([一二三四五六七八九十百千]+、)', text)
        if match:
            return match.group(1)
        match = re.match(r'^(（[一二三四五六七八九十百千]+）)', text)
        if match:
            return match.group(1)
        match = re.match(r'^(\d+）)', text)
        if match:
            return match.group(1)
        match = re.match(r'^([①-⑩])', text)
        if match:
            return match.group(1)
        return str(fallback)

    # ================== Excel 导入导出 ==================
    def export_to_excel(self):
        if not self.doc or not self.doc.tables:
            QMessageBox.warning(self, "警告", "文档中没有表格")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("选择要导出的表格")
        layout = QVBoxLayout(dialog)
        list_widget = QListWidget()
        list_widget.setSelectionMode(QListWidget.MultiSelection)
        for i, table in enumerate(self.doc.tables):
            item = QListWidgetItem(f"{i+1}: {get_table_name(table, self.doc, i)}")
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
        save_path, _ = QFileDialog.getSaveFileName(self, "保存 Excel", "", "Excel 文件 (*.xlsx)")
        if not save_path:
            return
        wb = Workbook()
        wb.remove(wb.active)
        for idx in selected:
            table = self.doc.tables[idx]
            ws = wb.create_sheet(title=f"表格{idx+1}")
            for r, row in enumerate(table.rows):
                for c, cell in enumerate(row.cells):
                    ws.cell(row=r+1, column=c+1, value=cell.text)
        wb.save(save_path)
        QMessageBox.information(self, "完成", f"已导出到 {save_path}")

    def import_from_excel(self):
        if not self.doc or not self.doc.tables:
            QMessageBox.warning(self, "警告", "文档中没有表格")
            return
        excel_path, _ = QFileDialog.getOpenFileName(self, "选择 Excel", "", "Excel 文件 (*.xlsx)")
        if not excel_path:
            return
        wb = load_workbook(excel_path, data_only=True)
        if not wb.sheetnames:
            QMessageBox.warning(self, "警告", "Excel 为空")
            return
        ws = wb[wb.sheetnames[0]]
        headers = [cell.value for cell in ws[1]]
        data_rows = list(ws.iter_rows(min_row=2, values_only=True))
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
            for c in range(len(table.columns)):
                header = table.rows[0].cells[c].text.strip() if table.rows else f"列{c+1}"
                combo_word_col.addItem(f"{c+1} - {header}", c)
        combo_table.currentIndexChanged.connect(lambda: self.update_import_cols(combo_table, combo_word_col))
        layout.addRow("Word 列：", combo_word_col)
        combo_excel_col = QComboBox()
        for i, h in enumerate(headers):
            combo_excel_col.addItem(f"{i+1} - {h if h else ''}", i)
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
        current = len(table.rows) - 1
        if current < needed:
            for _ in range(needed - current):
                table.add_row()
        for i, row_data in enumerate(data_rows):
            value = row_data[excel_col]
            table.rows[i+1].cells[word_col].text = str(value) if value is not None else ""
        new_path = self.get_new_save_path()
        if new_path:
            self.doc.save(new_path)
            QMessageBox.information(self, "完成", f"导入完成，新文件保存为：{new_path}")

    def update_import_cols(self, combo_table, combo_word_col):
        combo_word_col.clear()
        table_idx = combo_table.currentData()
        if table_idx is None:
            return
        table = self.doc.tables[table_idx]
        for c in range(len(table.columns)):
            header = table.rows[0].cells[c].text.strip() if table.rows else f"列{c+1}"
            combo_word_col.addItem(f"{c+1} - {header}", c)

# ================== 程序入口 ==================
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

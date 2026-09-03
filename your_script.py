import sys
import re
import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from openpyxl import Workbook, load_workbook
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QListWidget, QListWidgetItem, QLabel,
    QComboBox, QSpinBox, QCheckBox, QTableWidget, QTableWidgetItem,
    QTabWidget, QLineEdit, QMessageBox, QProgressBar, QTextEdit,
    QSplitter, QGroupBox, QFormLayout, QDialog, QDialogButtonBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# ================== 辅助函数 ==================

def is_heading_paragraph(para):
    """判断段落是否为标题，返回 (是否, 层级)"""
    text = para.text.strip()
    if not text:
        return False, 0

    # 正则匹配编号：1, 1.1, 1.1.1, 第一章, 第1章 等
    patterns = [
        r'^(\d+(\.\d+)*)\s+',               # 1, 1.1, 1.1.1
        r'^第[一二三四五六七八九十百千\d]+章\s*',  # 第一章
    ]
    for i, pat in enumerate(patterns):
        if re.match(pat, text):
            # 根据编号层级判断：数点数量+1
            if i == 0:  # 数字编号
                parts = text.split()[0].split('.')
                level = len(parts)
                # 最多三级
                if level > 3:
                    level = 3
                return True, level
            else:       # 第X章，视为一级
                return True, 1

    # 辅助：如果段落有加粗且字号较大，也可能为标题（但这里以编号为主，此处仅作参考）
    # 可以忽略，因为用户强调以编号为主
    return False, 0

def set_heading_style(paragraph, level):
    """将段落设置为内置标题样式"""
    style_name = f'Heading {level}'
    paragraph.style = paragraph.part.document.styles[style_name]
    # 可选：保留原有字体格式
    # for run in paragraph.runs:
    #     run.font.name = 'Calibri'
    #     run.font.size = Pt(14 - level)  # 简单调整

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
    """尝试获取表格名称：查找表格上方最近的段落中包含 '表' 或 'Table' 的文字"""
    # 通过遍历文档元素，找到表格之前的段落
    # 简化处理：返回 "表格 {index+1}"
    return f"表格 {index+1}"

# ================== 主窗口 ==================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word 智能标题与表格处理工具")
        self.setGeometry(100, 100, 900, 700)

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
        layout.addWidget(QLabel("识别到的标题（可多选，用于应用样式）："))
        layout.addWidget(self.list_headings)

        # 目录设置
        group_dir = QGroupBox("生成目录")
        dir_layout = QFormLayout()
        self.combo_dir_pos = QComboBox()
        self.combo_dir_pos.addItems(["文档开头", "指定页码", "替换现有目录"])
        self.spin_page = QSpinBox()
        self.spin_page.setRange(1, 1000)
        self.spin_page.setValue(1)
        self.spin_page.setEnabled(False)
        self.spin_levels = QSpinBox()
        self.spin_levels.setRange(1, 3)
        self.spin_levels.setValue(3)
        self.combo_dir_pos.currentIndexChanged.connect(
            lambda idx: self.spin_page.setEnabled(idx == 1)
        )
        dir_layout.addRow("插入位置：", self.combo_dir_pos)
        dir_layout.addRow("页码：", self.spin_page)
        dir_layout.addRow("显示级别：", self.spin_levels)
        self.btn_insert_toc = QPushButton("插入目录")
        self.btn_insert_toc.clicked.connect(self.insert_toc)
        dir_layout.addRow(self.btn_insert_toc)
        group_dir.setLayout(dir_layout)
        layout.addWidget(group_dir)

        # 应用样式按钮
        self.btn_apply_style = QPushButton("将选中标题应用为 Word 标题样式")
        self.btn_apply_style.clicked.connect(self.apply_heading_styles)
        layout.addWidget(self.btn_apply_style)

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
                item.setData(Qt.UserRole, len(self.headings)-1)  # 存储索引
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
        self.doc.save(self.current_doc_path)
        QMessageBox.information(self, "完成", "已应用标题样式并保存文档")

    def insert_toc(self):
        if not self.doc:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return
        pos = self.combo_dir_pos.currentIndex()
        levels = self.spin_levels.value()

        if pos == 0:  # 文档开头
            # 在第一个段落前插入
            first_para = self.doc.paragraphs[0]
            new_para = first_para.insert_paragraph_before()
            add_toc(new_para, levels)
        elif pos == 1:  # 指定页码
            page = self.spin_page.value()
            # 简化：无法准确获取页码，使用段落位置近似
            # 这里提示用户实际页码可能不准确
            QMessageBox.information(self, "提示", "指定页码功能为近似实现，将在文档前部插入目录，请手动调整位置。")
            first_para = self.doc.paragraphs[0]
            new_para = first_para.insert_paragraph_before()
            add_toc(new_para, levels)
        else:  # 替换现有目录
            # 查找现有目录域并替换（简化：删除第一个TOC域，插入新目录）
            # 此处简单实现：在第一个段落前插入新的目录，并提示用户删除旧目录
            QMessageBox.information(self, "提示", "请手动删除旧目录后重新插入。")
            first_para = self.doc.paragraphs[0]
            new_para = first_para.insert_paragraph_before()
            add_toc(new_para, levels)

        self.doc.save(self.current_doc_path)
        QMessageBox.information(self, "完成", "目录已插入，请在 Word 中更新域以显示页码")

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
        # 填充列选择下拉框
        self.combo_seq_col.clear()
        self.combo_title_col.clear()
        # 获取表头（第一行）
        headers = []
        if len(table.rows) > 0:
            for cell in table.rows[0].cells:
                headers.append(cell.text.strip())
        # 添加列号选项
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

        # 根据勾选层级收集标题
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
        # 确保表格有足够行：现有数据行数（除表头）与标题数量比较
        # 假设第一行是表头
        data_rows = len(table.rows) - 1
        needed_rows = len(titles_to_fill)
        if data_rows < needed_rows:
            # 添加行
            for _ in range(needed_rows - data_rows):
                table.add_row()
        # 填充
        for i, (level, text) in enumerate(titles_to_fill):
            row = table.rows[i + 1]  # 跳过表头
            # 序号列：写入完整编号（从文本中提取编号部分）
            # 标题文本中已经包含编号，可以直接用
            # 但若需区分层级，可以增加缩进，这里直接使用完整文本
            # 我们写入序号列：提取文本开头的编号
            match = re.match(r'^(\d+(\.\d+)*)\s+', text)
            if match:
                seq_text = match.group(1)
            else:
                # 没有编号则用序号
                seq_text = str(i+1)
            row.cells[seq_col].text = seq_text
            # 标题列：写入完整标题（含编号）
            row.cells[title_col].text = text

        self.doc.save(self.current_doc_path)
        QMessageBox.information(self, "完成", "表格填充完成并已保存")

    # ================== Excel 导出导入 ==================
    def export_to_excel(self):
        if not self.doc:
            QMessageBox.warning(self, "警告", "请先打开 Word 文档")
            return
        if not self.doc.tables:
            QMessageBox.warning(self, "警告", "文档中没有表格")
            return
        # 弹出对话框选择要导出的表格（多选）
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
        # 选择保存路径
        save_path, _ = QFileDialog.getSaveFileName(self, "保存 Excel 文件", "", "Excel 文件 (*.xlsx)")
        if not save_path:
            return
        wb = Workbook()
        wb.remove(wb.active)  # 删除默认工作表
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
        # 选择 Excel 文件
        excel_path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx)")
        if not excel_path:
            return
        wb = load_workbook(excel_path, data_only=True)
        sheet_names = wb.sheetnames
        if not sheet_names:
            QMessageBox.warning(self, "警告", "Excel 文件为空")
            return
        # 简单处理：选择第一个工作表
        ws = wb[sheet_names[0]]
        # 获取数据（第一行为表头）
        headers = [cell.value for cell in ws[1]]
        data_rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            data_rows.append(row)
        if not data_rows:
            QMessageBox.warning(self, "警告", "Excel 没有数据")
            return

        # 让用户选择目标 Word 表格和对应列
        dialog = QDialog(self)
        dialog.setWindowTitle("选择目标表格和列映射")
        layout = QFormLayout(dialog)
        combo_table = QComboBox()
        for i, table in enumerate(self.doc.tables):
            combo_table.addItem(f"{i+1}: {get_table_name(table, self.doc, i)}", i)
        layout.addRow("目标表格：", combo_table)

        combo_word_col = QComboBox()
        # 获取表头
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
        # 确保行数足够
        needed = len(data_rows)
        current_data_rows = len(table.rows) - 1
        if current_data_rows < needed:
            for _ in range(needed - current_data_rows):
                table.add_row()
        # 填充
        for i, row_data in enumerate(data_rows):
            value = row_data[excel_col]
            table.rows[i+1].cells[word_col].text = str(value) if value is not None else ""
        self.doc.save(self.current_doc_path)
        QMessageBox.information(self, "完成", "导入完成并已保存")

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

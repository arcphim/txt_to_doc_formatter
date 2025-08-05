import sys
import os
import json
import glob
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QListWidget, QTabWidget,
    QGroupBox, QFormLayout, QDoubleSpinBox, QSpinBox, QFontComboBox,
    QCheckBox, QMessageBox, QProgressBar, QSplitter, QFrame, QLineEdit,
    QComboBox, QDialog, QDialogButtonBox, QListWidgetItem, QAbstractItemView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPalette, QColor, QDragEnterEvent, QDropEvent
from PyQt5.QtCore import QMimeData, QUrl
from doc_formatter import DocFormatter
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from txt_to_word import txt_to_word

class DragDropListWidget(QListWidget):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DropOnly)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            files = []
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                    # 检查文件扩展名是否支持
                    if file_path.lower().endswith(('.docx', '.doc', '.txt', '.md')):
                        files.append(file_path)
            
            if files:
                # 获取父窗口的引用
                main_window = self.window()
                if hasattr(main_window, 'add_files_to_list'):
                    main_window.add_files_to_list(files)
            
            event.acceptProposedAction()
    
    def dragMoveEvent(self, event):
        event.acceptProposedAction()

class FormatThread(QThread):
    progress_updated = pyqtSignal(int)
    completed = pyqtSignal(str, list)  # 输出目录, 生成的文件列表
    error_occurred = pyqtSignal(str)
    
    def __init__(self, input_paths, output_dir, config_path):
        super().__init__()
        self.input_paths = input_paths
        self.output_dir = output_dir
        self.config_path = config_path
        self.formatter = DocFormatter(config_path)
        
    def process_txt_md_file(self, input_path, output_path):
        """处理txt/md文件"""
        print(f"在process_txt_md_file中处理文件: {input_path} -> {output_path}")
        # 调用txt_to_word模块的函数
        txt_to_word(input_path, output_path, self.config_path)
        print(f"文件处理完成: {input_path} -> {output_path}")
        
    def run(self):
        try:
            print(f"开始处理文件列表: {self.input_paths}")
            total = len(self.input_paths)
            print(f"总文件数: {total}")
            # 跟踪当前生成的文件
            generated_files = []
            
            for i, input_path in enumerate(self.input_paths):
                print(f"处理第{i+1}个文件: {input_path}")
                # 更新进度
                progress = int((i + 1) / total * 100)
                self.progress_updated.emit(progress)
                
                # 获取文件名并生成输出路径
                filename = os.path.basename(input_path)
                print(f"文件名: {filename}")
                # 避免重复添加formatted_前缀
                if filename.startswith("formatted_"):
                    output_filename = filename
                else:
                    # 对于txt/md文件，输出为docx格式
                    name, ext = os.path.splitext(filename)
                    if ext.lower() in ['.txt', '.md']:
                        output_filename = f"formatted_{name}.docx"
                    else:
                        output_filename = f"formatted_{filename}"
                
                # 检查是否存在同名文件，如果存在则添加序号
                output_path = os.path.join(self.output_dir, output_filename)
                print(f"输出路径: {output_path}")
                base_name, ext = os.path.splitext(output_filename)
                counter = 1
                while os.path.exists(output_path):
                    output_filename = f"{base_name}_{counter}{ext}"
                    output_path = os.path.join(self.output_dir, output_filename)
                    counter += 1
                
                # 根据文件扩展名选择处理方式
                # 应该检查输入文件的扩展名，而不是输出文件的扩展名
                _, input_ext = os.path.splitext(input_path)
                print(f"输入文件扩展名: {input_ext}")
                if input_ext.lower() in ['.txt', '.md']:
                    # 处理txt/md文件
                    print("调用process_txt_md_file处理txt/md文件")
                    self.process_txt_md_file(input_path, output_path)
                else:
                    # 格式化Word文档
                    print("调用format_document处理Word文档")
                    self.formatter.format_document(input_path, output_path)
                
                # 将生成的文件添加到列表中
                generated_files.append(os.path.basename(output_path))
            
            print("所有文件处理完成，发送completed信号")
            self.completed.emit(self.output_dir, generated_files)
        except Exception as e:
            print(f"捕获到异常: {e}")
            import traceback
            traceback.print_exc()
            self.error_occurred.emit(str(e))

class ConfigEditor(QWidget):
    def __init__(self, config_path, parent=None):
        super().__init__(parent)
        self.config_path = config_path
        self.config = self.load_config()
        self.init_ui()
        
    def load_config(self):
        with open(self.config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
        
    def save_config(self):
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 创建标签页
        self.tabs = QTabWidget()
        
        # 基本设置标签页
        basic_tab = QWidget()
        basic_layout = QFormLayout()
        
        # 页边距设置
        margin_group = QGroupBox("页边距设置 (cm)")
        margin_layout = QFormLayout()
        
        self.top_margin = QDoubleSpinBox()
        self.top_margin.setRange(0, 10)
        self.top_margin.setSingleStep(0.1)
        self.top_margin.setValue(self.config['margins']['top'])
        margin_layout.addRow("上页边距:", self.top_margin)
        
        self.bottom_margin = QDoubleSpinBox()
        self.bottom_margin.setRange(0, 10)
        self.bottom_margin.setSingleStep(0.1)
        self.bottom_margin.setValue(self.config['margins']['bottom'])
        margin_layout.addRow("下页边距:", self.bottom_margin)
        
        self.left_margin = QDoubleSpinBox()
        self.left_margin.setRange(0, 10)
        self.left_margin.setSingleStep(0.1)
        self.left_margin.setValue(self.config['margins']['left'])
        margin_layout.addRow("左页边距:", self.left_margin)
        
        self.right_margin = QDoubleSpinBox()
        self.right_margin.setRange(0, 10)
        self.right_margin.setSingleStep(0.1)
        self.right_margin.setValue(self.config['margins']['right'])
        margin_layout.addRow("右页边距:", self.right_margin)
        
        margin_group.setLayout(margin_layout)
        basic_layout.addRow(margin_group)
        
        # 行距设置
        spacing_group = QGroupBox("行距设置")
        spacing_layout = QFormLayout()
        
        self.line_spacing = QSpinBox()
        self.line_spacing.setRange(10, 50)
        self.line_spacing.setValue(self.config['spacing']['line_spacing'])
        spacing_layout.addRow("固定值 (磅):", self.line_spacing)
        
        spacing_group.setLayout(spacing_layout)
        basic_layout.addRow(spacing_group)
        
        basic_tab.setLayout(basic_layout)
        self.tabs.addTab(basic_tab, "基本设置")
        
        # 字体设置标签页
        font_tab = QWidget()
        font_layout = QFormLayout()
        
        # 标题字体
        title_font_group = QGroupBox("标题字体")
        title_font_layout = QFormLayout()
        
        self.title_font = QFontComboBox()
        self.title_font.setCurrentFont(QFont(self.config['title_font']['name']))
        title_font_layout.addRow("字体:", self.title_font)
        
        self.title_size = QSpinBox()
        self.title_size.setRange(6, 72)
        self.title_size.setValue(self.config['title_font']['size'])
        title_font_layout.addRow("字号 (磅):", self.title_size)
        
        title_font_group.setLayout(title_font_layout)
        font_layout.addRow(title_font_group)
        
        # 正文字体
        body_font_group = QGroupBox("正文字体")
        body_font_layout = QFormLayout()

        self.body_font = QFontComboBox()
        self.body_font.setCurrentFont(QFont(self.config['body_font']['name']))
        body_font_layout.addRow("中文字体:", self.body_font)

        self.digit_font = QFontComboBox()
        self.digit_font.setCurrentFont(QFont(self.config['body_font']['digit_font']))
        body_font_layout.addRow("数字字体:", self.digit_font)

        self.body_size = QSpinBox()
        self.body_size.setRange(6, 72)
        self.body_size.setValue(self.config['body_font']['size'])
        body_font_layout.addRow("字号 (磅):", self.body_size)

        body_font_group.setLayout(body_font_layout)
        font_layout.addRow(body_font_group)

        # 标题层级字体
        heading_font_group = QGroupBox("标题层级字体")
        heading_font_layout = QFormLayout()

        self.heading1_font = QFontComboBox()
        self.heading1_font.setCurrentFont(QFont(self.config['heading_levels'][0]['font']))
        heading_font_layout.addRow("一级标题字体:", self.heading1_font)

        self.heading1_size = QSpinBox()
        self.heading1_size.setRange(6, 72)
        self.heading1_size.setValue(self.config['heading_levels'][0]['size'])
        heading_font_layout.addRow("一级标题字号 (磅):", self.heading1_size)

        self.heading2_font = QFontComboBox()
        self.heading2_font.setCurrentFont(QFont(self.config['heading_levels'][1]['font']))
        heading_font_layout.addRow("二级标题字体:", self.heading2_font)

        self.heading2_size = QSpinBox()
        self.heading2_size.setRange(6, 72)
        self.heading2_size.setValue(self.config['heading_levels'][1]['size'])
        heading_font_layout.addRow("二级标题字号 (磅):", self.heading2_size)

        self.heading3_font = QFontComboBox()
        self.heading3_font.setCurrentFont(QFont(self.config['heading_levels'][2]['font']))
        heading_font_layout.addRow("三级标题字体:", self.heading3_font)

        self.heading3_size = QSpinBox()
        self.heading3_size.setRange(6, 72)
        self.heading3_size.setValue(self.config['heading_levels'][2]['size'])
        heading_font_layout.addRow("三级标题字号 (磅):", self.heading3_size)

        self.heading4_font = QFontComboBox()
        self.heading4_font.setCurrentFont(QFont(self.config['heading_levels'][3]['font']))
        heading_font_layout.addRow("四级标题字体:", self.heading4_font)

        self.heading4_size = QSpinBox()
        self.heading4_size.setRange(6, 72)
        self.heading4_size.setValue(self.config['heading_levels'][3]['size'])
        heading_font_layout.addRow("四级标题字号 (磅):", self.heading4_size)

        heading_font_group.setLayout(heading_font_layout)
        font_layout.addRow(heading_font_group)

        font_tab.setLayout(font_layout)
        self.tabs.addTab(font_tab, "字体设置")
        
        # 按钮
        btn_layout = QHBoxLayout()
        
        self.save_btn = QPushButton("保存设置")
        self.save_btn.clicked.connect(self.save_settings)
        btn_layout.addWidget(self.save_btn)
        
        self.reset_btn = QPushButton("重置为默认值")
        self.reset_btn.clicked.connect(self.reset_settings)
        btn_layout.addWidget(self.reset_btn)
        
        # 添加到主布局
        layout.addWidget(self.tabs)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
        
    def save_settings(self):
        # 更新配置
        self.config['margins']['top'] = self.top_margin.value()
        self.config['margins']['bottom'] = self.bottom_margin.value()
        self.config['margins']['left'] = self.left_margin.value()
        self.config['margins']['right'] = self.right_margin.value()
        
        self.config['spacing']['line_spacing'] = self.line_spacing.value()
        
        self.config['title_font']['name'] = self.title_font.currentFont().family()
        self.config['title_font']['size'] = self.title_size.value()
        
        self.config['body_font']['name'] = self.body_font.currentFont().family()
        self.config['body_font']['digit_font'] = self.digit_font.currentFont().family()
        self.config['body_font']['size'] = self.body_size.value()

        # 保存标题层级字体设置
        self.config['heading_levels'][0]['font'] = self.heading1_font.currentFont().family()
        self.config['heading_levels'][0]['size'] = self.heading1_size.value()
        self.config['heading_levels'][1]['font'] = self.heading2_font.currentFont().family()
        self.config['heading_levels'][1]['size'] = self.heading2_size.value()
        self.config['heading_levels'][2]['font'] = self.heading3_font.currentFont().family()
        self.config['heading_levels'][2]['size'] = self.heading3_size.value()
        self.config['heading_levels'][3]['font'] = self.heading4_font.currentFont().family()
        self.config['heading_levels'][3]['size'] = self.heading4_size.value()
        
        # 保存到文件
        self.save_config()
        QMessageBox.information(self, "成功", "设置已保存")
        
    def reset_settings(self):
        # 重新加载默认配置
        self.config = self.load_config()
        
        # 重置界面控件
        self.top_margin.setValue(self.config['margins']['top'])
        self.bottom_margin.setValue(self.config['margins']['bottom'])
        self.left_margin.setValue(self.config['margins']['left'])
        self.right_margin.setValue(self.config['margins']['right'])
        
        self.line_spacing.setValue(self.config['spacing']['line_spacing'])
        
        self.title_font.setCurrentFont(QFont(self.config['title_font']['name']))
        self.title_size.setValue(self.config['title_font']['size'])
        
        self.body_font.setCurrentFont(QFont(self.config['body_font']['name']))
        self.digit_font.setCurrentFont(QFont(self.config['body_font']['digit_font']))
        self.body_size.setValue(self.config['body_font']['size'])

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word自动排版工具")
        self.setGeometry(100, 100, 900, 600)
        
        # 配置文件路径
        self.config_path = 'config.json'
        
        # 加载配置
        self.config = self.load_config()
        
        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        
        # 创建菜单栏
        self.create_menu_bar()
        
        # 创建标签页
        self.tabs = QTabWidget()
        
        # 文档处理标签页
        self.process_tab = QWidget()
        self.init_process_tab()
        self.tabs.addTab(self.process_tab, "文档处理")
        
        # 设置标签页
        self.config_editor = ConfigEditor(self.config_path)
        self.tabs.addTab(self.config_editor, "排版设置")
        
        # 状态栏
        self.statusBar().showMessage("就绪")
        
        # 添加到主布局
        main_layout.addWidget(self.tabs)
        
        # 初始化输出目录下拉框
        self.init_output_dirs()
        
    def create_menu_bar(self):
        menubar = self.menuBar()
        
        # 文件菜单
        file_menu = menubar.addMenu('文件')
        
        exit_action = file_menu.addAction('退出')
        exit_action.triggered.connect(self.close)
        
        # 帮助菜单
        help_menu = menubar.addMenu('帮助')
        
        about_action = help_menu.addAction('关于')
        about_action.triggered.connect(self.show_about)
        
    def init_process_tab(self):
        layout = QVBoxLayout(self.process_tab)
        
        # 文件列表区域
        file_list_group = QGroupBox("待处理文件")
        file_list_layout = QVBoxLayout()
        
        self.file_list = DragDropListWidget()
        file_list_layout.addWidget(self.file_list)
        
        # 文件操作按钮
        file_btn_layout = QHBoxLayout()
        
        self.add_btn = QPushButton("添加文件")
        self.add_btn.clicked.connect(self.add_files)
        file_btn_layout.addWidget(self.add_btn)
        
        self.add_dir_btn = QPushButton("添加文件夹")
        self.add_dir_btn.clicked.connect(self.add_directory)
        file_btn_layout.addWidget(self.add_dir_btn)
        
        self.remove_btn = QPushButton("移除选中")
        self.remove_btn.clicked.connect(self.remove_files)
        file_btn_layout.addWidget(self.remove_btn)
        
        self.clear_btn = QPushButton("清空列表")
        self.clear_btn.clicked.connect(self.clear_files)
        file_btn_layout.addWidget(self.clear_btn)
        
        file_list_layout.addLayout(file_btn_layout)
        file_list_group.setLayout(file_list_layout)
        
        # 输出目录
        output_dir_group = QGroupBox("输出设置")
        output_dir_layout = QHBoxLayout()
        
        self.output_dir_combo = QComboBox()
        self.output_dir_combo.setEditable(True)
        self.output_dir_combo.setInsertPolicy(QComboBox.NoInsert)
        output_dir_layout.addWidget(self.output_dir_combo)
        
        self.browse_btn = QPushButton("浏览...")
        self.browse_btn.clicked.connect(self.browse_output_dir)
        output_dir_layout.addWidget(self.browse_btn)
        
        self.manage_dirs_btn = QPushButton("管理默认目录...")
        self.manage_dirs_btn.clicked.connect(self.manage_default_dirs)
        output_dir_layout.addWidget(self.manage_dirs_btn)
        
        output_dir_group.setLayout(output_dir_layout)
        
        # 处理按钮
        process_btn_layout = QHBoxLayout()
        
        self.process_btn = QPushButton("开始排版")
        self.process_btn.clicked.connect(self.process_files)
        process_btn_layout.addWidget(self.process_btn)
        
        self.preview_btn = QPushButton("预览效果")
        self.preview_btn.clicked.connect(self.preview_file)
        self.preview_btn.setEnabled(False)
        process_btn_layout.addWidget(self.preview_btn)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        
        # 添加到布局
        layout.addWidget(file_list_group)
        layout.addWidget(output_dir_group)
        layout.addLayout(process_btn_layout)
        layout.addWidget(self.progress_bar)
        
    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", "", "支持的文件 (*.docx *.doc *.txt *.md);;Word文档 (*.docx *.doc);;文本文件 (*.txt);;Markdown文件 (*.md)"
        )
        
        if files:
            self.add_files_to_list(files)
    
    def add_files_to_list(self, files):
        # 记录添加前的文件数量
        previous_count = self.file_list.count()
        
        for file in files:
            if self.file_list.findItems(file, Qt.MatchExactly):
                continue  # 跳过重复文件
            self.file_list.addItem(file)
        
        # 如果是添加单个文件，则自动选中该文件
        if len(files) == 1 and self.file_list.count() > previous_count:
            # 找到新添加的文件项并选中
            for i in range(self.file_list.count()):
                item = self.file_list.item(i)
                if item.text() == files[0]:
                    self.file_list.setCurrentItem(item)
                    break
        
        self.preview_btn.setEnabled(self.file_list.count() > 0)
        
    def add_directory(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        
        if dir_path:
            # 查找所有支持的文件
            supported_files = []
            for ext in ['docx', 'doc', 'txt', 'md']:
                supported_files.extend(glob.glob(os.path.join(dir_path, f"*.{ext}")))
                supported_files.extend(glob.glob(os.path.join(dir_path, f"**/*.{ext}"), recursive=True))
            
            if supported_files:
                for file in supported_files:
                    if not self.file_list.findItems(file, Qt.MatchExactly):
                        self.file_list.addItem(file)
                
                self.preview_btn.setEnabled(self.file_list.count() > 0)
            else:
                QMessageBox.information(self, "提示", "所选文件夹中没有支持的文件")
        
    def remove_files(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))
            
        self.preview_btn.setEnabled(self.file_list.count() > 0)
        
    def clear_files(self):
        self.file_list.clear()
        self.preview_btn.setEnabled(False)
        
    def browse_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if dir_path:
            # 检查目录是否已存在
            existing_dirs = [self.output_dir_combo.itemText(i) for i in range(self.output_dir_combo.count())]
            if dir_path not in existing_dirs:
                self.output_dir_combo.addItem(dir_path)
            self.output_dir_combo.setCurrentText(dir_path)
            
            # 更新配置文件中的当前目录
            self.update_current_dir_in_config(dir_path)
        
    def load_config(self):
        with open(self.config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def save_config(self):
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)
    
    def init_output_dirs(self):
        # 从配置文件加载默认目录
        default_dirs = self.config.get('output_dirs', {}).get('default_dirs', [])
        current_dir = self.config.get('output_dirs', {}).get('current_dir', '')
        
        # 添加默认目录到下拉框
        for dir_path in default_dirs:
            self.output_dir_combo.addItem(dir_path)
        
        # 设置当前目录
        if current_dir:
            # 检查当前目录是否已在下拉框中
            existing_dirs = [self.output_dir_combo.itemText(i) for i in range(self.output_dir_combo.count())]
            if current_dir not in existing_dirs:
                self.output_dir_combo.addItem(current_dir)
            self.output_dir_combo.setCurrentText(current_dir)
    
    def update_current_dir_in_config(self, dir_path):
        # 更新配置中的当前目录
        if 'output_dirs' not in self.config:
            self.config['output_dirs'] = {}
        self.config['output_dirs']['current_dir'] = dir_path
        self.save_config()
    
    def manage_default_dirs(self):
        # 创建管理默认目录对话框
        dialog = DefaultDirsDialog(self.config, self)
        if dialog.exec_() == QDialog.Accepted:
            # 更新配置
            self.config = dialog.get_updated_config()
            self.save_config()
            
            # 重新初始化输出目录下拉框
            self.output_dir_combo.clear()
            self.init_output_dirs()
    
    def process_files(self):
        # 检查是否有文件
        if self.file_list.count() == 0:
            QMessageBox.warning(self, "警告", "请先添加文件")
            return
        
        # 检查是否选择了输出目录
        output_dir = self.output_dir_combo.currentText()
        if not output_dir:
            QMessageBox.warning(self, "警告", "请选择输出目录")
            return
        
        # 获取文件列表
        file_paths = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        
        # 创建并启动处理线程
        self.process_thread = FormatThread(file_paths, output_dir, self.config_path)
        self.process_thread.progress_updated.connect(self.update_progress)
        self.process_thread.completed.connect(self.process_completed)
        self.process_thread.error_occurred.connect(self.process_error)
        
        # 禁用按钮
        self.process_btn.setEnabled(False)
        self.add_btn.setEnabled(False)
        self.add_dir_btn.setEnabled(False)
        self.remove_btn.setEnabled(False)
        self.clear_btn.setEnabled(False)
        self.preview_btn.setEnabled(False)
        
        # 开始处理
        self.process_thread.start()
        self.statusBar().showMessage("正在处理...")
        
    def update_progress(self, value):
        self.progress_bar.setValue(value)
        
    def process_completed(self, output_dir, generated_files):
        # 恢复按钮状态
        self.process_btn.setEnabled(True)
        self.add_btn.setEnabled(True)
        self.add_dir_btn.setEnabled(True)
        self.remove_btn.setEnabled(True)
        self.clear_btn.setEnabled(True)
        self.preview_btn.setEnabled(self.file_list.count() > 0)
        
        self.progress_bar.setValue(100)
        self.statusBar().showMessage("处理完成")
        
        # 构建文件列表信息
        if generated_files:
            file_list_text = "\n".join(generated_files[:10])  # 只显示前10个文件
            if len(generated_files) > 10:
                file_list_text += f"\n...及其他{len(generated_files) - 10}个文件"
        else:
            file_list_text = "未生成文件"
        
        QMessageBox.information(
            self, "成功", f"所有文件已处理完成\n输出目录: {output_dir}\n\n生成的文件:\n{file_list_text}"
        )
        
    def process_error(self, error_msg):
        # 恢复按钮状态
        self.process_btn.setEnabled(True)
        self.add_btn.setEnabled(True)
        self.add_dir_btn.setEnabled(True)
        self.remove_btn.setEnabled(True)
        self.clear_btn.setEnabled(True)
        self.preview_btn.setEnabled(self.file_list.count() > 0)
        
        self.statusBar().showMessage("处理出错")
        
        QMessageBox.critical(
            self, "错误", f"处理过程中出错:\n{error_msg}"
        )
        
    def preview_file(self):
        # 检查是否有选中文件
        if not self.file_list.currentItem():
            QMessageBox.warning(self, "警告", "请先选择一个文件")
            return
        
        # 检查是否选择了输出目录
        output_dir = self.output_dir_combo.currentText()
        if not output_dir:
            # 询问是否使用默认输出目录
            reply = QMessageBox.question(
                self, "输出目录", "未选择输出目录，是否使用临时目录?"
            )
            if reply != QMessageBox.Yes:
                return
            output_dir = os.path.join(os.path.dirname(__file__), "preview")
            os.makedirs(output_dir, exist_ok=True)
        
        # 获取选中文件
        file_path = self.file_list.currentItem().text()
        
        # 创建并启动处理线程
        self.preview_thread = FormatThread([file_path], output_dir, self.config_path)
        self.preview_thread.completed.connect(self.preview_completed)
        self.preview_thread.error_occurred.connect(self.process_error)
        
        # 禁用按钮
        self.process_btn.setEnabled(False)
        self.add_btn.setEnabled(False)
        self.add_dir_btn.setEnabled(False)
        self.remove_btn.setEnabled(False)
        self.clear_btn.setEnabled(False)
        self.preview_btn.setEnabled(False)
        
        # 开始处理
        self.preview_thread.start()
        self.statusBar().showMessage("正在生成预览...")
        self.progress_bar.setValue(0)
        
    def preview_completed(self, output_dir, generated_files):
        # 恢复按钮状态
        self.process_btn.setEnabled(True)
        self.add_btn.setEnabled(True)
        self.add_dir_btn.setEnabled(True)
        self.remove_btn.setEnabled(True)
        self.clear_btn.setEnabled(True)
        self.preview_btn.setEnabled(self.file_list.count() > 0)
        
        self.progress_bar.setValue(100)
        self.statusBar().showMessage("预览生成完成")
        
        # 获取实际生成的文件路径
        if generated_files:
            # 取第一个生成的文件
            output_filename = generated_files[0]
            output_path = os.path.join(output_dir, output_filename)
        else:
            output_path = ""
        
        # 询问是否打开文件
        if output_path and os.path.exists(output_path):
            reply = QMessageBox.question(
                self, "预览完成", f"预览文件已生成:\n{output_path}\n是否打开文件?"
            )
            if reply == QMessageBox.Yes:
                os.startfile(output_path)
        else:
            QMessageBox.warning(self, "文件未找到", f"找不到预览文件:\n{output_path}")
        
    def show_about(self):
        QMessageBox.about(
            self, "关于 Word自动排版工具",
            "Word自动排版工具 v1.0\n\n用于批量格式化Word文档的工具\n根据预设模板统一文档格式"
        )
        
    def add_directory(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        
        if dir_path:
            # 查找所有支持的文件
            supported_files = []
            for ext in ['docx', 'doc', 'txt', 'md']:
                supported_files.extend(glob.glob(os.path.join(dir_path, f"*.{ext}")))
                supported_files.extend(glob.glob(os.path.join(dir_path, f"**/*.{ext}"), recursive=True))
            
            if supported_files:
                for file in supported_files:
                    if not self.file_list.findItems(file, Qt.MatchExactly):
                        self.file_list.addItem(file)
                
                self.preview_btn.setEnabled(self.file_list.count() > 0)
            else:
                QMessageBox.information(self, "提示", "所选文件夹中没有支持的文件")
        
    def remove_files(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))
            
        self.preview_btn.setEnabled(self.file_list.count() > 0)
        
    def clear_files(self):
        self.file_list.clear()
        self.preview_btn.setEnabled(False)
        
class DefaultDirsDialog(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config.copy()  # 复制配置以避免直接修改
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("管理默认输出目录")
        self.setGeometry(200, 200, 500, 400)
        
        layout = QVBoxLayout()
        
        # 默认目录列表
        dirs_group = QGroupBox("默认输出目录")
        dirs_layout = QVBoxLayout()
        
        self.dirs_list = QListWidget()
        self.dirs_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        dirs_layout.addWidget(self.dirs_list)
        
        # 按钮布局
        btn_layout = QHBoxLayout()
        
        self.add_btn = QPushButton("添加目录")
        self.add_btn.clicked.connect(self.add_directory)
        btn_layout.addWidget(self.add_btn)
        
        self.remove_btn = QPushButton("删除选中")
        self.remove_btn.clicked.connect(self.remove_directories)
        btn_layout.addWidget(self.remove_btn)
        
        self.up_btn = QPushButton("上移")
        self.up_btn.clicked.connect(self.move_up)
        btn_layout.addWidget(self.up_btn)
        
        self.down_btn = QPushButton("下移")
        self.down_btn.clicked.connect(self.move_down)
        btn_layout.addWidget(self.down_btn)
        
        dirs_layout.addLayout(btn_layout)
        dirs_group.setLayout(dirs_layout)
        
        # 对话框按钮
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        # 添加到主布局
        layout.addWidget(dirs_group)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
        
        # 初始化列表
        self.init_list()
        
    def init_list(self):
        default_dirs = self.config.get('output_dirs', {}).get('default_dirs', [])
        for dir_path in default_dirs:
            self.dirs_list.addItem(dir_path)
        
    def add_directory(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if dir_path:
            # 检查目录是否已存在
            existing_items = [self.dirs_list.item(i).text() for i in range(self.dirs_list.count())]
            if dir_path not in existing_items:
                self.dirs_list.addItem(dir_path)
            else:
                QMessageBox.information(self, "提示", "该目录已存在")
        
    def remove_directories(self):
        selected_items = self.dirs_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选择要删除的目录")
            return
            
        for item in selected_items:
            self.dirs_list.takeItem(self.dirs_list.row(item))
        
    def move_up(self):
        current_row = self.dirs_list.currentRow()
        if current_row > 0:
            current_item = self.dirs_list.takeItem(current_row)
            self.dirs_list.insertItem(current_row - 1, current_item)
            self.dirs_list.setCurrentRow(current_row - 1)
        
    def move_down(self):
        current_row = self.dirs_list.currentRow()
        if 0 <= current_row < self.dirs_list.count() - 1:
            current_item = self.dirs_list.takeItem(current_row)
            self.dirs_list.insertItem(current_row + 1, current_item)
            self.dirs_list.setCurrentRow(current_row + 1)
        
    def get_updated_config(self):
        # 更新配置中的默认目录列表
        if 'output_dirs' not in self.config:
            self.config['output_dirs'] = {}
        
        default_dirs = []
        for i in range(self.dirs_list.count()):
            default_dirs.append(self.dirs_list.item(i).text())
            
        self.config['output_dirs']['default_dirs'] = default_dirs
        return self.config

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # 设置全局字体
    font = QFont("SimHei", 9)
    app.setFont(font)
    
    # 设置样式
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(240, 240, 240))
    palette.setColor(QPalette.WindowText, Qt.black)
    app.setPalette(palette)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
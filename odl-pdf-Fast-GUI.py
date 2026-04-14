import sys
import os
import re
import subprocess
import threading
import time
import shutil
import tempfile
import markdown
import fitz  # PyMuPDF
import requests
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog,
                             QListWidget, QComboBox, QCheckBox, QGroupBox, QTabWidget,
                             QProgressBar, QSplitter, QMessageBox, QListWidgetItem, QFrame,
                             QSpinBox, QDoubleSpinBox, QScrollArea, QGridLayout, QSizePolicy, QFormLayout)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QUrl, QSize
from PyQt5.QtGui import QPixmap, QIcon, QFont, QTextCursor, QImage
from PyQt5.QtWebEngineWidgets import QWebEngineView

try:
    import GPUtil
    HAS_GPUtil = True
except ImportError:
    HAS_GPUtil = False

# ==========================================
# Worker Thread for Running ODL-PDF
# ==========================================
class ODLWorker(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)
    progress_signal = pyqtSignal(int)

    def __init__(self, command, log_lines=None):
        super().__init__()
        self.command = command
        self._is_running = True
        self.log_lines = log_lines if log_lines is not None else []

    def run(self):
        self.log_signal.emit(f"[系统] 正在执行命令: {' '.join(self.command)}")
        self.log_signal.emit("-" * 50)
        try:
            process = subprocess.Popen(
                self.command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=False,  # 使用二进制模式
                cwd=os.getcwd()
            )
            while True:
                if not self._is_running:
                    process.terminate()
                    self.log_signal.emit("[系统] 任务已被用户终止。")
                    break
                output = process.stdout.readline()
                if output == b'' and process.poll() is not None:
                    break
                if output:
                    # 尝试多种编码解码
                    try:
                        # 首先尝试UTF-8
                        output_str = output.decode('utf-8').strip()
                    except UnicodeDecodeError:
                        try:
                            # 然后尝试GBK
                            output_str = output.decode('gbk').strip()
                        except UnicodeDecodeError:
                            # 最后使用replace模式
                            output_str = output.decode('utf-8', errors='replace').strip()
                    self.log_signal.emit(output_str)
                    self.log_lines.append(output_str)
            rc = process.poll()
            if rc == 0:
                self.finished_signal.emit(True, "转换任务完成。")
            else:
                self.finished_signal.emit(False, f"转换任务失败，退出码: {rc}")
        except Exception as e:
            self.finished_signal.emit(False, f"发生异常: {str(e)}")

    def stop(self):
        self._is_running = False

# ==========================================
# Main GUI Window
# ==========================================
class ODLGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ODL-PDF-GUI")
        self.resize(1400, 900)
        self.input_files = []
        self.worker = None
        self.current_preview_file = None
        self.current_source_page = 0
        self.source_total_pages = 0
        self.current_result_page = 0
        self.result_total_pages = 0
        self.log_lines = []
        # 批量处理状态
        self.batch_temp_input_dir = None
        self.batch_temp_output_dir = None
        self.batch_file_map = {}  # 映射: temp_filename -> original_full_path

        self.init_ui()
        self.gpu_timer = QTimer()
        self.gpu_timer.timeout.connect(self.update_gpu_info)
        self.gpu_timer.start(1000)

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        left_panel = QVBoxLayout()
        top_container = QWidget()
        top_layout = QVBoxLayout(top_container)
        top_layout.setContentsMargins(0, 0, 0, 0)

        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout(file_group)
        btn_layout = QHBoxLayout()
        self.btn_add_files = QPushButton("添加文件")
        self.btn_add_folder = QPushButton("添加文件夹")
        self.btn_clear = QPushButton("清空列表")
        self.btn_clear.setStyleSheet("background-color: #f44336; color: white; font-weight: bold; height: 30px;")
        btn_layout.addWidget(self.btn_add_files)
        btn_layout.addWidget(self.btn_add_folder)
        btn_layout.addWidget(self.btn_clear)
        self.lbl_file_count = QLabel("未添加文件")
        self.lbl_file_count.setStyleSheet("color: #666; font-size: 11px;")
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.SingleSelection)
        self.file_list.itemClicked.connect(self.preview_source_file)
        file_layout.addLayout(btn_layout)
        file_layout.addWidget(self.lbl_file_count)
        file_layout.addWidget(self.file_list)

        output_layout = QHBoxLayout()
        self.label_output = QLabel("输出目录:")
        self.txt_output_dir = QLineEdit()
        self.txt_output_dir.setPlaceholderText("默认为源文件所在目录")
        self.btn_browse_output = QPushButton("浏览...")
        self.btn_source_folder = QPushButton("同源文件夹")
        self.btn_source_folder.clicked.connect(self.set_output_to_source_folder)
        output_layout.addWidget(self.label_output)
        output_layout.addWidget(self.txt_output_dir)
        output_layout.addWidget(self.btn_browse_output)
        output_layout.addWidget(self.btn_source_folder)
        file_layout.addLayout(output_layout)

        auto_save_layout = QHBoxLayout()
        self.chk_auto_save = QCheckBox("自动保存")
        self.chk_auto_save.setChecked(True)
        self.chk_auto_save.setToolTip("转换成功后自动保存文件到输出目录")
        auto_save_layout.addWidget(self.chk_auto_save)
        auto_save_layout.addStretch()
        file_layout.addLayout(auto_save_layout)

        top_layout.addWidget(file_group)

        settings_tabs = QTabWidget()

        # ==========================================
        # 基础设置选项卡
        # ==========================================
        tab_basic = QWidget()
        layout_basic = QFormLayout(tab_basic)

        # 页面范围 (通用)
        page_layout = QHBoxLayout()
        self.txt_pages = QLineEdit()
        self.txt_pages.setPlaceholderText("留空则处理全部，例如: 1-22, 25-65")
        self.btn_clear_pages = QPushButton("清空")
        self.btn_clear_pages.setFixedWidth(50)
        self.btn_clear_pages.clicked.connect(lambda: self.txt_pages.clear())
        page_layout.addWidget(self.txt_pages)
        page_layout.addWidget(self.btn_clear_pages)
        layout_basic.addRow("页面范围:", page_layout)

        # 输出格式 (通用)
        format_layout = QVBoxLayout()
        self.chk_format_json = QCheckBox("JSON")
        self.chk_format_md = QCheckBox("Markdown")
        self.chk_format_md.setChecked(True)
        self.chk_format_html = QCheckBox("HTML")
        self.chk_format_pdf = QCheckBox("Annotated PDF")
        self.chk_format_text = QCheckBox("Text")
        format_layout.addWidget(self.chk_format_json)
        format_layout.addWidget(self.chk_format_md)
        format_layout.addWidget(self.chk_format_html)
        format_layout.addWidget(self.chk_format_pdf)
        format_layout.addWidget(self.chk_format_text)
        layout_basic.addRow("输出格式:", format_layout)

        self.chk_use_struct_tree = QCheckBox("使用结构标签 (保留作者意图的精确布局)")
        self.chk_use_struct_tree.setChecked(False)  # 修改默认不勾选
        layout_basic.addRow("", self.chk_use_struct_tree)

        self.chk_ai_safety = QCheckBox("人工智能安全 (即时注入保护)")
        self.chk_ai_safety.setChecked(False)  # 修改默认不勾选
        layout_basic.addRow("", self.chk_ai_safety)

        settings_tabs.addTab(tab_basic, "基础设置")

        top_layout.addWidget(settings_tabs)
        left_panel.addWidget(top_container)

        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout(log_group)
        self.log_window = QTextEdit()
        self.log_window.setReadOnly(True)
        self.log_window.setStyleSheet("background-color: #1e1e1e; color: #d4d4d4; font-family: Consolas;")
        self.log_window.setFont(QFont("Consolas", 12))
        log_layout.addWidget(self.log_window)
        self.btn_clear_log = QPushButton("清空Log")
        self.btn_clear_log.clicked.connect(self.clear_log)
        log_layout.addWidget(self.btn_clear_log)
        left_panel.addWidget(log_group, 1)

        gpu_group = QGroupBox("硬件监控")
        gpu_layout = QVBoxLayout(gpu_group)
        self.lbl_gpu_info = QLabel("正在检测显卡...")
        self.lbl_gpu_info.setStyleSheet("font-family: Consolas; background-color: black; color: #00FF00; padding: 5px;")
        self.lbl_gpu_info.setFont(QFont("Consolas", 10))
        gpu_layout.addWidget(self.lbl_gpu_info)
        left_panel.addWidget(gpu_group)

        main_layout.addLayout(left_panel, 3)

        right_panel = QVBoxLayout()
        preview_splitter = QSplitter(Qt.Horizontal)
        preview_splitter.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        source_container = QWidget()
        source_layout = QVBoxLayout(source_container)
        source_layout.setContentsMargins(0, 0, 0, 0)
        source_label = QLabel("源文件预览")
        source_label.setAlignment(Qt.AlignCenter)
        source_label.setStyleSheet("font-weight: bold;")
        source_label.setFixedHeight(35)
        source_layout.addWidget(source_label)

        self.source_scroll = QScrollArea()
        self.source_scroll.setWidgetResizable(True)
        self.source_scroll.setAlignment(Qt.AlignCenter)
        self.source_scroll.setStyleSheet("border: 1px solid #ccc;")
        self.source_preview_container = QWidget()
        self.source_preview_layout = QVBoxLayout(self.source_preview_container)
        self.source_preview_layout.setContentsMargins(0, 0, 0, 0)
        self.source_preview_layout.setAlignment(Qt.AlignCenter)
        self.source_preview_label = QLabel("请选择文件以预览\n支持 PDF 和图片")
        self.source_preview_label.setAlignment(Qt.AlignCenter)
        self.source_preview_label.setStyleSheet("background-color: #333; color: white;")
        self.source_preview_layout.addWidget(self.source_preview_label)
        self.source_scroll.setWidget(self.source_preview_container)
        source_layout.addWidget(self.source_scroll, 1)

        source_nav_layout = QHBoxLayout()
        self.btn_source_prev = QPushButton("上一页")
        self.btn_source_prev.setFixedWidth(80)
        self.btn_source_prev.setEnabled(False)
        self.lbl_source_page = QLabel("0/0")
        self.lbl_source_page.setAlignment(Qt.AlignCenter)
        self.lbl_source_page.setFixedHeight(30)
        self.lbl_source_page.setStyleSheet("border: 1px solid #ccc; background-color: #f5f5f5;")
        self.btn_source_next = QPushButton("下一页")
        self.btn_source_next.setFixedWidth(80)
        self.btn_source_next.setEnabled(False)
        source_nav_layout.addWidget(self.btn_source_prev)
        source_nav_layout.addWidget(self.lbl_source_page)
        source_nav_layout.addWidget(self.btn_source_next)
        source_layout.addLayout(source_nav_layout)
        source_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        preview_splitter.addWidget(source_container)

        result_container = QWidget()
        result_layout = QVBoxLayout(result_container)
        result_layout.setContentsMargins(0, 0, 0, 0)
        result_label = QLabel("结果预览")
        result_label.setAlignment(Qt.AlignCenter)
        result_label.setStyleSheet("font-weight: bold;")
        result_label.setFixedHeight(35)
        result_layout.addWidget(result_label)

        self.result_scroll = QScrollArea()
        self.result_scroll.setWidgetResizable(True)
        self.result_scroll.setStyleSheet("border: 1px solid #ccc;")
        self.result_preview = QWebEngineView()
        self.result_preview.setHtml("<html><body style='background-color:#fff; color:#333;'><h3>识别结果预览</h3><p>转换完成后将在此处显示 Markdown 渲染结果。</p></body></html>")
        self.result_scroll.setWidget(self.result_preview)
        result_layout.addWidget(self.result_scroll, 1)

        self.lbl_result_page = QLabel("0/0")
        self.lbl_result_page.setAlignment(Qt.AlignCenter)
        self.lbl_result_page.setFixedHeight(30)
        self.lbl_result_page.setStyleSheet("border: 1px solid #ccc; background-color: #f5f5f5;")
        result_layout.addWidget(self.lbl_result_page)
        result_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        preview_splitter.addWidget(result_container)

        right_panel.addWidget(preview_splitter, 1)
        preview_splitter.setStretchFactor(0, 1)
        preview_splitter.setStretchFactor(1, 1)
        QTimer.singleShot(0, lambda: self.set_splitter_equal_width(preview_splitter))
        source_container.setMinimumWidth(300)
        result_container.setMinimumWidth(300)

        btn_frame = QFrame()
        btn_layout = QHBoxLayout(btn_frame)
        btn_layout.setContentsMargins(0, 10, 0, 0)
        self.btn_start = QPushButton("转换全部")
        self.btn_start.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; height: 40px;")
        self.btn_convert_selected = QPushButton("转换所选文件")
        self.btn_convert_selected.setEnabled(False)
        self.btn_convert_selected.setStyleSheet("background-color: #FF9800; color: white; font-weight: bold; height: 40px;")
        self.btn_open_folder = QPushButton("打开输出目录")
        self.btn_open_folder.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; height: 40px;")
        self.btn_download = QPushButton("下载/另存为...")
        self.btn_download.setStyleSheet("background-color: #FF9800; color: white; font-weight: bold; height: 40px;")
        btn_layout.addWidget(self.btn_start)
        btn_layout.addWidget(self.btn_convert_selected)
        btn_layout.addWidget(self.btn_open_folder)
        btn_layout.addWidget(self.btn_download)
        right_panel.addWidget(btn_frame)

        main_layout.addLayout(right_panel, 7)

        self.btn_add_files.clicked.connect(self.add_files)
        self.btn_add_folder.clicked.connect(self.add_folder)
        self.btn_clear.clicked.connect(self.clear_file_list)
        self.btn_browse_output.clicked.connect(self.browse_output_dir)
        self.btn_start.clicked.connect(self.start_conversion)
        self.btn_convert_selected.clicked.connect(self.convert_selected_file)
        self.btn_open_folder.clicked.connect(self.open_output_folder)
        self.btn_download.clicked.connect(self.download_result)
        self.btn_source_prev.clicked.connect(self.prev_source_page)
        self.btn_source_next.clicked.connect(self.next_source_page)

    # ==========================================
    # Logic Implementation
    # ==========================================
    def get_tool_executable(self, name):
        """健壮地查找虚拟环境中的可执行文件路径，兼容 Conda 和标准 venv"""
        # 1. 与 sys.executable 同目录 (标准 venv: env/Scripts/python.exe)
        exe_dir = os.path.dirname(sys.executable)
        if sys.platform == "win32":
            exe_path = os.path.join(exe_dir, f"{name}.exe")
        else:
            exe_path = os.path.join(exe_dir, name)
        if os.path.exists(exe_path):
            return exe_path

        # 2. 同目录下的 Scripts 子目录 (Conda 或某些 venv: env/python.exe, exe in env/Scripts/)
        if sys.platform == "win32":
            exe_path_scripts = os.path.join(exe_dir, "Scripts", f"{name}.exe")
        else:
            exe_path_scripts = os.path.join(exe_dir, "bin", name)
        if os.path.exists(exe_path_scripts):
            return exe_path_scripts

        # 3. sys.prefix 下的 Scripts 目录
        prefix_dir = sys.prefix
        if sys.platform == "win32":
            exe_path_prefix = os.path.join(prefix_dir, "Scripts", f"{name}.exe")
        else:
            exe_path_prefix = os.path.join(prefix_dir, "bin", name)
        if os.path.exists(exe_path_prefix):
            return exe_path_prefix

        # 4. 如果都找不到，返回名称本身，依赖系统 PATH
        return name

    def clear_file_list(self):
        self.file_list.clear()
        self.input_files.clear()
        self.txt_output_dir.clear()  # 清空输出目录
        self.txt_pages.clear()  # 清空页面范围
        self.update_file_count_label()

    def set_output_to_source_folder(self):
        """清空输出目录，恢复为默认的源文件所在目录"""
        self.txt_output_dir.clear()

    def log(self, message):
        cursor = self.log_window.textCursor()
        cursor.movePosition(QTextCursor.End)
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        cursor.insertText(f"[{timestamp}] {message}\n")
        self.log_window.setTextCursor(cursor)
        self.log_window.ensureCursorVisible()
        self.log_lines.append(f"[{timestamp}] {message}")

    def clear_log(self):
        self.log_window.clear()
        self.log_lines.clear()
        self.log("日志已清空")

    def update_gpu_info(self):
        if not HAS_GPUtil:
            self.lbl_gpu_info.setText("GPUtil未安装")
            return
        try:
            gpus = GPUtil.getGPUs()
            if gpus:
                gpu = gpus[0]
                usage_percent = (gpu.memoryUsed / gpu.memoryTotal) * 100 if gpu.memoryTotal > 0 else 0
                gpu_info = f"GPU: {gpu.name}\n显存: {gpu.memoryTotal}MB (已用: {gpu.memoryUsed}MB, {usage_percent:.1f}% | 可用: {gpu.memoryFree}MB)"
                self.lbl_gpu_info.setText(gpu_info)
            else:
                self.lbl_gpu_info.setText("未检测到独立显卡 (使用CPU)")
        except Exception as e:
            self.lbl_gpu_info.setText(f"GPU监控错误: {str(e)}")

    def update_file_count_label(self):
        font = QFont()
        font.setPointSize(11)
        self.lbl_file_count.setFont(font)
        count = self.file_list.count()
        if count == 0:
            self.lbl_file_count.setText("未添加文件")
        elif count == 1:
            item = self.file_list.item(0)
            if item:
                file_path = item.text()
                if file_path.lower().endswith('.pdf'):
                    try:
                        doc = fitz.open(file_path)
                        page_count = doc.page_count
                        doc.close()
                        self.lbl_file_count.setText(f"已添加 1 个文件，共 {page_count} 页")
                    except:
                        self.lbl_file_count.setText("已添加 1 个文件")
                else:
                    self.lbl_file_count.setText("已添加 1 个图片文件")
        else:
            self.lbl_file_count.setText(f"已添加 {count} 个文件 (批量模式)")

    def set_splitter_equal_width(self, splitter):
        total_width = splitter.width()
        if total_width > 0:
            splitter.setSizes([total_width // 2, total_width // 2])
        else:
            splitter.setSizes([500, 500])

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择文件", "", "PDF Files (*.pdf);;Images (*.png *.jpg *.jpeg)")
        if files:
            self.file_list.addItems(files)
            self.input_files.extend(files)
            self.update_file_count_label()
            if len(files) == 1 and self.file_list.count() == 1:
                self.preview_source_file(self.file_list.item(0))

    def add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            count_before = self.file_list.count()
            for root, dirs, files in os.walk(folder):
                for f in files:
                    if f.lower().endswith('.pdf'):
                        full_path = os.path.join(root, f)
                        self.file_list.addItem(full_path)
                        self.input_files.append(full_path)
            if self.file_list.count() > count_before:
                self.update_file_count_label()
                if self.file_list.count() == 1:
                    self.preview_source_file(self.file_list.item(0))

    def browse_output_dir(self):
        folder = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if folder:
            self.txt_output_dir.setText(folder)

    def render_source_page(self, file_path, page_num):
        try:
            scroll_size = self.source_scroll.viewport().size()
            max_width = scroll_size.width() - 20
            max_height = scroll_size.height() - 20
            if file_path.lower().endswith(('.png', '.jpg', '.jpeg')):
                pixmap = QPixmap(file_path)
                scaled_pixmap = pixmap.scaled(max_width, max_height, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.source_preview_label.setPixmap(scaled_pixmap)
                self.source_preview_label.setFixedSize(scaled_pixmap.size())
                self.source_preview_container.setFixedSize(scaled_pixmap.size())
                self.lbl_source_page.setText("1/1")
                self.source_total_pages = 1
                self.btn_source_prev.setEnabled(False)
                self.btn_source_next.setEnabled(False)
            elif file_path.lower().endswith('.pdf'):
                doc = fitz.open(file_path)
                self.source_total_pages = doc.page_count
                if page_num >= self.source_total_pages:
                    page_num = self.source_total_pages - 1
                if page_num < 0:
                    page_num = 0
                self.current_source_page = page_num
                page = doc.load_page(page_num)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(img)
                scaled_pixmap = pixmap.scaled(max_width, max_height, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.source_preview_label.setPixmap(scaled_pixmap)
                self.source_preview_label.setFixedSize(scaled_pixmap.size())
                self.source_preview_container.setFixedSize(scaled_pixmap.size())
                self.lbl_source_page.setText(f"{self.current_source_page + 1}/{self.source_total_pages}")
                self.btn_source_prev.setEnabled(self.current_source_page > 0)
                self.btn_source_next.setEnabled(self.current_source_page < self.source_total_pages - 1)
                doc.close()
        except Exception as e:
            self.source_preview_label.setText(f"无法预览文件:\n{str(e)}")

    def prev_source_page(self):
        if self.current_source_page > 0:
            self.render_source_page(self.current_preview_file, self.current_source_page - 1)

    def next_source_page(self):
        if self.current_source_page < self.source_total_pages - 1:
            self.render_source_page(self.current_preview_file, self.current_source_page + 1)

    def preview_source_file(self, item):
        file_path = item.text()
        self.current_preview_file = file_path
        self.current_source_page = 0
        self.btn_convert_selected.setEnabled(True)
        self.log(f"预览源文件: {file_path}")
        self.render_source_page(file_path, 0)
        # 同时加载对应的结果预览
        self.load_result_preview(file_path)

    def load_result_preview(self, file_path):
        """根据源文件路径查找并加载对应的结果预览"""
        if not file_path:
            return
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_dir = self.txt_output_dir.text() if self.txt_output_dir.text() else os.path.dirname(file_path)

        # 如果设置了页面范围，添加到文件名中
        page_range = self.txt_pages.text().strip()
        if page_range:
            base_name_with_range = f"{base_name}（page_{page_range}）"
        else:
            base_name_with_range = base_name

        possible_paths = [
            os.path.join(output_dir, f"{base_name_with_range}.md"),
            os.path.join(output_dir, f"{base_name}.md"),
            os.path.join(output_dir, base_name, f"{base_name}.md")
        ]
        for path in possible_paths:
            if os.path.exists(path):
                self.preview_result(path)
                return
        # 如果没有找到转换结果，清空结果预览
        self.result_preview.setHtml("<html><body style='background-color:#fff; color:#333;'><h3>识别结果预览</h3><p>该文件尚未转换或未找到转换结果。</p></body></html>")

    def get_selected_formats(self):
        formats = []
        if self.chk_format_json.isChecked():
            formats.append("json")
        if self.chk_format_md.isChecked():
            formats.append("markdown")
        if self.chk_format_html.isChecked():
            formats.append("html")
        if self.chk_format_pdf.isChecked():
            formats.append("pdf")
        if self.chk_format_text.isChecked():
            formats.append("text")
        return ",".join(formats) if formats else "markdown"

    def start_conversion(self):
        if self.worker and self.worker.isRunning():
            QMessageBox.warning(self, "提示", "已有任务在运行中...")
            return
        if self.file_list.count() == 0:
            QMessageBox.warning(self, "提示", "请先添加要转换的文件。")
            return

        self.log_lines = []
        files_to_process = []
        for i in range(self.file_list.count()):
            files_to_process.append(self.file_list.item(i).text())

        # 获取当前虚拟环境下的可执行文件路径
        odl_exe = self.get_tool_executable("opendataloader-pdf")

        # 输出格式
        formats = self.get_selected_formats()

        # 使用结构标签
        use_struct = self.chk_use_struct_tree.isChecked()

        # AI安全
        ai_safety = self.chk_ai_safety.isChecked()

        # 页面范围
        pages = self.txt_pages.text().strip()

        # 检查输出目录
        output_dir = self.txt_output_dir.text()

        # 只有在设置了输出目录时才检查目录
        if not output_dir:
            output_dir = None  # 标记为未设置，使用源文件目录

        # 单文件处理：始终逐个处理每个文件
        if len(files_to_process) == 1:
            # 单文件模式，直接处理
            self.log(f"执行单文件处理")
            self.run_single_file_conversion(odl_exe, files_to_process[0], formats, pages, use_struct, ai_safety, output_dir)
        else:
            # 多文件处理：检查是否来自同一目录
            source_dirs = set()
            for file_path in files_to_process:
                source_dirs.add(os.path.dirname(file_path))

            if len(source_dirs) > 1:
                # 文件来自不同目录，需要分开单独处理
                self.log(f"检测到文件来自不同目录，将分别处理每个文件（输出到各自源目录）")
                self.run_separate_conversions(odl_exe, files_to_process, formats, pages, use_struct, ai_safety)
            else:
                # 文件来自同一目录，批量处理
                self.log(f"检测到文件来自同一目录，执行批量处理")
                effective_output_dir = output_dir if output_dir else list(source_dirs)[0]
                self.run_batch_conversion(odl_exe, files_to_process, formats, pages, use_struct, ai_safety, effective_output_dir)

    def run_single_file_conversion(self, odl_exe, file_path, formats, pages, use_struct, ai_safety, output_dir=None):
        """处理单个文件，支持自动命名和页面范围"""
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        source_dir = os.path.dirname(file_path)
        effective_output_dir = output_dir if output_dir else source_dir

        # 如果设置了页面范围，添加到文件名中
        if pages:
            base_name_with_range = f"{base_name}（page_{pages}）"
        else:
            base_name_with_range = base_name

        # 获取带序号的输出路径（避免覆盖同名文件）
        output_path = self.get_available_save_path(effective_output_dir, base_name_with_range, ".md")
        output_file_name = os.path.basename(output_path)
        output_base_name = os.path.splitext(output_file_name)[0]

        self.log(f"执行单文件处理")
        self.log(f"源文件: {file_path}")
        self.log(f"输出格式: {formats}")
        self.log(f"输出目录: {effective_output_dir}")
        if pages:
            self.log(f"页面范围: {pages}")
        self.log(f"输出文件: {output_file_name}")

        # 构建命令参数
        args = [odl_exe]
        args.extend(["-f", formats])
        args.extend(["-o", effective_output_dir])
        if pages:
            args.extend(["--pages", pages])
        if use_struct:
            args.append("--use-struct-tree")
        if ai_safety:
            args.append("--sanitize")
        args.append(file_path)

        self.log(f"[系统] 正在执行命令: {' '.join(args)}")
        self.log("-" * 50)

        # 存储目标路径用于后续重命名
        self._pending_rename = {
            'source_dir': effective_output_dir,
            'original_base': base_name,
            'target_base': output_base_name,
            'format': formats.split(',')[0] if ',' in formats else formats,
            'pages': pages
        }

        # 创建worker
        self.worker = ODLWorker(args, self.log_lines)
        self.worker.log_signal.connect(self.log)

        def on_single_finished(success, msg):
            self.log(msg)
            self.btn_start.setEnabled(True)
            self.btn_convert_selected.setEnabled(True)

            if success and hasattr(self, '_pending_rename'):
                # 重命名输出文件
                rename_info = self._pending_rename
                source_base = rename_info['original_base']
                target_base = rename_info['target_base']
                source_dir = rename_info['source_dir']
                output_format = rename_info['format']
                pages = rename_info.get('pages', '')

                # 获取扩展名
                ext_map = {
                    'json': '.json',
                    'markdown': '.md',
                    'html': '.html',
                    'pdf': '.pdf',
                    'text': '.txt'
                }
                ext = ext_map.get(output_format, '.md')

                # 查找原始输出文件
                original_path = os.path.join(source_dir, f"{source_base}{ext}")
                target_path = os.path.join(source_dir, f"{target_base}{ext}")

                if os.path.exists(original_path):
                    if original_path != target_path:
                        try:
                            if os.path.exists(target_path):
                                os.remove(target_path)
                            os.rename(original_path, target_path)
                            self.log(f"文件已重命名: {target_path}")
                        except Exception as e:
                            self.log(f"重命名失败: {str(e)}")

                    # 刷新预览
                    if self.current_preview_file:
                        self.load_result_preview(self.current_preview_file)
                else:
                    # 尝试其他可能的路径（包括带页面范围的原始文件名）
                    source_base_with_range = source_base + f"（page_{pages}）" if pages else source_base
                    possible_paths = [
                        os.path.join(source_dir, f"{source_base_with_range}{ext}")
                    ]
                    for path in possible_paths:
                        if os.path.exists(path) and path != target_path:
                            try:
                                if os.path.exists(target_path):
                                    os.remove(target_path)
                                os.rename(path, target_path)
                                self.log(f"文件已重命名: {target_path}")
                                if self.current_preview_file:
                                    self.load_result_preview(self.current_preview_file)
                                break
                            except:
                                pass

                delattr(self, '_pending_rename')

        self.worker.finished_signal.connect(on_single_finished)
        self.worker.start()
        self.btn_start.setEnabled(False)
        self.btn_convert_selected.setEnabled(False)

    def run_batch_conversion(self, odl_exe, files_to_process, formats, pages, use_struct, ai_safety, output_dir):
        """批量处理多个文件"""
        self.log(f"模式: 批量模式")
        self.log(f"输出格式: {formats}")
        self.log(f"输出目录: {output_dir}")
        if pages:
            self.log(f"页面范围: {pages}")
        self.log(f"文件数量: {len(files_to_process)}")

        # 准备批量处理队列
        self._batch_queue = []
        for file_path in files_to_process:
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            if pages:
                base_name_with_range = f"{base_name}（page_{pages}）"
            else:
                base_name_with_range = base_name
            target_path = self.get_available_save_path(output_dir, base_name_with_range, ".md")
            self._batch_queue.append({
                'file_path': file_path,
                'base_name': base_name,
                'target_path': target_path,
                'output_base': os.path.splitext(os.path.basename(target_path))[0]
            })

        self._batch_index = 0
        self._batch_total = len(self._batch_queue)
        self._batch_output_dir = output_dir
        self._batch_formats = formats
        self._batch_pages = pages
        self._batch_use_struct = use_struct
        self._batch_ai_safety = ai_safety
        self._odl_exe = odl_exe

        self._process_next_batch()

    def _process_next_batch(self):
        """处理批量队列中的下一个文件"""
        if self._batch_index >= self._batch_total:
            self.log("批量处理全部完成！")
            self.btn_start.setEnabled(True)
            self.btn_convert_selected.setEnabled(True)
            if self.current_preview_file:
                self.load_result_preview(self.current_preview_file)
            return

        batch_item = self._batch_queue[self._batch_index]
        file_path = batch_item['file_path']
        target_base = batch_item['output_base']
        self._batch_current_target_base = target_base

        self.log(f"处理文件 ({self._batch_index + 1}/{self._batch_total}): {os.path.basename(file_path)}")
        self.log(f"目标文件名: {target_base}.md")

        args = [self._odl_exe]
        args.extend(["-f", self._batch_formats])
        args.extend(["-o", self._batch_output_dir])
        if self._batch_pages:
            args.extend(["--pages", self._batch_pages])
        if self._batch_use_struct:
            args.append("--use-struct-tree")
        if self._batch_ai_safety:
            args.append("--sanitize")
        args.append(file_path)

        self.worker = ODLWorker(args, self.log_lines)
        self.worker.log_signal.connect(self.log)

        def on_batch_finished(success, msg):
            if success:
                # 重命名输出文件
                ext_map = {
                    'json': '.json',
                    'markdown': '.md',
                    'html': '.html',
                    'pdf': '.pdf',
                    'text': '.txt'
                }
                output_format = self._batch_formats.split(',')[0] if ',' in self._batch_formats else self._batch_formats
                ext = ext_map.get(output_format, '.md')

                original_path = os.path.join(self._batch_output_dir, f"{batch_item['base_name']}{ext}")
                target_path = os.path.join(self._batch_output_dir, f"{self._batch_current_target_base}{ext}")

                if os.path.exists(original_path):
                    if original_path != target_path:
                        try:
                            if os.path.exists(target_path):
                                os.remove(target_path)
                            os.rename(original_path, target_path)
                            self.log(f"文件已重命名: {target_path}")
                        except Exception as e:
                            self.log(f"重命名失败: {str(e)}")
                else:
                    self.log(f"警告: 未找到原始输出文件 {original_path}")

            # 处理下一个文件
            self._batch_index += 1
            self._process_next_batch()

        self.worker.finished_signal.connect(on_batch_finished)
        self.worker.start()

    def run_separate_conversions(self, odl_exe, files_to_process, formats, pages, use_struct, ai_safety):
        """当文件来自不同目录时，分别单独处理每个文件"""
        total_files = len(files_to_process)
        self.log(f"开始逐个处理文件（共 {total_files} 个）...")

        # 创建一个临时队列来跟踪处理进度
        self._pending_files = list(files_to_process)
        self._current_format = formats
        self._current_pages = pages
        self._current_use_struct = use_struct
        self._current_ai_safety = ai_safety
        self._odl_exe = odl_exe
        self.btn_start.setEnabled(False)
        self.btn_convert_selected.setEnabled(False)
        self._process_next_file()

    def _process_next_file(self):
        """处理队列中的下一个文件"""
        if not self._pending_files:
            self.log("所有文件处理完成！")
            self.btn_start.setEnabled(True)
            self.btn_convert_selected.setEnabled(True)
            # 刷新当前选中文件的预览
            if self.current_preview_file:
                self.load_result_preview(self.current_preview_file)
            return

        file_path = self._pending_files.pop(0)
        output_dir = os.path.dirname(file_path)
        processed_count = len(self._pending_files)
        total_count = processed_count + 1

        args = [self._odl_exe]
        args.extend(["-f", self._current_format])
        args.extend(["-o", output_dir])
        if self._current_pages:
            args.extend(["--pages", self._current_pages])
        if self._current_use_struct:
            args.append("--use-struct-tree")
        if self._current_ai_safety:
            args.append("--sanitize")
        args.append(file_path)

        self.log(f"处理文件 ({processed_count + 1}/{total_count}): {os.path.basename(file_path)}")

        # 创建worker处理单个文件
        self.worker = ODLWorker(args, self.log_lines)
        self.worker.log_signal.connect(self.log)

        # 完成后继续处理下一个
        def on_single_finished(success, msg):
            if success:
                self.log(f"文件 {os.path.basename(file_path)} 处理成功")
            else:
                self.log(f"文件 {os.path.basename(file_path)} 处理失败: {msg}")
            # 处理完成后继续处理下一个
            self._process_next_file()

        self.worker.finished_signal.connect(on_single_finished)
        self.worker.start()

    def run_worker(self, args, current_file_name):
        self.worker = ODLWorker(args, self.log_lines)
        self.worker.log_signal.connect(self.log)
        self.worker.finished_signal.connect(lambda success, msg: self.on_conversion_finished(success, msg, current_file_name))
        self.worker.start()
        self.btn_start.setEnabled(False)
        self.btn_convert_selected.setEnabled(False)

    def on_conversion_finished(self, success, message, current_file_name):
        self.log(message)
        self.btn_start.setEnabled(True)
        self.btn_convert_selected.setEnabled(True)

        # 单文件预览刷新
        if success:
            if current_file_name == "批量任务":
                # 批量模式下，刷新当前选中文件的预览
                if self.current_preview_file:
                    self.load_result_preview(self.current_preview_file)
            else:
                # 单文件模式下，只刷新当前文件的预览
                if self.current_preview_file == current_file_name:
                    base_name = os.path.splitext(os.path.basename(current_file_name))[0]
                    output_dir = self.txt_output_dir.text() if self.txt_output_dir.text() else os.path.dirname(current_file_name)

                    # 如果设置了页面范围，添加到文件名中
                    page_range = self.txt_pages.text().strip()
                    if page_range:
                        base_name_with_range = f"{base_name}（page_{page_range}）"
                    else:
                        base_name_with_range = base_name

                    possible_paths = [
                        os.path.join(output_dir, f"{base_name_with_range}.md"),
                        os.path.join(output_dir, f"{base_name}.md"),
                        os.path.join(output_dir, base_name, f"{base_name}.md")
                    ]
                    for path in possible_paths:
                        if os.path.exists(path):
                            self.preview_result(path)
                            if self.chk_auto_save.isChecked():
                                self.auto_save_result(path)
                            break

    def convert_selected_file(self):
        """转换文件列表中选中的文件"""
        if self.worker and self.worker.isRunning():
            QMessageBox.warning(self, "提示", "已有任务在运行中...")
            return
        if not self.current_preview_file:
            QMessageBox.warning(self, "提示", "请先在文件列表中选择要转换的文件。")
            return

        self.log_lines = []
        file_path = self.current_preview_file

        # 获取当前虚拟环境下的可执行文件路径
        odl_exe = self.get_tool_executable("opendataloader-pdf")

        # 输出格式
        formats = self.get_selected_formats()

        # 使用结构标签
        use_struct = self.chk_use_struct_tree.isChecked()

        # AI安全
        ai_safety = self.chk_ai_safety.isChecked()

        # 页面范围
        pages = self.txt_pages.text().strip()

        # 检查输出目录
        output_dir = self.txt_output_dir.text()

        # 只有在设置了输出目录时才检查目录
        if not output_dir:
            output_dir = None  # 标记为未设置，使用源文件目录

        self.run_single_file_conversion(odl_exe, file_path, formats, pages, use_struct, ai_safety, output_dir)

    def get_available_save_path(self, output_dir, base_name, extension=".md"):
        """获取可用的保存路径，如果文件已存在则自动添加序号"""
        save_path = os.path.join(output_dir, f"{base_name}{extension}")
        if not os.path.exists(save_path):
            return save_path

        counter = 1
        while True:
            save_path = os.path.join(output_dir, f"{base_name}-{counter:02d}{extension}")
            if not os.path.exists(save_path):
                return save_path
            counter += 1

    def auto_save_result(self, md_path):
        try:
            if not md_path.endswith('.md'):
                return
            if self.current_preview_file:
                original_name = os.path.splitext(os.path.basename(self.current_preview_file))[0]
                output_dir = self.txt_output_dir.text() if self.txt_output_dir.text() else os.path.dirname(self.current_preview_file)

                # 如果设置了页面范围，添加到文件名中
                page_range = self.txt_pages.text().strip()
                if page_range:
                    # 将页面范围格式化为（page:1-10）格式
                    base_name_with_range = f"{original_name}（page_{page_range}）"
                else:
                    base_name_with_range = original_name

                auto_save_path = self.get_available_save_path(output_dir, base_name_with_range)

                if os.path.abspath(md_path) == os.path.abspath(auto_save_path):
                    return

                shutil.copy(md_path, auto_save_path)
                self.log(f"自动保存成功: {auto_save_path}")
        except Exception as e:
            self.log(f"自动保存失败: {str(e)}")

    def preview_result(self, md_path):
        try:
            with open(md_path, 'r', encoding='utf-8') as f:
                md_content = f.read()
            line_count = len(md_content.split('\n'))
            self.result_total_pages = max(1, (line_count + 99) // 100)
            self.current_result_page = 1
            self.lbl_result_page.setText(f"{self.current_result_page}/{self.result_total_pages}")

            # 获取MD文件所在目录，用于解析图片相对路径
            md_dir = os.path.dirname(os.path.abspath(md_path))
            md_base_name = os.path.splitext(os.path.basename(md_path))[0]
            # 去掉页面范围后缀（如 "（page:1-10）"），获取原始文件名以定位 _images 目录
            clean_base_name = re.sub(r'[（(]page_[^）)]+[）)]', '', md_base_name).strip()

            # 处理图片路径：将相对路径转换为绝对 file:// URL
            def fix_image_path(match):
                img_src = match.group(1)
                # 如果已是绝对路径或网络地址，直接返回原值
                if img_src.startswith(('http://', 'https://', 'file://', 'data:')):
                    return f'src="{img_src}"'
                # 相对于MD文件目录的绝对路径
                img_full_path = os.path.join(md_dir, img_src)
                if not os.path.exists(img_full_path):
                    img_filename = os.path.basename(img_src)
                    # 优先用去掉页面范围后的原始名称查找 _images 目录
                    for base in (clean_base_name, md_base_name):
                        alt_path = os.path.join(md_dir, f"{base}_images", img_filename)
                        if os.path.exists(alt_path):
                            img_full_path = alt_path
                            break
                # 转换为 file:// URL（统一用正斜杠，兼容 Windows）
                file_url = QUrl.fromLocalFile(os.path.normpath(img_full_path)).toString()
                return f'src="{file_url}"'

            # 转换Markdown为HTML
            html_content = markdown.markdown(md_content, extensions=['tables', 'fenced_code'])

            # 替换 img 标签的 src 属性，将相对路径转为绝对 file:// URL
            html_content = re.sub(r'src="([^"]*)"', fix_image_path, html_content)

            full_html = f"""
            <html>
            <head>
            <style>
                body {{ font-family: 'Microsoft YaHei', sans-serif; margin: 20px; line-height: 1.6; }}
                pre {{ background-color: #f4f4f4; padding: 10px; border-radius: 5px; overflow-x: auto; }}
                code {{ font-family: Consolas; background-color: #f4f4f4; padding: 2px 4px; border-radius: 3px; }}
                table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                img {{ max-width: 100%; height: auto; }}
            </style>
            </head>
            <body>
            {html_content}
            </body>
            </html>
            """
            # 设置 baseUrl 为 MD 文件所在目录，确保相对路径资源可正常加载
            base_url = QUrl.fromLocalFile(md_dir + os.sep)
            self.result_preview.setHtml(full_html, base_url)
            self.log(f"已加载预览: {md_path}")
        except Exception as e:
            self.log(f"预览失败: {str(e)}")

    def open_output_folder(self):
        output_dir = self.txt_output_dir.text()
        if output_dir and os.path.exists(output_dir):
            os.startfile(output_dir)
            return
        if self.current_preview_file:
            output_dir = os.path.dirname(self.current_preview_file)
            if os.path.exists(output_dir):
                os.startfile(output_dir)
                return
        QMessageBox.warning(self, "提示", "请先设置输出目录或选择一个文件。")

    def download_result(self):
        if not self.current_preview_file:
            return
        base_name = os.path.splitext(os.path.basename(self.current_preview_file))[0]
        output_dir = self.txt_output_dir.text() if self.txt_output_dir.text() else os.path.dirname(self.current_preview_file)

        # 如果设置了页面范围，添加到文件名中
        page_range = self.txt_pages.text().strip()
        if page_range:
            # 将页面范围格式化为（page:1-10）格式
            base_name_with_range = f"{base_name}（page_{page_range}）"
        else:
            base_name_with_range = base_name

        src_path = None
        possible_paths = [
            os.path.join(output_dir, f"{base_name_with_range}.md"),
            os.path.join(output_dir, f"{base_name}.md"),
            os.path.join(output_dir, base_name, f"{base_name}.md")
        ]
        for p in possible_paths:
            if os.path.exists(p):
                src_path = p
                break
        if not src_path:
            QMessageBox.warning(self, "错误", "未找到转换结果文件。")
            return

        # 获取不重名的默认文件名
        default_save_name = os.path.basename(self.get_available_save_path(output_dir, base_name_with_range))

        save_path, _ = QFileDialog.getSaveFileName(self, "保存结果", default_save_name, "Markdown Files (*.md)")
        if save_path:
            try:
                shutil.copy(src_path, save_path)
                self.log(f"文件已保存至: {save_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", str(e))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    font = QFont("Microsoft YaHei", 9)
    app.setFont(font)
    window = ODLGUI()
    window.show()
    sys.exit(app.exec_())

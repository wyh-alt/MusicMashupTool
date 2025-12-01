"""
整合界面 - 统一的GUI界面
"""
import sys
import logging
from pathlib import Path
from typing import Optional, Dict, List
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit, QProgressBar,
    QFileDialog, QMessageBox, QGroupBox, QCheckBox, QTabWidget,
    QDoubleSpinBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon

from pipeline_worker import PipelineWorker

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class DropLineEdit(QLineEdit):
    """支持拖拽的文本框"""
    file_dropped = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            self.setText(file_path)
            self.file_dropped.emit(file_path)
            event.acceptProposedAction()


class IntegratedMainWindow(QMainWindow):
    """整合主窗口"""
    
    def __init__(self):
        super().__init__()
        self.worker: Optional[PipelineWorker] = None
        self.setWindowTitle("音乐串烧一键工具")
        self.resize(900, 750)
        
        # 设置窗口图标
        icon_path = Path(__file__).parent / 'app_icon.ico'
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        
        self.init_ui()
    
    def init_ui(self):
        """初始化界面"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # 输入设置组
        input_group = QGroupBox("输入设置")
        input_layout = QVBoxLayout()
        input_group.setLayout(input_layout)
        
        # 原始表格文件
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(QLabel("原始歌曲表格:"))
        self.excel_input = DropLineEdit()
        self.excel_input.setPlaceholderText("拖拽或选择包含歌曲信息的Excel文件...")
        excel_layout.addWidget(self.excel_input)
        browse_excel_btn = QPushButton("浏览...")
        browse_excel_btn.clicked.connect(self.browse_excel)
        excel_layout.addWidget(browse_excel_btn)
        input_layout.addLayout(excel_layout)
        
        # 音频文件目录
        audio_layout = QHBoxLayout()
        audio_layout.addWidget(QLabel("音频文件目录:"))
        self.audio_input = DropLineEdit()
        self.audio_input.setPlaceholderText("拖拽或选择包含音频文件的目录...")
        audio_layout.addWidget(self.audio_input)
        browse_audio_btn = QPushButton("浏览...")
        browse_audio_btn.clicked.connect(self.browse_audio)
        audio_layout.addWidget(browse_audio_btn)
        input_layout.addLayout(audio_layout)
        
        # 输出目录
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出目录:"))
        self.output_input = DropLineEdit()
        self.output_input.setPlaceholderText("选择输出目录...")
        output_layout.addWidget(self.output_input)
        browse_output_btn = QPushButton("浏览...")
        browse_output_btn.clicked.connect(self.browse_output)
        output_layout.addWidget(browse_output_btn)
        input_layout.addLayout(output_layout)
        
        main_layout.addWidget(input_group)
        
        # 参数设置组
        param_group = QGroupBox("参数设置")
        param_layout = QVBoxLayout()
        param_group.setLayout(param_layout)
        
        # 分类调号区间设置
        key_range_layout = QHBoxLayout()
        key_range_layout.addWidget(QLabel("[分类] 调号区间:"))
        self.key_range_spinbox = QDoubleSpinBox()
        self.key_range_spinbox.setRange(0, 6)
        self.key_range_spinbox.setSingleStep(1)
        self.key_range_spinbox.setValue(2)
        self.key_range_spinbox.setDecimals(0)
        self.key_range_spinbox.setSuffix(" 个半音")
        self.key_range_spinbox.setFixedWidth(120)
        key_range_layout.addWidget(self.key_range_spinbox)
        key_range_layout.addWidget(QLabel("（调号相差范围，默认±2即一个全音）"))
        key_range_layout.addStretch()
        param_layout.addLayout(key_range_layout)
        
        # 分类速度区间设置
        bpm_range_layout = QHBoxLayout()
        bpm_range_layout.addWidget(QLabel("[分类] 速度区间:"))
        self.bpm_range_spinbox = QDoubleSpinBox()
        self.bpm_range_spinbox.setRange(0, 20)
        self.bpm_range_spinbox.setSingleStep(1)
        self.bpm_range_spinbox.setValue(5)
        self.bpm_range_spinbox.setDecimals(0)
        self.bpm_range_spinbox.setSuffix(" BPM")
        self.bpm_range_spinbox.setFixedWidth(120)
        bpm_range_layout.addWidget(self.bpm_range_spinbox)
        bpm_range_layout.addWidget(QLabel("（速度相差范围，默认±5）"))
        bpm_range_layout.addStretch()
        param_layout.addLayout(bpm_range_layout)
        
        # 静音间隙设置
        gap_layout = QHBoxLayout()
        gap_layout.addWidget(QLabel("[拼接] 静音间隙:"))
        self.gap_spinbox = QDoubleSpinBox()
        self.gap_spinbox.setRange(0, 5.0)
        self.gap_spinbox.setSingleStep(0.1)
        self.gap_spinbox.setValue(0.5)
        self.gap_spinbox.setSuffix(" 秒")
        self.gap_spinbox.setFixedWidth(120)
        gap_layout.addWidget(self.gap_spinbox)
        gap_layout.addWidget(QLabel("（音频拼接时片段之间的静音时长）"))
        gap_layout.addStretch()
        param_layout.addLayout(gap_layout)
        
        main_layout.addWidget(param_group)
        
        # 处理流程显示
        process_group = QGroupBox("处理流程")
        process_layout = QVBoxLayout()
        process_group.setLayout(process_layout)
        
        # 三个步骤的状态显示
        self.step1_label = QLabel("步骤 1: 歌曲分类 - 等待开始")
        self.step1_label.setStyleSheet("padding: 5px;")
        process_layout.addWidget(self.step1_label)
        
        self.step1_progress = QProgressBar()
        self.step1_progress.setMaximum(100)
        self.step1_progress.setValue(0)
        process_layout.addWidget(self.step1_progress)
        
        self.step2_label = QLabel("步骤 2: 变调变速 - 等待开始")
        self.step2_label.setStyleSheet("padding: 5px;")
        process_layout.addWidget(self.step2_label)
        
        self.step2_progress = QProgressBar()
        self.step2_progress.setMaximum(100)
        self.step2_progress.setValue(0)
        process_layout.addWidget(self.step2_progress)
        
        self.step3_label = QLabel("步骤 3: 音频拼接 - 等待开始")
        self.step3_label.setStyleSheet("padding: 5px;")
        process_layout.addWidget(self.step3_label)
        
        self.step3_progress = QProgressBar()
        self.step3_progress.setMaximum(100)
        self.step3_progress.setValue(0)
        process_layout.addWidget(self.step3_progress)
        
        main_layout.addWidget(process_group)
        
        # 总进度
        total_progress_layout = QHBoxLayout()
        total_progress_layout.addWidget(QLabel("总进度:"))
        self.total_progress = QProgressBar()
        self.total_progress.setMaximum(100)
        self.total_progress.setValue(0)
        total_progress_layout.addWidget(self.total_progress)
        main_layout.addLayout(total_progress_layout)
        
        # 控制按钮
        button_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始处理")
        self.start_btn.setMinimumHeight(40)
        self.start_btn.setStyleSheet("font-size: 14px; font-weight: bold;")
        self.start_btn.clicked.connect(self.start_processing)
        button_layout.addWidget(self.start_btn)
        
        self.stop_btn = QPushButton("停止")
        self.stop_btn.setMinimumHeight(40)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_processing)
        button_layout.addWidget(self.stop_btn)
        
        main_layout.addLayout(button_layout)
        
        # 日志区域
        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout()
        log_group.setLayout(log_layout)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)
        
        main_layout.addWidget(log_group)
    
    def browse_excel(self):
        """浏览Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "",
            "Excel文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        if file_path:
            self.excel_input.setText(file_path)
    
    def browse_audio(self):
        """浏览音频目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择音频文件目录")
        if dir_path:
            self.audio_input.setText(dir_path)
    
    def browse_output(self):
        """浏览输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_input.setText(dir_path)
    
    def log(self, message: str, level: str = "INFO"):
        """记录日志"""
        self.log_text.append(f"[{level}] {message}")
        logger.log(getattr(logging, level), message)
    
    def start_processing(self):
        """开始处理"""
        # 验证输入
        excel_path = self.excel_input.text().strip()
        audio_dir = self.audio_input.text().strip()
        output_dir = self.output_input.text().strip()
        
        if not excel_path:
            QMessageBox.warning(self, "警告", "请选择原始歌曲表格文件！")
            return
        
        if not audio_dir:
            QMessageBox.warning(self, "警告", "请选择音频文件目录！")
            return
        
        if not output_dir:
            QMessageBox.warning(self, "警告", "请选择输出目录！")
            return
        
        if not Path(excel_path).exists():
            QMessageBox.critical(self, "错误", f"表格文件不存在：{excel_path}")
            return
        
        if not Path(audio_dir).exists():
            QMessageBox.critical(self, "错误", f"音频目录不存在：{audio_dir}")
            return
        
        # 创建输出目录
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        # 禁用开始按钮，启用停止按钮
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        
        # 重置进度
        self.step1_progress.setValue(0)
        self.step2_progress.setValue(0)
        self.step3_progress.setValue(0)
        self.total_progress.setValue(0)
        
        # 清空日志
        self.log_text.clear()
        self.log("=" * 60)
        self.log("开始处理音乐串烧流程...")
        self.log("=" * 60)
        
        # 创建工作线程
        gap_duration = self.gap_spinbox.value()
        key_range = int(self.key_range_spinbox.value())
        bpm_range = int(self.bpm_range_spinbox.value())
        
        self.log(f"分类参数：调号区间=±{key_range}个半音，速度区间=±{bpm_range} BPM")
        self.log(f"拼接参数：静音间隙={gap_duration}秒")
        
        self.worker = PipelineWorker(
            excel_path=Path(excel_path),
            audio_dir=Path(audio_dir),
            output_dir=Path(output_dir),
            gap_duration=gap_duration,
            key_range=key_range,
            bpm_range=bpm_range
        )
        
        # 连接信号
        self.worker.step1_progress.connect(self.on_step1_progress)
        self.worker.step2_progress.connect(self.on_step2_progress)
        self.worker.step3_progress.connect(self.on_step3_progress)
        self.worker.total_progress.connect(self.on_total_progress)
        self.worker.log_message.connect(self.log)
        self.worker.finished_signal.connect(self.on_processing_finished)
        self.worker.error_occurred.connect(self.on_error)
        
        # 启动线程
        self.worker.start()
    
    def stop_processing(self):
        """停止处理"""
        if self.worker:
            self.worker.cancel()
            self.log("正在取消处理...", "WARNING")
    
    def on_step1_progress(self, current: int, total: int, message: str):
        """步骤1进度更新"""
        progress = int((current / total) * 100) if total > 0 else 0
        self.step1_progress.setValue(progress)
        self.step1_label.setText(f"步骤 1: 歌曲分类 - {message}")
    
    def on_step2_progress(self, current: int, total: int, message: str):
        """步骤2进度更新"""
        progress = int((current / total) * 100) if total > 0 else 0
        self.step2_progress.setValue(progress)
        self.step2_label.setText(f"步骤 2: 变调变速 - {message}")
    
    def on_step3_progress(self, current: int, total: int, message: str):
        """步骤3进度更新"""
        progress = int((current / total) * 100) if total > 0 else 0
        self.step3_progress.setValue(progress)
        self.step3_label.setText(f"步骤 3: 音频拼接 - {message}")
    
    def on_total_progress(self, progress: int):
        """总进度更新"""
        self.total_progress.setValue(progress)
    
    def on_processing_finished(self, success: bool, message: str):
        """处理完成"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        
        if success:
            self.log("=" * 60)
            self.log("所有处理完成！")
            self.log(message)
            self.log("=" * 60)
            QMessageBox.information(self, "完成", message)
        else:
            self.log("处理被取消或失败", "WARNING")
    
    def on_error(self, error_msg: str):
        """错误处理"""
        self.log(error_msg, "ERROR")
        QMessageBox.critical(self, "错误", error_msg)
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)


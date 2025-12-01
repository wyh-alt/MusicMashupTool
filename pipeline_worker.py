"""
流程处理工作线程
按顺序执行：歌曲分类 → 变调变速 → 音频拼接
"""
import logging
from pathlib import Path
from typing import Optional
from PyQt6.QtCore import QThread, pyqtSignal

from step1_classifier import classify_songs_core
from step2_pitch_tempo import process_pitch_tempo_core
from step3_concat import concat_audio_core

logger = logging.getLogger(__name__)


class PipelineWorker(QThread):
    """流程处理工作线程"""
    
    # 信号定义
    step1_progress = pyqtSignal(int, int, str)  # 当前进度, 总数, 消息
    step2_progress = pyqtSignal(int, int, str)
    step3_progress = pyqtSignal(int, int, str)
    total_progress = pyqtSignal(int)  # 总进度百分比 (0-100)
    log_message = pyqtSignal(str, str)  # 消息, 级别
    finished_signal = pyqtSignal(bool, str)  # 是否成功, 消息
    error_occurred = pyqtSignal(str)  # 错误消息
    
    def __init__(self, excel_path: Path, audio_dir: Path, output_dir: Path, 
                 gap_duration: float, key_range: int = 2, bpm_range: int = 5):
        super().__init__()
        self.excel_path = excel_path
        self.audio_dir = audio_dir
        self.output_dir = output_dir
        self.gap_duration = gap_duration
        self.key_range = key_range  # 调号区间（±几个半音）
        self.bpm_range = bpm_range  # 速度区间（±几个BPM）
        self.is_cancelled = False
        
        # 中间文件路径
        self.classified_excel_path: Optional[Path] = None
        self.processed_audio_dir: Optional[Path] = None
    
    def cancel(self):
        """取消处理"""
        self.is_cancelled = True
    
    def run(self):
        """执行处理流程"""
        try:
            # 步骤1: 歌曲分类
            if not self.step1_classify():
                return
            
            if self.is_cancelled:
                self.finished_signal.emit(False, "处理已取消")
                return
            
            # 步骤2: 变调变速
            if not self.step2_pitch_tempo():
                return
            
            if self.is_cancelled:
                self.finished_signal.emit(False, "处理已取消")
                return
            
            # 步骤3: 音频拼接
            if not self.step3_concat():
                return
            
            # 完成后清理临时文件
            import shutil
            if self.processed_audio_dir and self.processed_audio_dir.exists():
                try:
                    shutil.rmtree(self.processed_audio_dir)
                    self.log_message.emit("已清理临时文件", "INFO")
                except Exception as e:
                    self.log_message.emit(f"清理临时文件失败: {e}", "WARNING")
            
            # 完成
            self.finished_signal.emit(
                True,
                f"所有处理完成！\n\n"
                f"分类表格：{self.classified_excel_path}\n"
                f"音频成品：{self.output_dir}/*.mp3"
            )
            
        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            error_msg = f"处理过程出错：{str(e)}\n\n详细信息：\n{error_detail}"
            logger.error(error_msg)
            self.error_occurred.emit(error_msg)
            self.finished_signal.emit(False, "处理失败")
    
    def step1_classify(self) -> bool:
        """步骤1: 歌曲分类"""
        try:
            self.log_message.emit("开始步骤1：歌曲分类", "INFO")
            self.step1_progress.emit(0, 100, "读取表格...")
            self.total_progress.emit(0)
            
            # 创建输出目录（直接在主目录）
            self.output_dir.mkdir(parents=True, exist_ok=True)
            
            # 生成输出文件名（直接保存在输出目录根目录）
            input_name = self.excel_path.stem
            self.classified_excel_path = self.output_dir / f"{input_name}_分类结果.xlsx"
            
            self.step1_progress.emit(10, 100, "分析歌曲数据...")
            
            # 调用分类核心函数
            def progress_callback(current, total, message):
                if self.is_cancelled:
                    return False
                progress = 10 + int((current / total) * 80) if total > 0 else 10
                self.step1_progress.emit(progress, 100, message)
                return True
            
            groups, df = classify_songs_core(
                self.excel_path,
                self.classified_excel_path,
                progress_callback,
                key_range=self.key_range,
                bpm_range=self.bpm_range
            )
            
            self.step1_progress.emit(100, 100, f"完成！生成了 {len(groups)} 个分类")
            self.total_progress.emit(33)
            
            self.log_message.emit(f"步骤1完成：生成了 {len(groups)} 个分类", "INFO")
            self.log_message.emit(f"分类表格已保存至：{self.classified_excel_path}", "INFO")
            
            return True
            
        except Exception as e:
            error_msg = f"步骤1（歌曲分类）失败：{str(e)}"
            self.error_occurred.emit(error_msg)
            return False
    
    def step2_pitch_tempo(self) -> bool:
        """步骤2: 变调变速"""
        try:
            self.log_message.emit("开始步骤2：变调变速处理", "INFO")
            self.step2_progress.emit(0, 100, "准备音频处理...")
            
            # 创建临时目录（用点开头隐藏）
            import tempfile
            step2_output_dir = self.output_dir / ".temp_audio"
            step2_output_dir.mkdir(parents=True, exist_ok=True)
            self.processed_audio_dir = step2_output_dir
            
            self.step2_progress.emit(5, 100, "解析分类表格...")
            
            # 调用变调变速核心函数
            def progress_callback(current, total, message):
                if self.is_cancelled:
                    return False
                progress = 5 + int((current / total) * 90) if total > 0 else 5
                self.step2_progress.emit(progress, 100, message)
                self.log_message.emit(f"处理进度：{message}", "INFO")
                return True
            
            success_count, total_count = process_pitch_tempo_core(
                self.classified_excel_path,
                self.audio_dir,
                step2_output_dir,
                progress_callback
            )
            
            self.step2_progress.emit(100, 100, f"完成！处理了 {success_count}/{total_count} 个文件")
            self.total_progress.emit(66)
            
            self.log_message.emit(
                f"步骤2完成：成功处理 {success_count}/{total_count} 个音频文件",
                "INFO"
            )
            
            return True
            
        except Exception as e:
            error_msg = f"步骤2（变调变速）失败：{str(e)}"
            self.error_occurred.emit(error_msg)
            return False
    
    def step3_concat(self) -> bool:
        """步骤3: 音频拼接"""
        try:
            self.log_message.emit("开始步骤3：音频拼接", "INFO")
            self.step3_progress.emit(0, 100, "准备拼接...")
            
            # 直接输出到主目录（不创建子文件夹）
            step3_output_dir = self.output_dir
            step3_output_dir.mkdir(parents=True, exist_ok=True)
            
            self.step3_progress.emit(5, 100, "读取分类表格...")
            
            # 调用音频拼接核心函数
            def progress_callback(current, total, message):
                if self.is_cancelled:
                    return False
                progress = 5 + int((current / total) * 90) if total > 0 else 5
                self.step3_progress.emit(progress, 100, message)
                self.log_message.emit(f"拼接进度：{message}", "INFO")
                return True
            
            success_count, total_count = concat_audio_core(
                self.classified_excel_path,
                self.processed_audio_dir,
                step3_output_dir,
                self.gap_duration,
                progress_callback
            )
            
            self.step3_progress.emit(100, 100, f"完成！拼接了 {success_count}/{total_count} 个成品")
            self.total_progress.emit(100)
            
            self.log_message.emit(
                f"步骤3完成：成功拼接 {success_count}/{total_count} 个音频成品",
                "INFO"
            )
            
            return True
            
        except Exception as e:
            error_msg = f"步骤3（音频拼接）失败：{str(e)}"
            self.error_occurred.emit(error_msg)
            return False


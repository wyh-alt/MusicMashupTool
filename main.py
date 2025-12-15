"""
音乐串烧一键工具 - 主程序入口
整合三个功能模块：歌曲分类 → 变调变速 → 音频拼接
"""
import sys
import platform
import subprocess
from pathlib import Path

# 预先导入 librosa，确保 asyncio 完成导入
# 这样后续修改 subprocess.Popen 就不会破坏 asyncio
import librosa  # noqa: F401

# Windows 下配置 subprocess.Popen 隐藏窗口（在 asyncio 导入完成后）
if platform.system() == 'Windows':
    _original_popen = subprocess.Popen
    
    class _PopenNoWindow(subprocess.Popen):
        """Windows 下隐藏子进程窗口的 Popen 子类"""
        def __init__(self, *args, **kwargs):
            # 设置 STARTUPINFO 隐藏窗口
            if 'startupinfo' not in kwargs:
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
                kwargs['startupinfo'] = startupinfo
            
            # 设置 CREATE_NO_WINDOW 标志
            creation_flags = kwargs.get('creationflags', 0)
            kwargs['creationflags'] = creation_flags | subprocess.CREATE_NO_WINDOW
            
            super().__init__(*args, **kwargs)
    
    # 替换全局 Popen（此时 asyncio 已完成导入，是安全的）
    subprocess.Popen = _PopenNoWindow

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon
from gui import IntegratedMainWindow


def main():
    """主函数"""
    app = QApplication(sys.argv)
    
    # 设置应用程序信息
    app.setApplicationName("音乐串烧一键工具")
    app.setApplicationVersion("1.1.0")
    
    # 设置应用图标
    icon_path = Path(__file__).parent / 'app_icon.ico'
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    
    # 创建主窗口
    window = IntegratedMainWindow()
    window.show()
    
    # 运行应用程序
    sys.exit(app.exec())


if __name__ == '__main__':
    main()



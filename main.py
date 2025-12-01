"""
音乐串烧一键工具 - 主程序入口
整合三个功能模块：歌曲分类 → 变调变速 → 音频拼接
"""
import sys
from pathlib import Path
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



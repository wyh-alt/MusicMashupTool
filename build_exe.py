"""
音乐串烧一键工具 - EXE打包脚本
使用PyInstaller将程序打包为独立可执行文件
"""
import subprocess
import sys
import os
from pathlib import Path


def check_pyinstaller():
    """检查PyInstaller是否安装"""
    try:
        import PyInstaller
        print(f"[OK] PyInstaller version: {PyInstaller.__version__}")
        return True
    except ImportError:
        print("[!] PyInstaller not found, installing...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
        return True


def build_exe():
    """执行打包"""
    # 获取当前目录（已在main中切换）
    current_dir = Path.cwd()
    
    print("=" * 60)
    print("Music Mashup Tool - EXE Builder")
    print("=" * 60)
    
    # 检查PyInstaller
    check_pyinstaller()
    
    # 检查图标文件
    icon_file = current_dir / 'app_icon.ico'
    if not icon_file.exists():
        print(f"[!] Warning: Icon file not found: {icon_file}")
        icon_arg = []
    else:
        print(f"[OK] Icon file: {icon_file}")
        icon_arg = ['--icon', str(icon_file)]
    
    # PyInstaller参数
    args = [
        sys.executable, '-m', 'PyInstaller',
        '--name', 'MusicMashupTool',      # 输出文件名
        '--onefile',                       # 单文件模式
        '--windowed',                      # 无控制台窗口
        '--noconfirm',                     # 覆盖输出目录
        '--clean',                         # 清理临时文件
        
        # 添加数据文件
        '--add-data', f'{icon_file};.',    # 包含图标
        
        # 隐藏导入（确保所有依赖被包含）
        '--hidden-import', 'PyQt6',
        '--hidden-import', 'PyQt6.QtCore',
        '--hidden-import', 'PyQt6.QtGui',
        '--hidden-import', 'PyQt6.QtWidgets',
        '--hidden-import', 'pandas',
        '--hidden-import', 'openpyxl',
        '--hidden-import', 'librosa',
        '--hidden-import', 'soundfile',
        '--hidden-import', 'numpy',
        '--hidden-import', 'pydub',
        '--hidden-import', 'audioread',
        '--hidden-import', 'sklearn',
        '--hidden-import', 'sklearn.utils._cython_blas',
        '--hidden-import', 'sklearn.neighbors.typedefs',
        '--hidden-import', 'sklearn.neighbors._partition_nodes',
        '--hidden-import', 'scipy.special._cdflib',
        
        # 收集所有子模块
        '--collect-all', 'librosa',
        '--collect-all', 'soundfile',
        '--collect-all', 'audioread',
        
        # 入口文件
        'main.py'
    ]
    
    # 添加图标参数
    args = args[:-1] + icon_arg + [args[-1]]
    
    print("\n[*] Starting build process...")
    print(f"[*] Command: {' '.join(args[2:])}\n")
    
    # 执行打包
    try:
        subprocess.check_call(args)
        print("\n" + "=" * 60)
        print("[OK] Build completed successfully!")
        print("=" * 60)
        
        # 显示输出位置
        dist_dir = current_dir / 'dist'
        exe_file = dist_dir / 'MusicMashupTool.exe'
        
        if exe_file.exists():
            size_mb = exe_file.stat().st_size / (1024 * 1024)
            print(f"\n[*] Output: {exe_file}")
            print(f"[*] Size: {size_mb:.1f} MB")
            print("\n[!] Note: You need to have ffmpeg installed on target machine")
            print("    Download: https://ffmpeg.org/download.html")
        
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"\n[ERROR] Build failed: {e}")
        return False


def main():
    """主函数"""
    print("\nThis script will build the Music Mashup Tool as a standalone EXE.\n")
    
    # 切换到脚本所在目录
    script_dir = Path(__file__).parent.resolve()
    os.chdir(script_dir)
    print(f"[*] Working directory: {script_dir}")
    
    # 检查是否在正确的目录
    if not Path('main.py').exists():
        print("[ERROR] main.py not found in script directory.")
        sys.exit(1)
    
    # 执行打包
    success = build_exe()
    
    if success:
        print("\n" + "=" * 60)
        print("IMPORTANT: FFmpeg Requirement")
        print("=" * 60)
        print("""
The packaged EXE requires ffmpeg to be installed on the target machine.

Option 1: Install ffmpeg system-wide
  - Download from: https://ffmpeg.org/download.html
  - Add to system PATH

Option 2: Place ffmpeg.exe in the same folder as the EXE
  - Download ffmpeg-release-essentials.zip
  - Extract ffmpeg.exe to the same folder

The EXE file is located in: dist/MusicMashupTool.exe
""")
    
    input("\nPress Enter to exit...")


if __name__ == '__main__':
    main()


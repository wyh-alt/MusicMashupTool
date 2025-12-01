@echo off
chcp 65001 >nul
title Music Mashup Tool - Build EXE

echo ============================================================
echo    Music Mashup Tool - EXE Builder
echo ============================================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python first.
    pause
    exit /b 1
)

:: Install/Update PyInstaller
echo [*] Checking PyInstaller...
pip install pyinstaller --upgrade -q

:: Build using spec file
echo [*] Building EXE...
echo.

pyinstaller MusicMashupTool.spec --noconfirm --clean

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed!
    pause
    exit /b 1
)

echo.
echo ============================================================
echo    Build Complete!
echo ============================================================
echo.
echo Output: dist\MusicMashupTool.exe
echo.
echo IMPORTANT: Target machine needs ffmpeg installed!
echo Download: https://ffmpeg.org/download.html
echo.
echo Or place ffmpeg.exe in the same folder as MusicMashupTool.exe
echo.
pause


@echo off
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    set "PYTHON=.venv\Scripts\python.exe"
) else (
    set "PYTHON=python"
)

"%PYTHON%" -m pip install -r requirements-dev.txt
if errorlevel 1 goto failed

"%PYTHON%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name PPT_Word_to_PDF ^
    --hidden-import pythoncom ^
    --hidden-import pywintypes ^
    --collect-submodules win32com ^
    gui.py
if errorlevel 1 goto failed

echo.
echo 打包完成：dist\PPT_Word_to_PDF.exe
pause
exit /b 0

:failed
echo.
echo 打包失败，请查看上方错误信息。
pause
exit /b 1

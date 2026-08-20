@echo off
setlocal
cd /d "%~dp0"

where makensis >nul 2>nul
if errorlevel 1 (
    echo 未找到 NSIS 编译器 makensis。
    echo 请先安装 NSIS：https://nsis.sourceforge.io/Download
    pause
    exit /b 1
)

if not exist "PPT_Word_to_PDF.exe" (
    echo 未找到 PPT_Word_to_PDF.exe。
    echo 请先运行 build_exe.bat 生成主程序，或将可执行文件放在项目根目录。
    pause
    exit /b 1
)

if not exist "release" mkdir "release"
makensis /DPROJECT_ROOT="%CD%" "installer\PPT_Word_to_PDF_Setup.nsi"
if errorlevel 1 goto failed

echo.
echo 安装包已生成：release\PPT_Word_to_PDF_Setup_v1.0.1.exe
pause
exit /b 0

:failed
echo.
echo 安装包构建失败，请查看上方错误信息。
pause
exit /b 1

@echo off
chcp 65001 >nul
:: 设置构建与链接环境时区为中国标准时间 (CST, UTC+8)
set TZ=CST-8
echo ===================================================================
echo   Excel 暗水印与保护工具 - MinGW 独立 GUI 静态编译脚本 (CST时区 v0.1.1)
echo ===================================================================

where g++ >nul 2>nul
if %errorlevel% neq 0 (
    echo [错误] 未在 PATH 中找到 g++.exe。请确认已安装 MinGW-w64。
    pause
    exit /b 1
)

echo [1/2] 编译 Windows 视觉清单资源 (UTF-8 代码页 65001)...
windres -c 65001 resource.rc -O coff -o resource.o

echo [2/2] 正在进行全静态 GUI 链接编译 (-mwindows)...
g++ -O2 -std=c++17 -DUNICODE -D_UNICODE -municode -mwindows ^
    gui_main.cpp excel_toolkit_core.cpp miniz.c resource.o ^
    -o excel-toolkit.exe ^
    -static -static-libgcc -static-libstdc++ ^
    -lcomctl32 -lshell32 -ladvapi32 -luser32 -lgdi32 -lcomdlg32 -luxtheme

if %errorlevel% equ 0 (
    echo 编译成功！生成独立 GUI 绿色版: excel-toolkit.exe
) else (
    echo 编译失败，请检查编译错误。
)
pause

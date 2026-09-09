@echo off
chcp 65001 >nul
echo ===================================================================
echo   Excel 暗水印与保护工具 - MSVC 独立 GUI 静态编译 (无需 VC++ 运行库)
echo ===================================================================

where cl >nul 2>nul
if %errorlevel% neq 0 (
    echo [错误] 未检测到 MSVC 编译器 (cl.exe)。
    echo 请在 "Developer Command Prompt for VS 2019/2022" 中运行此脚本！
    pause
    exit /b 1
)

echo [1/3] 编译资源文件 resource.rc...
rc /nologo resource.rc

echo [2/3] 正在使用 /MT 静态编译原生 GUI 程序 (无黑框 /SUBSYSTEM:WINDOWS)...
cl /nologo /O2 /MT /utf-8 /EHsc /std:c++17 ^
   gui_main.cpp excel_toolkit_core.cpp miniz.c resource.res ^
   /link /SUBSYSTEM:WINDOWS /OUT:excel-toolkit.exe ^
   comctl32.lib shell32.lib advapi32.lib user32.lib gdi32.lib comdlg32.lib

if %errorlevel% equ 0 (
    echo [3/3] 编译成功！
    echo 生成原生单文件绿色图形程序: excel-toolkit.exe
    echo 该程序可以在任何纯净 Windows 7/10/11 机器上直接双击运行，无黑色控制台，无需安装 VC++ 运行库。
) else (
    echo [错误] 编译失败，请检查上方日志。
)
pause

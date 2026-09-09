# Excel工具箱 (excel-toolkit) — C++ 原生独立版 (Windows 10/11 现代 Fluent 版)

对应完整底层业务逻辑的 C++ 原生重构版，提供现代 Windows 10/11 原生 GUI 桌面客户端与跨平台编译支持。

## 核心设计与模块架构

1. **卡片 1【加工】**：
   - **写入 veryHidden 隐性暗水印工作表**：在 Excel (.xlsx / .xls) 中写入深度隐藏工作表 `~sysdata_cfg`，记录防伪溯源编号、暗水印文字与时间戳。
   - **锁定工作簿结构**：写入 `lockStructure="1"`，置灰 Excel“取消隐藏”菜单，防止恶意窥探或篡改暗水印表。
   - **锁定工作表内容**：遍历所有工作表写入 `sheetProtection`，启用单元格只读防篡改保护。
2. **卡片 2【脱壳】**：
   - **工作表结构与暗水印扫描**：深度扫描所有 Sheet 及其可见性状态（正常可见 visible / 普通隐藏 hidden / 深度隐藏 veryHidden），解析隐性暗水印特征与溯源编号。
   - **一键脱壳解锁**：一键清除工作簿结构保护与所有工作表哈希密码，还原干净可编辑状态。

## 目录文件清单

- `app.ico`: Windows 应用程序立体多分辨率图标 (256/48/32/16)。
- `gui_main.cpp`: Windows 10/11 原生 GUI 图形界面程序，支持拖拽、高 DPI 自适应、全 Unicode 宽字符流、无控制台黑框。
- `excel_toolkit_core.h` / `excel_toolkit_core.cpp`: 核心 OpenXML 内存流解析引擎，零临时文件残留，杜绝崩溃。
- `main.cpp`: 控制台命令行版本。
- `miniz.h` / `miniz.c`: 内嵌 RFC1951 Deflate/ZIP 引擎。
- `resource.rc`: Windows 视觉清单、版本信息及主图标资源。
- `CMakeLists.txt`: CMake 跨平台构建清单。
- `build_msvc.bat`: MSVC /MT 静态单文件编译一键脚本 (无 VC++ 运行库依赖)。
- `build_mingw.bat`: MinGW-w64 编译脚本。
- `.github/workflows/build.yml`: GitHub Actions 自动编译与 Release 发布流水线。

## GitHub Actions 自动构建与发布

- **Push 触发**：自动编译输出 `excel-toolkit-{UNIX时间戳}-beta.exe` 并上传至 Artifacts 供直接下载；
- **Release 触发**：自动编译输出 `excel-toolkit-v{版本号}-beta.exe` 并自动上传到 Release 附件列表。

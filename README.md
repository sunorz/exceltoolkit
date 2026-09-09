# excel-toolkit — Excel 暗水印与保护工具 (C++ 原生独立版 / 零依赖)

对应 Python 版 excel_toolkit 完整逻辑的 C++ 原生重构版，提供现代 Windows 原生 GUI 桌面客户端与跨平台编译支持。

## 核心功能与 Checkbox 自由组合

1. **[✓] 添加暗水印**：在 Excel (.xlsx / .xls) 中写入 `veryHidden` 深度隐藏工作表 `~sysdata_cfg`，记录溯源编号、水印文字和 ISO 生成时间。
2. **[✓] 锁定工作簿结构**：写入 `lockStructure="1"`，彻底置灰 Excel“取消隐藏”菜单，防止窥探暗水印表。
3. **[✓] 锁定工作表内容**：遍历所有工作表写入 `sheetProtection`，禁止单元格编辑与排版修改。
4. **[ ] 一键解除保护**：一键清空工作簿与工作表的全部密码与保护限制。
5. **暗水印取证与查看**：扫描工作簿所有 Sheet 及其可见性状态（visible/hidden/veryHidden），解析隐藏表中的溯源编号与水印文字。

## 目录文件清单

- `gui_main.cpp`: Windows 原生 GUI 图形界面程序，支持拖拽、复选框自由组合、无控制台黑框。
- `excel_toolkit_core.h` / `excel_toolkit_core.cpp`: 核心 OpenXML 解析与处理模块，内存级 ZIP 读写，杜绝闪退。
- `main.cpp`: 控制台命令行版本。
- `miniz.h` / `miniz.c`: 内嵌轻量压缩解压库。
- `resource.rc`: Windows 资源文件与版本信息。
- `CMakeLists.txt`: CMake 构建清单。
- `build_msvc.bat`: MSVC /MT 静态编译一键脚本。
- `build_mingw.bat`: MinGW-w64 编译脚本。
- `.github/workflows/build.yml`: GitHub Actions 自动编译与 Release 发布流水线。

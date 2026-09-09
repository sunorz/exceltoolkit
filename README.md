# Excel 暗水印与保护工具 (C++ 原生独立版 / 零依赖)

对应 Python 版 `excel_toolkit.py` 完整逻辑的 C++ 原生重构版。

## 核心功能

1. **添加暗水印**：在 Excel (.xlsx / .xls) 中写入 `veryHidden` 深度隐藏工作表 `~sysdata_cfg`，记录溯源编号、水印文字和 ISO 生成时间。
2. **查看暗水印**：扫描工作簿所有 Sheet 及其可见性状态（visible/hidden/veryHidden），解析隐藏表中的溯源编号与水印文字。
3. **保护内容（禁止编辑 / 解除保护）**：
   - 开启保护：对所有工作表注入密码保护（禁止单元格编辑、格式修改、插入删除行），并锁定工作簿结构（`lockStructure="1"`），防止普通用户在 Excel 界面中取消隐藏暗水印表；
   - 解除保护：一键清空工作表与工作簿保护。

## GitHub Actions 自动构建与发布

本工程根目录内置 `.github/workflows/build.yml`：

- **Push 触发**：自动编译输出 `excel-toolkit-{UNIX时间戳}-beta.exe` 并上传至 Artifacts 供直接下载；
- **Release 触发**：自动编译输出 `excel-toolkit-v{版本号}-beta.exe` 并自动上传到 Release 附件列表。

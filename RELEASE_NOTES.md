# Release Notes

## v1.1.0

- 修复 Windows PowerPoint COM 路径正斜杠导致的 `0x80070003` 异常。
- 新增 Office 弹窗自动拦截（DisplayAlerts / Normal.dotm 提示），杜绝后台无响应卡死。
- 支持全系列格式（`.pptm`、`.pps`、`.ppsx`、`.ppsm`、`.docm`、`.dot`、`.dotx`、`.rtf`）。
- 新增“原地递归转换”选项，保留原有目录分类与层级结构。
- 新增任务中途“取消转换”与确定性进度条显示。
- 新增只读文件安全删除防护与 PDF 读写占用检测，严格校验完整性后才删除源文件。
- 优化 Office 故障重试机制与友好的中文环境缺失提示。

## v1.0.1

- Added a standard Windows `setup.exe` installer for the existing 64-bit application.
- Added per-user installation, desktop and Start menu shortcuts, and a standard uninstaller.
- Added SHA-256 verification for installer release assets.
- Added a GitHub Actions workflow that builds installer artifacts on Windows and publishes them for version tags.

### Installation

Download `PPT_Word_to_PDF_Setup_v1.0.1.exe` from the v1.0.1 release, then run the installer. Microsoft Word and PowerPoint must already be installed on the target Windows computer.

## v1.0.0

- Added a simple Windows GUI for non-technical users.
- Added batch conversion for PowerPoint and Word files to PDF.
- Added optional subfolder flattening before conversion.
- Added optional source-file deletion after successful conversion.
- Added one-click executable build script.

### Download

Use `PPT_Word_to_PDF.exe` from the release package.

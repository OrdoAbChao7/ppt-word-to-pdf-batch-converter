# Windows 安装程序

安装程序基于 [NSIS](https://nsis.sourceforge.io/)，以当前仓库根目录的 `PPT_Word_to_PDF.exe` 为唯一应用载荷。它会在当前用户的 `%LOCALAPPDATA%\PPT_Word_to_PDF` 安装主程序和使用说明，并创建桌面、开始菜单与卸载快捷方式。

## 本地构建

在 Windows 安装 NSIS 后，双击 `build_setup.bat`；也可以在命令行执行：

```bat
makensis installer\PPT_Word_to_PDF_Setup.nsi
```

输出文件位于 `release/PPT_Word_to_PDF_Setup_v1.0.1.exe`。构建前请确认根目录已有最新的 `PPT_Word_to_PDF.exe`。

## 发布自动化

`.github/workflows/windows-installer.yml` 会在推送 `v*` 标签时使用 Windows runner 构建安装包、生成 SHA-256 校验文件，并将两者附加到对应的 GitHub Release。也可在 Actions 页面手动运行工作流获取安装包 artifact。

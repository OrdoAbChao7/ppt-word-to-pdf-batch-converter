# Windows 安装程序

安装程序基于 [NSIS](https://nsis.sourceforge.io/)，以当前仓库根目录的 `PPT_Word_to_PDF.exe` 为唯一应用载荷。它会在当前用户的 `%LOCALAPPDATA%\PPT_Word_to_PDF` 安装主程序，并创建桌面、开始菜单与卸载快捷方式；完整使用说明保留在仓库 README 中。

## 本地构建

在 Windows 安装 NSIS 后，双击 `build_setup.bat`；也可以在命令行执行：

```bat
makensis installer\PPT_Word_to_PDF_Setup.nsi
```

输出文件位于 `release/PPT_Word_to_PDF_Setup_v1.1.0.exe`。版本号统一取自安装脚本的 `PRODUCT_VERSION`。构建前请确认根目录已有最新的 `PPT_Word_to_PDF.exe`。

## 发布自动化

`.github/workflows/windows-installer.yml` 会在推送 `v*` 标签时检查标签与安装程序版本一致，运行测试，并使用 Windows runner 从源码构建绿色版、安装包及各自的 SHA-256 校验文件。构建产物附加到对应的 GitHub Release 草稿，检查完成后再正式发布。也可在 Actions 页面手动运行工作流获取构建产物 artifact。

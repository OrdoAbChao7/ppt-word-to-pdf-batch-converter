# PPT / Word 批量转 PDF 工具

一个面向非技术人员的 Windows 桌面小工具，可以批量把文件夹中的 PowerPoint 和 Word 文件转换成 PDF。

<img width="565" height="440" alt="1" src="https://github.com/user-attachments/assets/6d671668-b753-4630-a573-b5bb365a433f" />

## 功能

- 支持 `.ppt`、`.pptx`、`.doc`、`.docx` 批量转换为 PDF
- 提供简洁 GUI，双击即可使用
- 可选：先整理子文件夹，把子文件夹里的文件移动到主文件夹
- 可选：转换成功后删除原始 PPT / Word 文件
- 实时显示处理记录

## 运行环境

- Windows
- 已安装 Microsoft PowerPoint 和 Microsoft Word
- 如果直接运行源码，需要 Python 3.10 或更高版本

> 转换功能依赖本机 Office 程序，因此不适用于没有安装 Word / PowerPoint 的电脑。

## 给普通用户使用

推荐从 [Releases](../../releases) 下载 `PPT_Word_to_PDF_Setup_v1.0.1.exe`，双击安装后通过桌面或开始菜单启动程序。安装程序会在当前用户目录安装，不需要管理员权限。

安装完成前，请确认电脑满足以下条件：

- 使用 64 位 Windows。
- 已安装可正常启动的 Microsoft Word 和 Microsoft PowerPoint。
- Office 程序可以打开需要转换的文档。

也可以下载仓库根目录中的绿色版程序后双击运行：

```text
PPT_Word_to_PDF.exe
```

然后：

1. 点击“选择文件夹”
2. 选择要处理的文件夹
3. 按需勾选整理子文件夹、删除源文件
4. 点击“开始转换”
5. 等待完成提示

## 从源码运行

```bash
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python gui.py
```

也可以直接双击：

```text
启动GUI.bat
```

## 打包 exe

```bash
build_exe.bat
```

打包完成后，程序会出现在：

```text
dist\PPT_Word_to_PDF.exe
```

## 打包 setup.exe

安装 [NSIS](https://nsis.sourceforge.io/Download) 后，双击：

```text
build_setup.bat
```

生成的安装程序位于：

```text
release\PPT_Word_to_PDF_Setup_v1.0.1.exe
```

安装程序会创建桌面和开始菜单快捷方式，同时提供标准卸载入口。发布流程、SHA-256 校验与 GitHub Actions 自动构建说明见 [installer/README.md](installer/README.md)。

## 项目结构

```text
.
├── gui.py              # 图形界面入口
├── run.py              # 转换和文件夹整理逻辑
├── build_exe.bat       # 打包 exe
├── build_setup.bat     # 打包 Windows 安装程序
├── installer/          # NSIS 安装程序配置与构建说明
├── 启动GUI.bat          # 从源码启动 GUI
├── 使用说明.txt         # 给普通操作者看的简短说明
├── requirements.txt    # Python 依赖
├── LICENSE             # 开源许可证
└── README.md           # 项目说明
```

## 注意事项

- 转换过程中请不要手动关闭 Word 或 PowerPoint。
- 如果不确定是否需要保留原文件，请不要勾选“转换成功后删除原文件”。
- 旧版 `.doc` / `.ppt` 文件能否成功转换，取决于本机 Office 是否能正常打开该文件。

## 许可证

本项目使用 MIT License。

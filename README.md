<div align="center">
  <h1>ppt-word-to-pdf-batch-converter</h1>
  <b>English</b> | <a href="./README_zh-CN.md"><b>中文</b></a>
</div>
<br>



<!-- portfolio-authenticity:start -->
## Project status

**Stage:** Windows desktop conversion utility.

**Why I built it:** I built this to turn a repeated Office-to-PDF conversion task into a reviewable batch process with predictable output locations.

**Boundary:** It is tested for Windows desktop use and depends on locally installed Microsoft Office components. Conversion fidelity, protected files, macros, fonts, and damaged documents can produce failures or different PDFs. It is not a headless server converter.

See [PROJECT_STATUS.md](./PROJECT_STATUS.md) for the evidence still needed and the maintenance rule.
<!-- portfolio-authenticity:end -->

A simple Windows desktop script with a basic GUI to batch convert PowerPoint and Word files to PDF using local Microsoft Office COM automation.

<img width="565" height="440" alt="1" src="https://github.com/user-attachments/assets/6d671668-b753-4630-a573-b5bb365a433f" />
- Real-time processing log display

## System requirements

- Windows
- Microsoft PowerPoint and Microsoft Word installed
- If running from source, Python 3.10 or later

> Conversion relies on locally installed Office applications and will not work on machines without Word/PowerPoint.

## For end users

It is recommended to download `PPT_Word_to_PDF_Setup_v1.0.1.exe` from [Releases](../../releases). Double-click to install, then start the program from the Desktop or Start Menu. The installer performs a per-user installation and does not require administrator privileges.

Before installing, make sure your PC meets the following:

- 64-bit Windows
- Microsoft Word and Microsoft PowerPoint are installed and can start normally
- Office apps can open the documents you need to convert

You can also download the portable executable from the repository root and double-click to run:

```text
PPT_Word_to_PDF.exe
```

Then:

1. Click "Choose Folder"
2. Select the folder to process
3. Check "Flatten subfolders" and/or "Delete source files" if needed
4. Click "Start Conversion"
5. Wait for the completion message

## Run from source

```bash
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python gui.py
```

You can also simply double-click:

```text
启动GUI.bat
```

## Build exe

```bash
build_exe.bat
```

After building, the program will appear at:

```text
dist\PPT_Word_to_PDF.exe
```

## Build setup.exe

After installing [NSIS](https://nsis.sourceforge.io/Download), double-click:

```text
build_setup.bat
```

The generated installer is located at:

```text
release\PPT_Word_to_PDF_Setup_v1.0.1.exe
```

The installer will create Desktop and Start Menu shortcuts and provides a standard uninstaller. For release process, SHA-256 verification, and GitHub Actions automation, see [installer/README.md](installer/README.md).

## Project structure

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

## Notes

- Do not manually close Word or PowerPoint during conversion.
- If you are not sure whether to keep the original files, do not check "Delete original files after successful conversion".
- Whether legacy `.doc`/`.ppt` files can be converted depends on whether your local Office can open them.

## License

This project uses the MIT License.

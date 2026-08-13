# PPT / Word Batch-to-PDF Converter

> A small Windows desktop utility for batch-converting PowerPoint and Word files to PDF.

[中文文档](README.md)

This project is designed for non-technical users who need to convert a folder of Microsoft PowerPoint and Word files without opening each document manually.

## Features

| Feature | Description |
|---|---|
| Batch conversion | Converts `.ppt`, `.pptx`, `.doc`, and `.docx` files to PDF. |
| Desktop GUI | Provides a simple graphical interface that can be launched by double-clicking the packaged application. |
| Folder organization | Optionally moves files from subfolders into the selected main folder before conversion. |
| Source-file control | Optionally deletes the original PowerPoint / Word files only after a successful conversion. |
| Activity log | Shows processing progress and results in the interface. |

## Requirements

| Use case | Requirement |
|---|---|
| Run the packaged application | Windows with Microsoft PowerPoint and Microsoft Word installed. |
| Run from source | Windows, Microsoft Office, and Python 3.10 or later. |
| Convert legacy documents | The local Office installation must be able to open the target `.doc` / `.ppt` file. |

> PDF conversion uses locally installed Microsoft Office applications. The tool is not suitable for systems without Word and PowerPoint.

## Installation

### For regular users

Download the release executable from the repository root or a GitHub release, then run:

```text
PPT_Word_to_PDF.exe
```

No Python installation is required for the packaged executable.

### From source

Clone the repository and create an isolated Python environment:

```powershell
git clone https://github.com/OrdoAbChao7/ppt-word-to-pdf-batch-converter.git
cd ppt-word-to-pdf-batch-converter

python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python gui.py
```

You can also start the GUI by double-clicking:

```text
启动GUI.bat
```

## How to use the desktop application

1. Open `PPT_Word_to_PDF.exe`.
2. Select the folder containing the files to convert.
3. Optionally choose whether to organize subfolder files into the selected folder.
4. Decide whether original source files should be removed after a successful conversion.
5. Select **Start Conversion** and wait for the completion message.
6. Review the activity log before closing the application.

## Safety notes

The optional source-file deletion setting is irreversible for files that are successfully converted. If you are unsure whether an original document should be retained, leave that option unchecked and verify the generated PDFs first.

Do not manually close Word or PowerPoint while a conversion is running. Conversion of older `.doc` or `.ppt` files depends on whether the installed Office version can open them successfully.

## Build the executable

Run the build script from a source checkout:

```powershell
build_exe.bat
```

The generated executable is written to:

```text
dist\PPT_Word_to_PDF.exe
```

## Project layout

```text
.
├── gui.py              # Graphical application entry point
├── run.py              # Conversion and folder-organization logic
├── build_exe.bat       # Executable packaging script
├── 启动GUI.bat          # Source-mode GUI launcher
├── 使用说明.txt         # Short end-user instructions in Chinese
├── requirements.txt    # Python dependencies
├── LICENSE             # Open-source license
└── README.md           # Chinese project documentation
```

## Contributing

Contributions are welcome. Please keep changes focused and test them against Office files in a safe disposable folder. Do not change deletion behavior without clearly documenting the impact and adding safeguards for original documents.

## License

This project is released under the [MIT License](LICENSE).

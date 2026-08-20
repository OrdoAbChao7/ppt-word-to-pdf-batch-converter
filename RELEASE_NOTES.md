# Release Notes

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

Use `dist/PPT_Word_to_PDF.exe` from the release package.

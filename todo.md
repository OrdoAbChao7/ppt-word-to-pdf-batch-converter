# Release TODO

- [x] Review the existing executable, runtime dependencies, versioning, and release assets.
- [x] Add a reproducible NSIS configuration that installs the existing executable with Start menu and desktop shortcuts.
- [x] Add installer build instructions and a GitHub Actions workflow for repeatable Windows installer artifacts.
- [x] Build and inspect the setup.exe installation payload and uninstaller metadata.
- [x] Update release notes and README with installation prerequisites and setup.exe usage.
- [x] Commit installer sources, create a GitHub Release, and upload the setup.exe asset.
- [x] Verify the public GitHub Release and deliver its download link.
- [x] Normalize NSIS source and output paths so the setup.exe compiles reproducibly on Linux and Windows hosts.
- [x] Diagnose and correct the Windows GitHub Actions setup.exe build failure, then rerun the v1.0.1 release workflow.

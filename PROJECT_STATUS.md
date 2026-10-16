# Project status and engineering notes

This document records the boundary between what this repository is intended to demonstrate, what has been verified in the repository, and what still needs evidence. It is deliberately more specific than a feature list.

## Current stage

**Windows desktop conversion utility**

## Why this exists

I built this to turn a repeated Office-to-PDF conversion task into a reviewable batch process with predictable output locations.

## Scope and known limitations

It is tested for Windows desktop use and depends on locally installed Microsoft Office components. Conversion fidelity, protected files, macros, fonts, and damaged documents can produce failures or different PDFs. It is not a headless server converter.

## Next evidence to collect

Record additional Office versions and representative fixtures for layout drift, missing-font failures, downloaded files in Protected View, and very large documents.

## Reliability verification (2026-10-15)

- Actual user-log regression: a course folder had 29/34 failures with `ExportAsFixedFormat: file is read-only`. The failing PPTX files have modification protection (`modifyVerifier`); ordinary generated fixtures had missed this case. Native `SaveAs` and `SaveCopyAs` also failed on the samples. Importing slides into a temporary presentation succeeded for both PPTX and legacy PPT. All 34 actual source documents then converted successfully; PDF page counts and slide sizes matched and every source SHA-256 remained unchanged.
- The rebuilt executable also passed the entire 34-file folder (453 PDF pages), including all 29 originally failing presentations. Verification outputs and a per-file JSON report are under the local ignored `dist` directory; user course material is not included in regression fixtures.
- `python -m tests.verify_real_folder --source SOURCE --output OUTPUT` verifies real files while writing only to a separate output directory and saves a JSON report. `--executable EXE` checks bundled workers. This supplements synthetic regression tests; those tests alone do not establish compatibility with a user's actual documents.
- The real Office smoke test now generates both modification-protected PPTX and PPT fixtures and requires compatibility export for them. Compatibility copying preserves individual slides and dimensions; presentation-level print configuration and custom shows are not reproduced. Protected files that cannot be opened or whose slides cannot be imported remain explicit failures.
- Local Windows / Microsoft Word and PowerPoint 16.0: real conversion checks cover DOCX, PPTX, RTF, Chinese/space-containing paths, nested folders, and identical source stems. Password-protected DOCX and damaged PPTX fail independently without producing a final PDF or deleting their source.
- `python -m unittest discover -v` covers atomic replacement, invalid PDFs, retries, failure totals, output collisions, changed-source retention, timeout/cancellation, shared Office ownership, and GUI retry/progress behavior. GUI tests skip explicitly if Tcl/Tk cannot access a desktop.
- `python -m tests.office_smoke` generates temporary real Office fixtures, checks PDF pages/text and per-file outcomes, and removes the fixtures afterwards. It requires a normal Windows desktop with Word and PowerPoint installed.
- pywin32 COM cache corruption (`CLSIDToClassMap`) was reproduced locally. The application now uses recursive dynamic dispatch without deleting or rebuilding the user's shared cache.
- PDF verification checks file structure and page presence, not visual fidelity. Existing target PDFs are replaced only after validation. Different source formats with the same stem receive names such as `report.docx.pdf` and `report.pptx.pdf`.
- A healthy Office instance is reused for speed. After failure the worker reconnects once. The default per-attempt deadline is 120 seconds, adjustable in the GUI; a file can therefore take two deadlines before being marked failed. Only a newly identified Office process can be terminated on timeout; existing user Office processes are preserved.
- The rebuilt portable executable passed GUI startup/clean shutdown and the same real Office smoke scenarios using its bundled worker entry point (`python -m tests.office_smoke --executable dist/PPT_Word_to_PDF.exe`).

## GUI refinement (2026-10-16)

- Light, compact Tkinter interface with a single primary conversion action. File-moving/deletion options and per-attempt timeout live in a collapsible settings area; active file changes stay visible in its summary.
- Results are the default tab. Completion updates the results and totals without a modal dialog; cancellation and failed-only retry appear when applicable. Logs and log export remain available in a separate tab.
- Keyboard shortcuts: Ctrl+O selects a folder, Ctrl+Enter starts conversion, and Enter opens selected file details. Result columns support horizontal scrolling; long progress filenames are shortened for display only.
- `python -m unittest discover -v`: all 33 tests passed, including six GUI tests, on the local Windows desktop. Window inspection covers the default 920×720 layout, expanded settings at 800×640, mixed results, and long progress filenames. No conversion backend changes.
- The portable executable was rebuilt with PyInstaller and its GUI startup and clean shutdown were verified. The root portable executable and `dist/PPT_Word_to_PDF.exe` contain the updated interface.
- v1.1.0 preparation: the packaged-worker real Office smoke test passed all seven scenarios (five successful exports plus the expected isolated failures for password-protected and damaged files), including modification-protected PPT/PPTX, same-stem files, Chinese paths, and RTF. Release automation derives package names from the installer version and stages portable/setup executables and their SHA-256 files in a Release draft for verification before publication.

## Maintenance rule

Future changes should describe one concrete behavior, include the smallest relevant verification step, and update this document whenever the project boundary changes.

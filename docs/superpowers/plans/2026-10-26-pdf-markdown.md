# PDF Markdown Implementation Plan

**Goal:** Add local Microsoft MarkItDown PDF conversion to the existing desktop utility.

**Architecture:** A dedicated spawned worker writes Markdown to a staged file. Batch conversion returns the existing result model; the GUI captures the conversion mode in each job and reuses progress and retry handling.

**Tech Stack:** Python, Tkinter, multiprocessing, markitdown[pdf] 0.1.9, PyInstaller.

**Spec:** ../specs/2026-10-26-pdf-markdown-design.md

## Constraints and review focus

Keep PDF sources; never activate Office moving/deletion settings in Markdown mode. Preserve existing Markdown on empty output, cancellation or errors. Capture mode for retry. Test corrupted/encrypted/blank PDFs, Unicode paths and timeout cleanup. Package dynamic converters and distribution metadata.

## Tasks

- [x] Add failing backend tests for recursion, source retention, atomic replacement, retry validation and empty output.
- [x] Implement `markdown_backend.convert_pdf_folder_report` and process-isolated conversion with cancellable deadlines.
- [x] Add failing GUI tests for mode routing and retries; implement mode selection and mode-specific settings.
- [x] Pin PDF dependencies, update both build paths and user documentation.
- [x] Run full unittest suite (40 passed), real English/Chinese MarkItDown PDF checks and packaged-worker checks; independent review found no blocking issue. Smoke fixtures use temporary directories and are removed automatically.

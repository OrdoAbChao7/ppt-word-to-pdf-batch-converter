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

- Local Windows / Microsoft Word and PowerPoint 16.0: real conversion checks cover DOCX, PPTX, RTF, Chinese/space-containing paths, nested folders, and identical source stems. Password-protected DOCX and damaged PPTX fail independently without producing a final PDF or deleting their source.
- `python -m unittest discover -v` covers atomic replacement, invalid PDFs, retries, failure totals, output collisions, changed-source retention, timeout/cancellation, shared Office ownership, and GUI retry/progress behavior. GUI tests skip explicitly if Tcl/Tk cannot access a desktop.
- `python -m tests.office_smoke` generates temporary real Office fixtures, checks PDF pages/text and per-file outcomes, and removes the fixtures afterwards. It requires a normal Windows desktop with Word and PowerPoint installed.
- pywin32 COM cache corruption (`CLSIDToClassMap`) was reproduced locally. The application now uses recursive dynamic dispatch without deleting or rebuilding the user's shared cache.
- PDF verification checks file structure and page presence, not visual fidelity. Existing target PDFs are replaced only after validation. Different source formats with the same stem receive names such as `report.docx.pdf` and `report.pptx.pdf`.
- A healthy Office instance is reused for speed. After failure the worker reconnects once. The default per-attempt deadline is 120 seconds, adjustable in the GUI; a file can therefore take two deadlines before being marked failed. Only a newly identified Office process can be terminated on timeout; existing user Office processes are preserved.
- The rebuilt portable executable passed GUI startup/clean shutdown and the same real Office smoke scenarios using its bundled worker entry point (`python -m tests.office_smoke --executable dist/PPT_Word_to_PDF.exe`).

## Maintenance rule

Future changes should describe one concrete behavior, include the smallest relevant verification step, and update this document whenever the project boundary changes.

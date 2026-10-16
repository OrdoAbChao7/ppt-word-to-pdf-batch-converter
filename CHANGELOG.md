# Changelog

This file starts tracking repository-facing changes from the current maintenance pass. Earlier commits remain available in Git history.

## v1.1.0 — 2026-10-16

- Simplify the light desktop interface: collapsible advanced settings, visible source-file change summaries, contextual cancel/retry actions, results-first completion, keyboard shortcuts, and improved small-window layout.
- Publish matching portable and installer packages with SHA-256 checksums. Derive installer metadata and release asset names from one version definition and reject mismatched version tags.
- Fix read-only export failures on modification-protected PPT/PPTX files: import slides into a temporary PowerPoint presentation, preserve slide size, verify imported slide counts, and export without modifying source files or protection settings. Record compatibility handling in the conversion log.
- Verify the reported 34-file course folder against real Office: 34 successful conversions, matching visible-slide counts/page sizes, and unchanged source SHA-256 hashes. Add generated write-protected PPT/PPTX smoke fixtures.
- Bypass damaged pywin32 generated caches for application and returned COM objects.
- Isolate Office conversion in reusable worker processes with a per-file timeout, bounded shutdown, cancellation, and one reconnect/retry.
- Report per-file failures accurately, reserve distinct PDF names across Word/PowerPoint formats, and export via validated temporary PDFs and atomic replacement.
- Preserve sources by default; delete only after PDF structure/page checks, retaining any source changed during conversion.
- Add GUI failure retry, detailed results, log export, adjustable timeouts, and consistent batch progress.
- Add regression tests and a real Office smoke test; rebuild installer executables from current source instead of bundling a stale binary.
- Added explicit project status, scope, known limitations, and next verification work in `PROJECT_STATUS.md`.

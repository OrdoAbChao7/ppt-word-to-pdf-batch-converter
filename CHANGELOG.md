# Changelog

This file starts tracking repository-facing changes from the current maintenance pass. Earlier commits remain available in Git history.

## Unreleased

- Bypass damaged pywin32 generated caches for application and returned COM objects.
- Isolate Office conversion in reusable worker processes with a per-file timeout, bounded shutdown, cancellation, and one reconnect/retry.
- Report per-file failures accurately, reserve distinct PDF names across Word/PowerPoint formats, and export via validated temporary PDFs and atomic replacement.
- Preserve sources by default; delete only after PDF structure/page checks, retaining any source changed during conversion.
- Add GUI failure retry, detailed results, log export, adjustable timeouts, and consistent batch progress.
- Add regression tests and a real Office smoke test; rebuild installer executables from current source instead of bundling a stale binary.
- Added explicit project status, scope, known limitations, and next verification work in `PROJECT_STATUS.md`.

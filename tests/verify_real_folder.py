"""Verify a real folder without writing to or deleting any source files.

python -m tests.verify_real_folder --source SOURCE --output OUTPUT [--executable EXE]
Outputs PDFs and a JSON report under OUTPUT. Requires a Windows Office desktop.
"""

import argparse
import hashlib
import json
import multiprocessing
import sys
from pathlib import Path
from types import SimpleNamespace

import pythoncom
from pypdf import PdfReader

from office_backend import OfficeSession, OfficeWorker
from run import PPT_EXTENSIONS, DOC_EXTENSIONS, _iter_files, _output_paths, _validate_pdf


def digest(path):
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def verify(source, output, executable=None):
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output or source in output.parents:
        raise ValueError("Verification output must be separate from the source folder")
    output.mkdir(parents=True, exist_ok=True)
    files = [source / relative for relative in _iter_files(str(source), PPT_EXTENSIONS + DOC_EXTENSIONS, True)]
    targets = _output_paths([str(path) for path in files])
    if executable:
        multiprocessing.set_executable(str(Path(executable).resolve()))
        sys.frozen = True
    pythoncom.CoInitialize()
    inspector = None
    workers = {}
    results = []
    try:
        for index, path in enumerate(files, 1):
            print(f"[{index}/{len(files)}] {path.name}", flush=True)
            item = {"file": str(path.relative_to(source)), "status": "failed"}
            before = digest(path)
            try:
                kind = "ppt" if path.suffix.lower() in PPT_EXTENSIONS else "doc"
                expected = None
                dimensions = None
                if kind == "ppt":
                    if inspector is None:
                        inspector = OfficeSession("ppt", SimpleNamespace(send=lambda _: None))
                    presentation = inspector.app.Presentations.Open(str(path), True, False, False)
                    try:
                        expected = sum(not presentation.Slides.Item(i).SlideShowTransition.Hidden
                                       for i in range(1, presentation.Slides.Count + 1))
                        dimensions = (presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight)
                    finally:
                        presentation.Close()
                target = output / Path(targets[str(path)]).relative_to(source)
                target.parent.mkdir(parents=True, exist_ok=True)
                if kind not in workers:
                    workers[kind] = OfficeWorker(kind, timeout=120)
                notice = workers[kind].convert(str(path), str(target))
                _validate_pdf(target)
                reader = PdfReader(target)
                pages = len(reader.pages)
                if expected is not None and pages != expected:
                    raise AssertionError(f"Expected {expected} visible slides, PDF has {pages} pages")
                if dimensions and any(abs(float(actual) - float(wanted)) > 2 for actual, wanted in zip(
                    (reader.pages[0].mediabox.width, reader.pages[0].mediabox.height), dimensions
                )):
                    raise AssertionError("PDF page dimensions differ from source slide size")
                item.update(status="success", output=str(target), pages=pages,
                            expected_pages=expected, notice=notice,
                            text_characters=sum(len(page.extract_text()) for page in reader.pages))
                print(f"  OK: {pages} pages" + ("; readonly compatibility" if notice else ""), flush=True)
            except Exception as exc:
                item["error"] = str(exc)
                print(f"  FAILED: {exc}", flush=True)
                for worker in workers.values():
                    worker.close()
                workers.clear()
            item["source_unchanged"] = digest(path) == before
            if not item["source_unchanged"]:
                item.update(status="failed", error="Source bytes changed during verification")
            results.append(item)
    finally:
        for worker in workers.values():
            worker.close()
        if inspector:
            inspector.close()
        pythoncom.CoUninitialize()
    succeeded = sum(item["status"] == "success" for item in results)
    report = {"total": len(files), "succeeded": succeeded, "failed": len(files) - succeeded, "results": results}
    (output / "verification-report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"RESULT: {succeeded}/{len(files)} succeeded; sources unchanged: {all(item['source_unchanged'] for item in results)}")
    return 0 if succeeded == len(files) else 1


if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--executable")
    args = parser.parse_args()
    raise SystemExit(verify(args.source, args.output, args.executable))

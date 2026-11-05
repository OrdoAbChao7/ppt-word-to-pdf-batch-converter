"""Verify MarkItDown conversion using source or a bundled GUI worker."""

import argparse
import multiprocessing
import sys
import tempfile
from pathlib import Path

from markdown_backend import convert_pdf_folder_report
from tests.test_markdown import create_text_pdf


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--executable")
    args = parser.parse_args()
    if args.executable:
        multiprocessing.set_executable(str(Path(args.executable).resolve()))
        sys.frozen = True
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        source = root / "中文 课程.pdf"
        create_text_pdf(source)
        original = source.read_bytes()
        (root / "broken.pdf").write_bytes(b"not a PDF")
        report = convert_pdf_folder_report(directory, timeout=45)
        assert (report.succeeded, report.failed) == (1, 1), report
        text = source.with_suffix(".md").read_text(encoding="utf-8")
        assert "MarkItDown PDF conversion" in text, repr(text)
        assert "中文课程转换" in text, repr(text)
        assert source.read_bytes() == original
        assert not list(root.glob(".md-convert-*"))
    print("MarkItDown smoke passed: English/Chinese text, Unicode path, damaged PDF, source retention.")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()

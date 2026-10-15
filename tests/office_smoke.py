"""Optional real Office check: python -m tests.office_smoke (Windows desktop)."""

import argparse
import multiprocessing
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace

import pythoncom
from pypdf import PdfReader

from office_backend import OfficeSession
from run import convert_folder_report


def main(executable=None):
    pythoncom.CoInitialize()
    try:
        with tempfile.TemporaryDirectory(prefix="office-smoke-", dir=Path(__file__).resolve().parents[1]) as temporary:
            root = Path(temporary)
            nested = root / "中文 空格目录"
            nested.mkdir()
            for kind in ("doc", "ppt"):
                session = OfficeSession(kind, SimpleNamespace(send=lambda _: None))
                try:
                    if kind == "doc":
                        document = session.app.Documents.Add()
                        try:
                            document.Content.Text = "转换验证 Office to PDF\n第二行：中文与空格路径。"
                            document.SaveAs2(str(nested / "同名.docx"), FileFormat=16)
                            document.SaveAs2(str(root / "word.rtf"), FileFormat=6)
                        finally:
                            document.Close(0)
                        protected = session.app.Documents.Add()
                        try:
                            protected.Content.Text = "Password-protected fixture"
                            protected.Password = "test123"
                            protected.SaveAs2(str(root / "密码保护.docx"), FileFormat=16)
                        finally:
                            protected.Close(0)
                    else:
                        presentation = session.app.Presentations.Add(False)
                        try:
                            slide = presentation.Slides.Add(1, 1)
                            slide.Shapes.Title.TextFrame.TextRange.Text = "转换验证 / PDF smoke test"
                            presentation.SaveAs(str(nested / "同名.pptx"), 24)
                        finally:
                            presentation.Close()
                finally:
                    session.close()
            assert (root / "密码保护.docx").read_bytes().startswith(b"\xd0\xcf\x11\xe0"), "Password fixture must be encrypted"
            (root / "损坏.pptx").write_bytes(b"not an office document\x00\xff")
            if executable:
                # Exercise the bundled worker entry point rather than spawning
                # the development Python. This mirrors Windows frozen startup.
                multiprocessing.set_executable(str(Path(executable).resolve()))
                sys.frozen = True
            report = convert_folder_report(str(root), recursive=True, timeout=30)
            assert (report.succeeded, report.failed) == (3, 2), report
            actual = {Path(item.source).name: item.status for item in report.results}
            assert actual == {"word.rtf": "success", "同名.docx": "success", "同名.pptx": "success",
                              "密码保护.docx": "failed", "损坏.pptx": "failed"}, actual
            assert not report.cancelled
            for item in report.results:
                assert Path(item.source).exists(), item
                if item.status == "success":
                    reader = PdfReader(item.output)
                    assert len(reader.pages) == 1, item
                    assert "smoke test" in reader.pages[0].extract_text() or "Office to PDF" in reader.pages[0].extract_text(), item
                else:
                    assert not Path(item.output).exists(), item
            print("PASS: real Office, Chinese paths, nested folders, same stems, RTF, password and damaged files"
                  + (" (packaged executable workers)" if executable else ""))
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser()
    parser.add_argument("--executable", help="Use a packaged GUI executable for conversion workers")
    main(parser.parse_args().executable)

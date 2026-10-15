import tempfile
import time
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import patch

import gui
import run
from tests.test_conversion import FakeWorker, pdf_bytes


class GuiTests(unittest.TestCase):
    def setUp(self):
        try:
            self.app = gui.PdfConverterApp()
        except tk.TclError as exc:
            self.skipTest(f"Desktop Tcl/Tk unavailable: {exc}")
        self.app.withdraw()
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        for target in ("gui.messagebox.showwarning", "gui.messagebox.showinfo", "gui.messagebox.showerror"):
            mock = patch(target)
            mock.start()
            self.addCleanup(mock.stop)
        worker = patch("run.OfficeWorker", FakeWorker)
        worker.start()
        self.addCleanup(worker.stop)
        FakeWorker.exporter = None

    def tearDown(self):
        if hasattr(self, "app"):
            self.app.cancel_event.set()
            if self.app.worker:
                self.app.worker.join(5)
            for after_id in self.app.tk.call("after", "info"):
                self.app.after_cancel(after_id)
            self.app.destroy()

    def wait_for_result(self):
        deadline = time.monotonic() + 5
        while time.monotonic() < deadline:
            self.app.update()
            if self.app.last_report is not None:
                return
            time.sleep(0.01)
        self.fail("GUI never displayed conversion results")

    def test_failure_results_retry_and_updated_timeout(self):
        (self.folder / "good.docx").write_bytes(b"source")
        (self.folder / "bad.pptx").write_bytes(b"source")
        def export(src, dst):
            if Path(src).name == "bad.pptx":
                raise RuntimeError("fixture failure")
            Path(dst).write_bytes(pdf_bytes())
        FakeWorker.exporter = staticmethod(export)
        self.app.folder_var.set(str(self.folder))
        self.app.start_button.invoke()
        self.wait_for_result()
        self.assertEqual(self.app.last_report.failed, 1)
        self.assertEqual(len(self.app.result_tree.get_children()), 2)
        self.assertEqual(str(self.app.retry_button["state"]), "normal")
        self.assertIn("失败 1", self.app.status_var.get())
        self.assertEqual(float(self.app.progress["value"]), 2)
        original = (self.folder / "good.pdf").stat().st_mtime_ns
        FakeWorker.exporter = None
        self.app.timeout_var.set("360")
        with patch("gui.convert_folder_report", wraps=run.convert_folder_report) as convert:
            self.app.retry_button.invoke()
            self.wait_for_result()
            self.assertEqual(convert.call_args.kwargs["timeout"], 360)
        self.assertEqual(self.app.last_report.succeeded, 1)
        self.assertEqual(self.app.last_report.failed, 0)
        self.assertEqual((self.folder / "good.pdf").stat().st_mtime_ns, original)
        self.assertEqual(str(self.app.retry_button["state"]), "disabled")

    def test_retry_honors_current_source_deletion_option(self):
        source = self.folder / "source.docx"
        source.write_bytes(b"source")
        FakeWorker.exporter = staticmethod(lambda src, dst: (_ for _ in ()).throw(RuntimeError("failure")))
        self.app.folder_var.set(str(self.folder))
        self.app.delete_source_var.set(True)
        self.app.start_button.invoke()
        self.wait_for_result()
        self.assertEqual(self.app.last_report.failed, 1)
        self.app.delete_source_var.set(False)
        FakeWorker.exporter = None
        self.app.retry_button.invoke()
        self.wait_for_result()
        self.assertEqual(self.app.last_report.succeeded, 1)
        self.assertTrue(source.exists())

    def test_empty_path_does_not_convert_current_directory(self):
        self.app.folder_var.set("")
        self.app.start_button.invoke()
        self.assertIsNone(self.app.worker)


if __name__ == "__main__":
    unittest.main()

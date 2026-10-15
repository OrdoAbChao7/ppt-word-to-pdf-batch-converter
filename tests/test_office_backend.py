import multiprocessing
import time
import unittest
import pythoncom
from types import SimpleNamespace
from unittest.mock import patch

from office_backend import ConversionCancelled, OfficeSession, OfficeWorker, _application_pid


def stalled_worker(connection):
    try:
        connection.recv()
        time.sleep(20)
    finally:
        connection.close()


class BackendTests(unittest.TestCase):
    def test_readonly_export_imports_all_slides_with_original_size_without_changing_source(self):
        source_closed = []
        copy_closed = []
        source = SimpleNamespace(
            ReadOnly=-1,
            Slides=SimpleNamespace(Count=3),
            PageSetup=SimpleNamespace(SlideWidth=960, SlideHeight=540),
            ExportAsFixedFormat=lambda *a, **kw: (_ for _ in ()).throw(
                pythoncom.com_error(-2147352567, "Exception", (0, "Microsoft PowerPoint", "该文件为只读文件，无法导出。", "", 0, -2147467259), None)),
            Close=lambda: source_closed.append(True),
        )
        observed = []
        working = SimpleNamespace(PageSetup=SimpleNamespace(SlideWidth=720, SlideHeight=540))
        def insert(path, position):
            self.assertEqual((path, position), ("readonly.pptx", 0))
            return 3
        working.Slides = SimpleNamespace(InsertFromFile=insert, Count=3)
        working.ExportAsFixedFormat = lambda path, fmt, **kw: observed.append(
            (path, fmt, working.PageSetup.SlideWidth, working.PageSetup.SlideHeight))
        working.Close = lambda: copy_closed.append(True)
        session = OfficeSession.__new__(OfficeSession)
        session.kind = "ppt"
        session.app = SimpleNamespace(Presentations=SimpleNamespace(Open=lambda *a: source, Add=lambda *a: working))
        session.export("readonly.pptx", "result.pdf")
        self.assertEqual(observed, [("result.pdf", 2, 960, 540)])
        self.assertEqual(source.ReadOnly, -1)
        self.assertEqual(source_closed, [True])
        self.assertEqual(copy_closed, [True])

    def test_unrelated_export_error_is_not_treated_as_readonly(self):
        error = pythoncom.com_error(-2147352567, "Exception", (0, "Microsoft PowerPoint", "Disk full", "", 0, -2147467259), None)
        document = SimpleNamespace(ReadOnly=-1, Close=lambda: None,
                                   ExportAsFixedFormat=lambda *a, **kw: (_ for _ in ()).throw(error))
        session = OfficeSession.__new__(OfficeSession)
        session.kind = "ppt"
        session.app = SimpleNamespace(Presentations=SimpleNamespace(
            Open=lambda *a: document, Add=lambda *a: self.fail("Unrelated errors must not import slides")))
        with self.assertRaises(pythoncom.com_error) as caught:
            session.export("source.pptx", "result.pdf")
        self.assertIs(caught.exception, error)

    def test_incomplete_readonly_import_is_rejected_and_closed(self):
        closed = []
        working = SimpleNamespace(PageSetup=SimpleNamespace(SlideWidth=0, SlideHeight=0),
                                  Slides=SimpleNamespace(Count=2, InsertFromFile=lambda *a: 2),
                                  Close=lambda: closed.append(True),
                                  ExportAsFixedFormat=lambda *a, **kw: self.fail("Incomplete slides must not be exported"))
        source = SimpleNamespace(PageSetup=SimpleNamespace(SlideWidth=960, SlideHeight=540), Slides=SimpleNamespace(Count=3))
        session = OfficeSession.__new__(OfficeSession)
        session.app = SimpleNamespace(Presentations=SimpleNamespace(Add=lambda *a: working))
        with self.assertRaisesRegex(RuntimeError, "导入不完整"):
            session._export_readonly_presentation(source, "source.ppt", "result.pdf")
        self.assertEqual(closed, [True])

    def test_unavailable_powerpoint_hwnd_falls_back_to_actual_window_and_restores_caption(self):
        application = SimpleNamespace(Caption="PowerPoint")
        application.HWND = lambda: (_ for _ in ()).throw(TypeError("HWND unsupported"))
        with patch("win32gui.EnumWindows", side_effect=lambda callback, data: callback(123, data)), \
             patch("win32gui.GetWindowText", side_effect=lambda hwnd: application.Caption), \
             patch("office_backend.win32process.GetWindowThreadProcessId", return_value=(1, 20)):
            self.assertEqual(_application_pid(application, "ppt"), 20)
        self.assertEqual(application.Caption, "PowerPoint")

    def test_powerpoint_method_window_handle_identifies_actual_process(self):
        application = SimpleNamespace(HWND=lambda: 123, AutomationSecurity=1, DisplayAlerts=2, Visible=1,
                                      Quit=lambda: None)
        def window_process(hwnd):
            if hwnd != 123:
                raise TypeError("Window handle must be an integer")
            return 1, 20
        with patch("office_backend.pythoncom.CoCreateInstanceEx", return_value=(application,)), \
             patch("office_backend._dynamic_dispatch", side_effect=lambda obj: obj), \
             patch("office_backend.win32process.EnumProcesses", return_value=[]), \
             patch("office_backend.win32process.GetWindowThreadProcessId", side_effect=window_process):
            session = OfficeSession("ppt", SimpleNamespace(send=lambda _: None))
            self.assertTrue(session.owned)
            session.close()

    def test_concurrent_powerpoint_start_does_not_own_existing_application(self):
        application = SimpleNamespace(HWND=123, AutomationSecurity=1, DisplayAlerts=2, Visible=1,
                                      user_document_open=True)
        application.Quit = lambda: setattr(application, "user_document_open", False)
        with patch("office_backend.pythoncom.CoCreateInstanceEx", return_value=(application,)), \
             patch("office_backend._dynamic_dispatch", side_effect=lambda obj: obj), \
             patch("office_backend.win32process.EnumProcesses", side_effect=[[10], [10, 20]]), \
             patch("office_backend._find_new_office_pid", return_value=20, create=True), \
             patch("office_backend.win32process.GetWindowThreadProcessId", return_value=(1, 10)):
            session = OfficeSession("ppt", SimpleNamespace(send=lambda _: None))
            session.close()
            self.assertTrue(application.user_document_open)

    def test_word_starts_without_hwnd_property_and_disables_macros(self):
        application = SimpleNamespace(AutomationSecurity=1, DisplayAlerts=-1, Quit=lambda: None)
        connection = SimpleNamespace(send=lambda message: None)
        with patch("office_backend.pythoncom.CoCreateInstanceEx", return_value=(application,)), \
             patch("office_backend._dynamic_dispatch", side_effect=lambda obj: obj), \
             patch("office_backend.win32process.EnumProcesses", return_value=[]), \
             patch("office_backend._application_pid", return_value=20):
            session = OfficeSession("doc", connection)
            self.assertTrue(session.owned)
            self.assertEqual(application.AutomationSecurity, 3)
            self.assertEqual(application.DisplayAlerts, 0)
            session.close()

    def test_shared_powerpoint_is_restored_without_closing_user_application(self):
        application = SimpleNamespace(HWND=123, AutomationSecurity=1, DisplayAlerts=2, Visible=1)
        application.Quit = lambda: self.fail("Must not quit a pre-existing Office process")
        with patch("office_backend.pythoncom.CoCreateInstanceEx", return_value=(application,)), \
             patch("office_backend._dynamic_dispatch", side_effect=lambda obj: obj), \
             patch("office_backend.win32process.EnumProcesses", return_value=[20]), \
             patch("office_backend._application_pid", return_value=20):
            session = OfficeSession("ppt", SimpleNamespace(send=lambda _: None))
            self.assertFalse(session.owned)
            session.close()
            self.assertEqual(application.AutomationSecurity, 1)
            self.assertEqual(application.DisplayAlerts, 2)

    def start_stalled_worker(self, timeout):
        worker = OfficeWorker("doc", timeout)
        context = multiprocessing.get_context("spawn")
        worker.connection, child = context.Pipe()
        worker.process = context.Process(target=stalled_worker, args=(child,))
        worker.process.start()
        child.close()
        process = worker.process
        self.addCleanup(worker.close)
        return worker, process

    def test_hung_conversion_times_out_and_reaps_worker(self):
        worker, process = self.start_stalled_worker(0.2)
        started = time.monotonic()
        with self.assertRaises(TimeoutError):
            worker.convert("source", "target")
        self.assertLess(time.monotonic() - started, 4)
        self.assertIsNone(worker.process)
        self.assertIsNone(worker.connection)

    def test_cancellation_interrupts_hung_conversion(self):
        worker, process = self.start_stalled_worker(30)
        started = time.monotonic()
        with self.assertRaises(ConversionCancelled):
            worker.convert("source", "target", lambda: True)
        self.assertLess(time.monotonic() - started, 4)
        self.assertIsNone(worker.process)


if __name__ == "__main__":
    unittest.main()

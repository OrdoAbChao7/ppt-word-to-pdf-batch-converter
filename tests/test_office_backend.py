import multiprocessing
import time
import unittest
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

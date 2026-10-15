"""Office automation in a restartable process, independent of the Tk UI."""

import multiprocessing
import time
import uuid

import pythoncom
import win32api
import win32com.client.dynamic
import win32event
import win32gui
import win32process


class ConversionCancelled(Exception):
    pass


class OfficeUnavailable(RuntimeError):
    pass


class _DynamicDispatch(win32com.client.dynamic.CDispatch):
    def _wrap_dispatch_(self, obj, userName=None, returnCLSID=None):
        return _dynamic_dispatch(obj, userName)

    def _make_method_(self, name):
        method = super()._make_method_(name)
        if method is not None:
            # pywin32's generated dynamic methods otherwise call the top-level
            # Dispatch again for returned Documents/Presentations/Shapes.
            method.__func__.__globals__["Dispatch"] = _dynamic_dispatch
        return method


def _dynamic_dispatch(obj, userName=None, resultCLSID=None):
    return win32com.client.dynamic.Dispatch(obj, userName, createClass=_DynamicDispatch)


def _application_pid(app, kind):
    probe = None
    try:
        if kind == "ppt":
            try:
                hwnd = _com_call(lambda: app.HWND)
                if callable(hwnd):
                    hwnd = _com_call(hwnd)
                return win32process.GetWindowThreadProcessId(hwnd)[1]
            except Exception:
                return _caption_pid(app)
        else:
            # Word exposes HWND on Window, not Application. A temporary blank
            # document gives the PID of THIS COM instance, even with other Word
            # instances starting concurrently. Macros are already disabled.
            probe = _com_call(lambda: app.Documents.Add(Visible=False))
            hwnd = _com_call(lambda: probe.Windows.Item(1).Hwnd)
        # Some PowerPoint type libraries expose HWND as a zero-argument method.
        if callable(hwnd):
            hwnd = _com_call(hwnd)
        return win32process.GetWindowThreadProcessId(hwnd)[1]
    except Exception:
        return None
    finally:
        if probe is not None:
            _com_call(lambda: probe.Close(0))


def _caption_pid(app):
    """Some PowerPoint builds advertise HWND but return MEMBERNOTFOUND."""
    previous = _com_call(lambda: app.Caption)
    marker = "PDF-convert-" + uuid.uuid4().hex
    try:
        _com_call(lambda: setattr(app, "Caption", marker))
        deadline = time.monotonic() + 1
        while True:
            matches = set()
            def inspect_window(hwnd, _):
                if marker in win32gui.GetWindowText(hwnd):
                    matches.add(win32process.GetWindowThreadProcessId(hwnd)[1])
            win32gui.EnumWindows(inspect_window, None)
            if matches:
                return next(iter(matches)) if len(matches) == 1 else None
            if time.monotonic() >= deadline:
                return None
            pythoncom.PumpWaitingMessages()
            time.sleep(0.05)
    finally:
        _com_call(lambda: setattr(app, "Caption", previous))


def _com_call(operation):
    """Office can briefly reject calls while starting or repaginating."""
    for attempt in range(4):
        try:
            return operation()
        except pythoncom.com_error as exc:
            if exc.hresult not in (-2147418111, -2147417846) or attempt == 3:
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(0.25 * (attempt + 1))


class OfficeSession:
    def __init__(self, kind, connection):
        self.kind = kind
        self.app = None
        self.owned = False
        self.previous = {}
        name = "PowerPoint" if kind == "ppt" else "Word"
        existing_pids = set(win32process.EnumProcesses())
        try:
            # Dispatch/DispatchEx consult gen_py. Raw IDispatch deliberately avoids
            # stale or incomplete generated caches (CLSIDToClassMap failures).
            raw = _com_call(lambda: pythoncom.CoCreateInstanceEx(
                f"{name}.Application", None, pythoncom.CLSCTX_LOCAL_SERVER,
                None, (pythoncom.IID_IDispatch,),
            )[0])
            self.app = _dynamic_dispatch(raw)
            settings = {"AutomationSecurity": 3, "DisplayAlerts": 1 if kind == "ppt" else 0}
            if kind == "ppt":
                settings["Visible"] = 1
            for key, value in settings.items():
                self.previous[key] = _com_call(lambda key=key: getattr(self.app, key))
                _com_call(lambda key=key, value=value: setattr(self.app, key, value))
            pid = _application_pid(self.app, kind)
            self.owned = pid is not None and pid not in existing_pids
            connection.send(("office", pid, self.owned))
        except Exception as exc:
            self.close()
            raise OfficeUnavailable(
                f"无法启动或配置 Microsoft {name}。请确认已安装并可手动打开，"
                f"完成首次启动/激活提示后重试。原始错误：{exc}"
            ) from exc

    def export(self, source, target):
        document = None
        notice = ""
        try:
            if self.kind == "ppt":
                document = _com_call(lambda: self.app.Presentations.Open(source, True, False, False))
                # pywin32 cannot marshal pythoncom.Missing into the optional
                # PrintRange IDispatch parameter; pass a null interface explicitly.
                try:
                    _com_call(lambda: document.ExportAsFixedFormat(target, 2, PrintRange=None))
                except pythoncom.com_error as exc:
                    description = (exc.excepinfo[2] or "") if exc.excepinfo else ""
                    read_only = _com_call(lambda: document.ReadOnly)
                    read_only = read_only() if callable(read_only) else read_only
                    if not read_only or not any(word in description.lower() for word in ("只读", "read-only", "read only")):
                        raise
                    self._export_readonly_presentation(document, source, target)
                    notice = "源课件限制直接导出，已通过临时演示文稿生成 PDF；原课件及保护设置未修改。"
            else:
                document = _com_call(lambda: self.app.Documents.Open(
                    FileName=source, ReadOnly=True, ConfirmConversions=False,
                    AddToRecentFiles=False, Visible=False, OpenAndRepair=True,
                    NoEncodingDialog=True, PasswordDocument="__pdf_converter_no_password__",
                    PasswordTemplate="__pdf_converter_no_password__",
                    WritePasswordDocument="__pdf_converter_no_password__",
                    WritePasswordTemplate="__pdf_converter_no_password__",
                ))
                _com_call(lambda: document.ExportAsFixedFormat(
                    OutputFileName=target, ExportFormat=17, OpenAfterExport=False,
                ))
        finally:
            if document is not None:
                if self.kind == "ppt":
                    _com_call(lambda: document.Close())
                else:
                    _com_call(lambda: document.Close(0))
        return notice

    def _export_readonly_presentation(self, source_document, source, target):
        working = _com_call(lambda: self.app.Presentations.Add(False))
        try:
            # Use PowerPoint's own import operation so the original slide masters,
            # objects and notes are imported rather than rasterizing slide images.
            width = _com_call(lambda: source_document.PageSetup.SlideWidth)
            height = _com_call(lambda: source_document.PageSetup.SlideHeight)
            _com_call(lambda: setattr(working.PageSetup, "SlideWidth", width))
            _com_call(lambda: setattr(working.PageSetup, "SlideHeight", height))
            expected = _com_call(lambda: source_document.Slides.Count)
            imported = _com_call(lambda: working.Slides.InsertFromFile(source, 0))
            actual = _com_call(lambda: working.Slides.Count)
            if imported != expected or actual != expected:
                raise RuntimeError(f"只读课件导入不完整：原有 {expected} 页，导入 {imported} 页，实际 {actual} 页；已保留源文件。")
            _com_call(lambda: working.ExportAsFixedFormat(target, 2, PrintRange=None))
        finally:
            _com_call(lambda: working.Close())

    def close(self):
        if self.app is None:
            return
        try:
            if self.owned:
                _com_call(lambda: self.app.Quit())
            else:
                # PowerPoint may reuse an interactive instance. Do not quit it.
                for key, value in self.previous.items():
                    try:
                        _com_call(lambda key=key, value=value: setattr(self.app, key, value))
                    except Exception:
                        pass
        except Exception:
            pass
        finally:
            self.app = None


def _office_main(kind, connection):
    pythoncom.CoInitialize()
    session = None
    try:
        session = OfficeSession(kind, connection)
        connection.send(("ready",))
        while True:
            command = connection.recv()
            if command[0] == "stop":
                break
            try:
                notice = session.export(command[1], command[2])
                connection.send(("done", notice))
            except Exception as exc:
                connection.send(("error", str(exc)))
                break
    except OfficeUnavailable as exc:
        connection.send(("unavailable", str(exc)))
    except (EOFError, BrokenPipeError, OSError):
        pass
    finally:
        if session is not None:
            session.close()
        connection.close()
        pythoncom.CoUninitialize()


class OfficeWorker:
    """Reuse a healthy Office instance; bound startup, export, and shutdown."""

    def __init__(self, kind, timeout=120):
        self.kind = kind
        self.timeout = timeout
        self.process = None
        self.connection = None
        self.office_handle = None

    def _receive(self, deadline, cancel_callback):
        while True:
            if cancel_callback and cancel_callback():
                raise ConversionCancelled()
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise TimeoutError(f"单文件处理超过 {self.timeout:g} 秒；可调大超时时间，或先用 Office 手动检查此文件。")
            if self.connection.poll(min(0.1, remaining)):
                try:
                    message = self.connection.recv()
                except EOFError as exc:
                    raise RuntimeError("Office 转换进程意外退出，将重新连接后重试。") from exc
                if message[0] == "office":
                    if message[2]:
                        # A held handle avoids killing an unrelated process if a PID
                        # gets reused. Never kill all WINWORD/POWERPNT processes.
                        self.office_handle = win32api.OpenProcess(0x100001, False, message[1])
                    continue
                if message[0] == "unavailable":
                    raise OfficeUnavailable(message[1])
                if message[0] == "error":
                    raise RuntimeError(message[1])
                return message
            if not self.process.is_alive():
                raise RuntimeError("Office 转换进程意外退出，将重新连接后重试。")

    def convert(self, source, target, cancel_callback=None):
        deadline = time.monotonic() + self.timeout
        try:
            if self.process is None:
                context = multiprocessing.get_context("spawn")
                self.connection, child = context.Pipe()
                self.process = context.Process(target=_office_main, args=(self.kind, child), daemon=True)
                try:
                    self.process.start()
                finally:
                    child.close()
                self._receive(deadline, cancel_callback)
            self.connection.send(("convert", source, target))
            message = self._receive(deadline, cancel_callback)
            return message[1] if len(message) > 1 else ""
        except Exception:
            self.close()
            raise

    def close(self):
        if self.process is not None:
            try:
                if self.process.is_alive():
                    try:
                        self.connection.send(("stop",))
                    except (OSError, BrokenPipeError):
                        pass
                    self.process.join(1)
                    if self.process.is_alive():
                        self.process.terminate()
                        self.process.join(1)
                self.process.close()
            finally:
                self.process = None
        if self.connection is not None:
            self.connection.close()
            self.connection = None
        if self.office_handle is not None:
            try:
                if win32event.WaitForSingleObject(self.office_handle, 0) == win32event.WAIT_TIMEOUT:
                    win32api.TerminateProcess(self.office_handle, 1)
                    win32event.WaitForSingleObject(self.office_handle, 1000)
            finally:
                self.office_handle.Close()
                self.office_handle = None

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from run import convert_folder


class PdfConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PPT / Word 批量转 PDF (安全增强版)")
        self.geometry("780x590")
        self.minsize(700, 500)

        self.folder_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=True)
        self.flatten_var = tk.BooleanVar(value=False)
        self.delete_source_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="请选择要处理的文件夹")
        self.log_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel_event = threading.Event()

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._drain_log_queue)

    def _on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("提示", "当前有转换任务正在进行中，确认要退出吗？"):
                return
            self.cancel_event.set()
        self.destroy()

    def _build_ui(self) -> None:
        self.configure(bg="#f6f7f9")

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f6f7f9")
        style.configure("TLabel", background="#f6f7f9", font=("Microsoft YaHei UI", 10))
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Hint.TLabel", foreground="#5f6b7a")
        style.configure("TButton", font=("Microsoft YaHei UI", 10), padding=(14, 8))
        style.configure("Accent.TButton", font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("TCheckbutton", background="#f6f7f9", font=("Microsoft YaHei UI", 10))

        container = ttk.Frame(self, padding=24)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="PPT / Word 批量转 PDF", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            container,
            text="选择文件夹，一键批量把 PPT、PPTX、PPTM、PPS、PPSX、DOC、DOCX、DOCM、RTF 转成 PDF。",
            style="Hint.TLabel",
        ).pack(anchor=tk.W, pady=(6, 16))

        path_frame = ttk.Frame(container)
        path_frame.pack(fill=tk.X)

        self.path_entry = ttk.Entry(path_frame, textvariable=self.folder_var, font=("Microsoft YaHei UI", 10))
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6)
        ttk.Button(path_frame, text="选择文件夹", command=self._choose_folder).pack(side=tk.LEFT, padx=(10, 0))

        options_frame = ttk.Frame(container)
        options_frame.pack(fill=tk.X, pady=(14, 8))

        ttk.Checkbutton(
            options_frame,
            text="包含子文件夹：原地转换并保留原有文件夹目录结构 (推荐)",
            variable=self.recursive_var,
        ).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(
            options_frame,
            text="整理子文件夹：把子文件夹里的 Office 文件集中移动到当前主文件夹",
            variable=self.flatten_var,
        ).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(
            options_frame,
            text="转换成功后安全删除源文件 (严格校验 PDF 生成完整性后才删除)",
            variable=self.delete_source_var,
        ).pack(anchor=tk.W, pady=2)

        actions_frame = ttk.Frame(container)
        actions_frame.pack(fill=tk.X, pady=(12, 10))

        self.start_button = ttk.Button(
            actions_frame,
            text="开始转换",
            style="Accent.TButton",
            command=self._start_convert,
        )
        self.start_button.pack(side=tk.LEFT)

        self.cancel_button = ttk.Button(
            actions_frame,
            text="取消转换",
            command=self._cancel_convert,
            state=tk.DISABLED,
        )
        self.cancel_button.pack(side=tk.LEFT, padx=(10, 0))

        ttk.Button(actions_frame, text="打开输出文件夹", command=self._open_folder).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Label(actions_frame, textvariable=self.status_var, style="Hint.TLabel").pack(side=tk.LEFT, padx=(16, 0))

        self.progress = ttk.Progressbar(container, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=(4, 12))

        ttk.Label(container, text="处理记录").pack(anchor=tk.W)
        self.log_text = scrolledtext.ScrolledText(
            container,
            height=13,
            wrap=tk.WORD,
            font=("Consolas", 10),
            relief=tk.FLAT,
            borderwidth=1,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.log_text.configure(state=tk.DISABLED)

    def _choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择要转换的文件夹")
        if folder:
            normalized_folder = os.path.normpath(folder)
            self.folder_var.set(normalized_folder)
            self.status_var.set("已选择文件夹")

    def _open_folder(self) -> None:
        folder = os.path.normpath(self.folder_var.get().strip().strip('"'))
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择一个有效的文件夹。")
            return
        os.startfile(folder)

    def _cancel_convert(self) -> None:
        if self.worker and self.worker.is_alive():
            self.cancel_event.set()
            self.cancel_button.configure(state=tk.DISABLED)
            self.status_var.set("正在取消转换，等待当前文件完成...")
            self._append_log(">>> 收到用户取消指令，正在安全退出...")

    def _start_convert(self) -> None:
        folder = os.path.normpath(self.folder_var.get().strip().strip('"'))
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择一个有效的文件夹。")
            return

        if self.worker and self.worker.is_alive():
            return

        self.cancel_event.clear()
        self._set_running(True)
        self._clear_log()
        self._append_log("开始处理，请不要关闭 PowerPoint 或 Word 窗口。")

        self.worker = threading.Thread(
            target=self._convert_in_background,
            args=(
                folder,
                self.recursive_var.get(),
                self.flatten_var.get(),
                self.delete_source_var.get(),
            ),
            daemon=True,
        )
        self.worker.start()

    def _convert_in_background(
        self,
        folder: str,
        recursive: bool,
        flatten: bool,
        delete_source: bool,
    ) -> None:
        try:
            ppt_count, doc_count = convert_folder(
                folder,
                recursive=recursive,
                flatten=flatten,
                delete_source=delete_source,
                log_callback=lambda message: self.log_queue.put(("log", message)),
                cancel_callback=lambda: self.cancel_event.is_set(),
                progress_callback=lambda cur, tot, f: self.log_queue.put(("progress", (cur, tot, f))),
            )
            if self.cancel_event.is_set():
                self.log_queue.put(("cancelled", f"转换已取消：已完成 PPT {ppt_count} 个，Word {doc_count} 个。"))
            else:
                self.log_queue.put(("done", f"全部转换完成：PPT {ppt_count} 个，Word {doc_count} 个。"))
        except Exception as exc:
            self.log_queue.put(("error", str(exc)))

    def _drain_log_queue(self) -> None:
        try:
            while True:
                kind, data = self.log_queue.get_nowait()
                if kind == "log":
                    self._append_log(str(data))
                elif kind == "progress":
                    cur, tot, file_name = data
                    self.progress.stop()
                    self.progress.configure(mode="determinate", maximum=tot, value=cur)
                    self.status_var.set(f"正在转换 ({cur}/{tot})")
                elif kind == "done":
                    self._append_log(str(data))
                    self._set_running(False)
                    messagebox.showinfo("完成", str(data))
                elif kind == "cancelled":
                    self._append_log(str(data))
                    self._set_running(False)
                    messagebox.showwarning("已取消", str(data))
                elif kind == "error":
                    self._append_log(f"处理失败：{data}")
                    self._set_running(False)
                    messagebox.showerror("处理失败", str(data))
        except queue.Empty:
            pass

        self.after(100, self._drain_log_queue)

    def _set_running(self, running: bool) -> None:
        if running:
            self.start_button.configure(state=tk.DISABLED)
            self.cancel_button.configure(state=tk.NORMAL)
            self.progress.configure(mode="indeterminate")
            self.progress.start(10)
            self.status_var.set("正在启动 Office 进行转换...")
        else:
            self.start_button.configure(state=tk.NORMAL)
            self.cancel_button.configure(state=tk.DISABLED)
            self.progress.stop()
            self.status_var.set("就绪")

    def _clear_log(self) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)


if __name__ == "__main__":
    app = PdfConverterApp()
    app.mainloop()

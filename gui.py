import os
import multiprocessing
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from run import ConversionReport, convert_folder_report


class PdfConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PPT / Word 批量转 PDF")
        self.geometry("980x740")
        self.minsize(820, 620)

        self.folder_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=True)
        self.flatten_var = tk.BooleanVar(value=False)
        self.delete_source_var = tk.BooleanVar(value=False)
        self.timeout_var = tk.StringVar(value="120")
        self.status_var = tk.StringVar(value="请选择要处理的文件夹")
        self.log_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel_event = threading.Event()
        self.last_report: ConversionReport | None = None
        self.last_job = None
        self.active_folder = ""
        self.close_requested = False
        self.option_widgets = []

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._drain_log_queue)

    def _on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            if self.close_requested:
                return
            if not messagebox.askyesno("提示", "当前有转换任务正在进行中，确认要退出吗？"):
                return
            self.close_requested = True
            self.cancel_event.set()
            self.status_var.set("正在停止转换并清理 Office，请稍候...")
            return
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
            text="批量转换演示文稿和 Word 文档；默认保留源文件，失败后可单独重试。",
            style="Hint.TLabel",
        ).pack(anchor=tk.W, pady=(6, 16))

        path_frame = ttk.Frame(container)
        path_frame.pack(fill=tk.X)

        self.path_entry = ttk.Entry(path_frame, textvariable=self.folder_var, font=("Microsoft YaHei UI", 10))
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6)
        self.choose_button = ttk.Button(path_frame, text="选择文件夹", command=self._choose_folder)
        self.choose_button.pack(side=tk.LEFT, padx=(10, 0))

        options_frame = ttk.Frame(container)
        options_frame.pack(fill=tk.X, pady=(14, 8))

        for text, variable in (
            ("包含子文件夹：在原文件旁生成 PDF，保留目录结构", self.recursive_var),
            ("整理子文件夹：先将子目录中的 Office 源文件移动到主文件夹", self.flatten_var),
            ("转换成功后删除源文件（PDF 结构和页面校验通过后执行）", self.delete_source_var),
        ):
            widget = ttk.Checkbutton(options_frame, text=text, variable=variable)
            widget.pack(anchor=tk.W, pady=2)
            self.option_widgets.append(widget)
        timeout_frame = ttk.Frame(options_frame)
        timeout_frame.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(timeout_frame, text="单文件超时（秒）").pack(side=tk.LEFT)
        timeout_entry = ttk.Spinbox(timeout_frame, from_=10, to=1800, increment=30,
                                    textvariable=self.timeout_var, width=7)
        timeout_entry.pack(side=tk.LEFT, padx=(8, 12))
        self.option_widgets.append(timeout_entry)
        ttk.Label(timeout_frame, text="超时或失败时重新连接 Office 重试一次；大型文档可调大。",
                  style="Hint.TLabel").pack(side=tk.LEFT)

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

        self.retry_button = ttk.Button(actions_frame, text="仅重试失败文件", command=self._retry_failed,
                                       state=tk.DISABLED)
        self.retry_button.pack(side=tk.LEFT, padx=(10, 0))

        ttk.Button(actions_frame, text="打开输出文件夹", command=self._open_folder).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(actions_frame, text="保存日志", command=self._save_log).pack(side=tk.RIGHT)

        ttk.Label(container, textvariable=self.status_var, style="Hint.TLabel").pack(anchor=tk.W, pady=(4, 4))

        self.progress = ttk.Progressbar(container, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=(4, 12))

        notebook = ttk.Notebook(container)
        notebook.pack(fill=tk.BOTH, expand=True)
        log_frame = ttk.Frame(notebook)
        results_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text="处理记录")
        notebook.add(results_frame, text="文件结果（双击查看详情）")
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=13,
            wrap=tk.WORD,
            font=("Consolas", 10),
            relief=tk.FLAT,
            borderwidth=1,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)
        self.result_tree = ttk.Treeview(results_frame, columns=("status", "file", "detail"), show="headings")
        for column, title, width in (("status", "结果", 65), ("file", "源文件", 260), ("detail", "输出 / 失败原因", 460)):
            self.result_tree.heading(column, text=title)
            self.result_tree.column(column, width=width, stretch=column != "status")
        result_scroll = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        result_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_tree.configure(yscrollcommand=result_scroll.set)
        self.result_tree.pack(fill=tk.BOTH, expand=True)
        self.result_tree.tag_configure("failed", foreground="#b42318")
        self.result_tree.bind("<Double-1>", self._show_result_detail)

    def _choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择要转换的文件夹")
        if folder:
            normalized_folder = os.path.normpath(folder)
            self.folder_var.set(normalized_folder)
            self.status_var.set("已选择文件夹")

    def _open_folder(self) -> None:
        folder = self.folder_var.get().strip().strip('"')
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择一个有效的文件夹。")
            return
        try:
            os.startfile(folder)
        except OSError as exc:
            messagebox.showerror("无法打开文件夹", str(exc))

    def _save_log(self) -> None:
        path = filedialog.asksaveasfilename(title="保存处理日志", defaultextension=".txt",
                                          initialfile="转换日志.txt", filetypes=[("文本文件", "*.txt")])
        if path:
            try:
                with open(path, "w", encoding="utf-8-sig") as stream:
                    stream.write(self.log_text.get("1.0", tk.END))
            except OSError as exc:
                messagebox.showerror("保存失败", str(exc))

    def _show_result_detail(self, event=None) -> None:
        selection = self.result_tree.selection()
        if selection and self.last_report:
            result = self.last_report.results[int(selection[0])]
            detail = result.error or result.warning or "转换成功"
            messagebox.showinfo("文件详情", f"源文件：{result.source}\n输出：{result.output}\n\n{detail}")

    def _cancel_convert(self) -> None:
        if self.worker and self.worker.is_alive():
            self.cancel_event.set()
            self.cancel_button.configure(state=tk.DISABLED)
            self.status_var.set("正在取消转换并清理 Office...")
            self._append_log(">>> 收到用户取消指令，正在安全退出...")

    def _start_convert(self) -> None:
        folder = self.folder_var.get().strip().strip('"')
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择一个有效的文件夹。")
            return

        try:
            timeout = int(self.timeout_var.get())
            if not 10 <= timeout <= 1800:
                raise ValueError()
        except ValueError:
            messagebox.showwarning("超时时间无效", "请输入 10 到 1800 之间的秒数。")
            return
        job = (os.path.abspath(folder), self.recursive_var.get(), self.flatten_var.get(),
               self.delete_source_var.get(), timeout)
        self._start_job(job)

    def _retry_failed(self) -> None:
        if self.last_report and self.last_report.failed_files and self.last_job:
            try:
                timeout = int(self.timeout_var.get())
                if not 10 <= timeout <= 1800:
                    raise ValueError()
            except ValueError:
                messagebox.showwarning("超时时间无效", "请输入 10 到 1800 之间的秒数。")
                return
            outputs = {item.source: item.output for item in self.last_report.results if item.status == "failed"}
            job = (*self.last_job[:3], self.delete_source_var.get(), timeout)
            self._start_job(job, self.last_report.failed_files, outputs)

    def _start_job(self, job, only_files=None, output_paths=None) -> None:

        if self.worker and self.worker.is_alive():
            return

        self.cancel_event.clear()
        self.active_folder = job[0]
        self.folder_var.set(self.active_folder)
        self.last_job = job
        self.last_report = None
        self._set_running(True)
        self._clear_log()
        self.result_tree.delete(*self.result_tree.get_children())
        self._append_log(f"开始处理：{self.active_folder}")
        self._append_log("PDF 将先写入临时文件，校验后替换同名输出。请勿手动关闭转换使用的 Office。")

        self.worker = threading.Thread(
            target=self._convert_in_background,
            args=(*job, only_files, output_paths),
            daemon=False,
        )
        self.worker.start()

    def _convert_in_background(
        self,
        folder: str,
        recursive: bool,
        flatten: bool,
        delete_source: bool,
        timeout: int,
        only_files=None,
        output_paths=None,
    ) -> None:
        try:
            report = convert_folder_report(
                folder,
                recursive=recursive,
                flatten=flatten,
                delete_source=delete_source,
                timeout=timeout,
                only_files=only_files,
                output_paths=output_paths,
                log_callback=lambda message: self.log_queue.put(("log", message)),
                cancel_callback=lambda: self.cancel_event.is_set(),
                progress_callback=lambda cur, tot, f: self.log_queue.put(("progress", (cur, tot, f))),
            )
            self.log_queue.put(("result", report))
        except Exception as exc:
            self.log_queue.put(("error", str(exc)))

    def _drain_log_queue(self) -> None:
        try:
            for _ in range(200):
                kind, data = self.log_queue.get_nowait()
                if kind == "log":
                    self._append_log(str(data))
                elif kind == "progress":
                    cur, tot, file_name = data
                    self.progress.stop()
                    self.progress.configure(mode="determinate", maximum=max(tot, 1), value=cur)
                    self.status_var.set(f"已处理 {cur}/{tot} 个文件" + (f" · {file_name}" if file_name else ""))
                elif kind == "result":
                    self.last_report = data
                    self._set_running(False)
                    self.status_var.set(data.summary())
                    for index, item in enumerate(data.results):
                        self.result_tree.insert("", tk.END, iid=str(index), tags=(item.status,), values=(
                            "成功" if item.status == "success" else "失败",
                            os.path.relpath(item.source, self.active_folder),
                            item.error or item.warning or os.path.relpath(item.output, self.active_folder),
                        ))
                    if self.close_requested:
                        self.destroy()
                        return
                    if data.cancelled or data.failed:
                        messagebox.showwarning("已取消" if data.cancelled else "部分文件转换失败", data.summary())
                    else:
                        messagebox.showinfo("处理完成", data.summary())
                elif kind == "error":
                    self._append_log(f"处理失败：{data}")
                    self._set_running(False)
                    self.status_var.set("处理失败，请查看日志")
                    if self.close_requested:
                        self.destroy()
                        return
                    messagebox.showerror("处理失败", str(data))
        except queue.Empty:
            pass

        self.after(100, self._drain_log_queue)

    def _set_running(self, running: bool) -> None:
        for widget in (self.path_entry, self.choose_button, *self.option_widgets):
            widget.configure(state=tk.DISABLED if running else tk.NORMAL)
        self.retry_button.configure(state=tk.NORMAL if not running and self.last_report and self.last_report.failed else tk.DISABLED)
        if running:
            self.start_button.configure(state=tk.DISABLED)
            self.cancel_button.configure(state=tk.NORMAL)
            self.progress.configure(mode="indeterminate")
            self.progress.configure(value=0)
            self.progress.start(10)
            self.status_var.set("正在扫描文件并启动 Office...")
        else:
            self.start_button.configure(state=tk.NORMAL)
            self.cancel_button.configure(state=tk.DISABLED)
            self.progress.stop()
            if self.last_report:
                self.progress.configure(mode="determinate", maximum=max(self.last_report.total, 1),
                                        value=self.last_report.succeeded + self.last_report.failed)
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
    multiprocessing.freeze_support()
    app = PdfConverterApp()
    app.mainloop()

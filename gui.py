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
        self.geometry("920x720")
        self.minsize(800, 640)

        self.folder_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=True)
        self.flatten_var = tk.BooleanVar(value=False)
        self.delete_source_var = tk.BooleanVar(value=False)
        self.timeout_var = tk.StringVar(value="120")
        self.status_var = tk.StringVar(value="选择文件夹后即可开始转换")
        self.settings_summary_var = tk.StringVar()
        self.progress_count_var = tk.StringVar(value="")
        self.log_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel_event = threading.Event()
        self.last_report: ConversionReport | None = None
        self.last_job = None
        self.active_folder = ""
        self.close_requested = False
        self.option_widgets = []

        self._build_ui()
        for variable in (self.flatten_var, self.delete_source_var, self.timeout_var):
            variable.trace_add("write", self._update_settings_summary)
        self._update_settings_summary()
        self.bind("<Control-o>", lambda event: self.choose_button.invoke())
        self.bind("<Control-Return>", lambda event: self.start_button.invoke())
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
        self.configure(bg="#ffffff")
        font = ("Microsoft YaHei UI", 10)
        ink, muted, line = "#202936", "#606b79", "#dce1e8"
        accent = "#245bdb"
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=font, background="#ffffff", foreground=ink)
        style.configure("TFrame", background="#ffffff")
        style.configure("TLabel", background="#ffffff", foreground=ink)
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 19, "bold"))
        style.configure("Section.TLabel", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Hint.TLabel", foreground=muted)
        style.configure("Success.TLabel", foreground="#23704b")
        style.configure("Warning.TLabel", foreground="#a14516")
        style.configure("TButton", padding=(14, 8), background="#f6f7f9",
                        bordercolor=line, lightcolor="#f6f7f9", darkcolor="#f6f7f9",
                        focuscolor=accent, relief=tk.FLAT)
        style.map("TButton", background=[("disabled", "#f6f7f9"), ("pressed", "#e7ebf1"),
                                          ("active", "#edf0f5")],
                  foreground=[("disabled", "#929aa5")],
                  bordercolor=[("focus", accent)])
        style.configure("Accent.TButton", font=("Microsoft YaHei UI", 10, "bold"),
                        background=accent, foreground="#ffffff", bordercolor=accent,
                        lightcolor=accent, darkcolor=accent, focuscolor="#ffffff",
                        padding=(24, 10))
        style.map("Accent.TButton", background=[("disabled", "#e9edf3"),
                                                 ("pressed", "#1945ae"), ("active", "#1e4fbe")],
                  foreground=[("disabled", "#7b8593")],
                  bordercolor=[("disabled", "#e9edf3"), ("!disabled", accent)],
                  lightcolor=[("disabled", "#e9edf3"), ("!disabled", accent)],
                  darkcolor=[("disabled", "#e9edf3"), ("!disabled", accent)])
        style.configure("Quiet.TButton", background="#ffffff", bordercolor="#ffffff",
                        lightcolor="#ffffff", darkcolor="#ffffff", foreground=accent,
                        padding=(10, 4))
        style.map("Quiet.TButton", background=[("active", "#eff4ff")],
                  bordercolor=[("focus", accent), ("active", "#eff4ff")])
        style.configure("TEntry", padding=(10, 9), fieldbackground="#ffffff",
                        bordercolor=line, lightcolor=line, darkcolor=line)
        style.map("TEntry", bordercolor=[("focus", accent)],
                  fieldbackground=[("disabled", "#f6f7f9")])
        style.configure("TCheckbutton", font=font, background="#ffffff",
                        focuscolor=accent, indicatormargin=(0, 0, 8, 0))
        style.map("TCheckbutton", background=[("active", "#ffffff")],
                  foreground=[("disabled", "#929aa5")])
        style.configure("TSpinbox", padding=(6, 5), arrowsize=14, bordercolor=line)
        style.map("TSpinbox", bordercolor=[("focus", accent)])
        style.configure("TSeparator", background=line, bordercolor=line, lightcolor=line, darkcolor=line)
        style.configure("Horizontal.TProgressbar", background=accent, troughcolor="#edf0f5",
                        bordercolor="#edf0f5", lightcolor=accent, darkcolor=accent, arrowsize=4)
        style.configure("TNotebook", background="#ffffff", borderwidth=0,
                        bordercolor=line, tabmargins=(0, 0, 0, 0))
        style.configure("TNotebook.Tab", padding=(16, 6), background="#f6f7f9",
                        bordercolor=line, lightcolor="#ffffff", darkcolor=line)
        style.map("TNotebook.Tab", background=[("selected", "#ffffff"), ("active", "#edf0f5")],
                  foreground=[("selected", accent)])
        style.configure("Treeview", font=font, rowheight=36, fieldbackground="#ffffff",
                        background="#ffffff", borderwidth=0, bordercolor=line,
                        lightcolor="#ffffff", darkcolor="#ffffff", foreground=ink)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"),
                        background="#f6f7f9", foreground=muted, padding=(10, 10),
                        relief=tk.FLAT, borderwidth=0)
        style.map("Treeview", background=[("selected", "#e8efff")],
                  foreground=[("selected", "#1945ae")])
        style.map("Treeview.Heading", background=[("active", "#edf0f5")])
        for orientation in ("Horizontal", "Vertical"):
            style.configure(f"{orientation}.TScrollbar", background="#d6dce5", troughcolor="#f6f7f9",
                            bordercolor="#f6f7f9", lightcolor="#d6dce5", darkcolor="#d6dce5",
                            arrowcolor=muted, arrowsize=12, gripcount=0)

        container = ttk.Frame(self, padding=24)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="PPT / Word 批量转 PDF", style="Title.TLabel").pack(anchor=tk.W, pady=(0, 20))
        ttk.Label(container, text="文件夹", style="Section.TLabel").pack(anchor=tk.W, pady=(0, 8))
        path_frame = ttk.Frame(container)
        path_frame.pack(fill=tk.X)
        self.path_entry = ttk.Entry(path_frame, textvariable=self.folder_var, font=font)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.choose_button = ttk.Button(path_frame, text="选择文件夹", command=self._choose_folder)
        self.choose_button.pack(side=tk.LEFT, padx=(10, 0))

        options_frame = ttk.Frame(container)
        options_frame.pack(fill=tk.X, pady=(10, 0))
        recursive_check = ttk.Checkbutton(options_frame, text="包含子文件夹", variable=self.recursive_var)
        recursive_check.pack(side=tk.LEFT)
        self.option_widgets.append(recursive_check)
        self.advanced_button = ttk.Button(options_frame, text="高级设置", style="Quiet.TButton",
                                          command=self._toggle_advanced)
        self.advanced_button.pack(side=tk.RIGHT)
        self.advanced_frame = ttk.Frame(container, padding=(0, 12, 0, 4))
        file_options = ttk.Frame(self.advanced_frame)
        file_options.pack(fill=tk.X, pady=(0, 8))
        for text, variable in (
            ("将子文件夹的源文件移到主文件夹", self.flatten_var),
            ("PDF 校验成功后删除源文件", self.delete_source_var),
        ):
            widget = ttk.Checkbutton(file_options, text=text, variable=variable)
            widget.pack(side=tk.LEFT, padx=(0, 24))
            self.option_widgets.append(widget)
        timeout_frame = ttk.Frame(self.advanced_frame)
        timeout_frame.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(timeout_frame, text="单次尝试超时").pack(side=tk.LEFT)
        timeout_entry = ttk.Spinbox(timeout_frame, from_=10, to=1800, increment=30,
                                    textvariable=self.timeout_var, width=6, font=font)
        timeout_entry.pack(side=tk.LEFT, padx=(10, 6))
        self.option_widgets.append(timeout_entry)
        ttk.Label(timeout_frame, text="秒 · 失败后自动重试一次", style="Hint.TLabel").pack(side=tk.LEFT)
        self.settings_summary_label = ttk.Label(container, textvariable=self.settings_summary_var,
                                                style="Hint.TLabel", wraplength=800)
        self.settings_summary_label.pack(fill=tk.X, pady=(8, 0))
        actions_frame = ttk.Frame(container)
        actions_frame.pack(fill=tk.X, pady=(16, 16))
        self.start_button = ttk.Button(
            actions_frame, text="开始转换", style="Accent.TButton", command=self._start_convert,
        )
        self.start_button.pack(side=tk.LEFT)
        self.cancel_button = ttk.Button(
            actions_frame, text="取消转换", command=self._cancel_convert, state=tk.DISABLED,
        )
        self.retry_button = ttk.Button(actions_frame, text="仅重试失败文件", command=self._retry_failed,
                                       state=tk.DISABLED)
        ttk.Label(actions_frame, text="Ctrl+O 选择 · Ctrl+Enter 转换", style="Hint.TLabel").pack(side=tk.RIGHT)
        ttk.Separator(container).pack(fill=tk.X, pady=(0, 12))
        status_frame = ttk.Frame(container)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        self.progress_count_label = ttk.Label(status_frame, textvariable=self.progress_count_var,
                                              style="Hint.TLabel")
        self.progress_count_label.pack(side=tk.RIGHT, anchor=tk.N, padx=(16, 0))
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var,
                                      style="Hint.TLabel", wraplength=700)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        status_frame.bind("<Configure>", self._resize_status)
        container.bind("<Configure>", lambda event: self.settings_summary_label.configure(
            wraplength=max(200, event.width - 48)))
        self.progress = ttk.Progressbar(container, mode="determinate", value=0)
        self.progress.pack(fill=tk.X, pady=(0, 20))
        self.notebook = ttk.Notebook(container)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.results_frame = ttk.Frame(self.notebook)
        self.log_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.results_frame, text="文件结果")
        self.notebook.add(self.log_frame, text="处理日志")
        result_footer = ttk.Frame(self.results_frame, padding=(10, 8, 0, 0))
        result_footer.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(result_footer, text="双击文件或按 Enter 查看详情", style="Hint.TLabel").pack(side=tk.LEFT)
        ttk.Button(result_footer, text="打开输出文件夹", style="Quiet.TButton",
                   command=self._open_folder).pack(side=tk.RIGHT)
        log_footer = ttk.Frame(self.log_frame, padding=(10, 8, 0, 0))
        log_footer.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(log_footer, text="记录转换过程与失败原因", style="Hint.TLabel").pack(side=tk.LEFT)
        ttk.Button(log_footer, text="保存日志", style="Quiet.TButton",
                   command=self._save_log).pack(side=tk.RIGHT)
        self.log_text = scrolledtext.ScrolledText(
            self.log_frame, height=6, wrap=tk.WORD, font=font, relief=tk.FLAT,
            borderwidth=0, background="#ffffff", foreground=ink, padx=12, pady=12,
            selectbackground="#e8efff", selectforeground="#1945ae",
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)
        result_body = ttk.Frame(self.results_frame)
        result_body.pack(fill=tk.BOTH, expand=True)
        result_body.rowconfigure(0, weight=1)
        result_body.columnconfigure(0, weight=1)
        self.result_tree = ttk.Treeview(result_body, columns=("status", "file", "detail"),
                                        show="headings", height=5, selectmode="browse")
        for column, title, width in (("status", "结果", 72), ("file", "源文件", 280), ("detail", "输出 / 失败原因", 380)):
            self.result_tree.heading(column, text=title)
            self.result_tree.column(column, width=width, minwidth=width if column == "status" else 160,
                                    stretch=column != "status", anchor=tk.CENTER if column == "status" else tk.W)
        result_scroll = ttk.Scrollbar(result_body, orient=tk.VERTICAL, command=self.result_tree.yview)
        result_scroll.grid(row=0, column=1, sticky="ns")
        horizontal_scroll = ttk.Scrollbar(result_body, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        horizontal_scroll.grid(row=1, column=0, sticky="ew")
        self.result_tree.configure(yscrollcommand=result_scroll.set, xscrollcommand=horizontal_scroll.set)
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        self.result_tree.tag_configure("failed", foreground="#b42318", background="#fff5f3")
        self.result_tree.bind("<Double-1>", self._show_result_detail)
        self.result_tree.bind("<Return>", self._show_result_detail)
        self.empty_state = ttk.Frame(result_body, padding=16)
        self.empty_state.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        self.empty_title = ttk.Label(self.empty_state, text="准备好要转换的文件夹", style="Section.TLabel")
        self.empty_title.pack()
        self.empty_hint = ttk.Label(self.empty_state, text="转换完成后，在这里查看每个文件的结果。",
                                     style="Hint.TLabel")
        self.empty_hint.pack(pady=(8, 0))

    def _toggle_advanced(self) -> None:
        if self.advanced_frame.winfo_manager():
            self.advanced_frame.pack_forget()
            self.advanced_button.configure(text="高级设置")
        else:
            self.advanced_frame.pack(fill=tk.X, before=self.settings_summary_label)
            self.advanced_button.configure(text="收起设置")

    def _update_settings_summary(self, *_args) -> None:
        details = ["移动源文件到主文件夹" if self.flatten_var.get() else "PDF 保存在原文件旁",
                   "成功后删除源文件" if self.delete_source_var.get() else "保留源文件"]
        if self.timeout_var.get() != "120":
            details.append(f"超时 {self.timeout_var.get()} 秒")
        self.settings_summary_var.set(" · ".join(details))
        self.settings_summary_label.configure(
            style="Warning.TLabel" if self.flatten_var.get() or self.delete_source_var.get() else "Hint.TLabel")

    def _resize_status(self, event) -> None:
        self.status_label.configure(wraplength=max(200, event.width - self.progress_count_label.winfo_reqwidth() - 20))

    def _choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择要转换的文件夹")
        if folder:
            normalized_folder = os.path.normpath(folder)
            self.folder_var.set(normalized_folder)
            self.status_var.set("已选择文件夹")

    def _open_folder(self) -> None:
        # Results belong to the last job even if the user edits the next input.
        folder = self.active_folder or self.folder_var.get().strip().strip('"')
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
        if self.advanced_frame.winfo_manager():
            self._toggle_advanced()
        self._set_running(True)
        self._clear_log()
        self.result_tree.delete(*self.result_tree.get_children())
        self.notebook.select(self.results_frame)
        self.empty_title.configure(text="正在处理文件")
        self.empty_hint.configure(text="完成后显示文件结果，可随时取消转换。")
        self.empty_state.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
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
                    self.progress_count_var.set(f"{cur} / {tot}")
                    display_name = file_name if len(file_name) <= 80 else f"{file_name[:38]}…{file_name[-38:]}"
                    self.status_var.set(f"已处理 {cur}/{tot} 个文件" + (f" · {display_name}" if display_name else ""))
                elif kind == "result":
                    self.last_report = data
                    self._set_running(False)
                    self.status_var.set(data.summary())
                    self.status_label.configure(style="Warning.TLabel" if data.cancelled or data.failed else "Success.TLabel")
                    self.notebook.select(self.results_frame)
                    if data.results:
                        self.empty_state.place_forget()
                    else:
                        self.empty_title.configure(text="任务已取消" if data.cancelled else "未找到可转换的文件")
                        self.empty_hint.configure(text="可重新开始转换。" if data.cancelled else "请选择包含 PPT 或 Word 文档的文件夹。")
                    for index, item in enumerate(data.results):
                        self.result_tree.insert("", tk.END, iid=str(index), tags=(item.status,), values=(
                            "成功" if item.status == "success" else "失败",
                            os.path.relpath(item.source, self.active_folder),
                            item.error or item.warning or os.path.relpath(item.output, self.active_folder),
                        ))
                    if self.close_requested:
                        self.destroy()
                        return
                elif kind == "error":
                    self._append_log(f"处理失败：{data}")
                    self._set_running(False)
                    self.status_var.set("处理失败，请查看日志")
                    self.status_label.configure(style="Warning.TLabel")
                    self.empty_title.configure(text="转换未完成")
                    self.empty_hint.configure(text="请在处理日志中查看原因后重试。")
                    self.notebook.select(self.log_frame)
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
        can_retry = bool(not running and self.last_report and self.last_report.failed)
        self.retry_button.configure(state=tk.NORMAL if can_retry else tk.DISABLED)
        if can_retry:
            self.retry_button.pack(side=tk.LEFT, padx=(10, 0), after=self.start_button)
        else:
            self.retry_button.pack_forget()
        self.status_label.configure(style="Hint.TLabel")
        if running:
            self.start_button.configure(state=tk.DISABLED, text="正在转换…")
            self.cancel_button.configure(state=tk.NORMAL)
            self.cancel_button.pack(side=tk.LEFT, padx=(10, 0), after=self.start_button)
            self.progress.configure(mode="indeterminate")
            self.progress.configure(value=0)
            self.progress.start(10)
            self.progress_count_var.set("")
            self.status_var.set("正在扫描文件并启动 Office...")
        else:
            self.start_button.configure(state=tk.NORMAL, text="开始转换")
            self.cancel_button.configure(state=tk.DISABLED)
            self.cancel_button.pack_forget()
            self.progress.stop()
            if self.last_report:
                self.progress.configure(mode="determinate", maximum=max(self.last_report.total, 1),
                                        value=self.last_report.succeeded + self.last_report.failed)
                self.progress_count_var.set(f"{self.last_report.succeeded + self.last_report.failed} / {self.last_report.total}")
            else:
                self.progress.configure(mode="determinate", value=0)
                self.progress_count_var.set("")
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

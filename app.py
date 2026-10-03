"""Small native GUI for macOS development and Windows distribution."""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from scanner import CHINESE_HEADERS, KOREAN_HEADERS, export_exceptions, scan_folder
from ignored_sheets import default_ignore_file, load_ignored, save_ignored, sheet_key
from matcher import match_folder


def open_path(path):
    if sys.platform == "win32":
        os.startfile(str(path))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class CheckboxTreeview(ttk.Treeview):
    """Checks stay independent of the highlighted row used to view details."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, show="tree headings", selectmode="browse", **kwargs)
        self.checked = set()
        self.unchecked_image = self.checkbox_image(False)
        self.checked_image = self.checkbox_image(True)
        self.heading("#0", text="勾选")
        self.column("#0", width=52, minwidth=52, stretch=False)
        self.bind("<Button-1>", self.click_checkbox)
        self.bind("<space>", self.toggle_focused)

    def checkbox_image(self, checked):
        image = tk.PhotoImage(master=self, width=18, height=18)
        image.put("#2563eb" if checked else "#64748b", to=(1, 1, 17, 17))
        image.put("#2563eb" if checked else "#ffffff", to=(2, 2, 16, 16))
        if checked:
            for x, y in ((4, 8), (5, 9), (6, 10), (7, 11), (8, 10), (9, 9), (10, 8), (11, 7), (12, 6)):
                image.put("#ffffff", to=(x, y, x + 2, y + 2))
        return image

    def insert(self, parent, index, iid=None, **kwargs):
        kwargs.setdefault("image", self.unchecked_image)
        return super().insert(parent, index, iid=iid, **kwargs)

    def delete(self, *items):
        self.checked.difference_update(items)
        return super().delete(*items)

    def checked_items(self):
        return tuple(item for item in self.get_children() if item in self.checked)

    def set_checked(self, items, checked=True):
        for item in items:
            if self.exists(item):
                if checked:
                    self.checked.add(item)
                else:
                    self.checked.discard(item)
                self.item(item, image=self.checked_image if checked else self.unchecked_image)
        self.event_generate("<<ChecksChanged>>", when="tail")

    def check_all(self):
        self.set_checked(self.get_children())

    def uncheck_all(self):
        self.set_checked(self.get_children(), False)

    def click_checkbox(self, event):
        item = self.identify_row(event.y)
        if item and self.identify_column(event.x) == "#0":
            self.set_checked((item,), item not in self.checked)
            self.focus(item)
            self.selection_set(item)
            return "break"

    def toggle_focused(self, event):
        item = self.focus()
        if item:
            self.set_checked((item,), item not in self.checked)
        return "break"


class Application:
    def __init__(self, root, ignore_file=None):
        self.root = root
        self.report = None
        self.running = False
        self.match_report = None
        self.events = queue.Queue()
        self.ignore_file = Path(ignore_file) if ignore_file is not None else default_ignore_file()
        self.ignore_error = ""
        try:
            self.ignored = load_ignored(self.ignore_file)
        except (OSError, ValueError) as error:
            self.ignored = set()
            self.ignore_error = str(error)
        root.title("中韩表头检查")
        root.geometry("1100x760")
        root.minsize(820, 520)
        root.protocol("WM_DELETE_WINDOW", self.close)
        frame = ttk.Frame(root, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="中韩译文检索", font=("", 18, "bold")).pack(anchor="w")
        ttk.Label(frame, text="第一步检查表头，第二步检索译文。扫描范围可调整；结果写入副本。").pack(anchor="w", pady=(6, 14))

        folder_row = ttk.Frame(frame)
        folder_row.pack(fill="x")
        self.folder = tk.StringVar()
        ttk.Label(folder_row, text="源文件夹").pack(side="left")
        self.folder_entry = ttk.Entry(folder_row, textvariable=self.folder)
        self.folder_entry.pack(side="left", fill="x", expand=True, padx=8)
        self.browse_button = ttk.Button(folder_row, text="选择文件夹", command=self.browse)
        self.browse_button.pack(side="left")

        settings = ttk.LabelFrame(frame, text="检查规则（表头名称每行一个）", padding=10)
        settings.pack(fill="x", pady=12)
        settings.columnconfigure(1, weight=1)
        settings.columnconfigure(3, weight=1)
        ttk.Label(settings, text="中文表头").grid(row=0, column=0, sticky="nw", padx=(0, 8))
        self.chinese = tk.Text(settings, height=6, width=24, wrap="none")
        self.chinese.insert("1.0", "\n".join(CHINESE_HEADERS))
        self.chinese.grid(row=0, column=1, sticky="ew")
        ttk.Label(settings, text="韩语表头").grid(row=0, column=2, sticky="nw", padx=8)
        self.korean = tk.Text(settings, height=6, width=24, wrap="none")
        self.korean.insert("1.0", "\n".join(KOREAN_HEADERS))
        self.korean.grid(row=0, column=3, sticky="ew")
        ttk.Label(settings, text="每张工作表检查前几行").grid(row=1, column=0, columnspan=2, sticky="w", pady=(10, 0))
        self.rows = tk.StringVar(value="10")
        self.row_entry = ttk.Entry(settings, textvariable=self.rows, width=10)
        self.row_entry.grid(row=1, column=2, sticky="w", padx=8, pady=(10, 0))
        options = ttk.Frame(settings)
        options.grid(row=2, column=0, columnspan=4, sticky="w", pady=(10, 0))
        self.recursive = tk.BooleanVar(value=True)
        self.include_hidden_sheets = tk.BooleanVar(value=True)
        self.include_hidden_rows = tk.BooleanVar(value=True)
        self.scan_switches = []
        for title, variable in (("扫描子文件夹", self.recursive), ("扫描隐藏工作表", self.include_hidden_sheets), ("扫描隐藏行", self.include_hidden_rows)):
            switch = ttk.Checkbutton(options, text=title, variable=variable)
            switch.pack(side="left", padx=(0, 24))
            self.scan_switches.append(switch)

        actions = ttk.Frame(frame)
        actions.pack(fill="x")
        self.start_button = ttk.Button(actions, text="开始检查", command=self.start)
        self.start_button.pack(side="left")
        self.export_button = ttk.Button(actions, text="导出异常清单", command=self.export, state="disabled")
        self.export_button.pack(side="left", padx=8)
        self.open_button = ttk.Button(actions, text="打开首个所选文件所在文件夹", command=self.open_selected, state="disabled")
        self.open_button.pack(side="left")
        self.ignore_button = ttk.Button(actions, text="忽略勾选 Sheet", command=self.ignore_selected, state="disabled")
        self.ignore_button.pack(side="left", padx=8)
        self.status = tk.StringVar(value="请选择客户源文件夹。无中韩表头的参考表也会列出，需客户确认是否为翻译表。")
        ttk.Label(frame, textvariable=self.status, wraplength=1000).pack(anchor="w", pady=(10, 5))
        self.progress = ttk.Progressbar(frame)
        self.progress.pack(fill="x", pady=(0, 8))

        self.tabs = ttk.Notebook(frame)
        self.tabs.pack(fill="both", expand=True)
        self.exception_tab = ttk.Frame(self.tabs, padding=6)
        self.tabs.add(self.exception_tab, text="异常列表")
        controls = ttk.Frame(self.exception_tab)
        controls.pack(fill="x", pady=(0, 6))
        ttk.Button(controls, text="勾选全部", command=lambda: self.table.check_all()).pack(side="left")
        ttk.Button(controls, text="取消勾选", command=lambda: self.table.uncheck_all()).pack(side="left", padx=8)
        ttk.Label(controls, text="点击左侧方框勾选，可选择不相邻行；双击文字查看详情。").pack(side="left", padx=10)
        table_frame = ttk.Frame(self.exception_tab)
        table_frame.pack(fill="both", expand=True)
        columns = ("path", "sheet", "issue", "chinese", "korean")
        self.table = CheckboxTreeview(table_frame, columns=columns)
        for name, title, width in zip(columns, ("相对路径", "工作表", "异常原因", "中文表头位置", "韩语表头位置"), (300, 180, 320, 180, 180)):
            self.table.heading(name, text=title)
            self.table.column(name, width=width, minwidth=100, stretch=False)
        vertical = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        horizontal = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.table.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.table.bind("<Double-1>", self.double_click_details)
        self.table.bind("<Return>", lambda event: self.show_details())
        self.ignored_tab = ttk.Frame(self.tabs, padding=6)
        self.tabs.add(self.ignored_tab, text="已忽略列表")
        controls = ttk.Frame(self.ignored_tab)
        controls.pack(fill="x", pady=(0, 6))
        ttk.Button(controls, text="勾选全部", command=lambda: self.ignored_table.check_all()).pack(side="left")
        ttk.Button(controls, text="取消勾选", command=lambda: self.ignored_table.uncheck_all()).pack(side="left", padx=8)
        self.restore_button = ttk.Button(controls, text="移除勾选忽略记录", command=self.restore_selected)
        self.restore_button.pack(side="left", padx=8)
        ttk.Label(self.ignored_tab, text="移除的是本地忽略记录，重新检查后恢复处理这些 Sheet。").pack(anchor="w", pady=(0, 6))
        ttk.Label(self.ignored_tab, text=f"本地记录：{self.ignore_file}", wraplength=1000).pack(anchor="w", pady=(0, 6))
        ignored_frame = ttk.Frame(self.ignored_tab)
        ignored_frame.pack(fill="both", expand=True)
        self.ignored_table = CheckboxTreeview(ignored_frame, columns=("path", "sheet"))
        for name, label, width in (("path", "文件完整路径", 760), ("sheet", "Sheet 名称", 220)):
            self.ignored_table.heading(name, text=label)
            self.ignored_table.column(name, width=width, minwidth=120, stretch=False)
        vertical = ttk.Scrollbar(ignored_frame, orient="vertical", command=self.ignored_table.yview)
        horizontal = ttk.Scrollbar(ignored_frame, orient="horizontal", command=self.ignored_table.xview)
        self.ignored_table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.ignored_table.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        ignored_frame.rowconfigure(0, weight=1)
        ignored_frame.columnconfigure(0, weight=1)
        self.match_tab = ttk.Frame(self.tabs, padding=12)
        self.tabs.add(self.match_tab, text="第二步：译文检索")
        self.baseline_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.match_inputs = []
        for label, variable in (("基准文件夹", self.baseline_folder), ("输出文件夹", self.output_folder)):
            line = ttk.Frame(self.match_tab)
            line.pack(fill="x", pady=4)
            ttk.Label(line, text=label, width=12).pack(side="left")
            entry = ttk.Entry(line, textvariable=variable)
            entry.pack(side="left", fill="x", expand=True, padx=8)
            button = ttk.Button(line, text="选择文件夹", command=lambda v=variable, t=label: self.browse_match(v, t))
            button.pack(side="left")
            self.match_inputs.extend((entry, button))
        ttk.Label(self.match_tab, text="先完成第一步检查；已忽略及异常 Sheet 自动跳过。\n按 G 列 Text 检索，中文去除首尾空白后全文匹配。1～3 条展开，超过 3 条只写条数。\n基准原件不改动，副本 R～Z 写三组中韩文本及来源，AA 写匹配条数。", wraplength=980).pack(anchor="w", pady=8)
        line = ttk.Frame(self.match_tab)
        line.pack(fill="x")
        self.match_button = ttk.Button(line, text="开始检索并生成结果", command=self.start_match, state="disabled")
        self.match_button.pack(side="left")
        self.match_open_button = ttk.Button(line, text="打开结果文件夹", command=self.open_match_output, state="disabled")
        self.match_open_button.pack(side="left", padx=8)
        self.match_summary = tk.StringVar(value="请先检查源文件表头，再选择 0_基准文件 和单独的输出文件夹。")
        ttk.Label(self.match_tab, textvariable=self.match_summary, wraplength=980, justify="left").pack(anchor="w", pady=10)
        for table in (self.table, self.ignored_table):
            table.bind("<<ChecksChanged>>", lambda event: self.set_running(self.running))
        self.refresh_ignored()
        self.set_running(False)
        if self.ignore_error:
            self.status.set("本地忽略记录读取失败，检查已禁用。请修复记录文件后重启。")
            root.after(0, lambda: messagebox.showerror("忽略记录读取失败", self.ignore_error))
        root.after(100, self.poll)

    def browse(self):
        folder = filedialog.askdirectory(title="选择客户源文件夹")
        if folder:
            self.folder.set(folder)

    def set_running(self, running):
        self.running = running
        for widget in (self.start_button, self.browse_button, self.folder_entry, self.row_entry, *self.match_inputs, *self.scan_switches):
            widget.configure(state="disabled" if running else "normal")
        if self.ignore_error:
            self.start_button.configure(state="disabled")
        for widget in (self.chinese, self.korean):
            widget.configure(state="disabled" if running else "normal")
        available = not running and self.report is not None
        self.export_button.configure(state="normal" if available else "disabled")
        self.open_button.configure(state="normal" if available and self.report.exceptions else "disabled")
        self.ignore_button.configure(state="normal" if available and self.table.checked_items() and not self.ignore_error else "disabled")
        self.restore_button.configure(state="normal" if not running and self.ignored_table.checked_items() and not self.ignore_error else "disabled")
        self.match_button.configure(state="normal" if available and self.report.passed_sheets and not self.ignore_error else "disabled")
        self.match_open_button.configure(state="normal" if not running and self.match_report else "disabled")

    def browse_match(self, variable, title):
        folder = filedialog.askdirectory(title=f"选择{title}")
        if folder:
            variable.set(folder)

    def start_match(self):
        if self.running or self.report is None or self.ignore_error:
            return
        try:
            from scanner import normalize
            current = (Path(self.folder.get().strip()).expanduser().resolve(), int(self.rows.get()),
                       {normalize(v) for v in self.chinese.get("1.0", "end").splitlines() if normalize(v)},
                       {normalize(v) for v in self.korean.get("1.0", "end").splitlines() if normalize(v)},
                       self.recursive.get(), self.include_hidden_sheets.get(), self.include_hidden_rows.get())
            expected = (self.report.root, self.report.header_rows, set(self.report.chinese_headers), set(self.report.korean_headers),
                        self.report.recursive, self.report.include_hidden_sheets, self.report.include_hidden_rows)
            if current != expected:
                raise ValueError("源文件夹或检查规则已改变，请先重新执行第一步检查")
            baseline = self.baseline_folder.get().strip()
            output = self.output_folder.get().strip()
            if not baseline or not output:
                raise ValueError("请选择基准文件夹和输出文件夹")
        except ValueError as error:
            messagebox.showerror("无法开始检索", str(error))
            return
        report, ignored, events = self.report, set(self.ignored), self.events
        self.match_report = None
        self.match_summary.set("正在检索，完成后显示结果位置。")
        self.set_running(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start(15)

        def work():
            try:
                result = match_folder(report, baseline, output, ignored_sheets=ignored,
                                      progress=lambda text: events.put(("match_progress", text)))
                events.put(("match_done", result))
            except Exception as error:
                events.put(("match_error", str(error)))

        threading.Thread(target=work, daemon=True).start()

    def open_match_output(self):
        if self.match_report:
            try:
                open_path(self.match_report.destination)
            except OSError as error:
                messagebox.showerror("无法打开结果文件夹", str(error))

    def start(self):
        if self.running or self.ignore_error:
            return
        try:
            rows = int(self.rows.get())
            if rows < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("检查规则有误", "检查行数必须是大于等于 1 的整数。")
            return
        folder = self.folder.get().strip()
        if not folder:
            messagebox.showerror("请选择文件夹", "请先选择客户源文件夹。")
            return
        chinese = self.chinese.get("1.0", "end").splitlines()
        korean = self.korean.get("1.0", "end").splitlines()
        self.report = None
        self.table.delete(*self.table.get_children())
        self.set_running(True)
        self.progress.configure(value=0, maximum=1)
        self.status.set("正在查找 Excel 文件……")
        ignored = set(self.ignored)
        scan_options = dict(recursive=self.recursive.get(), include_hidden_sheets=self.include_hidden_sheets.get(),
                            include_hidden_rows=self.include_hidden_rows.get())
        events = self.events

        def work():
            try:
                report = scan_folder(folder, rows, chinese, korean,
                                     progress=lambda i, total, path: events.put(("progress", (i, total, path))),
                                     ignored_sheets=ignored, **scan_options)
                events.put(("done", report))
            except Exception as error:
                events.put(("error", str(error)))

        threading.Thread(target=work, daemon=True).start()

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "progress":
                    index, total, path = value
                    self.progress.configure(maximum=total, value=index - 1)
                    self.status.set(f"正在检查 {index}/{total}：{path.name}")
                elif kind == "done":
                    self.report = value
                    self.refresh_exceptions()
                    self.progress.configure(value=value.files)
                    self.set_running(False)
                elif kind == "error":
                    self.set_running(False)
                    self.status.set("检查未完成，请处理错误后重试。")
                    messagebox.showerror("检查失败", value)
                elif kind == "match_progress":
                    self.status.set(value)
                elif kind in ("match_done", "match_error"):
                    self.progress.stop()
                    self.progress.configure(mode="determinate", value=0)
                    if kind == "match_done":
                        self.match_report = value
                        self.report = value.check_report
                        self.refresh_exceptions()
                        summary = (f"检索完成：基准 {value.baseline_rows} 行，匹配 {value.matched_rows} 行；"
                                   f"源文件参与 {value.source_rows} 行，匹配 {value.source_matched} 行；跳过 {value.skipped} 项。\n"
                                   f"结果位置：{value.destination}\n基准匹配副本位于“基准匹配结果”子文件夹；统计与跳过原因见“源文件匹配统计.xlsx”。")
                        self.match_summary.set(summary)
                        self.status.set(summary.split("\n")[0])
                    else:
                        self.match_summary.set("检索失败，未生成本次结果。请处理错误后重试。")
                        self.status.set("检索未完成。")
                        messagebox.showerror("检索失败", value)
                    self.set_running(False)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def refresh_exceptions(self):
        self.table.delete(*self.table.get_children())
        if self.report is None:
            return
        report = self.report
        for index, result in enumerate(report.exceptions):
            self.table.insert("", "end", iid=str(index), values=(str(result.path.relative_to(report.root)), result.sheet, result.issue, result.chinese, result.korean))
        self.status.set(f"已检查 {report.files} 个文件；{report.passed_sheets} 张表通过；{report.abnormal_files} 个文件需确认，共 {len(report.exceptions)} 项异常；已忽略 {report.ignored_sheets} 张表；按开关跳过 {len(report.skipped_sheets)} 张表。")

    def refresh_ignored(self):
        self.ignored_table.delete(*self.ignored_table.get_children())
        self.ignored_rows = sorted(self.ignored)
        for index, (filename, sheet) in enumerate(self.ignored_rows):
            self.ignored_table.insert("", "end", iid=str(index), values=(filename, sheet))
        self.tabs.tab(self.ignored_tab, text=f"已忽略列表（{len(self.ignored_rows)}）")

    def update_ignored(self, records):
        try:
            save_ignored(records, self.ignore_file)
        except OSError as error:
            messagebox.showerror("忽略记录保存失败", str(error))
            return False
        self.ignored = records
        self.refresh_ignored()
        return True

    def ignore_selected(self):
        if self.running or self.report is None or self.ignore_error:
            return
        selected = [self.report.exceptions[int(index)] for index in self.table.checked_items()]
        records = {sheet_key(result.path, result.sheet) for result in selected if result.sheet}
        if not records:
            messagebox.showinfo("请勾选 Sheet", "请勾选需要忽略的异常工作表。文件级读取失败或不支持格式的记录没有 Sheet 名称，不能按 Sheet 忽略。")
            return
        if not self.update_ignored(self.ignored | records):
            return
        before = len(self.report.results)
        self.report.results = [result for result in self.report.results
                               if not result.sheet or sheet_key(result.path, result.sheet) not in records]
        self.report.ignored_sheets += before - len(self.report.results)
        self.refresh_exceptions()
        self.set_running(False)
        if any(not result.sheet for result in selected):
            self.status.set(self.status.get() + " 文件级异常已保留，不能按 Sheet 忽略。")

    def restore_selected(self):
        if self.running or self.ignore_error:
            return
        records = {self.ignored_rows[int(index)] for index in self.ignored_table.checked_items()}
        if not records:
            messagebox.showinfo("请勾选忽略记录", "请先在已忽略列表中勾选要移除的记录。")
            return
        if not self.update_ignored(self.ignored - records):
            return
        self.report = None
        self.refresh_exceptions()
        self.progress.configure(value=0)
        self.set_running(False)
        self.status.set(f"已移除 {len(records)} 条忽略记录。请重新检查，恢复处理这些 Sheet。")

    def selected(self):
        selection = self.table.selection()
        return self.report.exceptions[int(selection[0])] if selection and self.report else None

    def double_click_details(self, event):
        if self.table.identify_column(event.x) == "#0":
            return "break"
        if self.table.identify_row(event.y):
            self.show_details()

    def show_details(self):
        result = self.selected()
        if result:
            passed = [item.sheet for item in self.report.results if item.path == result.path and item.passed]
            window = tk.Toplevel(self.root)
            window.title("异常原因及检查依据")
            window.geometry("900x600")
            window.transient(self.root)
            frame = ttk.Frame(window, padding=12)
            frame.pack(fill="both", expand=True)
            text = tk.Text(frame, wrap="word", padx=8, pady=8)
            scroll = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
            text.configure(yscrollcommand=scroll.set)
            text.pack(side="left", fill="both", expand=True)
            scroll.pack(side="right", fill="y")
            text.insert("1.0", f"文件：{result.path}\n工作表：{result.sheet or '未能读取'}\n显示状态：{result.visibility or '未能读取'}\n\n具体原因：{result.issue}\n\n该文件已通过的工作表：{'；'.join(passed) or '无'}\n（以上原因针对当前工作表；其他工作表可能已经通过。）\n\n{result.details}")
            text.configure(state="disabled")
            ttk.Button(window, text="关闭", command=window.destroy).pack(pady=(0, 12))

    def open_selected(self):
        result = self.selected()
        if not result:
            messagebox.showinfo("请选择异常", "请先在列表中选择一项异常。")
            return
        try:
            open_path(result.path.parent)
        except OSError as error:
            messagebox.showerror("无法打开文件夹", str(error))

    def export(self):
        destination = filedialog.asksaveasfilename(
            title="保存异常清单", initialfile=f"表头异常清单_{datetime.now():%Y%m%d_%H%M%S}.csv",
            defaultextension=".csv", filetypes=[("Excel 可打开的 CSV", "*.csv")],
        )
        if destination:
            try:
                export_exceptions(self.report, destination)
            except (OSError, ValueError) as error:
                messagebox.showerror("导出失败", str(error))
                return
            messagebox.showinfo("已导出", f"异常清单已保存：\n{destination}\n\n可用 Excel 打开。")

    def close(self):
        if self.running:
            messagebox.showinfo("正在处理", "请等待检查或检索结束后关闭窗口。")
        else:
            self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    Application(root)
    root.mainloop()

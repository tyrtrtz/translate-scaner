"""Small native GUI for macOS development and Windows distribution."""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk

from scanner import export_exceptions, scan_folder
from ignored_sheets import default_ignore_file, load_ignored, save_ignored, sheet_key
from matcher import match_folder
from settings import default_settings, load_settings, save_settings


COLORS = dict(background="#f5f5f7", surface="#ffffff", ink="#1d1d1f",
              muted="#6e6e73", accent="#0071e3", accent_hover="#0066cc",
              soft="#eaf3ff", border="#e5e5ea", amber="#b7791f", stripe="#f8f8fa")


def rounded_image(root, fill, size=24, radius=8, outline=None):
    """Small stretchable Tk images replace the theme's bevels and square borders."""
    image = tk.PhotoImage(master=root, width=size, height=size)

    def inside(x, y, inset=0):
        r = radius - inset
        dx = max(radius - x - 0.5, 0, x + 0.5 - (size - radius))
        dy = max(radius - y - 0.5, 0, y + 0.5 - (size - radius))
        return inset <= x < size - inset and inset <= y < size - inset and dx * dx + dy * dy <= r * r

    for y in range(size):
        pixels = [x for x in range(size) if inside(x, y)]
        if not pixels:
            continue
        image.put(outline or fill, to=(pixels[0], y, pixels[-1] + 1, y + 1))
        if outline:
            inner = [x for x in pixels if inside(x, y, 1)]
            if inner:
                image.put(fill, to=(inner[0], y, inner[-1] + 1, y + 1))
    return image


def configure_theme(root):
    """Use stock Tk widgets with one consistent, cross-platform palette."""
    families = set(tkfont.families(root))
    body = next((name for name in ("PingFang SC", "Microsoft YaHei UI", "Noto Sans CJK SC") if name in families),
                tkfont.nametofont("TkDefaultFont").actual("family"))
    root.configure(background=COLORS["background"])
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", font=(body, 11), background=COLORS["background"], foreground=COLORS["ink"])
    style.configure("TFrame", background=COLORS["background"])
    style.configure("Card.TFrame", background=COLORS["surface"])
    style.configure("TLabel", background=COLORS["background"])
    style.configure("Card.TLabel", background=COLORS["surface"])
    style.configure("Muted.TLabel", foreground=COLORS["muted"], font=(body, 10))
    style.configure("CardMuted.TLabel", background=COLORS["surface"], foreground=COLORS["muted"], font=(body, 10))
    style.configure("Title.TLabel", font=(body, 20, "bold"))
    style.configure("Section.TLabel", background=COLORS["surface"], font=(body, 13, "bold"))
    style.configure("Number.TLabel", background=COLORS["surface"], font=(body, 24), foreground=COLORS["ink"])
    style.configure("WarningNumber.TLabel", background=COLORS["surface"], font=(body, 24), foreground=COLORS["amber"])
    style.configure("Badge.TLabel", padding=(12, 7), background=COLORS["soft"], foreground=COLORS["accent"], font=(body, 10))
    style.configure("TButton", padding=(11, 6), width=0, background=COLORS["surface"], bordercolor=COLORS["border"],
                    lightcolor=COLORS["surface"], darkcolor=COLORS["surface"], focuscolor=COLORS["accent"], relief="flat")
    style.map("TButton", background=[("!disabled", COLORS["surface"])], foreground=[("disabled", "#9b9ba1")])
    style.configure("Compact.TButton", padding=(8, 2))
    style.configure("Primary.TButton", background=COLORS["surface"], foreground="white", bordercolor=COLORS["accent"],
                    lightcolor=COLORS["accent"], darkcolor=COLORS["accent"], font=(body, 11, "bold"))
    style.map("Primary.TButton", background=[("disabled", COLORS["surface"]), ("!disabled", COLORS["surface"])],
              foreground=[("disabled", COLORS["muted"]), ("!disabled", "white")],
              bordercolor=[("disabled", COLORS["border"]), ("!disabled", COLORS["accent"])])
    style.configure("TEntry", padding=5, borderwidth=0, background=COLORS["surface"], fieldbackground=COLORS["surface"], bordercolor=COLORS["border"],
                    lightcolor=COLORS["surface"], darkcolor=COLORS["surface"], insertcolor=COLORS["ink"])
    style.map("TEntry", bordercolor=[("focus", COLORS["accent"])], fieldbackground=[("disabled", COLORS["stripe"])])
    style.configure("TCheckbutton", background=COLORS["surface"], padding=(0, 3), indicatorbackground=COLORS["surface"],
                    indicatorforeground=COLORS["accent"], font=(body, 10))
    style.map("TCheckbutton", background=[("active", COLORS["surface"])],
              indicatorbackground=[("selected", COLORS["accent"]), ("!selected", COLORS["surface"])])
    # Remove the native notebook client and tree field/focus frames, not just their border widths.
    style.layout("TNotebook", [("Notebook.padding", {"sticky": "nswe"})])
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    style.layout("TEntry", [("Entry.padding", {"sticky": "nswe", "children": [
        ("Entry.textarea", {"sticky": "nswe"})]})])
    style.configure("TNotebook", background=COLORS["background"], borderwidth=0, tabmargins=(0, 0, 0, 10))
    style.configure("TNotebook.Tab", padding=(12, 4), font=(body, 10))
    style.map("TNotebook.Tab", foreground=[("selected", COLORS["accent"]), ("!selected", COLORS["muted"])])
    style.configure("Treeview", background=COLORS["surface"], fieldbackground=COLORS["surface"],
                    rowheight=38, borderwidth=0, font=(body, 10))
    style.configure("Treeview.Heading", background=COLORS["stripe"], foreground=COLORS["muted"],
                    padding=(10, 10), borderwidth=0, font=(body, 10, "bold"), relief="flat")
    style.map("Treeview", background=[("selected", COLORS["soft"])], foreground=[("selected", COLORS["ink"])])
    style.map("Treeview.Heading", background=[("active", COLORS["soft"])])

    # Keep these images alive for the lifetime of the Tcl theme.
    root.theme_images = images = {
        "button": rounded_image(root, "#f0f0f3"),
        "hover": rounded_image(root, "#e5effb"),
        "disabled": rounded_image(root, "#f5f5f7"),
        "focus": rounded_image(root, "#f0f0f3", outline=COLORS["accent"]),
        "primary": rounded_image(root, COLORS["accent"]),
        "primary_hover": rounded_image(root, COLORS["accent_hover"]),
        "primary_focus": rounded_image(root, COLORS["accent"], outline="#003f88"),
        "primary_disabled": rounded_image(root, COLORS["border"]),
        "tab": rounded_image(root, "#ececf0"),
        "tab_selected": rounded_image(root, COLORS["surface"], outline="#d9d9df"),
        "tab_focus": rounded_image(root, COLORS["surface"], outline=COLORS["accent"]),
        "thumb": rounded_image(root, "#c4c4ca", size=10, radius=5),
        "thumb_hover": rounded_image(root, "#96969e", size=10, radius=5),
        "check": rounded_image(root, COLORS["accent"], size=18, radius=4),
        "unchecked": rounded_image(root, COLORS["surface"], size=18, radius=4, outline="#c7c7cc"),
    }
    for x, y in ((4, 8), (5, 9), (6, 10), (7, 11), (8, 10), (9, 9), (10, 8), (11, 7), (12, 6)):
        images["check"].put("#ffffff", to=(x, y, x + 2, y + 2))
    style.element_create("FlatCheck.indicator", "image", images["unchecked"], ("selected", images["check"]),
                         padding=(0, 0, 7, 0), sticky="w")
    style.layout("TCheckbutton", [("Checkbutton.padding", {"sticky": "nswe", "children": [
        ("FlatCheck.indicator", {"side": "left", "sticky": "w"}),
        ("Checkbutton.focus", {"side": "left", "sticky": "w", "children": [("Checkbutton.label", {"sticky": "w"})]})]})])
    line = tk.PhotoImage(master=root, width=1, height=1)
    line.put(COLORS["border"], to=(0, 0, 1, 1))
    images["separator"] = line
    style.element_create("Hairline.separator", "image", line, sticky="we")
    style.layout("TSeparator", [("Hairline.separator", {"sticky": "we"})])
    for prefix, normal, states in (
        ("Pill", "button", (("disabled", "disabled"), ("pressed", "hover"), ("focus", "focus"), ("active", "hover"))),
        ("PrimaryPill", "primary", (("disabled", "primary_disabled"), ("pressed", "primary_hover"), ("focus", "primary_focus"), ("active", "primary_hover"))),
    ):
        style.element_create(prefix + ".background", "image", images[normal],
                             *((state, images[key]) for state, key in states), border=8, padding=0, sticky="nswe")
        layout = [(prefix + ".background", {"sticky": "nswe", "children": [
            ("Button.padding", {"sticky": "nswe", "children": [("Button.label", {"sticky": "nswe"})]})]})]
        style.layout("Primary.TButton" if prefix == "PrimaryPill" else "TButton", layout)
    style.element_create("Segment.background", "image", images["tab"],
                         ("selected", "focus", images["tab_focus"]), ("selected", images["tab_selected"]),
                         border=8, padding=0, sticky="nswe")
    style.layout("TNotebook.Tab", [("Segment.background", {"sticky": "nswe", "children": [
        ("Notebook.padding", {"sticky": "nswe", "children": [("Notebook.label", {"sticky": "nswe"})]})]})])
    style.element_create("Slim.Scrollbar.thumb", "image", images["thumb"], ("active", images["thumb_hover"]),
                         ("pressed", images["thumb_hover"]), border=4, padding=0, sticky="nswe")
    trough = tk.PhotoImage(master=root, width=1, height=1)
    trough.put(COLORS["surface"], to=(0, 0, 1, 1))
    images["trough"] = trough
    style.element_create("Slim.Scrollbar.trough", "image", trough, sticky="nswe")
    for orientation in ("Vertical", "Horizontal"):
        style.layout(orientation + ".TScrollbar", [("Slim.Scrollbar.trough", {"sticky": "nswe", "children": [
            ("Slim.Scrollbar.thumb", {"expand": 1, "sticky": "nswe"})]})])
        style.configure(orientation + ".TScrollbar", width=10, borderwidth=0)
    for name, color in (("trough", COLORS["border"]), ("pbar", COLORS["accent"])):
        strip = tk.PhotoImage(master=root, width=1, height=3)
        strip.put(color, to=(0, 0, 1, 3))
        images["progress_" + name] = strip
        style.element_create("FlatProgress." + name, "image", strip, sticky="nswe")
    style.layout("Horizontal.TProgressbar", [("FlatProgress.trough", {"sticky": "we", "children": [
        ("FlatProgress.pbar", {"side": "left", "sticky": "ns"})]})])
    style.configure("Horizontal.TProgressbar", borderwidth=0, padding=0)
    style.configure("TSeparator", background=COLORS["border"])
    return body


class RoundedPanel(tk.Frame):
    """A canvas border gives native content rounded corners without dependencies."""

    def __init__(self, parent, padding=6, radius=10):
        try:
            backing = parent.cget("background")
        except tk.TclError:
            backing = ttk.Style(parent).lookup(parent.cget("style") or parent.winfo_class(), "background")
        backing = backing or COLORS["background"]
        super().__init__(parent, background=backing, borderwidth=0, highlightthickness=0)
        self.radius = radius
        self.border = COLORS["border"]
        self.canvas = tk.Canvas(self, background=backing, highlightthickness=0, borderwidth=0)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.content = ttk.Frame(self, style="Card.TFrame", padding=padding)
        self.content.pack(fill="both", expand=True, padx=radius, pady=radius)
        self.canvas.bind("<Configure>", self.draw)

    def draw(self, event=None):
        width, height, r = self.winfo_width() - 1, self.winfo_height() - 1, self.radius
        self.canvas.delete("border")
        self.canvas.create_polygon(1+r, 1, width-r, 1, width, 1, width, 1+r,
                                   width, height-r, width, height, width-r, height,
                                   1+r, height, 1, height, 1, height-r, 1, 1+r, 1, 1,
                                   smooth=True, splinesteps=24, fill=COLORS["surface"],
                                   outline=self.border, tags="border")

    def focus_border(self, focused):
        self.border = COLORS["accent"] if focused else COLORS["border"]
        self.draw()


class AutoScrollbar(ttk.Scrollbar):
    """Show only when content overflows; keep native dragging and keyboard behavior."""

    def set(self, first, last):
        if float(first) <= 0 and float(last) >= 1:
            self.grid_remove()
        else:
            self.grid()
        super().set(first, last)


def rounded_entry(parent, variable, width=None):
    panel = RoundedPanel(parent, padding=0, radius=7)
    options = dict(textvariable=variable)
    if width is not None:
        options["width"] = width
    entry = ttk.Entry(panel.content, **options)
    entry.pack(fill="x", expand=True)
    entry.bind("<FocusIn>", lambda event: panel.focus_border(True))
    entry.bind("<FocusOut>", lambda event: panel.focus_border(False))
    return panel, entry


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
        image.put(COLORS["accent"] if checked else COLORS["border"], to=(1, 1, 17, 17))
        image.put(COLORS["accent"] if checked else COLORS["surface"], to=(2, 2, 16, 16))
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
    def __init__(self, root, ignore_file=None, settings_file=None):
        self.root = root
        self.report = None
        self.running = False
        self.match_report = None
        self.events = queue.Queue()
        self.ignore_file = Path(ignore_file) if ignore_file is not None else default_ignore_file()
        self.settings_file = Path(settings_file) if settings_file is not None else self.ignore_file.with_name("settings.json")
        self.settings_error = ""
        try:
            preferences = load_settings(self.settings_file)
        except (OSError, ValueError) as error:
            preferences = default_settings()
            self.settings_error = str(error)
        self.ignore_error = ""
        try:
            self.ignored = load_ignored(self.ignore_file)
        except (OSError, ValueError) as error:
            self.ignored = set()
            self.ignore_error = str(error)
        root.title("中韩表头检查")
        root.geometry("1240x840")
        root.minsize(1060, 760)
        root.protocol("WM_DELETE_WINDOW", self.close)
        self.body_font = configure_theme(root)
        frame = ttk.Frame(root, padding=(24, 20, 24, 14))
        frame.pack(fill="both", expand=True)
        header = ttk.Frame(frame)
        header.pack(fill="x", pady=(0, 10))
        tk.Label(header, text="译", background=COLORS["accent"], foreground="white",
                 font=(self.body_font, 22, "bold"), padx=14, pady=6).pack(side="left", padx=(0, 14))
        wordmark = ttk.Frame(header)
        wordmark.pack(side="left")
        ttk.Label(wordmark, text="中韩译文工作台", style="Title.TLabel").pack(anchor="w")
        ttk.Label(wordmark, text="TRANSLATION DESK  /  表头检查 · 译文检索", style="Muted.TLabel").pack(anchor="w", pady=(4, 0))
        ttk.Label(header, text="源文件只读  /  结果另存副本", style="Badge.TLabel").pack(side="right")

        body = ttk.Frame(frame)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        sidebar_panel = RoundedPanel(body, padding=4)
        sidebar_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 18))
        sidebar = sidebar_panel.content
        sidebar.columnconfigure(0, weight=1)
        ttk.Label(sidebar, text="客户源文件夹", style="Section.TLabel").grid(row=1, column=0, sticky="w", pady=(0, 8))
        self.folder = tk.StringVar(value=preferences["source_folder"])
        folder_field, self.folder_entry = rounded_entry(sidebar, self.folder, width=28)
        folder_field.grid(row=3, column=0, sticky="ew")
        source_actions = ttk.Frame(sidebar, style="Card.TFrame")
        source_actions.grid(row=4, column=0, sticky="ew", pady=(4, 8))
        ttk.Label(source_actions, text="支持 .xlsx / .xlsm", style="CardMuted.TLabel").pack(side="left")
        self.browse_button = ttk.Button(source_actions, text="选择文件夹", command=self.browse)
        self.browse_button.pack(side="right")
        ttk.Separator(sidebar).grid(row=5, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(sidebar, text="表头名称 · 每行一个", style="Section.TLabel").grid(row=6, column=0, sticky="w")
        headers = ttk.Frame(sidebar, style="Card.TFrame")
        headers.grid(row=8, column=0, sticky="ew", pady=(8, 0))
        for column in (0, 1):
            headers.columnconfigure(column, weight=1, uniform="headers")
        for column, title, key in ((0, "中文  /  ZH", "chinese_headers"), (1, "韩语  /  KO", "korean_headers")):
            ttk.Label(headers, text=title, style="CardMuted.TLabel").grid(row=0, column=column, sticky="w", pady=(0, 6))
            editor_panel = RoundedPanel(headers, padding=0, radius=7)
            editor_panel.grid(row=1, column=column, sticky="ew", padx=(0, 6) if column == 0 else (6, 0))
            editor = tk.Text(editor_panel.content, height=6, width=13, wrap="none", font=(self.body_font, 10),
                             background=COLORS["surface"], foreground=COLORS["ink"], insertbackground=COLORS["accent"],
                             relief="flat", highlightthickness=0, padx=2, pady=2, undo=True)
            editor.insert("1.0", "\n".join(preferences[key]))
            editor.pack(fill="both", expand=True)
            editor.bind("<FocusIn>", lambda event, panel=editor_panel: panel.focus_border(True))
            editor.bind("<FocusOut>", lambda event, panel=editor_panel: panel.focus_border(False))
            if column == 0:
                self.chinese = editor
            else:
                self.korean = editor
        scope = ttk.Frame(sidebar, style="Card.TFrame")
        scope.grid(row=9, column=0, sticky="ew", pady=(6, 2))
        ttk.Label(scope, text="检查前", style="Card.TLabel").pack(side="left")
        self.rows = tk.StringVar(value=str(preferences["rows"]))
        row_field, self.row_entry = rounded_entry(scope, self.rows, width=4)
        row_field.pack(side="left", padx=8)
        ttk.Label(scope, text="行表头", style="Card.TLabel").pack(side="left")
        self.recursive = tk.BooleanVar(value=preferences["recursive"])
        self.include_hidden_sheets = tk.BooleanVar(value=preferences["include_hidden_sheets"])
        self.include_hidden_rows = tk.BooleanVar(value=preferences["include_hidden_rows"])
        self.scan_switches = []
        for row, (title, variable) in enumerate((("包含子文件夹", self.recursive), ("包含隐藏工作表", self.include_hidden_sheets),
                                                ("包含隐藏行", self.include_hidden_rows)), 10):
            switch = ttk.Checkbutton(sidebar, text=title, variable=variable)
            switch.grid(row=row, column=0, sticky="w")
            self.scan_switches.append(switch)
        sidebar.rowconfigure(13, weight=1)
        ttk.Separator(sidebar).grid(row=14, column=0, sticky="ew", pady=(8, 8))
        self.start_button = ttk.Button(sidebar, text="开始表头检查  →", style="Primary.TButton", command=self.start)
        self.start_button.grid(row=15, column=0, sticky="ew")
        self.save_settings_button = ttk.Button(sidebar, text="保存当前设置", style="Compact.TButton", command=lambda: self.store_settings(notify=True))
        self.save_settings_button.grid(row=16, column=0, sticky="ew", pady=(6, 0))

        workspace = ttk.Frame(body)
        workspace.grid(row=0, column=1, sticky="nsew")
        metrics = ttk.Frame(workspace)
        metrics.pack(fill="x", pady=(0, 16))
        self.metric_values = {}
        for column, (key, title) in enumerate((("files", "已检查文件"), ("passed", "通过工作表"), ("exceptions", "待确认异常"), ("ignored", "忽略记录"))):
            metrics.columnconfigure(column, weight=1, uniform="metrics")
            card_panel = RoundedPanel(metrics, padding=(6, 2))
            card_panel.grid(row=0, column=column, sticky="nsew", padx=(0, 10) if column < 3 else 0)
            card = card_panel.content
            value = tk.StringVar(value="—" if key != "ignored" else str(len(self.ignored)))
            self.metric_values[key] = value
            ttk.Label(card, textvariable=value, style="WarningNumber.TLabel" if key == "exceptions" else "Number.TLabel").pack(anchor="w")
            ttk.Label(card, text=title, style="CardMuted.TLabel").pack(anchor="w", pady=(4, 0))
        self.tabs = ttk.Notebook(workspace)
        self.tabs.pack(fill="both", expand=True)
        self.exception_tab = ttk.Frame(self.tabs)
        exception_panel = RoundedPanel(self.exception_tab, padding=6)
        exception_panel.pack(fill="both", expand=True)
        exception_content = exception_panel.content
        self.tabs.add(self.exception_tab, text="表头检查")
        title_row = ttk.Frame(exception_content, style="Card.TFrame")
        title_row.pack(fill="x", pady=(0, 12))
        ttk.Label(title_row, text="待确认工作表", style="Section.TLabel").pack(side="left")
        self.selection_count = tk.StringVar(value="已勾选 0 项")
        ttk.Label(title_row, textvariable=self.selection_count, style="CardMuted.TLabel").pack(side="right")
        controls = ttk.Frame(exception_content, style="Card.TFrame")
        controls.pack(fill="x", pady=(0, 10))
        ttk.Button(controls, text="全选", command=lambda: self.table.check_all()).pack(side="left")
        ttk.Button(controls, text="取消", command=lambda: self.table.uncheck_all()).pack(side="left", padx=(6, 0))
        self.export_button = ttk.Button(controls, text="导出清单", command=self.export, state="disabled")
        self.export_button.pack(side="left", padx=(16, 6))
        self.open_button = ttk.Button(controls, text="打开所在文件夹", command=self.open_selected, state="disabled")
        self.open_button.pack(side="left")
        self.ignore_button = ttk.Button(controls, text="忽略勾选", command=self.ignore_selected, state="disabled")
        self.ignore_button.pack(side="right")
        table_frame = ttk.Frame(exception_content, style="Card.TFrame")
        table_frame.pack(fill="both", expand=True)
        columns = ("path", "sheet", "issue", "chinese", "korean")
        self.table = CheckboxTreeview(table_frame, columns=columns)
        for name, title, width in zip(columns, ("相对路径", "工作表", "异常原因", "中文表头位置", "韩语表头位置"), (230, 135, 300, 160, 160)):
            self.table.heading(name, text=title)
            self.table.column(name, width=width, minwidth=100, stretch=False)
        self.layout_table(table_frame, self.table)
        self.table.bind("<Double-1>", self.double_click_details)
        self.table.bind("<Return>", lambda event: self.show_details())
        self.table.bind("<<TreeviewSelect>>", lambda event: self.set_running(self.running))
        self.exception_empty, self.empty_title, self.empty_description = self.empty_state(
            table_frame, "准备好，开始第一次检查", "选择左侧源文件夹并检查表头。\n需要人工确认的工作表会显示在这里。")
        ttk.Label(exception_content, text="方框支持多选 · 双击行或按 Enter 查看依据 · 空格切换勾选", style="CardMuted.TLabel").pack(side="bottom", anchor="w", pady=(12, 0), before=table_frame)

        self.match_tab = ttk.Frame(self.tabs)
        match_panel = RoundedPanel(self.match_tab, padding=8)
        match_panel.pack(fill="both", expand=True)
        match_content = match_panel.content
        self.tabs.add(self.match_tab, text="译文检索")
        ttk.Label(match_content, text="从检查，到译文匹配", style="Section.TLabel").pack(anchor="w")
        self.match_hint = ttk.Label(match_content, text="先完成表头检查，通过的工作表即可参与检索。", style="CardMuted.TLabel", wraplength=620)
        self.match_hint.pack(anchor="w", pady=(4, 8))
        self.baseline_folder = tk.StringVar(value=preferences["baseline_folder"])
        self.output_folder = tk.StringVar(value=preferences["output_folder"])
        self.match_inputs = []
        for label, variable, hint in (("基准文件夹", self.baseline_folder, "支持多个基准文件，按 G 列中文匹配，并结合前后三句推荐译文"),
                                      ("输出文件夹", self.output_folder, "选择独立目录，每次检索生成新的结果副本")):
            ttk.Label(match_content, text=label, style="Card.TLabel").pack(anchor="w", pady=(0, 6))
            line = ttk.Frame(match_content, style="Card.TFrame")
            line.pack(fill="x")
            field, entry = rounded_entry(line, variable)
            field.pack(side="left", fill="x", expand=True, padx=(0, 8))
            button = ttk.Button(line, text="选择文件夹", command=lambda v=variable, t=label: self.browse_match(v, t))
            button.pack(side="left")
            self.match_inputs.extend((entry, button))
            ttk.Label(match_content, text=hint, style="CardMuted.TLabel").pack(anchor="w", pady=(6, 14))
        ttk.Separator(match_content).pack(fill="x", pady=(0, 10))
        line = ttk.Frame(match_content, style="Card.TFrame")
        line.pack(fill="x")
        self.match_button = ttk.Button(line, text="检索并生成结果  →", style="Primary.TButton", command=self.start_match, state="disabled")
        self.match_button.pack(side="left")
        self.match_open_button = ttk.Button(line, text="打开结果文件夹", command=self.open_match_output, state="disabled")
        self.match_open_button.pack(side="left", padx=8)
        self.match_details_button = ttk.Button(line, text="打开匹配明细", command=self.open_match_details, state="disabled")
        self.match_details_button.pack(side="left")
        ttk.Label(match_content, text="检索结果", style="Section.TLabel").pack(anchor="w", pady=(20, 8))
        self.match_summary = tk.StringVar(value="等待检索。完成后将在这里显示匹配统计与结果位置。")
        summary_frame = ttk.Frame(match_content, style="Card.TFrame")
        summary_frame.pack(fill="both", expand=True)
        self.match_summary_text = tk.Text(summary_frame, height=5, width=1, wrap="word", relief="flat",
                                          highlightthickness=0, background=COLORS["surface"], foreground=COLORS["ink"],
                                          font=(self.body_font, 11), padx=0, pady=4, spacing3=6)
        summary_scroll = AutoScrollbar(summary_frame, orient="vertical", command=self.match_summary_text.yview)
        self.match_summary_text.configure(yscrollcommand=summary_scroll.set)
        self.match_summary_text.grid(row=0, column=0, sticky="nswe")
        summary_scroll.grid(row=0, column=1, sticky="ns", padx=(6, 0))
        summary_frame.rowconfigure(0, weight=1)
        summary_frame.columnconfigure(0, weight=1)
        self.match_summary.trace_add("write", self.render_match_summary)
        self.render_match_summary()
        self.match_rules_label = ttk.Label(match_content, text="异常及已忽略工作表自动跳过。候选译文写入副本 R～Z 列，AA～AC 列记录条数、缺译与冲突；完整来源另存匹配明细。", style="CardMuted.TLabel", wraplength=620, justify="left")
        self.match_rules_label.pack(side="bottom", anchor="w", fill="x", pady=(16, 0), before=summary_frame)
        match_content.bind("<Configure>", self.resize_match_text)

        self.ignored_tab = ttk.Frame(self.tabs)
        ignored_panel = RoundedPanel(self.ignored_tab, padding=6)
        ignored_panel.pack(fill="both", expand=True)
        ignored_content = ignored_panel.content
        self.tabs.add(self.ignored_tab, text="忽略管理")
        ttk.Label(ignored_content, text="已忽略的工作表", style="Section.TLabel").pack(anchor="w", pady=(0, 6))
        ttk.Label(ignored_content, text="移除记录后重新检查，即可恢复处理对应工作表。", style="CardMuted.TLabel").pack(anchor="w", pady=(0, 12))
        controls = ttk.Frame(ignored_content, style="Card.TFrame")
        controls.pack(fill="x", pady=(0, 10))
        ttk.Button(controls, text="全选", command=lambda: self.ignored_table.check_all()).pack(side="left")
        ttk.Button(controls, text="取消", command=lambda: self.ignored_table.uncheck_all()).pack(side="left", padx=6)
        self.restore_button = ttk.Button(controls, text="移除勾选记录", command=self.restore_selected)
        self.restore_button.pack(side="right")
        ignored_frame = ttk.Frame(ignored_content, style="Card.TFrame")
        ignored_frame.pack(fill="both", expand=True)
        self.ignored_table = CheckboxTreeview(ignored_frame, columns=("path", "sheet"))
        for name, label, width in (("path", "文件完整路径", 520), ("sheet", "工作表", 180)):
            self.ignored_table.heading(name, text=label)
            self.ignored_table.column(name, width=width, minwidth=120, stretch=False)
        self.layout_table(ignored_frame, self.ignored_table)
        self.ignored_empty, _, _ = self.empty_state(ignored_frame, "没有忽略记录", "确认无需处理的参考表后，\n可在表头检查列表中勾选并忽略。")
        self.ignore_location_label = ttk.Label(ignored_content, text=f"本地记录：{self.ignore_file}", style="CardMuted.TLabel", wraplength=620)
        self.ignore_location_label.pack(side="bottom", anchor="w", fill="x", pady=(12, 0), before=ignored_frame)
        ignored_content.bind("<Configure>", lambda event: self.ignore_location_label.configure(wraplength=max(200, event.width - 32)))

        footer = ttk.Frame(frame)
        footer.pack(side="bottom", fill="x", pady=(10, 0), before=body)
        self.progress = ttk.Progressbar(footer)
        self.progress.pack(fill="x", pady=(0, 9))
        self.status = tk.StringVar(value="就绪。选择客户源文件夹，开始表头检查；参考表也会列出，供人工确认。")
        self.status_label = ttk.Label(footer, textvariable=self.status, style="Muted.TLabel", wraplength=1160, justify="left")
        self.status_label.pack(anchor="w", fill="x")
        footer.bind("<Configure>", lambda event: self.status_label.configure(wraplength=max(200, event.width)))
        for table in (self.table, self.ignored_table):
            table.bind("<<ChecksChanged>>", lambda event: self.set_running(self.running))
        self.refresh_ignored()
        self.set_running(False)
        if self.ignore_error:
            self.status.set("本地忽略记录读取失败，检查已禁用。请修复记录文件后重启。")
            root.after(0, lambda: messagebox.showerror("忽略记录读取失败", self.ignore_error))
        if self.settings_error:
            root.after(0, lambda: messagebox.showerror("本地设置读取失败", self.settings_error + "\n已使用默认设置；保存已停用，原文件不会被覆盖。"))
        root.after(100, self.poll)

    def layout_table(self, frame, table):
        vertical = AutoScrollbar(frame, orient="vertical", command=table.yview)
        horizontal = AutoScrollbar(frame, orient="horizontal", command=table.xview)
        table.horizontal_scrollbar = horizontal
        table.configure(yscrollcommand=vertical.set,
                        xscrollcommand=lambda first, last: horizontal.set(first, last) if table.get_children() else horizontal.set(0, 1))
        table.tag_configure("alternate", background=COLORS["stripe"])
        table.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns", padx=(6, 0))
        horizontal.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def refresh_horizontal_scrollbar(self, table):
        # Row changes do not always trigger xscrollcommand when column widths stay fixed.
        table.horizontal_scrollbar.set(*(table.xview() if table.get_children() else (0, 1)))

    def empty_state(self, parent, title, description):
        panel = ttk.Frame(parent, style="Card.TFrame", padding=12)
        ttk.Label(panel, text="文  →  한", style="Number.TLabel").pack(pady=(0, 10))
        heading = ttk.Label(panel, text=title, style="Section.TLabel")
        heading.pack()
        detail = ttk.Label(panel, text=description, style="CardMuted.TLabel", justify="center")
        detail.pack(pady=(8, 0))
        panel.place(relx=0.5, rely=0.5, anchor="center")
        return panel, heading, detail

    def render_match_summary(self, *args):
        self.match_summary_text.configure(state="normal")
        self.match_summary_text.delete("1.0", "end")
        self.match_summary_text.insert("1.0", self.match_summary.get())
        self.match_summary_text.configure(state="disabled")

    def resize_match_text(self, event):
        for label in (self.match_hint, self.match_rules_label):
            label.configure(wraplength=max(200, event.width - 36))

    def browse(self):
        folder = filedialog.askdirectory(title="选择客户源文件夹")
        if folder:
            self.folder.set(folder)

    def update_overview(self):
        self.refresh_horizontal_scrollbar(self.table)
        report = self.report
        for key, value in (("files", report.files if report else "—"), ("passed", report.passed_sheets if report else "—"),
                           ("exceptions", len(report.exceptions) if report else "—"), ("ignored", len(self.ignored))):
            self.metric_values[key].set(str(value))
        self.selection_count.set(f"已勾选 {len(self.table.checked_items())} 项")
        if report and report.exceptions:
            self.exception_empty.place_forget()
        else:
            if report and not report.files:
                title, description = "没有找到表格文件", "请确认源文件夹与扫描范围，\n再重新执行表头检查。"
            elif report:
                title, description = "检查完成，未发现异常", "通过的工作表可进入译文检索。\n表头通过不代表译文质量已经验证。"
            elif self.running:
                title, description = "正在检查表头…", "检查完成后，需要确认的工作表会显示在这里。"
            else:
                title, description = "准备好，开始检查", "选择左侧源文件夹并检查表头。\n需要人工确认的工作表会显示在这里。"
            self.empty_title.configure(text=title)
            self.empty_description.configure(text=description)
            self.exception_empty.place(relx=0.5, rely=0.5, anchor="center")
        ready = report is not None and report.passed_sheets > 0
        self.match_hint.configure(text=(f"已有 {report.passed_sheets} 张工作表通过检查，可检索译文。异常与忽略项自动跳过。" if ready
                                       else "先完成表头检查，通过的工作表即可参与检索。"))

    def set_running(self, running):
        self.running = running
        self.update_overview()
        self.start_button.configure(text="正在检查…" if running and self.report is None else "开始表头检查  →")
        self.match_button.configure(text="正在检索…" if running and self.report is not None else "检索并生成结果  →")
        for widget in (self.start_button, self.browse_button, self.folder_entry, self.row_entry, *self.match_inputs, *self.scan_switches):
            widget.configure(state="disabled" if running else "normal")
        if self.ignore_error:
            self.start_button.configure(state="disabled")
        for widget in (self.chinese, self.korean):
            widget.configure(state="disabled" if running else "normal")
        available = not running and self.report is not None
        self.export_button.configure(state="normal" if available else "disabled")
        self.open_button.configure(state="normal" if available and self.table.selection() else "disabled")
        self.ignore_button.configure(state="normal" if available and self.table.checked_items() and not self.ignore_error else "disabled")
        self.restore_button.configure(state="normal" if not running and self.ignored_table.checked_items() and not self.ignore_error else "disabled")
        self.match_button.configure(state="normal" if available and self.report.passed_sheets and not self.ignore_error else "disabled")
        self.match_open_button.configure(state="normal" if not running and self.match_report else "disabled")
        self.match_details_button.configure(state="normal" if not running and self.match_report else "disabled")
        self.save_settings_button.configure(state="normal" if not running and not self.settings_error else "disabled")

    def store_settings(self, notify=False):
        if self.settings_error:
            return False
        try:
            try:
                rows = int(self.rows.get())
            except ValueError:
                raise ValueError("检查行数必须是大于等于 1 的整数") from None
            preferences = dict(version=1, source_folder=self.folder.get().strip(),
                               baseline_folder=self.baseline_folder.get().strip(), output_folder=self.output_folder.get().strip(),
                               rows=rows,
                               chinese_headers=[v for v in self.chinese.get("1.0", "end").splitlines() if v.strip()],
                               korean_headers=[v for v in self.korean.get("1.0", "end").splitlines() if v.strip()],
                               recursive=self.recursive.get(), include_hidden_sheets=self.include_hidden_sheets.get(),
                               include_hidden_rows=self.include_hidden_rows.get())
            save_settings(preferences, self.settings_file)
        except (OSError, ValueError) as error:
            messagebox.showerror("设置保存失败", f"{error}\n原有设置已保留。")
            return False
        if notify:
            messagebox.showinfo("设置已保存", "文件夹、表头、检查行数及扫描开关已保存，重启后恢复。")
        return True

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
        self.store_settings()
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

    def open_match_details(self):
        if self.match_report:
            try:
                open_path(self.match_report.destination / "匹配明细.xlsx")
            except OSError as error:
                messagebox.showerror("无法打开匹配明细", str(error))

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
        self.store_settings()
        self.report = None
        self.tabs.select(self.exception_tab)
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
                        summary = (f"检索完成：基准 {value.baseline_files} 个文件、{value.baseline_rows} 行，匹配 {value.matched_rows} 行，推荐 {value.recommended_rows} 行；"
                                   f"源文件参与 {value.source_rows} 行，匹配 {value.source_matched} 行；跳过 {value.skipped} 项。\n"
                                   f"缺译 {value.missing_translations} 行；冲突中文 {value.conflict_texts} 条；完整匹配对应 {value.match_pairs} 条。\n"
                                   f"结果位置：{value.destination}\n逐基准统计见“基准文件匹配统计.xlsx”；源文件统计见“源文件匹配统计.xlsx”；上下文依据及各类明细见“匹配明细.xlsx”。")
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
            self.update_overview()
            return
        report = self.report
        for index, result in enumerate(report.exceptions):
            self.table.insert("", "end", iid=str(index), values=(str(result.path.relative_to(report.root)), result.sheet, result.issue, result.chinese, result.korean), tags=("alternate",) if index % 2 else ())
        self.update_overview()
        self.status.set(f"已检查 {report.files} 个文件；{report.passed_sheets} 张表通过；{report.abnormal_files} 个文件需确认，共 {len(report.exceptions)} 项异常；已忽略 {report.ignored_sheets} 张表；按开关跳过 {len(report.skipped_sheets)} 张表。")

    def refresh_ignored(self):
        self.ignored_table.delete(*self.ignored_table.get_children())
        self.ignored_rows = sorted(self.ignored)
        for index, (filename, sheet) in enumerate(self.ignored_rows):
            self.ignored_table.insert("", "end", iid=str(index), values=(filename, sheet), tags=("alternate",) if index % 2 else ())
        self.refresh_horizontal_scrollbar(self.ignored_table)
        self.tabs.tab(self.ignored_tab, text=f"忽略管理 · {len(self.ignored_rows)}")
        self.metric_values["ignored"].set(str(len(self.ignored_rows)))
        if self.ignored_rows:
            self.ignored_empty.place_forget()
        else:
            self.ignored_empty.place(relx=0.5, rely=0.5, anchor="center")

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
            window.configure(background=COLORS["background"])
            frame = ttk.Frame(window, padding=18)
            frame.pack(fill="both", expand=True)
            text = tk.Text(frame, wrap="word", padx=16, pady=16, font=(self.body_font, 11),
                           spacing1=3, spacing3=6, relief="flat", background=COLORS["surface"],
                           foreground=COLORS["ink"], highlightthickness=1, highlightbackground=COLORS["border"])
            scroll = AutoScrollbar(frame, orient="vertical", command=text.yview)
            text.configure(yscrollcommand=scroll.set)
            text.grid(row=0, column=0, sticky="nswe")
            scroll.grid(row=0, column=1, sticky="ns", padx=(6, 0))
            frame.rowconfigure(0, weight=1)
            frame.columnconfigure(0, weight=1)
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
            self.store_settings()
            self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    Application(root)
    root.mainloop()

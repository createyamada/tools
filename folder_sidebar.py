import ctypes
from ctypes import wintypes
from datetime import datetime, timedelta
import html
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox
import webbrowser
import xml.etree.ElementTree as ET
from urllib.parse import unquote

# ------------------------
# EXE対応・設定ファイル
# ------------------------

BASE_DIR = os.path.dirname(
    sys.executable
    if getattr(sys, "frozen", False)
    else os.path.abspath(__file__)
)
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

DEFAULT_CONFIG = {
    "target_folder": r"C:\Work",
    "window": {
        "width": 1100,
        "height": 720,
        "x": 0,
        "y": 0,
    },
    "column_widths": [330, 350, 400],
    "alpha": 0.97,
    "topmost": False,
    "user_name": "",
    "onenote_pages": [],
}

def save_config(config):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as file:
            json.dump(config, file, indent=4, ensure_ascii=False)
    except OSError:
        pass

def load_config():
    config = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as file:
                config = json.load(file)
        except (OSError, json.JSONDecodeError):
            pass
    merged = {**DEFAULT_CONFIG, **config}
    merged["window"] = {**DEFAULT_CONFIG["window"], **config.get("window", {})}
    if not os.path.exists(CONFIG_FILE):
        save_config(merged)
    return merged

CONFIG = load_config()


# ------------------------
# COLOR
# ------------------------

BG_COLOR = "#0F172A"
CARD_COLOR = "#111827"
LIST_BG = "#1E293B"
TEXT_COLOR = "#F8FAFC"
SUB_TEXT = "#94A3B8"
SELECT_BG = "#2563EB"
BUTTON_BG = "#334155"
BUTTON_ACTIVE = "#475569"

def hotkey_listener(app):
    """Ctrl + Alt + Spaceでウィンドウの表示を切り替える。"""
    user32 = ctypes.windll.user32
    user32.RegisterHotKey(None, 1, 0x0001 | 0x0002, 0x20)
    msg = wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
        if msg.message == 0x0312:
            app.root.after(0, app.toggle_sidebar)

class FolderSidebar:
    def __init__(self, root):
        self.root, self.config = root, CONFIG
        self.current_date = datetime.now().date()
        self.schedule_urls, self.display_to_file = {}, {}
        self.resize_start = self.save_job = None
        self.initial_sashes_set = False
        window = self.config["window"]
        root.overrideredirect(True)
        root.attributes("-topmost", bool(self.config["topmost"]))
        root.attributes("-alpha", float(self.config["alpha"]))
        root.geometry(f'{window["width"]}x{window["height"]}+{window["x"]}+{window["y"]}')
        root.minsize(720, 360)
        root.configure(bg=BG_COLOR)
        self.is_topmost = bool(self.config["topmost"])
        self.build_header()
        self.panes = tk.PanedWindow(root, orient="horizontal", bg="white", sashwidth=5,
                                    sashrelief="flat", borderwidth=0, opaqueresize=True)
        self.panes.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        folder_panel = self.make_panel("📁 フォルダー一覧")
        schedule_panel = self.make_panel("📅 スケジュール")
        onenote_panel = self.make_panel("🟣 OneNote タスク管理")
        self.panes.add(folder_panel, minsize=190)
        self.panes.add(schedule_panel, minsize=220)
        self.panes.add(onenote_panel, minsize=260)
        self.build_folder_panel(folder_panel)
        self.build_schedule_panel(schedule_panel)
        self.build_onenote_panel(onenote_panel)
        self.build_resize_grip()
        root.bind("<Escape>", lambda _e: root.destroy())
        root.bind("<Control-Shift-space>", self.toggle_topmost)
        root.bind("<Configure>", self.schedule_config_save)
        self.panes.bind("<ButtonRelease-1>", self.save_layout)
        root.after(100, self.restore_sashes)
        self.refresh()

    def build_header(self):
        # ------------------------
        # HEADER
        # ------------------------

        header = tk.Frame(self.root, bg=BG_COLOR)
        header.pack(fill="x", padx=14, pady=(10, 8))
        header.bind("<ButtonPress-1>", self.start_move)
        header.bind("<B1-Motion>", self.move_window)
        title = tk.Label(header, text="WORKSPACE SIDEBAR", bg=BG_COLOR, fg="white",
                         font=("Yu Gothic UI", 15, "bold"))
        title.pack(side="left")
        title.bind("<ButtonPress-1>", self.start_move)
        title.bind("<B1-Motion>", self.move_window)
        tk.Button(header, text="✕", command=self.root.destroy, bg=BG_COLOR, fg=SUB_TEXT,
                  activebackground=BG_COLOR, borderwidth=0).pack(side="right")
        self.pin_label = tk.Label(header, text="📌 ON" if self.is_topmost else "📌 OFF",
                                  bg=BG_COLOR, fg="#60A5FA" if self.is_topmost else SUB_TEXT)
        self.pin_label.pack(side="right", padx=8)
        self.pin_label.bind("<Button-1>", self.toggle_topmost)

    def make_panel(self, title):
        panel = tk.Frame(self.panes, bg=CARD_COLOR)

        tk.Label(
            panel,
            text=title,
            bg=CARD_COLOR,
            fg=TEXT_COLOR,
            font=("Yu Gothic UI", 11, "bold"),
            anchor="w",
        ).pack(
            fill="x",
            padx=12,
            pady=(12, 8),
        )

        return panel

    def make_entry(self, parent, width=None):
        return tk.Entry(parent, width=width, bg="#374151", fg="white", insertbackground="white",
                        borderwidth=0, font=("Yu Gothic UI", 9))

    def small_button(self, parent, text, command):
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=BUTTON_BG,
            fg="white",
            borderwidth=0,
        )

    def build_folder_panel(self, panel):
        # ------------------------
        # フォルダー一覧
        # ------------------------

        self.listbox = tk.Listbox(panel, bg=LIST_BG, fg=TEXT_COLOR, selectbackground=SELECT_BG,
                                  activestyle="none", borderwidth=0, highlightthickness=0,
                                  font=("Yu Gothic UI", 10))
        self.listbox.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        self.listbox.bind("<Double-Button-1>", self.open_file)
        register = tk.Frame(panel, bg=CARD_COLOR)
        register.pack(fill="x", padx=10, pady=(0, 8))
        self.path_entry = self.make_entry(register)
        self.path_entry.pack(side="left", fill="x", expand=True, ipady=6)
        tk.Button(register, text="＋", command=self.create_shortcut, bg=SELECT_BG, fg="white",
                  borderwidth=0, width=4).pack(side="right", padx=(6, 0))
        tk.Button(
            panel,
            text="↻ 更新",
            command=self.refresh,
            bg=BUTTON_BG,
            fg="white",
            activebackground=BUTTON_ACTIVE,
            borderwidth=0,
        ).pack(
            anchor="w",
            padx=10,
            pady=(0, 10),
        )

    def build_schedule_panel(self, panel):
        # ------------------------
        # Outlookスケジュール
        # ------------------------

        nav = tk.Frame(panel, bg=CARD_COLOR)
        nav.pack(fill="x", padx=10, pady=(0, 8))
        self.small_button(nav, "本日", self.show_today).pack(side="left")
        self.small_button(nav, "＜", self.prev_day).pack(side="left", padx=(5, 2))
        self.schedule_title = tk.Label(
            nav,
            bg=CARD_COLOR,
            fg="white",
            font=("Yu Gothic UI", 10, "bold"),
        )
        self.schedule_title.pack(side="left", expand=True)
        self.small_button(nav, "＞", self.next_day).pack(side="right")
        self.schedule_listbox = tk.Listbox(panel, bg=LIST_BG, fg=TEXT_COLOR, selectbackground=SELECT_BG,
                                           activestyle="none", borderwidth=0, highlightthickness=0,
                                           font=("Yu Gothic UI", 10))
        self.schedule_listbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.schedule_listbox.bind("<Double-Button-1>", self.open_meeting)
        self.update_date_label()

    def build_onenote_panel(self, panel):
        # ------------------------
        # OneNoteタスク管理
        # 担当者とページリンクはconfig.jsonに記載する
        # ------------------------

        configured_user = self.config.get("user_name") or "未設定"

        tk.Label(
            panel,
            text=f"担当者: {configured_user}",
            bg=CARD_COLOR,
            fg=SUB_TEXT,
            anchor="w",
        ).pack(fill="x", padx=10, pady=(0, 6))

        tk.Button(panel, text="↻ OneNote更新", command=self.load_onenote_tasks,
                  bg=BUTTON_BG, fg="white", borderwidth=0).pack(anchor="w", padx=10, pady=(0, 6))
        container = tk.Frame(panel, bg=LIST_BG)
        container.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.note_canvas = tk.Canvas(container, bg=LIST_BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=self.note_canvas.yview)
        self.note_content = tk.Frame(self.note_canvas, bg=LIST_BG)
        self.note_window = self.note_canvas.create_window(
            (0, 0),
            window=self.note_content,
            anchor="nw",
        )
        self.note_canvas.configure(yscrollcommand=scrollbar.set)
        self.note_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.note_content.bind(
            "<Configure>",
            lambda _e: self.note_canvas.configure(
                scrollregion=self.note_canvas.bbox("all")
            ),
        )
        self.note_canvas.bind(
            "<Configure>",
            lambda event: self.note_canvas.itemconfigure(
                self.note_window,
                width=event.width,
            ),
        )

    def build_resize_grip(self):
        grip = tk.Label(self.root, text="◢", bg=BG_COLOR, fg=SUB_TEXT, cursor="size_nw_se")
        grip.place(relx=1, rely=1, anchor="se")
        grip.bind("<ButtonPress-1>", self.start_resize)
        grip.bind("<B1-Motion>", self.resize_window)

    def start_move(self, event):
        self.move_start = (
            event.x_root - self.root.winfo_x(),
            event.y_root - self.root.winfo_y(),
        )

    def move_window(self, event):
        self.root.geometry(f"+{event.x_root - self.move_start[0]}+{event.y_root - self.move_start[1]}")

    def start_resize(self, event):
        self.resize_start = (
            event.x_root,
            event.y_root,
            self.root.winfo_width(),
            self.root.winfo_height(),
        )

    def resize_window(self, event):
        if self.resize_start:
            x, y, width, height = self.resize_start
            self.root.geometry(f"{max(720, width + event.x_root - x)}x{max(360, height + event.y_root - y)}")

    def restore_sashes(self):
        widths = self.config.get("column_widths", DEFAULT_CONFIG["column_widths"])
        if len(widths) >= 2:
            self.panes.sash_place(0, int(widths[0]), 0)
            self.panes.sash_place(1, int(widths[0]) + int(widths[1]) + 5, 0)
        self.initial_sashes_set = True

    def schedule_config_save(self, _event=None):
        if self.save_job:
            self.root.after_cancel(self.save_job)
        self.save_job = self.root.after(400, self.save_layout)

    def save_layout(self, _event=None):
        # ------------------------
        # ウィンドウ・列幅保存
        # ------------------------

        if not self.initial_sashes_set:
            return
        total, first = self.panes.winfo_width(), self.panes.sash_coord(0)[0]
        second = self.panes.sash_coord(1)[0]
        self.config["column_widths"] = [first, max(1, second-first-5), max(1, total-second-5)]
        self.config["window"] = {
            "width": self.root.winfo_width(),
            "height": self.root.winfo_height(),
            "x": self.root.winfo_x(),
            "y": self.root.winfo_y(),
        }
        save_config(self.config)

    def update_date_label(self):
        weekdays = ["月", "火", "水", "木", "金", "土", "日"]
        d = self.current_date
        self.schedule_title.config(text=f"{d:%Y/%m/%d}（{weekdays[d.weekday()]}）")

    def prev_day(self):
        self.current_date -= timedelta(days=1)
        self.update_date_label()
        self.load_schedule()

    def next_day(self):
        self.current_date += timedelta(days=1)
        self.update_date_label()
        self.load_schedule()

    def show_today(self):
        self.current_date = datetime.now().date()
        self.update_date_label()
        self.load_schedule()

    def refresh(self):
        # ------------------------
        # 全データ更新
        # ------------------------

        self.listbox.delete(0, tk.END)
        self.display_to_file = {}

        folder = Path(self.config["target_folder"])

        if folder.exists():
            for path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
                display = path.stem if path.suffix.lower() == ".lnk" else path.name
                self.listbox.insert(tk.END, f"{self.get_icon(path)}  {display}")
                self.display_to_file[display] = path.name
        else:
            self.listbox.insert(tk.END, f"フォルダーが見つかりません: {folder}")

        self.load_schedule()
        self.load_onenote_tasks()

    def get_icon(self, path):
        if path.is_dir():
            return "📁"

        icons = {
            ".png": "🖼",
            ".jpg": "🖼",
            ".jpeg": "🖼",
            ".txt": "📄",
            ".xlsx": "📊",
            ".xls": "📊",
            ".pdf": "📕",
            ".py": "🐍",
        }

        return icons.get(path.suffix.lower(), "📄")

    def open_file(self, _event):
        selection = self.listbox.curselection()
        if selection:
            display = self.listbox.get(selection[0]).split("  ", 1)[-1]
            path = Path(self.config["target_folder"]) / self.display_to_file.get(display, display)
            if path.exists():
                os.startfile(str(path))

    def toggle_topmost(self, _event=None):
        self.is_topmost = not self.is_topmost
        self.root.attributes("-topmost", self.is_topmost)
        self.pin_label.config(text="📌 ON" if self.is_topmost else "📌 OFF",
                              fg="#60A5FA" if self.is_topmost else SUB_TEXT)
        self.config["topmost"] = self.is_topmost
        save_config(self.config)

    def create_shortcut(self):
        target_text = self.path_entry.get().strip().strip('"')
        if not target_text or not os.path.exists(target_text):
            messagebox.showwarning("登録", "存在するファイルまたはフォルダーのパスを入力してください。")
            return
        target = Path(target_text)
        shortcut = Path(self.config["target_folder"]) / (target.stem + ".lnk")
        esc_target = str(target).replace("'", "''")
        esc_shortcut = str(shortcut).replace("'", "''")
        esc_work = str(target.parent).replace("'", "''")

        script = (
            f"$w=New-Object -ComObject WScript.Shell; "
            f"$s=$w.CreateShortcut('{esc_shortcut}'); "
            f"$s.TargetPath='{esc_target}'; "
            f"$s.WorkingDirectory='{esc_work}'; "
            "$s.Save()"
        )

        try:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    script,
                ],
                check=True,
                capture_output=True,
            )
            self.path_entry.delete(0, tk.END)
            self.refresh()
        except subprocess.CalledProcessError:
            messagebox.showerror("登録", "ショートカットを作成できませんでした。")

    def load_schedule(self):
        # ------------------------
        # Outlook予定取得
        # ------------------------

        self.schedule_listbox.delete(0, tk.END)
        self.schedule_urls = {}

        target = self.current_date
        next_date = self.current_date + timedelta(days=1)

        script = r'''
$ErrorActionPreference = "Stop"

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$items = $namespace.GetDefaultFolder(9).Items

$items.Sort("[Start]")
$items.IncludeRecurrences = $true

$targetDate = Get-Date "__START__"
$nextDate = Get-Date "__END__"

$filter =
    "[Start] >= '" +
    $targetDate.ToString("g") +
    "' AND [Start] < '" +
    $nextDate.ToString("g") +
    "'"

$appointments = $items.Restrict($filter)

foreach ($appointment in $appointments)
{
    $body = ""

    try
    {
        $body = $appointment.Body
    }
    catch
    {
    }

    $urls =
        [regex]::Matches($body, 'https?://[^\s<>"'']+') |
        ForEach-Object { $_.Value }

    $url =
        $urls |
        Where-Object { $_ -match 'zoom\.us' } |
        Select-Object -First 1

    if (-not $url)
    {
        $url =
            $urls |
            Where-Object { $_ -match 'teams\.(microsoft|live)\.com' } |
            Select-Object -First 1
    }

    Write-Output (
        $appointment.Start.ToString("HH:mm") + "|" +
        $appointment.End.ToString("HH:mm") + "|" +
        $appointment.Subject.Replace("|", " ") + "|" +
        $url
    )
}
'''

        script = script.replace(
            "__START__",
            target.strftime("%Y-%m-%d"),
        ).replace(
            "__END__",
            next_date.strftime("%Y-%m-%d"),
        )

        try:
            result = subprocess.check_output(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    script,
                ],
                text=True,
                encoding="utf-8",
                errors="ignore",
                timeout=30,
            )

            for line in result.splitlines():
                parts = line.strip().split("|", 3)

                if len(parts) == 4:
                    start, end, title, url = parts
                    index = self.schedule_listbox.size()

                    self.schedule_listbox.insert(tk.END, f"{start}～{end}  {title}")

                    if url:
                        self.schedule_urls[index] = url
                        self.schedule_listbox.itemconfig(index, fg="#60A5FA")
        except Exception:
            self.schedule_listbox.insert(tk.END, "Outlookの予定を取得できませんでした")

    def open_meeting(self, _event):
        selection = self.schedule_listbox.curselection()

        if selection and selection[0] in self.schedule_urls:
            webbrowser.open(self.schedule_urls[selection[0]])

    def load_onenote_tasks(self):
        # ------------------------
        # OneNoteタスク取得・表示
        # ------------------------

        for child in self.note_content.winfo_children():
            child.destroy()

        pages = self.config.get("onenote_pages", [])

        if not pages:
            self.note_message(
                "config.jsonのonenote_pagesにページを登録してください",
                SUB_TEXT,
            )
            return

        if not self.config.get("user_name"):
            self.note_message(
                "config.jsonのuser_nameに担当者名を設定してください",
                "#FBBF24",
            )
            return
        for page in pages:
            link = page.get("link", "")
            title = tk.Label(self.note_content, text=f'▸ {page.get("name", "OneNote")}', bg=LIST_BG,
                             fg="#A78BFA", cursor="hand2", anchor="w", font=("Yu Gothic UI", 10, "bold"))
            title.pack(fill="x", padx=10, pady=(9, 3))
            title.bind("<Button-1>", lambda _e, url=link: self.open_onenote(url))
            try:
                tasks = self.get_onenote_tasks(link, self.config["user_name"])

                if tasks:
                    for task, due in tasks:
                        self.note_message(f"    {task}    {due}", TEXT_COLOR)
                else:
                    self.note_message("    該当するタスクはありません", SUB_TEXT)
            except Exception as error:
                self.note_message(f"    取得できません: {error}", "#FCA5A5")

    def note_message(self, text, color):
        tk.Label(self.note_content, text=text, bg=LIST_BG, fg=color, anchor="w", justify="left",
                 font=("Yu Gothic UI", 9), wraplength=340).pack(fill="x", padx=10, pady=2)

    def open_onenote(self, link):
        try:
            os.startfile(link)
        except OSError:
            webbrowser.open(link)

    def get_onenote_tasks(self, link, user_name):
        decoded_link = unquote(link)
        match = re.search(r"page-id=(\{?[0-9a-fA-F-]{36}\}?)", decoded_link, re.I)

        if not match:
            raise ValueError("リンクにpage-idがありません")

        page_id = "{" + match.group(1).strip("{}").upper() + "}"

        script = (
            "$ErrorActionPreference='Stop'; "
            "$oneNote=New-Object -ComObject OneNote.Application; "
            "$xml=''; "
            f"$oneNote.GetPageContent('{page_id}',[ref]$xml); "
            "[Console]::OutputEncoding=[Text.Encoding]::UTF8; "
            "[Console]::Write($xml)"
        )

        xml_text = subprocess.check_output(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                script,
            ],
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
        )

        return self.parse_onenote_tables(xml_text, user_name)

    @staticmethod
    def parse_onenote_tables(xml_text, user_name):
        """OneNote XML内の表から担当者に一致するタスクを抽出する。"""

        root = ET.fromstring(xml_text.lstrip("\ufeff"))
        tasks = []

        for table in root.iter():
            if not table.tag.endswith("Table"):
                continue

            rows = []

            for row in list(table):
                if not row.tag.endswith("Row"):
                    continue

                cells = []

                for cell in list(row):
                    if cell.tag.endswith("Cell"):
                        raw = " ".join(cell.itertext())
                        text = html.unescape(
                            re.sub(r"<[^>]+>", "", raw)
                        ).strip()
                        cells.append(text)

                if cells:
                    rows.append(cells)

            if not rows:
                continue

            headers = [re.sub(r"\s+", "", value) for value in rows[0]]

            try:
                assignee = headers.index("担当者")
                task_column = headers.index("タスク名")
                due_date = headers.index("期限")
            except ValueError:
                continue

            for row in rows[1:]:
                if max(assignee, task_column, due_date) >= len(row):
                    continue

                names = [
                    name.strip().casefold()
                    for name in re.split(r"[,、;/\n]", row[assignee])
                ]

                if user_name.strip().casefold() in names:
                    tasks.append(
                        (
                            row[task_column] or "（タスク名なし）",
                            row[due_date] or "期限なし",
                        )
                    )

        return tasks

    def toggle_sidebar(self):
        if self.root.state() == "withdrawn":
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        else:
            self.root.withdraw()


# ------------------------
# START
# ------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = FolderSidebar(root)

    threading.Thread(
        target=hotkey_listener,
        args=(app,),
        daemon=True,
    ).start()

    root.mainloop()

import os
import tkinter as tk
from pathlib import Path
import subprocess
import sys
import threading
import json
import base64
from datetime import datetime, timedelta
import re
import webbrowser
import ctypes
from ctypes import wintypes
import html
import xml.etree.ElementTree as ET
from urllib.parse import unquote
from tkinter import messagebox

# ------------------------
# EXE対応・設定ファイル
# ------------------------

# ------------------------
# EXE対応
# ------------------------

if getattr(sys, 'frozen', False):

    BASE_DIR = os.path.dirname(
        sys.executable
    )

else:

    BASE_DIR = os.path.dirname(
        os.path.abspath(__file__)
    )


# ------------------------
# CONFIG
# ------------------------

CONFIG_FILE = os.path.join(
    BASE_DIR,
    "config.json"
)


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

# ------------------------
# CONFIG SAVE
# ------------------------

def save_config(config):

    try:

        with open(CONFIG_FILE, "w", encoding="utf-8") as file:

            json.dump(config, file, indent=4, ensure_ascii=False)

        return True

    except OSError:

        return False

# ------------------------
# CONFIG LOAD
# ------------------------

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
CONTROL_AREA_HEIGHT = 72

# ------------------------
# HOTKEY
# ------------------------

def hotkey_listener(app):

    user32 = ctypes.windll.user32

    user32.RegisterHotKey(None, 1, 0x0001 | 0x0002, 0x20)

    msg = wintypes.MSG()

    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:

        if msg.message == 0x0312:

            app.root.after(0, app.toggle_sidebar)

# ------------------------
# MAIN WINDOW
# ------------------------

class FolderSidebar:

    def __init__(self, root):

        self.root = root
        self.config = CONFIG

        self.current_date = datetime.now().date()

        self.schedule_urls = {}
        self.display_to_file = {}

        self.resize_start = None
        self.save_job = None
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

        self.panes = tk.PanedWindow(
            root,
            orient="horizontal",
            bg="white",
            sashwidth=2,
            sashrelief="flat",
            borderwidth=0,
            opaqueresize=True
        )

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

        title = tk.Label(
            header,
            text="WORKSPACE WIDGET",
            bg=BG_COLOR,
            fg="white",
            font=("Yu Gothic UI", 15, "bold")
        )

        title.pack(side="left")

        title.bind("<ButtonPress-1>", self.start_move)
        title.bind("<B1-Motion>", self.move_window)

        close_button = tk.Button(
            header,
            text="✕",
            command=self.root.destroy,
            bg=BG_COLOR,
            fg=SUB_TEXT,
            activebackground=BG_COLOR,
            borderwidth=0
        )

        close_button.pack(side="right")

        self.pin_label = tk.Label(
            header,
            text="📌 ON" if self.is_topmost else "📌 OFF",
            bg=BG_COLOR,
            fg="#60A5FA" if self.is_topmost else SUB_TEXT
        )

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

        return tk.Entry(
            parent,
            width=width,
            bg="#374151",
            fg="white",
            insertbackground="white",
            borderwidth=0,
            font=("Yu Gothic UI", 9)
        )

    # ------------------------
    # BUTTON
    # ------------------------

    def small_button(self, parent, text, command):
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=BUTTON_BG,
            fg="white",
            borderwidth=0,
        )

    def update_button(self, parent, text, command):

        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=BUTTON_BG,
            fg="white",
            activebackground=BUTTON_ACTIVE,
            activeforeground="white",
            borderwidth=0
        )

        button.pack(
            anchor="w",
            padx=0,
            pady=(0, 6)
        )

    def control_area(self, panel):

        area = tk.Frame(
            panel,
            bg=CARD_COLOR,
            height=CONTROL_AREA_HEIGHT
        )

        area.pack(
            fill="x",
            padx=10
        )

        area.pack_propagate(False)

        return area

    def build_folder_panel(self, panel):
        # ------------------------
        # フォルダー一覧
        # ------------------------

        control = self.control_area(panel)

        folder_area = tk.Frame(
            control,
            bg=CARD_COLOR
        )

        folder_area.pack(
            fill="x",
            pady=(0, 6)
        )

        tk.Button(
            folder_area,
            text="↻ フォルダー更新",
            command=self.refresh_folder,
            bg=BUTTON_BG,
            fg="white",
            activebackground=BUTTON_ACTIVE,
            activeforeground="white",
            borderwidth=0
        ).pack(side="left")

        self.target_folder_var = tk.StringVar(
            value=self.config["target_folder"]
        )

        target_entry = tk.Entry(
            folder_area,
            textvariable=self.target_folder_var,
            state="readonly",
            readonlybackground="#374151",
            fg="white",
            borderwidth=0,
            font=("Yu Gothic UI", 9)
        )

        target_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(6, 0),
            ipady=4
        )

        register = tk.Frame(control, bg=CARD_COLOR)
        register.pack(fill="x")
        self.path_entry = self.make_entry(register)
        self.path_entry.pack(side="left", fill="x", expand=True, ipady=6)
        tk.Button(register, text="＋", command=self.create_shortcut, bg=SELECT_BG, fg="white",
                  borderwidth=0, width=4).pack(side="right", padx=(6, 0))

        self.listbox = tk.Listbox(panel, bg=LIST_BG, fg=TEXT_COLOR, selectbackground=SELECT_BG,
                                  activestyle="none", borderwidth=0, highlightthickness=0,
                                  font=("Yu Gothic UI", 10))
        self.listbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.listbox.bind("<Double-Button-1>", self.open_file)

    def build_schedule_panel(self, panel):
        # ------------------------
        # Outlookスケジュール
        # ------------------------

        control = self.control_area(panel)

        self.update_button(
            control,
            "↻ スケジュール更新",
            self.load_schedule
        )

        nav = tk.Frame(control, bg=CARD_COLOR)
        nav.pack(fill="x")
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

        control = self.control_area(panel)

        button_area = tk.Frame(
            control,
            bg=CARD_COLOR
        )

        button_area.pack(
            fill="x",
            pady=(0, 6)
        )

        tk.Button(
            button_area,
            text="↻ OneNote更新",
            command=self.load_onenote_tasks,
            bg=BUTTON_BG,
            fg="white",
            activebackground=BUTTON_ACTIVE,
            activeforeground="white",
            borderwidth=0
        ).pack(side="left")

        tk.Button(
            button_area,
            text="＋ ページ追加",
            command=self.show_onenote_page_dialog,
            bg=SELECT_BG,
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            borderwidth=0
        ).pack(side="left", padx=(6, 0))

        tk.Label(
            control,
            text=f"担当者: {configured_user}",
            bg=CARD_COLOR,
            fg=SUB_TEXT,
            anchor="w",
        ).pack(fill="x")

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

    # ------------------------
    # WINDOW MOVE
    # ------------------------

    def start_move(self, event):
        self.move_start = (
            event.x_root - self.root.winfo_x(),
            event.y_root - self.root.winfo_y(),
        )

    def move_window(self, event):
        self.root.geometry(f"+{event.x_root - self.move_start[0]}+{event.y_root - self.move_start[1]}")

    # ------------------------
    # WINDOW RESIZE
    # ------------------------

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

    # ------------------------
    # LAYOUT RESTORE
    # ------------------------

    def restore_sashes(self):
        widths = self.config.get("column_widths", DEFAULT_CONFIG["column_widths"])
        if len(widths) >= 2:
            self.panes.sash_place(0, int(widths[0]), 0)
            self.panes.sash_place(1, int(widths[0]) + int(widths[1]) + 5, 0)
        self.initial_sashes_set = True

    def schedule_config_save(self, event=None):

        if event is not None and event.widget is not self.root:
            return

        if self.save_job:
            self.root.after_cancel(self.save_job)
        self.save_job = self.root.after(400, self.save_layout)

    # ------------------------
    # LAYOUT SAVE
    # ------------------------

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

    # ------------------------
    # SCHEDULE DATE
    # ------------------------

    def update_date_label(self):

        weekdays = ["月", "火", "水", "木", "金", "土", "日"]

        d = self.current_date

        self.schedule_title.config(
            text=f"{d:%Y/%m/%d}（{weekdays[d.weekday()]}）"
        )

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

    # ------------------------
    # REFRESH
    # ------------------------

    def refresh(self):

        self.refresh_folder()
        self.load_schedule()
        self.load_onenote_tasks()

    def refresh_folder(self):

        try:

            with open(CONFIG_FILE, "r", encoding="utf-8") as file:

                saved_config = json.load(file)

            saved_target = saved_config.get("target_folder")

            if isinstance(saved_target, str) and saved_target.strip():

                self.config["target_folder"] = saved_target

        except (OSError, json.JSONDecodeError):

            pass

        self.target_folder_var.set(
            self.config["target_folder"]
        )

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

    # ------------------------
    # ICON
    # ------------------------

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

    # ------------------------
    # OPEN FILE
    # ------------------------

    def open_file(self, _event):
        selection = self.listbox.curselection()
        if selection:
            display = self.listbox.get(selection[0]).split("  ", 1)[-1]
            path = Path(self.config["target_folder"]) / self.display_to_file.get(display, display)
            if path.exists():
                os.startfile(str(path))

    # ------------------------
    # TOPMOST
    # ------------------------

    def toggle_topmost(self, _event=None):
        self.is_topmost = not self.is_topmost
        self.root.attributes("-topmost", self.is_topmost)
        self.pin_label.config(text="📌 ON" if self.is_topmost else "📌 OFF",
                              fg="#60A5FA" if self.is_topmost else SUB_TEXT)
        self.config["topmost"] = self.is_topmost
        save_config(self.config)

    # ------------------------
    # CREATE SHORTCUT
    # ------------------------

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

    # ------------------------
    # OUTLOOK SCHEDULE
    # ------------------------

    def load_schedule(self):

        target_date = self.current_date
        next_date = target_date + timedelta(days=1)

        self.schedule_listbox.delete(0, tk.END)

        self.schedule_urls = {}

        ps_script = r'''
    $ErrorActionPreference = "SilentlyContinue"

    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9)

    $items = $calendar.Items

    # 定期予定展開に必須
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true

    $today = Get-Date "__TARGET_DATE__"
    $tomorrow = Get-Date "__NEXT_DATE__"

    # 今日の予定のみ
    $filter =
    "[Start] >= '" +
    $today.ToString("g") +
    "' AND [Start] < '" +
    $tomorrow.ToString("g") +
    "'"

    $restricted = $items.Restrict($filter)

    $item = $restricted.GetFirst()

    function Encode-Text($text)
    {
        if($null -eq $text)
        {
            $text = ""
        }

        return [Convert]::ToBase64String(
            [Text.Encoding]::UTF8.GetBytes(
                [string]$text
            )
        )
    }

    while($item -ne $null)
    {
        try
        {
            $subject = $item.Subject

            $startTime = $item.Start.ToString("HH:mm")
            $endTime   = $item.End.ToString("HH:mm")

            $body = ""

            try
            {
                $body = $item.Body
            }
            catch
            {
            }

            $url = ""

            if($body)
            {
                $matches = [regex]::Matches(
                    $body,
                    'https?://[^\s<>"'']+'
                )

                # Zoom URLを優先
                $url = ($matches | Where-Object { $_.Value -match 'zoom\.us' } | Select-Object -First 1).Value

                # ZoomがなければTeams URL
                if (-not $url)
                {
                    $url = ($matches | Where-Object { $_.Value -match 'teams\.microsoft\.com' -or $_.Value -match 'teams\.live\.com' } | Select-Object -First 1).Value
                }

                # ZoomとTeamsがなければ最初のURL
                if (-not $url -and $matches.Count -gt 0)
                {
                    $url = $matches[0].Value
                }
            }

            Write-Output (
                (Encode-Text $startTime) + "|" +
                (Encode-Text $endTime) + "|" +
                (Encode-Text $subject) + "|" +
                (Encode-Text $url)
            )
        }
        catch
        {
        }

        $item = $restricted.GetNext()
    }
    '''

        ps_script = ps_script.replace(
            "__TARGET_DATE__",
            target_date.strftime("%Y-%m-%d")
        )

        ps_script = ps_script.replace(
            "__NEXT_DATE__",
            next_date.strftime("%Y-%m-%d")
        )

        try:

            result = subprocess.check_output(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    ps_script
                ],
                text=True,
                encoding="utf-8",
                errors="ignore"
            )

            for line in result.splitlines():

                line = line.strip()

                if not line:
                    continue

                parts = line.split("|", 3)

                if len(parts) != 4:
                    continue

                try:

                    start_time, end_time, title, url = [
                        base64.b64decode(value).decode("utf-8")
                        for value in parts
                    ]

                except (ValueError, UnicodeDecodeError):

                    continue

                display_text = (
                    f"{start_time}～{end_time}  {title}"
                )

                index = self.schedule_listbox.size()

                self.schedule_listbox.insert(
                    tk.END,
                    display_text
                )

                if url:

                    self.schedule_urls[index] = url

                    self.schedule_listbox.itemconfig(
                        index,
                        fg="#60A5FA"
                    )

        except subprocess.CalledProcessError as e:

            print("PowerShell Error:")
            print(e)

        except Exception as e:

            print("Outlook Error:", e)

    def open_meeting(self, _event):
        selection = self.schedule_listbox.curselection()

        if selection and selection[0] in self.schedule_urls:
            webbrowser.open(self.schedule_urls[selection[0]])

    # ------------------------
    # ONENOTE TASK
    # ------------------------

    def show_onenote_page_dialog(self):

        dialog = tk.Toplevel(self.root)

        dialog.title("OneNoteページ追加")
        dialog.configure(bg=CARD_COLOR)
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        width = 520
        height = 230

        self.root.update_idletasks()

        x = self.root.winfo_x() + (self.root.winfo_width() - width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - height) // 2

        dialog.geometry(
            f"{width}x{height}+{x}+{y}"
        )

        tk.Label(
            dialog,
            text="OneNoteの表示名",
            bg=CARD_COLOR,
            fg=TEXT_COLOR,
            anchor="w"
        ).pack(fill="x", padx=18, pady=(18, 5))

        name_entry = self.make_entry(dialog)
        name_entry.pack(fill="x", padx=18, ipady=7)

        tk.Label(
            dialog,
            text="OneNoteページを右クリックして取得したリンク",
            bg=CARD_COLOR,
            fg=TEXT_COLOR,
            anchor="w"
        ).pack(fill="x", padx=18, pady=(14, 5))

        link_entry = self.make_entry(dialog)
        link_entry.pack(fill="x", padx=18, ipady=7)

        button_area = tk.Frame(
            dialog,
            bg=CARD_COLOR
        )

        button_area.pack(
            fill="x",
            padx=18,
            pady=(18, 0)
        )

        def register_page():

            page_name = name_entry.get().strip()
            page_link = link_entry.get().strip()

            if not page_name:

                messagebox.showwarning(
                    "OneNoteページ追加",
                    "OneNoteの表示名を入力してください。",
                    parent=dialog
                )

                name_entry.focus_set()
                return

            if not page_link:

                messagebox.showwarning(
                    "OneNoteページ追加",
                    "OneNoteページのリンクを入力してください。",
                    parent=dialog
                )

                link_entry.focus_set()
                return

            pages = self.config.setdefault(
                "onenote_pages",
                []
            )

            page_data = {
                "name": page_name,
                "link": page_link
            }

            pages.append(page_data)

            if not save_config(self.config):

                pages.pop()

                messagebox.showerror(
                    "OneNoteページ追加",
                    "config.jsonへ保存できませんでした。",
                    parent=dialog
                )

                return

            dialog.destroy()

            self.load_onenote_tasks()

        tk.Button(
            button_area,
            text="登録",
            command=register_page,
            bg=SELECT_BG,
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            borderwidth=0,
            width=10
        ).pack(side="right")

        tk.Button(
            button_area,
            text="キャンセル",
            command=dialog.destroy,
            bg=BUTTON_BG,
            fg="white",
            activebackground=BUTTON_ACTIVE,
            activeforeground="white",
            borderwidth=0,
            width=10
        ).pack(side="right", padx=(0, 8))

        dialog.bind(
            "<Return>",
            lambda _event: register_page()
        )

        dialog.bind(
            "<Escape>",
            lambda _event: dialog.destroy()
        )

        name_entry.focus_set()

    def load_onenote_tasks(self):
        for child in self.note_content.winfo_children():
            child.destroy()

        pages = self.config.get("onenote_pages", [])

        if not pages:
            self.note_message(
                "config.jsonのonenote_pagesにページを登録してください",
                SUB_TEXT,
            )
            return

        for page in pages:

            if isinstance(page, str):

                link = page
                page_name = "OneNote"

            else:

                link = page.get("link", "")
                page_name = page.get("name", "OneNote")

            title = tk.Label(self.note_content, text=f"▸ {page_name}", bg=LIST_BG,
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

        link = self.normalize_onenote_link(link)

        try:

            os.startfile(link)

        except OSError:

            webbrowser.open(link)

    def get_onenote_tasks(self, link, user_name):

        decoded_link = self.normalize_onenote_link(link)

        match = re.search(
            r"page-id\s*=?\s*((?:\{[^}]+\})+|[^&#\s]+)",
            decoded_link,
            re.I
        )

        if not match:

            raise ValueError("リンクにpage-idがありません")

        page_id = match.group(1)

        section_match = re.match(
            r"^onenote:/*(.*?\.one)(?:#|$)",
            decoded_link,
            re.I
        )

        section_path = ""

        if section_match:

            section_path = section_match.group(1)

            if section_path.startswith("\\"):

                section_path = "\\\\" + section_path.lstrip("\\")

        page_name_match = re.search(
            r"\.one#([^&]+)",
            decoded_link,
            re.I
        )

        page_name = ""

        if page_name_match:

            page_name = page_name_match.group(1).strip()

        page_id_data = base64.b64encode(
            page_id.encode("utf-8")
        ).decode("ascii")

        section_path_data = base64.b64encode(
            section_path.encode("utf-8")
        ).decode("ascii")

        page_name_data = base64.b64encode(
            page_name.encode("utf-8")
        ).decode("ascii")

        script = (
            "$ErrorActionPreference='Stop'; "
            "[Console]::OutputEncoding=[Text.Encoding]::UTF8; "
            f"$pageId=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{page_id_data}')); "
            f"$sectionPath=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{section_path_data}')); "
            f"$pageName=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{page_name_data}')); "
            "$oneNote=New-Object -ComObject OneNote.Application; "
            "$sectionId=''; "
            "if($sectionPath){ "
            "$oneNote.OpenHierarchy($sectionPath,'',[ref]$sectionId,0); "
            "}; "
            "$hierarchy=''; "
            "$oneNote.GetHierarchy($sectionId,4,[ref]$hierarchy,2); "
            "[xml]$hierarchyXml=$hierarchy; "
            "$pages=@($hierarchyXml.SelectNodes(\"//*[local-name()='Page']\")); "
            "$pageNode=$pages | Where-Object { $_.ID -eq $pageId } | Select-Object -First 1; "
            "if(-not $pageNode -and $pageName){ "
            "$pageNode=$pages | Where-Object { $_.name -eq $pageName } | Select-Object -First 1; "
            "}; "
            "if(-not $pageNode){ "
            "throw (\"OneNoteページが見つかりません: \" + $pageName); "
            "}; "
            "$actualPageId=$pageNode.ID; "
            "$xml=''; "
            "$oneNote.GetPageContent($actualPageId,[ref]$xml,0,2); "
            "[Console]::Write($xml)"
        )

        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                script,
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
        )

        if result.returncode != 0:

            error_text = result.stderr.strip()

            if not error_text:

                error_text = result.stdout.strip()

            if not error_text:

                error_text = "OneNoteからページを取得できませんでした"

            raise RuntimeError(error_text)

        xml_text = result.stdout

        if not xml_text.strip():

            raise RuntimeError("OneNoteページの内容が空です")

        return self.parse_onenote_tables(xml_text, user_name)

    @staticmethod
    def normalize_onenote_link(link):

        link = html.unescape(unquote(link.strip()))

        return (
            link.replace("￥", "\\")
                .replace("＃", "#")
        )

    @staticmethod
    def parse_onenote_tables(xml_text, user_name):
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

                if (
                    not user_name.strip()
                    or user_name.strip().casefold() in names
                ):
                    tasks.append(
                        (
                            row[task_column] or "（タスク名なし）",
                            row[due_date] or "期限なし",
                        )
                    )

        return tasks

    # ------------------------
    # SIDEBAR DISPLAY
    # ------------------------

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

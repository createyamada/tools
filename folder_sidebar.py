import os
import win32com.client
import tkinter as tk
from pathlib import Path
import subprocess
import sys
import os
import json
from datetime import datetime
import re
import webbrowser
import ctypes
from ctypes import wintypes
import threading

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
    "width": 320,
    "alpha": 0.96,
    "topmost": False
}

def load_config():

    # ------------------------
    # config.json が無ければ生成
    # ------------------------

    if not os.path.exists(CONFIG_FILE):

        with open(
            CONFIG_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                DEFAULT_CONFIG,
                f,
                indent=4,
                ensure_ascii=False
            )

        return DEFAULT_CONFIG

    # ------------------------
    # config.json 読み込み
    # ------------------------

    try:

        with open(
            CONFIG_FILE,
            "r",
            encoding="utf-8"
        ) as f:

            config = json.load(f)

            return {
                **DEFAULT_CONFIG,
                **config
            }

    except Exception as e:

        print("Config Load Error:", e)

        return DEFAULT_CONFIG

CONFIG = load_config()

def hotkey_listener(app):

    user32 = ctypes.windll.user32

    MOD_ALT = 0x0001
    MOD_CONTROL = 0x0002

    VK_SPACE = 0x20

    HOTKEY_ID = 1

    user32.RegisterHotKey(
        None,
        HOTKEY_ID,
        MOD_CONTROL | MOD_ALT,
        VK_SPACE
    )

    msg = wintypes.MSG()

    while user32.GetMessageW(
        ctypes.byref(msg),
        None,
        0,
        0
    ) != 0:

        if msg.message == 0x0312:

            app.root.after(
                0,
                app.toggle_sidebar
            )

TARGET_FOLDER = CONFIG["target_folder"]

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

class FolderSidebar:

    def __init__(self, root):

        self.root = root

        self.root.overrideredirect(True)
        self.root.attributes("-topmost", False)
        self.root.attributes("-alpha", 0.97)

        width = 340

        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        TASKBAR_MARGIN = 60

        x = 0

        self.root.geometry(
            f"{width}x{screen_height - TASKBAR_MARGIN}+{x}+0"
        )

        self.root.configure(bg=BG_COLOR)

        self.is_topmost = False

        # ------------------------
        # HEADER
        # ------------------------

        header = tk.Frame(
            root,
            bg=BG_COLOR
        )

        header.pack(
            fill="x",
            pady=(18, 8),
            padx=18
        )

        title = tk.Label(
            header,
            text="📁 WORK FILES",
            bg=BG_COLOR,
            fg="white",
            font=("Yu Gothic UI", 18, "bold")
        )

        title.pack(anchor="w")

        subtitle = tk.Label(
            header,
            text="Workspace Sidebar",
            bg=BG_COLOR,
            fg=SUB_TEXT,
            font=("Yu Gothic UI", 9)
        )

        subtitle.pack(anchor="w")

        # ------------------------
        # MAIN CARD
        # ------------------------

        card = tk.Frame(
            root,
            bg=CARD_COLOR
        )

        card.pack(
            fill="both",
            expand=True,
            padx=14,
            pady=(0, 12)
        )

        # ------------------------
        # LISTBOX
        # ------------------------

        self.listbox = tk.Listbox(
            card,
            bg=LIST_BG,
            fg=TEXT_COLOR,
            selectbackground=SELECT_BG,
            selectforeground="white",
            activestyle="none",
            borderwidth=0,
            highlightthickness=0,
            font=("Yu Gothic UI", 11),
            relief="flat"
        )

        self.listbox.pack(
            fill="both",
            expand=True,
            padx=12,
            pady=12
        )
        
        # ------------------------
        # OUTLOOK TITLE
        # ------------------------
        self.schedule_urls = {}
        schedule_title = tk.Label(
            root,
            text="📅 TODAY SCHEDULE",
            bg=BG_COLOR,
            fg="white",
            font=("Yu Gothic UI", 11, "bold")
        )

        schedule_title.pack(
            fill="x",
            padx=10,
            pady=(0, 5)
        )

        # ------------------------
        # SCHEDULE LIST
        # ------------------------

        self.schedule_listbox = tk.Listbox(
            root,
            bg=LIST_BG,
            fg=TEXT_COLOR,
            selectbackground=SELECT_BG,
            activestyle="none",
            borderwidth=0,
            highlightthickness=0,
            font=("Yu Gothic UI", 10),
            height=8
        )

        self.schedule_listbox.pack(
            fill="x",
            padx=10,
            pady=(0, 10)
        )

        self.schedule_listbox.bind(
            "<Double-Button-1>",
            self.open_meeting
        )

        # ------------------------
        # FOOTER
        # ------------------------

        footer = tk.Frame(
            card,
            bg=CARD_COLOR
        )

        footer.pack(
            fill="x",
            padx=12,
            pady=(0, 12)
        )

        # ------------------------
        # 登録エリア
        # ------------------------

        register_frame = tk.Frame(
            root,
            bg=BG_COLOR
        )

        register_frame.pack(
            fill="x",
            padx=10,
            pady=(0, 10)
        )

        self.path_entry = tk.Entry(
            register_frame,
            bg="#374151",
            fg="white",
            insertbackground="white",
            borderwidth=0,
            font=("Yu Gothic UI", 10)
        )

        self.path_entry.pack(
            side="left",
            fill="x",
            expand=True,
            ipady=8,
            padx=(0, 6)
        )

        register_button = tk.Button(
            register_frame,
            text="＋",
            command=self.create_shortcut,
            bg="#2563EB",
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            borderwidth=0,
            font=("Yu Gothic UI", 11, "bold"),
            width=4,
            cursor="hand2"
        )

        register_button.pack(
            side="right"
        )

        refresh_button = tk.Button(
            footer,
            text="↻ Refresh",
            command=self.refresh,
            bg=BUTTON_BG,
            fg="white",
            activebackground=BUTTON_ACTIVE,
            activeforeground="white",
            borderwidth=0,
            relief="flat",
            font=("Yu Gothic UI", 9, "bold"),
            padx=12,
            pady=8,
            cursor="hand2"
        )

        refresh_button.pack(
            side="left"
        )

        self.pin_label = tk.Label(
            footer,
            text="📌 OFF",
            bg=CARD_COLOR,
            fg=SUB_TEXT,
            font=("Yu Gothic UI", 9)
        )

        self.pin_label.pack(
            side="right"
        )

        # ------------------------
        # EVENT
        # ------------------------

        self.listbox.bind(
            "<Double-Button-1>",
            self.open_file
        )

        self.root.bind(
            "<Escape>",
            lambda e: root.destroy()
        )

        self.root.bind(
            "<Control-Shift-space>",
            self.toggle_topmost
        )

        self.refresh()

    # ------------------------
    # REFRESH
    # ------------------------

    def refresh(self):

        self.listbox.delete(0, tk.END)

        folder = Path(TARGET_FOLDER)

        if folder.exists():

            for f in sorted(folder.iterdir()):

                if f.is_file():

                    icon = self.get_icon(f)

                    self.listbox.insert(
                        tk.END,
                        f"{icon}  {f.name}"
                    )
        self.load_schedule()

    # ------------------------
    # ICON
    # ------------------------

    def get_icon(self, file):

        ext = file.suffix.lower()

        if ext in [".png", ".jpg", ".jpeg"]:
            return "🖼"

        if ext in [".txt"]:
            return "📄"

        if ext in [".xlsx", ".xls"]:
            return "📊"

        if ext in [".pdf"]:
            return "📕"

        if ext in [".py"]:
            return "🐍"

        return "📁"

    # ------------------------
    # OPEN FILE
    # ------------------------

    def open_file(self, event):

        selection = self.listbox.curselection()

        if not selection:
            return

        text = self.listbox.get(selection[0])

        filename = text.split("  ", 1)[1]

        filepath = Path(TARGET_FOLDER) / filename

        print(filepath)

        try:

            os.startfile(str(filepath))

        except Exception as e:

            print("OPEN ERROR:", e)

    # ------------------------
    # TOPMOST
    # ------------------------

    def toggle_topmost(self, event=None):

        self.is_topmost = not self.is_topmost

        self.root.attributes(
            "-topmost",
            self.is_topmost
        )

        if self.is_topmost:

            self.pin_label.config(
                text="📌 ON",
                fg="#60A5FA"
            )

        else:

            self.pin_label.config(
                text="📌 OFF",
                fg=SUB_TEXT
            )


    def create_shortcut(self):

        target_path = self.path_entry.get().strip()

        if not target_path:
            return

        if not os.path.exists(target_path):
            print("Path not found")
            return

        target = Path(target_path)

        shortcut_name = target.stem + ".lnk"

        shortcut_path = Path(TARGET_FOLDER) / shortcut_name

        shell = win32com.client.Dispatch("WScript.Shell")

        shortcut = shell.CreateShortCut(
            str(shortcut_path)
        )

        shortcut.Targetpath = str(target)

        shortcut.WorkingDirectory = str(
            target.parent
        )

        shortcut.save()

        self.path_entry.delete(0, tk.END)

        self.refresh()


    def load_schedule(self):

        self.schedule_listbox.delete(0, tk.END)

        self.schedule_urls = {}

        try:

            outlook = win32com.client.Dispatch(
                "Outlook.Application"
            )

            namespace = outlook.GetNamespace("MAPI")

            calendar = namespace.GetDefaultFolder(9)

            items = calendar.Items

            items.IncludeRecurrences = True

            items.Sort("[Start]")

            today = datetime.now().date()

            for item in items:

                try:

                    start = item.Start
                    end = item.End

                    if start.date() != today:
                        continue

                    title = item.Subject

                    body = item.Body or ""

                    url = None

                    zoom_match = re.search(
                        r"https://[\w\.-]*zoom\.us/j/\S+",
                        body
                    )

                    if zoom_match:
                        url = zoom_match.group(0)

                    # Teams URL
                    teams_match = re.search(
                        r"https://teams\S+",
                        body
                    )

                    if teams_match:
                        url = teams_match.group(0)

                    # Zoom URL
                    if not url:

                        # Teams URL
                        teams_match = re.search(
                            r"https://teams\S+",
                            body
                        )

                        if teams_match:
                            url = teams_match.group(0)

                    time_text = (
                        f"{start.strftime('%H:%M')}"
                        f"～"
                        f"{end.strftime('%H:%M')}"
                    )

                    display_text = (
                        f"{time_text}  {title}"
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

                except Exception as e:

                    print(e)

        except Exception as e:

            print("Outlook Error:", e)


    def show_window(self):

        self.root.deiconify()

        self.root.attributes(
            "-topmost",
            True
        )

        self.root.lift()

        self.root.focus_force()

        self.is_topmost = True

    def toggle_sidebar(self):

        if self.root.state() == "withdrawn":

            self.root.deiconify()

            self.root.attributes(
                "-topmost",
                True
            )

            self.root.lift()

        else:

            self.root.withdraw()

    def open_meeting(self, event):

        selection = self.schedule_listbox.curselection()

        if not selection:
            return

        index = selection[0]

        url = self.schedule_urls.get(index)
        print(url)
        if not url:
            return

        webbrowser.open(url)
# ------------------------
# START
# ------------------------

root = tk.Tk()

app = FolderSidebar(root)

threading.Thread(
    target=hotkey_listener,
    args=(app,),
    daemon=True
).start()

root.mainloop()

import tkinter as tk
from pathlib import Path
import subprocess

TARGET_FOLDER = r"C:\Work"

BG_COLOR = "#111827"
LIST_BG = "#1F2937"
TEXT_COLOR = "#F9FAFB"
SELECT_BG = "#2563EB"

class FolderSidebar:

    def __init__(self, root):

        self.root = root

        # ------------------------
        # ウィンドウ設定
        # ------------------------

        self.root.overrideredirect(True)  # ← ×消す
        self.root.attributes("-topmost", False)

        # 半透明
        self.root.attributes("-alpha", 0.96)

        width = 320

        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        TASKBAR_MARGIN = 50

        x = 0

        self.root.geometry(
            f"{width}x{screen_height - TASKBAR_MARGIN}+{x}+0"
        )

        self.root.configure(bg=BG_COLOR)

        # ------------------------
        # タイトル
        # ------------------------

        title = tk.Label(
            root,
            text="📁 WORK FILES",
            bg=BG_COLOR,
            fg="white",
            font=("Yu Gothic UI", 16, "bold"),
            pady=15
        )

        title.pack(fill="x")

        # ------------------------
        # リスト
        # ------------------------

        self.listbox = tk.Listbox(
            root,
            bg=LIST_BG,
            fg=TEXT_COLOR,
            selectbackground=SELECT_BG,
            activestyle="none",
            borderwidth=0,
            highlightthickness=0,
            font=("Yu Gothic UI", 11),
        )

        self.listbox.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=(0, 10)
        )


        refresh_button = tk.Button(
            root,
            text="↻ Refresh",
            command=self.refresh,
            bg="#374151",
            fg="white",
            activebackground="#4B5563",
            activeforeground="white",
            borderwidth=0,
            font=("Yu Gothic UI", 9),
            pady=6,
            cursor="hand2"
        )

        refresh_button.pack(
            fill="x",
            padx=10,
            pady=(0, 10)
        )

        self.listbox.bind(
            "<Double-Button-1>",
            self.open_file
        )

        # ESCで終了
        self.root.bind(
            "<Escape>",
            lambda e: root.destroy()
        )
        self.is_topmost = False

        self.root.bind(
            "<Control-Shift-space>",
            self.toggle_topmost
        )


        self.refresh()

    # ------------------------
    # ファイル一覧更新
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


    # ------------------------
    # アイコン
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
    # ファイル開く
    # ------------------------

    def open_file(self, event):

        selection = self.listbox.curselection()

        if not selection:
            return

        text = self.listbox.get(selection[0])

        filename = text.split("  ", 1)[1]

        filepath = Path(TARGET_FOLDER) / filename

        try:

            subprocess.Popen(
                str(filepath),
                shell=True
            )

        except Exception as e:

            print(e)


    def toggle_topmost(self, event=None):

        self.is_topmost = not self.is_topmost

        self.root.attributes(
            "-topmost",
            self.is_topmost
        )

        print("Topmost:", self.is_topmost)

# ------------------------
# 起動
# ------------------------

root = tk.Tk()

app = FolderSidebar(root)

root.mainloop()
"""従来版 Outlook の予定から、複数人に共通する空き時間を表示する。"""

from __future__ import annotations

import json
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk


SLOT_MINUTES = 30
MAX_DAYS = 31
HELPER_FILE = Path(__file__).with_name("outlook_freebusy.ps1")
EMAIL_PATTERN = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")


class AppError(Exception):
    """画面に表示できるエラー。"""


class SearchCancelled(AppError):
    pass


def validate_search(emails: list[str], start_text: str, end_text: str,
                    from_text: str, to_text: str) -> tuple[list[str], date, date, int, int]:
    unique = []
    seen = set()
    for raw in emails:
        email = raw.strip()
        if not email:
            continue
        if not EMAIL_PATTERN.fullmatch(email):
            raise AppError(f"メールアドレスの形式を確認してください: {email}")
        if email.casefold() not in seen:
            unique.append(email)
            seen.add(email.casefold())
    if not unique:
        raise AppError("メールアドレスを1件以上入力してください。")
    try:
        start_day = date.fromisoformat(start_text.strip())
        end_day = date.fromisoformat(end_text.strip())
    except ValueError as exc:
        raise AppError("日付は YYYY-MM-DD 形式で入力してください。") from exc
    if end_day < start_day:
        raise AppError("終了日は開始日以降にしてください。")
    if (end_day - start_day).days + 1 > MAX_DAYS:
        raise AppError(f"検索期間は最大{MAX_DAYS}日です。")

    def minutes(value: str, allow_midnight_end: bool = False) -> int:
        if not re.fullmatch(r"\d{2}:\d{2}", value.strip()):
            raise AppError("時間は HH:MM 形式で入力してください。")
        hour, minute = map(int, value.strip().split(":"))
        if (hour > 23 and not (allow_midnight_end and hour == 24 and minute == 0)) or minute > 59 or minute % SLOT_MINUTES:
            raise AppError("時間は30分刻みで入力してください。終了時刻は24:00も指定できます。")
        return hour * 60 + minute

    from_minute = minutes(from_text)
    to_minute = minutes(to_text, allow_midnight_end=True)
    if from_minute >= to_minute:
        raise AppError("終了時刻は開始時刻より後にしてください。")
    return unique, start_day, end_day, from_minute, to_minute


def common_free_periods(views: list[str], day: date, from_minute: int,
                        to_minute: int, now: datetime | None = None) -> list[tuple[datetime, datetime]]:
    """全員が '0' の30分枠を、連続する時間帯にまとめる。"""
    count = (to_minute - from_minute) // SLOT_MINUTES
    if not views or any(len(view) != count or any(code not in "01234" for code in view)
                        for view in views):
        raise AppError("予定の応答が不完全です。空き時間を確定できません。")
    current = now if now is not None else datetime.now()
    if current.tzinfo is not None:
        current = current.astimezone().replace(tzinfo=None)
    base = datetime(day.year, day.month, day.day) + timedelta(minutes=from_minute)
    periods = []
    open_start = None
    for index in range(count):
        slot_start = base + timedelta(minutes=index * SLOT_MINUTES)
        free = slot_start >= current and all(view[index] == "0" for view in views)
        if free and open_start is None:
            open_start = slot_start
        if not free and open_start is not None:
            periods.append((open_start, slot_start))
            open_start = None
    if open_start is not None:
        periods.append((open_start, base + timedelta(minutes=count * SLOT_MINUTES)))
    return periods


def get_views_from_outlook(emails: list[str], first: date, last: date,
                           cancel: threading.Event) -> list[str]:
    """Windows PowerShell を介して従来版 Outlook の COM 機能を利用する。"""
    if sys.platform != "win32":
        raise AppError("この方式は Windows の従来版 Outlook でのみ利用できます。")
    powershell = shutil.which("powershell.exe")
    if powershell is None or not HELPER_FILE.exists():
        raise AppError("Windows PowerShell または outlook_freebusy.ps1 が見つかりません。")
    command = [powershell, "-NoProfile", "-NonInteractive", "-Sta",
               "-ExecutionPolicy", "Bypass", "-File", str(HELPER_FILE)]
    payload = json.dumps({"emails": emails, "start_date": first.isoformat(),
                          "end_date": last.isoformat()}, ensure_ascii=False)
    try:
        process = subprocess.Popen(
            command, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, encoding="utf-8", errors="replace",
            creationflags=subprocess.CREATE_NO_WINDOW)
    except OSError as exc:
        raise AppError(f"Outlook 取得処理を開始できません: {exc}") from exc
    deadline = time.monotonic() + 180
    try:
        try:
            output, errors = process.communicate(input=payload, timeout=0.5)
        except subprocess.TimeoutExpired:
            while True:
                if cancel.is_set():
                    process.kill()
                    process.communicate()
                    raise SearchCancelled("検索を中止しました。")
                if time.monotonic() >= deadline:
                    process.kill()
                    process.communicate()
                    raise AppError("Outlook からの応答が3分以内に得られませんでした。")
                try:
                    output, errors = process.communicate(timeout=0.5)
                    break
                except subprocess.TimeoutExpired:
                    continue
    finally:
        if process.poll() is None:
            process.kill()
            process.communicate()
    if cancel.is_set():
        raise SearchCancelled("検索を中止しました。")
    try:
        response = json.loads(output.strip().lstrip("\ufeff"))
    except (ValueError, AttributeError) as exc:
        detail = (errors or output or "Outlook から正しい応答がありませんでした。").strip()
        raise AppError(detail[:500]) from exc
    if not isinstance(response, dict) or response.get("ok") is not True:
        detail = response.get("error", "予定を取得できませんでした。") if isinstance(response, dict) else "予定を取得できませんでした。"
        raise AppError(str(detail))
    values = response.get("values")
    if not isinstance(values, list):
        raise AppError("Outlook から全員分の予定を取得できませんでした。")
    by_email = {}
    expected_length = ((last - first).days + 1) * 48
    self_view = response.get("self_view")
    if (not isinstance(self_view, str) or len(self_view) != expected_length
            or any(code not in "01234" for code in self_view)):
        raise AppError("Outlook から自分の予定を取得できませんでした。")
    for item in values:
        if not isinstance(item, dict):
            raise AppError("Outlook の応答が不完全です。")
        email = str(item.get("email", "")).casefold()
        view = item.get("view")
        if email in by_email or not isinstance(view, str) or len(view) != expected_length:
            raise AppError("Outlook の空き時間データが不完全です。")
        if any(code not in "01234" for code in view):
            raise AppError("Outlook の空き時間データに不明な値があります。")
        by_email[email] = view
    if set(by_email) != {email.casefold() for email in emails}:
        raise AppError("Outlook から全員分の予定を取得できませんでした。")
    return [self_view] + [by_email[email.casefold()] for email in emails]


def views_for_day(full_views: list[str], first: date, day: date,
                  from_minute: int, to_minute: int) -> list[str]:
    offset = (day - first).days * 48
    start = offset + from_minute // SLOT_MINUTES
    end = offset + to_minute // SLOT_MINUTES
    return [view[start:end] for view in full_views]


class FreeTimeApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Outlook 共通空き時間検索")
        self.root.minsize(590, 690)
        self.root.geometry("690x760")
        self.events: queue.Queue = queue.Queue()
        self.cancel = threading.Event()
        self.build_ui()
        self.root.after(100, self.drain_events)

    def build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Outlook 共通空き時間検索", font=("Yu Gothic UI", 15, "bold")).pack(anchor="w")
        ttk.Label(frame, text="自分と共通の空き時間を調べる人のメールアドレスを入力してください（最大10人）。").pack(anchor="w", pady=(4, 8))

        grid = ttk.Frame(frame)
        grid.pack(fill="x")
        self.email_entries = []
        for index in range(10):
            ttk.Label(grid, text=f"{index + 1:2d}.").grid(row=index, column=0, sticky="e", padx=(0, 8), pady=2)
            entry = ttk.Entry(grid)
            entry.grid(row=index, column=1, sticky="ew", pady=2)
            self.email_entries.append(entry)
        grid.columnconfigure(1, weight=1)

        period = ttk.LabelFrame(frame, text="検索条件（Windows / Outlook の表示時刻）", padding=10)
        period.pack(fill="x", pady=(14, 8))
        today = date.today()
        self.start_date = tk.StringVar(value=today.isoformat())
        self.end_date = tk.StringVar(value=(today + timedelta(days=6)).isoformat())
        self.from_time = tk.StringVar(value="09:00")
        self.to_time = tk.StringVar(value="18:00")
        ttk.Label(period, text="開始日").grid(row=0, column=0, sticky="w")
        ttk.Entry(period, textvariable=self.start_date, width=13).grid(row=0, column=1, padx=(5, 12))
        ttk.Label(period, text="終了日").grid(row=0, column=2, sticky="w")
        ttk.Entry(period, textvariable=self.end_date, width=13).grid(row=0, column=3, padx=5)
        ttk.Label(period, text="時間").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(period, textvariable=self.from_time, width=13).grid(row=1, column=1, padx=(5, 12), pady=(8, 0))
        ttk.Label(period, text="～").grid(row=1, column=2, sticky="w", pady=(8, 0))
        ttk.Entry(period, textvariable=self.to_time, width=13).grid(row=1, column=3, padx=5, pady=(8, 0))
        ttk.Label(period, text="日付: YYYY-MM-DD　時間: 30分刻み").grid(row=2, column=0, columnspan=4, sticky="w", pady=(8, 0))

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(4, 10))
        self.cancel_button = ttk.Button(buttons, text="中止", command=self.cancel.set, state="disabled")
        self.cancel_button.pack(side="right")
        self.search_button = ttk.Button(buttons, text="検索", command=self.search)
        self.search_button.pack(side="right", padx=(0, 8))

        self.status = tk.StringVar(value="メールアドレスと期間を入力して検索してください。")
        ttk.Label(frame, textvariable=self.status, wraplength=630).pack(anchor="w", pady=(0, 5))
        self.result = scrolledtext.ScrolledText(frame, height=12, wrap="word", state="disabled", font=("Yu Gothic UI", 10))
        self.result.pack(fill="both", expand=True)

    def set_result(self, text: str) -> None:
        self.result.configure(state="normal")
        self.result.delete("1.0", "end")
        self.result.insert("1.0", text)
        self.result.configure(state="disabled")

    def search(self) -> None:
        try:
            emails, first, last, from_minute, to_minute = validate_search(
                [entry.get() for entry in self.email_entries], self.start_date.get(),
                self.end_date.get(), self.from_time.get(), self.to_time.get())
        except AppError as exc:
            messagebox.showerror("入力エラー", str(exc), parent=self.root)
            return
        self.cancel.clear()
        self.search_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.set_result("")
        self.status.set("検索を開始しています…")
        threading.Thread(target=self.run_search,
                         args=(emails, first, last, from_minute, to_minute),
                         daemon=True).start()

    def run_search(self, emails: list[str], first: date, last: date,
                   from_minute: int, to_minute: int) -> None:
        try:
            self.events.put(("status", "従来版 Outlook から予定を取得中…"))
            full_views = get_views_from_outlook(emails, first, last, self.cancel)
            periods = []
            day = first
            while day <= last:
                views = views_for_day(full_views, first, day, from_minute, to_minute)
                periods.extend(common_free_periods(views, day, from_minute, to_minute))
                day += timedelta(days=1)
            if self.cancel.is_set():
                raise SearchCancelled("検索を中止しました。")
            lines = [f"対象: 自分 + 入力した{len(emails)}人 / {first} ～ {last} / Windows・Outlook の表示時刻", ""]
            if periods:
                weekdays = ("月", "火", "水", "木", "金", "土", "日")
                for start, end in periods:
                    end_label = end.strftime("%Y-%m-%d %H:%M") if end.date() != start.date() else end.strftime("%H:%M")
                    lines.append(f"{start:%Y-%m-%d} ({weekdays[start.weekday()]}) {start:%H:%M} ～ {end_label}")
                lines += ["", "※ 30分単位で全員が『空き』の枠を表示しています。仮予定も空き扱いしません。"]
            else:
                lines.append("指定期間に全員共通の空き時間はありません。")
            self.events.put(("done", "\n".join(lines)))
        except SearchCancelled as exc:
            self.events.put(("cancelled", str(exc)))
        except AppError as exc:
            self.events.put(("error", str(exc)))
        except Exception as exc:
            self.events.put(("error", f"予期しないエラー: {exc}"))

    def drain_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "status":
                    self.status.set(event[1])
                elif kind in ("done", "error", "cancelled"):
                    self.search_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    if kind == "done":
                        self.set_result(event[1])
                        self.status.set("検索が完了しました。")
                    else:
                        self.status.set(event[1])
                        if kind == "error":
                            messagebox.showerror("検索エラー", event[1], parent=self.root)
        except queue.Empty:
            pass
        self.root.after(100, self.drain_events)


def main() -> None:
    root = tk.Tk()
    FreeTimeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

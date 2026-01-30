import csv
import math
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.font as tkfont
import io
import webbrowser
import sqlite3
import json
from datetime import date, datetime
import random
import atexit
import msvcrt
import uuid
import zipfile
import copy
import hashlib
import tempfile
import urllib.request
from urllib.error import URLError, HTTPError
import threading
import subprocess
import time
from PIL import Image, ImageTk

DEFAULT_WORDS = [
    "abundant",
    "curious",
    "diligent",
    "elegant",
    "frugal",
    "graceful",
    "humble",
    "insight",
    "jubilant",
    "kindle",
    "lucid",
    "meticulous",
    "novel",
    "optimistic",
    "prudent",
    "resilient",
    "serene",
    "tenacious",
    "vivid",
    "wise",
]

INTERVAL_SECONDS = 15
NOTIFY_DURATION_SECONDS = 15
def resource_path(relative_path: str) -> str:
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)


APP_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, "frozen", False):
    DATA_DIR = os.path.dirname(sys.executable)
else:
    DATA_DIR = APP_DIR
WORDS_FILE = os.path.join(DATA_DIR, "words.csv")
LEGACY_WORDS_FILE = os.path.join(DATA_DIR, "words.txt")

APPDATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA", DATA_DIR), "VocabBuddy")
os.makedirs(APPDATA_DIR, exist_ok=True)
DB_PATH = os.path.join(APPDATA_DIR, "vocabbuddy.db")
BACKUP_DIR = os.path.join(APPDATA_DIR, "backups")
os.makedirs(BACKUP_DIR, exist_ok=True)
NOTIFY_MIN_WIDTH = 280
NOTIFY_MIN_HEIGHT = 200
NOTIFY_MAX_WIDTH = 360
NOTIFY_MAX_HEIGHT = 320
DONATE_ICON_PATH = resource_path("Donate-Icon-PNG-Clipart-Background.png")
DONATE_URL = "https://www.paypal.com/ncp/payment/ATH8S659Y8T6C"
HEADER_LOGO_PATH = resource_path("logo hd.png")
APP_ICON_PATH = resource_path("application.png")
BODY_LOGO_PATH = resource_path("logo hd.png")
ICON_DIR = resource_path("icon")
LOCK_FILE_PATH = os.path.join(APPDATA_DIR, "VocabBuddy.lock")
# App version for update checks.
CURRENT_VERSION = "0.8"
# GitHub-hosted update metadata (latest.json).
UPDATE_ENDPOINT = "https://raw.githubusercontent.com/AEgunz/VocabBuddy/main/latest.json"
# GitHub Releases API for the latest release (preferred source).
GITHUB_RELEASES_API = "https://api.github.com/repos/AEgunz/VocabBuddy/releases/latest"

THEME = {
    "bg": "#F6F7FB",
    "card": "#FFFFFF",
    "text": "#1F2937",
    "muted": "#6B7280",
    "accent": "#2563EB",
    "accent_hover": "#1D4ED8",
    "danger": "#EF4444",
    "danger_hover": "#DC2626",
    "border": "#E5E7EB",
    "shadow": "#000000",
}

_lock_file_handle = None


def ensure_single_instance() -> bool:
    global _lock_file_handle
    try:
        lock_file = open(LOCK_FILE_PATH, "a+")
        msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
        lock_file.seek(0)
        lock_file.truncate(0)
        lock_file.write(str(os.getpid()))
        lock_file.flush()
        _lock_file_handle = lock_file
        return True
    except OSError:
        try:
            temp_root = tk.Tk()
            temp_root.withdraw()
            messagebox.showerror("Already running", "The application is already running.")
            temp_root.destroy()
        except Exception:
            pass
        try:
            if "lock_file" in locals():
                lock_file.close()
        except Exception:
            pass
        return False


def release_single_instance() -> None:
    global _lock_file_handle
    if _lock_file_handle is not None:
        try:
            _lock_file_handle.seek(0)
            msvcrt.locking(_lock_file_handle.fileno(), msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
        try:
            _lock_file_handle.close()
        except OSError:
            pass
        _lock_file_handle = None
        try:
            if os.path.exists(LOCK_FILE_PATH):
                os.remove(LOCK_FILE_PATH)
        except OSError:
            pass

class RoundedButton(tk.Canvas):
    def __init__(
        self,
        parent: tk.Widget,
        text: str = "",
        command=None,
        bg: str = "#E5E5E5",
        fg: str = "#111827",
        hover: str = "#DCDCDC",
        radius: int = 10,
        font=("Segoe UI", 10),
        padx: int = 12,
        pady: int = 8,
        image: tk.PhotoImage | None = None,
        compound: str = "left",
    ) -> None:
        super().__init__(parent, highlightthickness=0, bd=0, bg=parent.cget("bg"))
        self._text = text
        self._command = command
        self._bg = bg
        self._hover = hover
        self._fg = fg
        self._radius = radius
        self._padx = padx
        self._pady = pady
        self._font = tkfont.Font(font=font)
        self._image = image
        self._compound = compound
        self._current_bg = bg

        height = self._font.metrics("linespace") + (self._pady * 2)
        self.configure(height=height)
        self.bind("<Configure>", lambda _e: self._redraw())
        self.bind("<Enter>", lambda _e: self._set_hover(True))
        self.bind("<Leave>", lambda _e: self._set_hover(False))
        self.bind("<Button-1>", self._on_click)
        self.configure(cursor="hand2")

    def _set_hover(self, hovering: bool) -> None:
        self._current_bg = self._hover if hovering else self._bg
        self._redraw()

    def _on_click(self, _event=None) -> None:
        if callable(self._command):
            self._command()

    def _rounded_rect(self, x1: int, y1: int, x2: int, y2: int, r: int, fill: str) -> None:
        self.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, fill=fill, outline=fill)
        self.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, fill=fill, outline=fill)
        self.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, fill=fill, outline=fill)
        self.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, fill=fill, outline=fill)
        self.create_rectangle(x1 + r, y1, x2 - r, y2, fill=fill, outline=fill)
        self.create_rectangle(x1, y1 + r, x2, y2 - r, fill=fill, outline=fill)

    def _redraw(self) -> None:
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= 1 or h <= 1:
            return
        r = min(self._radius, h // 2)
        self._rounded_rect(1, 1, w - 1, h - 1, r, self._current_bg)

        text_w = self._font.measure(self._text) if self._text else 0
        img_w = self._image.width() if self._image is not None else 0
        gap = 8 if self._image is not None and self._text else 0
        total = img_w + gap + text_w
        start_x = max(self._padx, (w - total) / 2)
        center_y = h / 2

        if self._image is not None:
            self.create_image(start_x + (img_w / 2), center_y, image=self._image)
            start_x += img_w + gap
        if self._text:
            self.create_text(start_x + (text_w / 2), center_y, text=self._text, font=self._font, fill=self._fg)

    def configure(self, cnf=None, **kwargs):
        if cnf:
            kwargs.update(cnf)
        pady_changed = False
        if "text" in kwargs:
            self._text = kwargs.pop("text")
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if "bg" in kwargs:
            self._bg = kwargs.pop("bg")
            self._current_bg = self._bg
        if "fg" in kwargs:
            self._fg = kwargs.pop("fg")
        if "activebackground" in kwargs:
            self._hover = kwargs.pop("activebackground")
        if "activeforeground" in kwargs:
            kwargs.pop("activeforeground")
        if "image" in kwargs:
            self._image = kwargs.pop("image")
        if "compound" in kwargs:
            self._compound = kwargs.pop("compound")
        if "padx" in kwargs:
            self._padx = kwargs.pop("padx")
        if "pady" in kwargs:
            self._pady = kwargs.pop("pady")
            pady_changed = True
        if "radius" in kwargs:
            self._radius = kwargs.pop("radius")
        if "font" in kwargs:
            self._font = tkfont.Font(font=kwargs.pop("font"))
            pady_changed = True
        super().configure(**kwargs)
        if pady_changed:
            height = self._font.metrics("linespace") + (self._pady * 2)
            super().configure(height=height)
        self._redraw()

    config = configure


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS words (id INTEGER PRIMARY KEY, word TEXT, meaning TEXT, image TEXT, image_blob BLOB)"
    )
    conn.execute(
        "CREATE TABLE IF NOT EXISTS config (key TEXT PRIMARY KEY, value TEXT)"
    )
    # Ensure new column exists for older databases
    try:
        conn.execute("ALTER TABLE words ADD COLUMN image_blob BLOB")
    except sqlite3.OperationalError:
        pass
    return conn


def migrate_from_files(conn: sqlite3.Connection) -> None:
    words: list[dict[str, str]] = []
    if os.path.exists(WORDS_FILE):
        with open(WORDS_FILE, "r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                word = (row.get("word") or "").strip()
                meaning = (row.get("meaning") or "").strip()
                image = (row.get("image") or "").strip()
                if word:
                    image_blob = None
                    if image and os.path.exists(image):
                        try:
                            with open(image, "rb") as img_f:
                                image_blob = img_f.read()
                        except OSError:
                            image_blob = None
                    words.append({"word": word, "meaning": meaning, "image": image, "image_blob": image_blob})
    elif os.path.exists(LEGACY_WORDS_FILE):
        with open(LEGACY_WORDS_FILE, "r", encoding="utf-8") as f:
            for line in f:
                w = line.strip()
                if w:
                    words.append({"word": w, "meaning": "", "image": "", "image_blob": None})

    if not words:
        words = [{"word": w, "meaning": "", "image": "", "image_blob": None} for w in DEFAULT_WORDS]

    conn.executemany(
        "INSERT INTO words (word, meaning, image, image_blob) VALUES (?, ?, ?, ?)",
        [(w["word"], w["meaning"], w["image"], w.get("image_blob")) for w in words],
    )
    conn.execute("INSERT OR REPLACE INTO config (key, value) VALUES (?, ?)", ("db_initialized", "1"))
    conn.commit()


def load_words(path: str) -> list[dict[str, str]]:
    conn = get_conn()
    cur = conn.execute("SELECT id, word, meaning, image, image_blob FROM words ORDER BY id ASC")
    rows = cur.fetchall()
    if not rows:
        cfg = load_config(DB_PATH)
        if cfg.get("db_initialized") == "1":
            conn.close()
            return []
        migrate_from_files(conn)
        cur = conn.execute("SELECT id, word, meaning, image, image_blob FROM words ORDER BY id ASC")
        rows = cur.fetchall()
    conn.close()
    return [{"id": r[0], "word": r[1], "meaning": r[2], "image": r[3], "image_blob": r[4]} for r in rows]


def save_words(path: str, words: list[dict[str, str]]) -> None:
    # De-duplicate to avoid accidental double-saves
    seen: set[tuple[str, str, str]] = set()
    deduped: list[dict[str, str]] = []
    for w in words:
        key = (w.get("word", ""), w.get("meaning", ""), w.get("image", ""))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(w)

    conn = get_conn()
    conn.execute("DELETE FROM words")
    conn.executemany(
        "INSERT OR REPLACE INTO words (id, word, meaning, image, image_blob) VALUES (?, ?, ?, ?, ?)",
        [
            (
                w.get("id"),
                w.get("word", ""),
                w.get("meaning", ""),
                w.get("image", ""),
                w.get("image_blob"),
            )
            for w in deduped
        ],
    )
    conn.execute("INSERT OR REPLACE INTO config (key, value) VALUES (?, ?)", ("db_initialized", "1"))
    conn.commit()
    conn.close()


def load_config(path: str) -> dict[str, int]:
    conn = get_conn()
    cur = conn.execute("SELECT key, value FROM config")
    rows = cur.fetchall()
    conn.close()
    return {k: v for k, v in rows if k and v is not None}


def save_config(path: str, config: dict[str, int]) -> None:
    conn = get_conn()
    for k, v in config.items():
        conn.execute("INSERT OR REPLACE INTO config (key, value) VALUES (?, ?)", (k, str(v)))
    conn.commit()
    conn.close()


class WordRotatorApp:
    def __init__(self, root: tk.Tk, words: list[dict[str, str]]):
        self.root = root
        self.words = words
        self.index = 0
        self.timer_id: str | None = None
        self.countdown_id: str | None = None
        self.interval_seconds = INTERVAL_SECONDS
        self.notify_duration_seconds = NOTIFY_DURATION_SECONDS
        self.remaining_seconds = self.interval_seconds
        self.notify_hide_id: str | None = None
        self.notify_image_ref: tk.PhotoImage | None = None
        self.donate_image_ref: tk.PhotoImage | None = None
        self.logo_image_ref: tk.PhotoImage | None = None
        self.button_icons: dict[str, tk.PhotoImage] = {}
        self.edit_index: int | None = None
        self.backup_timer_id: str | None = None
        self.backup_interval_seconds = 60 * 60
        self.undo_stack: list[list[dict[str, str]]] = []
        self._text_undo_stacks: dict[tk.Widget, list[str]] = {}
        self._text_undo_suspended: set[tk.Widget] = set()
        self._text_undo_initialized = False
        self.zoom_factor = 1.0
        self.base_scaling = float(self.root.tk.call("tk", "scaling"))
        self.is_fullscreen = False
        self._zoom_fonts: dict[tk.Widget, dict[str, int | str]] = {}
        self._zoom_btn_fonts: dict[RoundedButton, dict[str, int | str]] = {}
        cfg = load_config(DB_PATH)
        try:
            self.interval_seconds = int(cfg.get("interval_seconds", self.interval_seconds))
        except (TypeError, ValueError):
            self.interval_seconds = INTERVAL_SECONDS
        self.words_per_day = int(cfg.get("words_per_day", 10))
        self.remaining_seconds = self.interval_seconds
        self.active_words: list[dict[str, str]] = []
        self.active_date: str | None = None

        # Hide root; use separate windows for manager and notification.
        self.root.withdraw()

        self.manager = tk.Toplevel(self.root)
        self.manager.title("VocabBuddy")
        self.manager.geometry("820x780")
        self.manager.minsize(820, 780)
        self.manager.resizable(True, True)
        self.manager.configure(bg=THEME["bg"])
        self.manager.protocol("WM_DELETE_WINDOW", self.on_manager_close)
        self._build_menu_bar()

        self.editor: tk.Toplevel | None = None

        self.notify = tk.Toplevel(self.root)
        self.notify.withdraw()
        self.notify.overrideredirect(True)
        self.notify.attributes("-topmost", True)
        self.apply_app_icon()

        self.word_var = tk.StringVar()
        self.meaning_var = tk.StringVar()
        self.status_var = tk.StringVar()

        header = tk.Frame(self.manager, bg=THEME["bg"])
        header.pack(fill="x", padx=24, pady=(16, 6))
        header.columnconfigure(0, weight=1)
        header.columnconfigure(1, weight=0)
        header.columnconfigure(2, weight=1)

        brand_logo = tk.Label(header, bg=THEME["bg"])
        brand_logo.grid(row=0, column=1)
        self.load_logo(brand_logo, max_w=300, max_h=123, path=HEADER_LOGO_PATH)

        donate_link = self.make_button(header, text="Donate", command=self.open_donate, kind="danger")
        donate_link.configure(
            width=110,
            padx=6,
            pady=4,
            font=("Segoe UI", 8),
            radius=12,
            bg="#F43F5E",
            activebackground="#E11D48",
            fg="#FFFFFF",
        )
        self._set_btn_icon_path(donate_link, DONATE_ICON_PATH, size=14, force_white=True)
        donate_link.grid(row=0, column=2, sticky="e")

        self.page_title_var = tk.StringVar(value="Test")
        self.page_title_entry = tk.Entry(
            self.manager,
            textvariable=self.page_title_var,
            font=("Segoe UI", 20, "bold"),
            justify="center",
            bg=THEME["bg"],
            fg=THEME["text"],
            relief="flat",
            highlightthickness=0,
        )
        self.page_title_entry.pack(fill="x", padx=48, pady=(4, 6))

        self.next_word_var = tk.StringVar()
        pill = tk.Frame(self.manager, bg="#E8EDFF")
        pill.pack(pady=(4, 10))
        self.next_word_label = tk.Label(
            pill,
            textvariable=self.next_word_var,
            font=("Segoe UI", 10, "bold"),
            bg="#E8EDFF",
            fg=THEME["text"],
            padx=14,
            pady=6,
        )
        self.next_word_label.pack()

        stats = tk.Frame(self.manager, bg=THEME["bg"])
        stats.pack(pady=(6, 10))

        self.interval_value = tk.StringVar()
        self.total_value = tk.StringVar()
        self.per_day_value = tk.StringVar()

        def make_stat(parent, title, var):
            card_w, card_h = 160, 90
            canvas = tk.Canvas(parent, width=card_w, height=card_h, bg=THEME["bg"], highlightthickness=0)
            canvas.pack(side="left", padx=8)
            self._rounded_rect(canvas, 4, 4, card_w - 4, card_h - 4, 14, THEME["card"], "")
            canvas.create_text(
                card_w // 2,
                28,
                text=title,
                fill=THEME["muted"],
                font=("Segoe UI", 9),
            )
            value_id = canvas.create_text(
                card_w // 2,
                60,
                text=var.get(),
                fill=THEME["text"],
                font=("Segoe UI", 16, "bold"),
            )
            if not hasattr(self, "stat_texts"):
                self.stat_texts = []
            self.stat_texts.append((var, canvas, value_id))

        make_stat(stats, "Interval", self.interval_value)
        make_stat(stats, "Total words", self.total_value)
        make_stat(stats, "Per day", self.per_day_value)

        content = tk.Frame(self.manager, bg=THEME["bg"])
        content.pack(padx=24, pady=(8, 12))

        words_wrap = tk.Canvas(content, bg=THEME["bg"], highlightthickness=0, width=360, height=420)
        words_wrap.grid(row=0, column=0, padx=(0, 24))
        words_card = self._card_container(words_wrap, width=360, height=420, radius=16)
        words_wrap.create_window((0, 0), window=words_card, anchor="nw")
        tk.Label(words_card, text="Words", bg=THEME["card"], fg=THEME["text"], font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(10, 6))

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(
            words_card,
            textvariable=self.search_var,
            bg=THEME["bg"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            highlightthickness=0,
            bd=0,
        )
        self.search_entry.pack(fill="x", padx=12, pady=(0, 8))
        self.search_entry.insert(0, "Search words...")
        self.search_entry.bind("<FocusIn>", self._clear_search_placeholder)
        self.search_entry.bind("<FocusOut>", self._restore_search_placeholder)
        self.search_var.trace_add("write", lambda *args: self.refresh_listbox())

        self.words_list = tk.Listbox(
            words_card,
            height=10,
            bg=THEME["card"],
            fg=THEME["text"],
            highlightthickness=0,
            bd=0,
            selectbackground=THEME["accent"],
            selectforeground="#FFFFFF",
            font=("Segoe UI", 11),
        )
        self.words_list.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        self.count_var = tk.StringVar()
        tk.Label(words_card, textvariable=self.count_var, bg=THEME["card"], fg=THEME["muted"], font=("Segoe UI", 9)).pack(anchor="w", padx=12, pady=(0, 6))

        actions_row = tk.Frame(words_card, bg=THEME["card"])
        actions_row.pack(pady=(0, 12), padx=12, fill="x")
        actions_row.columnconfigure(0, weight=1, uniform="wordactions")
        actions_row.columnconfigure(1, weight=1, uniform="wordactions")

        self.remove_btn = self.make_button(actions_row, text="Remove selected", command=self.remove_selected, kind="secondary")
        self.remove_btn.grid(row=0, column=0, padx=(0, 6), sticky="ew")

        self.edit_btn = self.make_button(actions_row, text="Edit selected", command=self.edit_selected, kind="secondary")
        self.edit_btn.grid(row=0, column=1, padx=(6, 0), sticky="ew")

        controls_wrap = tk.Canvas(content, bg=THEME["bg"], highlightthickness=0, width=360, height=420)
        controls_wrap.grid(row=0, column=1)
        controls_card = self._card_container(controls_wrap, width=360, height=420, radius=16)
        controls_wrap.create_window((0, 0), window=controls_card, anchor="nw")
        tk.Label(controls_card, text="Controls", bg=THEME["card"], fg=THEME["text"], font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(10, 6))

        controls = tk.Frame(controls_card, bg=THEME["card"])
        controls.pack(padx=12, pady=(0, 12), anchor="center")
        controls.columnconfigure(0, weight=1, uniform="controls")
        controls.columnconfigure(1, weight=1, uniform="controls")

        self.start_btn = self.make_button(controls, text="Start", command=self.start, kind="primary")
        self.start_btn.grid(row=0, column=0, padx=6, pady=6, sticky="ew")

        self.stop_btn = self.make_button(controls, text="Stop", command=self.stop, kind="danger")
        self.stop_btn.grid(row=0, column=1, padx=6, pady=6, sticky="ew")

        self.next_btn = self.make_button(controls, text="Notify", command=self.show_next, kind="secondary")
        self.next_btn.grid(row=1, column=0, padx=6, pady=6, sticky="ew")

        self.editor_btn = self.make_button(controls, text="Open Editor", command=self.open_editor, kind="secondary")
        self.editor_btn.grid(row=1, column=1, padx=6, pady=6, sticky="ew")

        self.save_btn = self.make_button(controls, text="Save Words", command=self.save_words_clicked, kind="secondary")
        self.save_btn.configure(
            bg="#3B82F6",
            fg="#FFFFFF",
            activebackground="#2563EB",
            activeforeground="#FFFFFF",
        )
        self.save_btn.grid(row=2, column=0, padx=6, pady=6, sticky="ew")

        self.clear_btn = self.make_button(controls, text="Clear All", command=self.clear_all, kind="danger_outline")
        self.clear_btn.grid(row=2, column=1, padx=6, pady=6, sticky="ew")

        footer = tk.Label(
            self.manager,
            text="Powered by EDDARIF AYOUB",
            font=("Segoe UI", 9),
            bg=THEME["bg"],
            fg=THEME["muted"],
        )
        footer.pack(pady=(4, 12))

        self._apply_button_icons()

        self.update_ui(initial=True)
        self._maybe_check_for_updates()

    def update_ui(self, initial: bool = False) -> None:
        self.interval_value.set(f"{self.interval_seconds}s")
        self.total_value.set(str(len(self.words)))
        self.per_day_value.set(str(self.words_per_day))
        if hasattr(self, "stat_texts"):
            for var, canvas, text_id in self.stat_texts:
                canvas.itemconfigure(text_id, text=var.get())

        if not self.words:
            self.word_var.set("No words found")
            self.meaning_var.set("")
            self.page_title_var.set("No words found")
            self.next_word_var.set("Next word: stopped")
            self.refresh_listbox()
            return

        if initial:
            active = self.get_active_words()
            if active:
                self.word_var.set(active[self.index]["word"])
                self.meaning_var.set(active[self.index].get("meaning", ""))
        if self.word_var.get():
            self.page_title_var.set(self.word_var.get())
        self.refresh_listbox()
        self.update_countdown_label(running=self.timer_id is not None)

    def show_next(self) -> None:
        active = self.get_active_words()
        if not active:
            return
        self.index = (self.index + 1) % len(active)
        item = active[self.index]
        word = item["word"]
        meaning = item.get("meaning", "")
        image_path = item.get("image", "")
        image_blob = item.get("image_blob")
        self.word_var.set(word)
        self.page_title_var.set(word)
        self.meaning_var.set(meaning)
        self.status_var.set(
            f"Interval: {self.interval_seconds} seconds | Total words: {len(self.words)} | Per day: {self.words_per_day}"
        )
        self.remaining_seconds = self.interval_seconds
        self.update_countdown_label(running=True)
        self.show_notification(word, meaning, image_path, image_blob)

    def _tick(self) -> None:
        self.show_next()
        self.timer_id = self.root.after(self.interval_seconds * 1000, self._tick)

    def start(self) -> None:
        if self.timer_id is None:
            # Reset countdown at start so the display is accurate.
            self.remaining_seconds = self.interval_seconds
            self.update_countdown_label(running=True)
            self.timer_id = self.root.after(self.interval_seconds * 1000, self._tick)
            self.start_countdown()
        else:
            messagebox.showinfo("Already running", "The word rotation is already running.")

    def stop(self) -> None:
        if self.timer_id is not None:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None
        if self.countdown_id is not None:
            self.root.after_cancel(self.countdown_id)
            self.countdown_id = None
        self.update_countdown_label(running=False)

    def add_word(self) -> None:
        w = self.word_entry.get().strip()
        if not w:
            messagebox.showwarning("Empty word", "Please type a word first.")
            return
        edited_existing = False
        save_after = True
        meaning = self.meaning_entry.get("1.0", tk.END).strip()
        image = self.image_entry.get().strip()
        image_blob = None
        if image and os.path.exists(image):
            try:
                with open(image, "rb") as img_f:
                    image_blob = img_f.read()
            except OSError:
                image_blob = None
        if self.edit_index is None:
            for item in self.words:
                if item.get("word", "") == w and item.get("meaning", "") == meaning and item.get("image", "") == image:
                    messagebox.showinfo("Duplicate", "This word is already in the list.")
                    return
            self._push_undo_state()
            conn = get_conn()
            cur = conn.execute(
                "INSERT INTO words (word, meaning, image, image_blob) VALUES (?, ?, ?, ?)",
                (w, meaning, image, image_blob),
            )
            new_id = cur.lastrowid
            conn.commit()
            conn.close()
            self.words.append({"id": new_id, "word": w, "meaning": meaning, "image": image, "image_blob": image_blob})
        else:
            # Update existing entry in place
            if 0 <= self.edit_index < len(self.words):
                self._push_undo_state()
                old_item = self.words[self.edit_index]
                current_blob = old_item.get("image_blob")
                word_id = old_item.get("id")
                self.words[self.edit_index] = {
                    "id": word_id,
                    "word": w,
                    "meaning": meaning,
                    "image": image,
                    "image_blob": image_blob if image_blob is not None else current_blob,
                }
                if word_id is not None:
                    conn = get_conn()
                    conn.execute(
                        "UPDATE words SET word=?, meaning=?, image=?, image_blob=? WHERE id=?",
                        (w, meaning, image, image_blob if image_blob is not None else current_blob, word_id),
                    )
                    conn.commit()
                    conn.close()
                    save_after = False
                else:
                    save_words(WORDS_FILE, self.words)
                    save_after = False
                for i, item in enumerate(self.active_words):
                    if word_id is not None and item.get("id") == word_id:
                        self.active_words[i] = self.words[self.edit_index]
                        break
                    if word_id is None and item == old_item:
                        self.active_words[i] = self.words[self.edit_index]
                        break
                edited_existing = True
            self.edit_index = None
            self.add_btn.configure(text="Add word")
        self.word_entry.delete(0, tk.END)
        self.meaning_entry.delete("1.0", tk.END)
        self.image_entry.configure(state="normal")
        self.image_entry.delete(0, tk.END)
        self.image_entry.configure(state="readonly")
        if self._text_undo_initialized:
            self._reset_text_undo_stacks()
        if not edited_existing:
            new_index = len(self.words) - 1
            if hasattr(self, "_ensure_today_includes_new_word"):
                self._ensure_today_includes_new_word(new_index)
        if len(self.words) == 1:
            self.index = 0
            active = self.get_active_words()
            if active:
                self.word_var.set(active[self.index]["word"])
                self.meaning_var.set(active[self.index].get("meaning", ""))
                self.page_title_var.set(active[self.index]["word"])
        self.refresh_listbox()
        self.status_var.set(
            f"Interval: {self.interval_seconds} seconds | Total words: {len(self.words)} | Per day: {self.words_per_day}"
        )
        if save_after:
            save_words(WORDS_FILE, self.words)
        if not edited_existing:
            messagebox.showinfo("Saved", "Word added successfully!")
        self._persist_today_words()
        self.update_ui()

    def browse_image(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose an image",
            filetypes=[
                ("Image files", "*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.image_entry.configure(state="normal")
            self.image_entry.delete(0, tk.END)
            self.image_entry.insert(0, path)
            self.image_entry.configure(state="readonly")

    def _load_icon(self, filename: str, size: int = 18, force_white: bool = False) -> tk.PhotoImage | None:
        path = os.path.join(ICON_DIR, filename)
        if not os.path.exists(path):
            return None
        try:
            img = Image.open(path).convert("RGBA")
            img = img.resize((size, size), Image.LANCZOS)
            if force_white:
                _, _, _, a = img.split()
                white = Image.new("RGBA", img.size, (255, 255, 255, 255))
                white.putalpha(a)
                img = white
            tk_img = ImageTk.PhotoImage(img)
            self.button_icons[filename] = tk_img
            return tk_img
        except Exception:
            return None

    def _load_icon_path(self, path: str, size: int = 18, force_white: bool = False) -> tk.PhotoImage | None:
        if not os.path.exists(path):
            return None
        try:
            img = Image.open(path).convert("RGBA")
            img = img.resize((size, size), Image.LANCZOS)
            if force_white:
                _, _, _, a = img.split()
                white = Image.new("RGBA", img.size, (255, 255, 255, 255))
                white.putalpha(a)
                img = white
            tk_img = ImageTk.PhotoImage(img)
            self.button_icons[path] = tk_img
            return tk_img
        except Exception:
            return None

    def _set_btn_icon(self, btn: tk.Button, filename: str, size: int = 18, force_white: bool = False) -> None:
        icon = self._load_icon(filename, size, force_white)
        if icon:
            btn.configure(image=icon, compound="left", padx=10)

    def _set_btn_icon_path(self, btn: tk.Button, path: str, size: int = 18, force_white: bool = False) -> None:
        icon = self._load_icon_path(path, size, force_white)
        if icon:
            btn.configure(image=icon, compound="left", padx=10)

    def _apply_button_icons(self) -> None:
        self._set_btn_icon(self.start_btn, "play-button.png", force_white=True)
        self._set_btn_icon(self.stop_btn, "stop-button.png", force_white=True)
        self._set_btn_icon(self.next_btn, "Notify.png")
        self._set_btn_icon(self.editor_btn, "document-editor_36.png", size=18)
        self._set_btn_icon(self.save_btn, "OpenEditor.png", force_white=True)
        self._set_btn_icon(self.clear_btn, "clean all red.png")
        self.clear_btn.configure(fg=THEME["danger"], activeforeground=THEME["danger"])
        self._set_btn_icon(self.remove_btn, "Remove selected.png")
        if hasattr(self, "edit_btn"):
            self._set_btn_icon(self.edit_btn, "document-editor.png")
        if hasattr(self, "add_btn"):
            self._set_btn_icon(self.add_btn, "add.png", force_white=True)
        if hasattr(self, "apply_all_btn"):
            self._set_btn_icon(self.apply_all_btn, "check (1).png", size=18, force_white=True)
        if hasattr(self, "image_browse_btn"):
            self._set_btn_icon(self.image_browse_btn, "up-loading.png", force_white=True)

    def _clear_search_placeholder(self, _event=None) -> None:
        if self.search_entry.get() == "Search words...":
            self.search_entry.delete(0, tk.END)

    def _restore_search_placeholder(self, _event=None) -> None:
        if not self.search_entry.get():
            self.search_entry.insert(0, "Search words...")

    def _rounded_rect(
        self,
        canvas: tk.Canvas,
        x1: int,
        y1: int,
        x2: int,
        y2: int,
        r: int,
        fill: str,
        outline: str,
    ) -> None:
        outline = "" if not outline else outline
        canvas.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, fill=fill, outline=outline)
        canvas.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, fill=fill, outline=outline)
        canvas.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, fill=fill, outline=outline)
        canvas.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, fill=fill, outline=outline)
        canvas.create_rectangle(x1 + r, y1, x2 - r, y2, fill=fill, outline=outline)
        canvas.create_rectangle(x1, y1 + r, x2, y2 - r, fill=fill, outline=outline)

    def _card_container(self, canvas: tk.Canvas, width: int, height: int, radius: int = 16) -> tk.Frame:
        self._rounded_rect(canvas, 4, 4, width - 4, height - 4, radius, THEME["card"], "")
        inner = tk.Frame(canvas, bg=THEME["card"], width=width, height=height)
        inner.pack_propagate(False)
        return inner
    def style_button(self, btn: tk.Button, kind: str = "primary") -> None:
        if kind == "primary":
            bg = "#22C55E"
            hover = "#16A34A"
            fg = "#FFFFFF"
        elif kind == "danger":
            bg = "#EF4444"
            hover = "#DC2626"
            fg = "#FFFFFF"
        else:
            bg = "#E5E5E5"
            hover = "#DCDCDC"
            fg = THEME["danger"] if kind == "danger_outline" else THEME["text"]
        btn.configure(
            bg=bg,
            fg=fg,
            activebackground=hover,
            activeforeground=fg,
            relief="flat",
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            padx=12,
            pady=8,
        )

    def make_button(self, parent: tk.Widget, text: str, command, kind: str = "secondary") -> RoundedButton:
        btn = RoundedButton(parent, text=text, command=command)
        self.style_button(btn, kind=kind)
        return btn

    def open_donate(self) -> None:
        webbrowser.open_new_tab(DONATE_URL)

    def load_donate_icon(self) -> None:
        if os.path.exists(DONATE_ICON_PATH):
            try:
                img = tk.PhotoImage(file=DONATE_ICON_PATH)
                max_w, max_h = 28, 28
                scale = max(img.width() / max_w, img.height() / max_h)
                if scale >= 2:
                    img = img.subsample(int(math.ceil(scale)))
                self.donate_image_ref = img
                self.donate_btn.configure(image=self.donate_image_ref, compound="left", padx=8, pady=4)
            except tk.TclError:
                pass

    def apply_app_icon(self) -> None:
        if os.path.exists(APP_ICON_PATH):
            try:
                icon = tk.PhotoImage(file=APP_ICON_PATH)
                self.manager.iconphoto(True, icon)
                self.notify.iconphoto(True, icon)
                if self.editor is not None:
                    self.editor.iconphoto(True, icon)
            except tk.TclError:
                pass

    def load_logo(self, label: tk.Label, max_w: int, max_h: int, path: str) -> None:
        if os.path.exists(path):
            try:
                img = tk.PhotoImage(file=path)
                scale = max(img.width() / max_w, img.height() / max_h)
                if scale >= 2:
                    img = img.subsample(int(math.ceil(scale)))
                self.logo_image_ref = img
                label.configure(image=self.logo_image_ref)
            except tk.TclError:
                label.configure(text="VocabBudy", fg=THEME["text"])

    def save_words_clicked(self) -> None:
        save_words(WORDS_FILE, self.words)
        self.update_ui()
        messagebox.showinfo("Saved", f"Saved {len(self.words)} words to {WORDS_FILE}.")

    def _push_undo_state(self) -> None:
        snapshot = copy.deepcopy(self.words)
        self.undo_stack.append(snapshot)
        if len(self.undo_stack) > 20:
            self.undo_stack.pop(0)

    def _replace_words(self, words: list[dict[str, str]]) -> None:
        self.words = words
        save_words(WORDS_FILE, self.words)
        save_config(DB_PATH, {"cycle_order": "", "cycle_index": 0, "daily_set": "", "cycle_date": ""})
        self.active_date = None
        self.active_words = []
        self.index = 0
        self.refresh_listbox()
        active = self.get_active_words()
        if active:
            self.word_var.set(active[self.index]["word"])
            self.meaning_var.set(active[self.index].get("meaning", ""))
            self.page_title_var.set(active[self.index]["word"])
        else:
            self.word_var.set("No words found")
            self.meaning_var.set("")
            self.page_title_var.set("No words found")
        self.update_ui()

    def undo_last_action(self) -> None:
        if not self.undo_stack:
            messagebox.showinfo("Undo", "Nothing to undo.")
            return
        previous = self.undo_stack.pop()
        self._replace_words(previous)
        messagebox.showinfo("Undo", "Last action has been undone.")

    def _handle_undo_event(self, event) -> str | None:
        widget = event.widget
        if isinstance(widget, (tk.Entry, tk.Text)):
            if self._undo_text_widget(widget):
                return "break"
        self.undo_last_action()
        return "break"

    def _handle_ctrl_mousewheel(self, event) -> str:
        if event.delta > 0:
            self.zoom_in()
        elif event.delta < 0:
            self.zoom_out()
        return "break"

    def _handle_reset_zoom_event(self, _event) -> str:
        self.reset_zoom()
        return "break"

    def _apply_zoom(self) -> None:
        scale = self.base_scaling * self.zoom_factor
        self.root.tk.call("tk", "scaling", scale)
        self._ensure_zoom_fonts()
        for widget, base in self._zoom_fonts.items():
            size = int(base["size"] * self.zoom_factor)
            size = max(8, size)
            font = tkfont.Font(
                family=base["family"],
                size=size,
                weight=base["weight"],
                slant=base["slant"],
                underline=base["underline"],
                overstrike=base["overstrike"],
            )
            try:
                widget.configure(font=font)
            except tk.TclError:
                pass
        for btn, base in self._zoom_btn_fonts.items():
            size = int(base["size"] * self.zoom_factor)
            size = max(8, size)
            try:
                btn.configure(font=(base["family"], size, base["weight"]))
            except tk.TclError:
                pass
        if self.editor is not None and self.editor.winfo_exists():
            self.editor.update_idletasks()
        self.manager.update_idletasks()

    def zoom_in(self) -> None:
        self.zoom_factor = min(2.0, self.zoom_factor * 1.1)
        self._apply_zoom()

    def zoom_out(self) -> None:
        self.zoom_factor = max(0.6, self.zoom_factor / 1.1)
        self._apply_zoom()

    def reset_zoom(self) -> None:
        self.zoom_factor = 1.0
        self._apply_zoom()

    def toggle_fullscreen(self) -> None:
        self.is_fullscreen = not self.is_fullscreen
        self.manager.attributes("-fullscreen", self.is_fullscreen)

    def _ensure_zoom_fonts(self) -> None:
        for root in (self.manager, self.editor):
            if root is None or not root.winfo_exists():
                continue
            for widget in self._iter_widgets(root):
                if isinstance(widget, RoundedButton):
                    if widget not in self._zoom_btn_fonts:
                        actual = widget._font.actual()
                        self._zoom_btn_fonts[widget] = {
                            "family": actual.get("family", "Segoe UI"),
                            "size": int(actual.get("size", 10)) or 10,
                            "weight": actual.get("weight", "normal"),
                            "slant": actual.get("slant", "roman"),
                            "underline": int(actual.get("underline", 0)),
                            "overstrike": int(actual.get("overstrike", 0)),
                        }
                    continue
                try:
                    font_value = widget.cget("font")
                except tk.TclError:
                    continue
                if not font_value:
                    continue
                if widget in self._zoom_fonts:
                    continue
                actual = tkfont.Font(font=font_value).actual()
                base_size = int(actual.get("size", 10)) or 10
                self._zoom_fonts[widget] = {
                    "family": actual.get("family", "Segoe UI"),
                    "size": base_size,
                    "weight": actual.get("weight", "normal"),
                    "slant": actual.get("slant", "roman"),
                    "underline": int(actual.get("underline", 0)),
                    "overstrike": int(actual.get("overstrike", 0)),
                }

    def _iter_widgets(self, root: tk.Widget) -> list[tk.Widget]:
        widgets = []
        stack = [root]
        while stack:
            current = stack.pop()
            widgets.append(current)
            try:
                stack.extend(current.winfo_children())
            except tk.TclError:
                continue
        return widgets

    def _maybe_check_for_updates(self) -> None:
        # Skip automatic checks if we've already checked today.
        cfg = load_config(DB_PATH)
        last = cfg.get("last_update_check", "")
        today = date.today().isoformat()
        if last == today:
            return
        self.check_for_updates(manual=False)

    def check_for_updates(self, manual: bool = False) -> None:
        # Run update checks in a background thread to keep UI responsive.
        def worker():
            try:
                # Always bypass cache to avoid stale latest.json or release data.
                info = self._fetch_update_info(bypass_cache=True)
            except Exception as exc:
                if manual:
                    self.root.after(0, lambda: messagebox.showerror("Update check failed", str(exc)))
                return

            latest = (info.get("latest_version") or "").strip()
            download_url = (info.get("download_url") or "").strip()
            release_notes = (info.get("release_notes") or "").strip()
            sha256 = (info.get("sha256") or "").strip().lower()

            if not latest or not download_url:
                if manual:
                    self.root.after(0, lambda: messagebox.showerror("Update check failed", "Update data is incomplete."))
                return

            if not self._is_newer_version(latest, CURRENT_VERSION):
                if manual:
                    self.root.after(0, lambda: messagebox.showinfo("Updates", "You are up to date."))
                save_config(DB_PATH, {"last_update_check": date.today().isoformat()})
                return

            def prompt():
                if self._show_update_prompt(latest, release_notes):
                    self._download_and_update(download_url, sha256, latest)
                else:
                    save_config(DB_PATH, {"last_update_check": date.today().isoformat()})

            self.root.after(0, prompt)

        threading.Thread(target=worker, daemon=True).start()

    def _show_update_prompt(self, latest: str, notes: str) -> bool:
        # Custom dialog so buttons match "Update now" / "Later".
        result = {"update": False}
        win = tk.Toplevel(self.manager)
        win.title("Update Available")
        win.geometry("520x360")
        win.minsize(480, 300)
        win.configure(bg=THEME["bg"])
        win.transient(self.manager)
        win.grab_set()

        container = tk.Frame(win, bg=THEME["bg"])
        container.pack(fill="both", expand=True, padx=16, pady=16)

        tk.Label(
            container,
            text="Update Available",
            font=("Segoe UI", 13, "bold"),
            bg=THEME["bg"],
            fg=THEME["text"],
        ).pack(anchor="w")

        tk.Label(
            container,
            text=f"Current version: {CURRENT_VERSION}\\nNew version: {latest}",
            font=("Segoe UI", 10),
            bg=THEME["bg"],
            fg=THEME["text"],
            justify="left",
        ).pack(anchor="w", pady=(8, 8))

        text_frame = tk.Frame(container, bg=THEME["card"], highlightthickness=1, highlightbackground=THEME["border"])
        text_frame.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")

        text = tk.Text(
            text_frame,
            wrap="word",
            yscrollcommand=scrollbar.set,
            bg=THEME["card"],
            fg=THEME["text"],
            relief="flat",
            highlightthickness=0,
            padx=10,
            pady=8,
            font=("Segoe UI", 10),
        )
        text.pack(fill="both", expand=True)
        scrollbar.config(command=text.yview)
        text.insert("1.0", notes or "No release notes provided.")
        text.configure(state="disabled")

        btn_row = tk.Frame(container, bg=THEME["bg"])
        btn_row.pack(fill="x", pady=(10, 0))

        def do_update():
            result["update"] = True
            win.destroy()

        def do_later():
            result["update"] = False
            win.destroy()

        update_btn = self.make_button(btn_row, text="Update now", command=do_update, kind="primary")
        update_btn.pack(side="right", padx=(8, 0))
        later_btn = self.make_button(btn_row, text="Later", command=do_later, kind="secondary")
        later_btn.pack(side="right")

        win.wait_window()
        return result["update"]

    def _fetch_update_info(self, bypass_cache: bool = False) -> dict[str, str]:
        # Try GitHub Releases API first for freshest data; fall back to latest.json.
        info = self._fetch_update_info_from_releases(bypass_cache=bypass_cache)
        if info:
            return info
        return self._fetch_update_info_from_latest_json(bypass_cache=bypass_cache)

    def _fetch_update_info_from_releases(self, bypass_cache: bool = False) -> dict[str, str] | None:
        url = self._cache_bust_url(GITHUB_RELEASES_API) if bypass_cache else GITHUB_RELEASES_API
        req = urllib.request.Request(url, headers=self._update_headers(bypass_cache))
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                raw = resp.read().decode("utf-8")
            data = json.loads(raw)
        except Exception:
            return None

        tag = (data.get("tag_name") or "").lstrip("v").strip()
        notes = (data.get("body") or "").strip()
        assets = data.get("assets") or []
        latest_json_url = ""
        installer_url = ""
        sha256_url = ""
        for asset in assets:
            name = asset.get("name") or ""
            url = asset.get("browser_download_url") or ""
            if name.lower() == "latest.json":
                latest_json_url = url
            if name.lower().endswith(".exe") and "setup" in name.lower():
                installer_url = url
            if name.lower().endswith(".sha256") and "setup" in name.lower():
                sha256_url = url

        if latest_json_url:
            info = self._fetch_update_info_from_url(latest_json_url, bypass_cache=bypass_cache)
            if info:
                return info

        sha256_value = ""
        if sha256_url:
            try:
                sha_raw = self._fetch_text_url(sha256_url, bypass_cache=bypass_cache)
                sha256_value = sha_raw.strip().split()[0]
            except Exception:
                sha256_value = ""

        if tag and installer_url:
            return {
                "latest_version": tag,
                "download_url": installer_url,
                "release_notes": notes,
                "sha256": sha256_value,
            }
        return None

    def _fetch_update_info_from_latest_json(self, bypass_cache: bool = False) -> dict[str, str]:
        url = self._cache_bust_url(UPDATE_ENDPOINT) if bypass_cache else UPDATE_ENDPOINT
        return self._fetch_update_info_from_url(url, bypass_cache=bypass_cache)

    def _fetch_update_info_from_url(self, url: str, bypass_cache: bool = False) -> dict[str, str]:
        req = urllib.request.Request(url, headers=self._update_headers(bypass_cache))
        with urllib.request.urlopen(req, timeout=15) as resp:
            raw = resp.read().decode("utf-8")
        return json.loads(raw)

    def _fetch_text_url(self, url: str, bypass_cache: bool = False) -> str:
        req = urllib.request.Request(url, headers=self._update_headers(bypass_cache))
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.read().decode("utf-8")

    def _update_headers(self, bypass_cache: bool) -> dict[str, str]:
        headers = {"User-Agent": "VocabBuddy"}
        if bypass_cache:
            headers["Cache-Control"] = "no-cache"
            headers["Pragma"] = "no-cache"
        return headers

    def _cache_bust_url(self, url: str) -> str:
        sep = "&" if "?" in url else "?"
        return f"{url}{sep}t={int(time.time())}"

    def _is_newer_version(self, latest: str, current: str) -> bool:
        def parse(v: str) -> list[int]:
            parts = []
            for p in v.strip().split("."):
                try:
                    parts.append(int(p))
                except ValueError:
                    parts.append(0)
            return parts
        a = parse(latest)
        b = parse(current)
        max_len = max(len(a), len(b))
        a += [0] * (max_len - len(a))
        b += [0] * (max_len - len(b))
        return a > b

    def _download_and_update(self, url: str, expected_sha256: str, latest: str) -> None:
        # Download and verify the installer, then launch it on success.
        def worker():
            try:
                fd, temp_path = tempfile.mkstemp(prefix="VocabBuddy_", suffix=".exe")
                os.close(fd)
                with urllib.request.urlopen(url, timeout=30) as resp, open(temp_path, "wb") as f:
                    while True:
                        chunk = resp.read(1024 * 128)
                        if not chunk:
                            break
                        f.write(chunk)
                actual = self._sha256_file(temp_path)
                if expected_sha256 and actual.lower() != expected_sha256.lower():
                    raise ValueError("Downloaded file failed integrity check.")
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror("Update failed", str(exc)))
                return

            def launch():
                save_config(DB_PATH, {"last_update_check": date.today().isoformat()})
                self._run_installer_and_exit(temp_path)

            self.root.after(0, launch)

        threading.Thread(target=worker, daemon=True).start()

    def _sha256_file(self, path: str) -> str:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()

    def _run_installer_and_exit(self, installer_path: str) -> None:
        try:
            subprocess.Popen([installer_path], shell=False)
        except Exception:
            try:
                os.startfile(installer_path)
            except Exception as exc:
                messagebox.showerror("Update failed", str(exc))
                return
        self.on_manager_close()

    def _open_help_window(self, title: str, body: str, logo_path: str | None = None) -> None:
        win = tk.Toplevel(self.manager)
        win.title(title)
        win.geometry("640x520")
        win.minsize(520, 360)
        win.configure(bg=THEME["bg"])
        win.transient(self.manager)

        container = tk.Frame(win, bg=THEME["bg"])
        container.pack(fill="both", expand=True, padx=16, pady=16)

        if logo_path and os.path.exists(logo_path):
            try:
                img = tk.PhotoImage(file=logo_path)
                win._help_logo_ref = img
                tk.Label(container, image=img, bg=THEME["bg"]).pack(anchor="center", pady=(0, 8))
            except tk.TclError:
                pass

        tk.Label(
            container,
            text=title,
            font=("Segoe UI", 14, "bold"),
            bg=THEME["bg"],
            fg=THEME["text"],
        ).pack(anchor="w", pady=(0, 8))

        text_frame = tk.Frame(container, bg=THEME["card"], highlightthickness=1, highlightbackground=THEME["border"])
        text_frame.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")

        text = tk.Text(
            text_frame,
            wrap="word",
            yscrollcommand=scrollbar.set,
            bg=THEME["card"],
            fg=THEME["text"],
            relief="flat",
            highlightthickness=0,
            padx=10,
            pady=8,
            font=("Segoe UI", 10),
        )
        text.pack(fill="both", expand=True)
        scrollbar.config(command=text.yview)
        text.insert("1.0", body)
        text.configure(state="disabled")

        btn_row = tk.Frame(container, bg=THEME["bg"])
        btn_row.pack(fill="x", pady=(10, 0))
        close_btn = self.make_button(btn_row, text="Close", command=win.destroy, kind="secondary")
        close_btn.pack(side="right")

    def _open_formatted_window(
        self,
        title: str,
        blocks: list[tuple[str, str]],
        logo_path: str | None = None,
    ) -> None:
        win = tk.Toplevel(self.manager)
        win.title(title)
        win.geometry("640x520")
        win.minsize(520, 360)
        win.configure(bg=THEME["bg"])
        win.transient(self.manager)

        container = tk.Frame(win, bg=THEME["bg"])
        container.pack(fill="both", expand=True, padx=18, pady=16)

        if logo_path and os.path.exists(logo_path):
            try:
                img = tk.PhotoImage(file=logo_path)
                win._help_logo_ref = img
                tk.Label(container, image=img, bg=THEME["bg"]).pack(anchor="center", pady=(0, 10))
            except tk.TclError:
                pass

        tk.Label(
            container,
            text=title,
            font=("Segoe UI", 14, "bold"),
            bg=THEME["bg"],
            fg=THEME["text"],
        ).pack(anchor="w", pady=(0, 8))

        text_frame = tk.Frame(container, bg=THEME["card"], highlightthickness=1, highlightbackground=THEME["border"])
        text_frame.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")

        text = tk.Text(
            text_frame,
            wrap="word",
            yscrollcommand=scrollbar.set,
            bg=THEME["card"],
            fg=THEME["text"],
            relief="flat",
            highlightthickness=0,
            padx=12,
            pady=10,
            font=("Segoe UI", 10),
        )
        text.pack(fill="both", expand=True)
        scrollbar.config(command=text.yview)

        text.tag_configure("heading", font=("Segoe UI", 11, "bold"), spacing1=8, spacing3=4)
        text.tag_configure("body", font=("Segoe UI", 10), spacing1=2, spacing3=8)
        text.tag_configure("meta", font=("Segoe UI", 10), spacing1=2, spacing3=6)

        for tag, content in blocks:
            text.insert("end", content, tag)
            if not content.endswith("\n"):
                text.insert("end", "\n")
            text.insert("end", "\n")

        text.configure(state="disabled")

        btn_row = tk.Frame(container, bg=THEME["bg"])
        btn_row.pack(fill="x", pady=(10, 0))
        close_btn = self.make_button(btn_row, text="Close", command=win.destroy, kind="secondary")
        close_btn.pack(side="right")

    def open_guide(self) -> None:
        blocks = [
            ("heading", "Getting Started"),
            ("body", "1) Click 'Open Editor' to add or edit words."),
            ("body", "2) Enter a word, meaning/example, and optionally an image."),
            ("body", "3) Click 'Add word' to save. The list updates immediately."),
            ("body", "4) Use 'Edit selected' or 'Remove selected' to manage entries."),
            ("heading", "Daily Learning"),
            ("body", "1) Set your interval and words per day in the Editor header."),
            ("body", "2) Click 'Apply changes' to update rotation settings."),
            ("body", "3) Use Start/Stop/Notify from the main window to control prompts."),
            ("heading", "Backup & Share"),
            ("body", "1) Use File > Export .vb to share your vocabulary."),
            ("body", "2) Use File > Import .vb to restore entries on another device."),
            ("body", "3) Use File > Create Backup to save snapshots automatically."),
            ("heading", "Tips"),
            ("body", "- Search words using the search box."),
            ("body", "- Ctrl+Z undoes recent changes; F1 opens this guide."),
        ]
        self._open_formatted_window("Guide", blocks)

    def open_support(self) -> None:
        blocks = [
            ("heading", "Support VocabBuddy"),
            ("body", "Need help or have suggestions?"),
            ("body", "- Use the Help menu to review the Guide and About info."),
            ("body", "- If something isn't working, try restarting the app first."),
            ("body", "- Export your words before major changes or updates."),
            ("heading", "Usage Tips"),
            ("body", "- Keep images small for faster notifications."),
            ("body", "- Use meaningful examples to improve recall."),
            ("body", "- Set a comfortable interval for focused learning."),
            ("heading", "Contact"),
            ("body", "- For support, share a screenshot and a short description of the issue."),
        ]
        self._open_formatted_window("Support VocabBuddy", blocks)

    def open_about(self) -> None:
        blocks = [
            ("heading", "Description"),
            (
                "body",
                "VocabBuddy® is a smart reminder-based vocabulary application designed to help users remember words through regular and timely notifications.",
            ),
            ("heading", "Goal"),
            (
                "body",
                "The goal of VocabBuddy is simple: to keep reminding you of your vocabulary so you never forget it.",
            ),
            (
                "body",
                "VocabBuddy helps turn repetition into an easy daily habit, making vocabulary learning natural and effective.",
            ),
            ("heading", "Author"),
            ("meta", "Created by: Ayoub Eddarif"),
            ("heading", "Contact"),
            ("meta", "Contact: +212 6 48 34 40 89"),
            ("heading", "Version"),
            ("meta", f"Version: {CURRENT_VERSION}"),
            ("heading", "Platform"),
            ("meta", "Platform: Python Desktop Application"),
        ]
        self._open_formatted_window("About VocabBuddy", blocks, logo_path=BODY_LOGO_PATH)

    def _init_text_undo_bindings(self) -> None:
        if self._text_undo_initialized:
            return
        self.word_entry.bind("<KeyRelease>", self._on_entry_modified)
        self.word_entry.bind("<FocusOut>", self._on_entry_modified)
        self.meaning_entry.bind("<<Modified>>", self._on_text_modified)
        self._text_undo_initialized = True
        self._reset_text_undo_stacks()

    def _reset_text_undo_stacks(self) -> None:
        self._text_undo_stacks[self.word_entry] = [self.word_entry.get()]
        self._text_undo_stacks[self.meaning_entry] = [
            self.meaning_entry.get("1.0", "end-1c")
        ]

    def _on_entry_modified(self, event) -> None:
        self._record_text_state(event.widget)

    def _on_text_modified(self, event) -> None:
        widget = event.widget
        if widget in self._text_undo_suspended:
            widget.edit_modified(False)
            return
        if widget.edit_modified():
            self._record_text_state(widget)
            widget.edit_modified(False)

    def _record_text_state(self, widget: tk.Widget) -> None:
        if widget in self._text_undo_suspended:
            return
        if isinstance(widget, tk.Entry):
            value = widget.get()
        elif isinstance(widget, tk.Text):
            value = widget.get("1.0", "end-1c")
        else:
            return
        stack = self._text_undo_stacks.setdefault(widget, [value])
        if stack and stack[-1] == value:
            return
        stack.append(value)
        if len(stack) > 200:
            stack.pop(0)

    def _undo_text_widget(self, widget: tk.Widget) -> bool:
        stack = self._text_undo_stacks.get(widget)
        if not stack or len(stack) <= 1:
            return False
        stack.pop()
        value = stack[-1]
        self._text_undo_suspended.add(widget)
        try:
            if isinstance(widget, tk.Entry):
                widget.delete(0, tk.END)
                widget.insert(0, value)
            elif isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
                widget.insert("1.0", value)
        finally:
            self._text_undo_suspended.discard(widget)
        return True

    def _write_words_vb(self, path: str, words: list[dict[str, str]]) -> None:
        payload = {"version": 1, "words": []}
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for idx, item in enumerate(words):
                word = item.get("word", "")
                meaning = item.get("meaning", "")
                image_path = item.get("image", "")
                image_blob = item.get("image_blob")
                image_name = ""
                if image_blob is None and image_path and os.path.exists(image_path):
                    try:
                        with open(image_path, "rb") as img_f:
                            image_blob = img_f.read()
                    except OSError:
                        image_blob = None
                if image_blob:
                    ext = os.path.splitext(image_path)[1].lower() if image_path else ""
                    if not ext:
                        ext = ".png"
                    image_name = f"images/{idx}_{uuid.uuid4().hex}{ext}"
                    zf.writestr(image_name, image_blob)
                payload["words"].append(
                    {
                        "word": word,
                        "meaning": meaning,
                        "image_name": image_name,
                    }
                )
            zf.writestr("words.json", json.dumps(payload, ensure_ascii=True))

    def export_words_vb(self) -> None:
        if not self.words:
            messagebox.showinfo("Export", "No words to export.")
            return
        path = filedialog.asksaveasfilename(
            title="Export vocabulary",
            defaultextension=".vb",
            filetypes=[("VocabBuddy backup", "*.vb")],
        )
        if not path:
            return
        try:
            self._write_words_vb(path, self.words)
            messagebox.showinfo("Export", "Export completed successfully.")
        except (OSError, zipfile.BadZipFile) as exc:
            messagebox.showerror("Export failed", f"Unable to export file.\n{exc}")

    def _load_words_from_vb(self, path: str) -> list[dict[str, str]]:
        with zipfile.ZipFile(path, "r") as zf:
            if "words.json" not in zf.namelist():
                raise ValueError("Missing words.json")
            raw = zf.read("words.json").decode("utf-8")
            data = json.loads(raw)
            entries = data.get("words", [])
            imported: list[dict[str, str]] = []
            for entry in entries:
                word = (entry.get("word") or "").strip()
                meaning = entry.get("meaning") or ""
                image_name = entry.get("image_name") or ""
                image_blob = None
                if image_name:
                    try:
                        image_blob = zf.read(image_name)
                    except KeyError:
                        image_blob = None
                if word:
                    imported.append(
                        {
                            "word": word,
                            "meaning": meaning,
                            "image": "",
                            "image_blob": image_blob,
                        }
                    )
        return imported

    def _apply_imported_words(self, imported: list[dict[str, str]]) -> None:
        self._replace_words(imported)

    def import_words_vb(self) -> None:
        path = filedialog.askopenfilename(
            title="Import vocabulary",
            filetypes=[("VocabBuddy backup", "*.vb")],
        )
        if not path:
            return
        if self.words:
            if not messagebox.askyesno(
                "Replace words?",
                "Import will replace your current words. Continue?",
            ):
                return
        self._push_undo_state()
        try:
            imported = self._load_words_from_vb(path)
        except (OSError, zipfile.BadZipFile, ValueError, json.JSONDecodeError) as exc:
            messagebox.showerror("Import failed", f"Unable to import file.\n{exc}")
            return
        self._apply_imported_words(imported)
        messagebox.showinfo("Import", "Import completed successfully.")

    def _backup_filename(self) -> str:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return os.path.join(BACKUP_DIR, f"backup_{timestamp}.vb")

    def _schedule_next_backup(self) -> None:
        if self.backup_timer_id is not None:
            self.root.after_cancel(self.backup_timer_id)
        self.backup_timer_id = self.root.after(self.backup_interval_seconds * 1000, self._auto_backup_tick)

    def _auto_backup_tick(self) -> None:
        self._create_backup(silent=True)
        self._schedule_next_backup()

    def _create_backup(self, silent: bool = False) -> None:
        if not self.words:
            if not silent:
                messagebox.showinfo("Backup", "No words to back up.")
            return
        path = self._backup_filename()
        try:
            self._write_words_vb(path, self.words)
            if not silent:
                messagebox.showinfo("Backup", f"Backup created:\n{path}")
        except (OSError, zipfile.BadZipFile) as exc:
            if not silent:
                messagebox.showerror("Backup failed", f"Unable to create backup.\n{exc}")

    def create_backup(self) -> None:
        self._create_backup(silent=False)
        self._schedule_next_backup()

    def revert_to_backup(self) -> None:
        path = filedialog.askopenfilename(
            title="Revert to backup",
            initialdir=BACKUP_DIR,
            filetypes=[("VocabBuddy backup", "*.vb")],
        )
        if not path:
            return
        if self.words:
            if not messagebox.askyesno(
                "Replace words?",
                "Reverting will replace your current words. Continue?",
            ):
                return
        self._push_undo_state()
        try:
            imported = self._load_words_from_vb(path)
        except (OSError, zipfile.BadZipFile, ValueError, json.JSONDecodeError) as exc:
            messagebox.showerror("Revert failed", f"Unable to restore backup.\n{exc}")
            return
        self._apply_imported_words(imported)
        messagebox.showinfo("Revert", "Backup restored successfully.")

    def open_editor(self, reset_fields: bool = True) -> None:
        if self.editor is not None and self.editor.winfo_exists():
            self.editor.deiconify()
            self.editor.lift()
            if reset_fields:
                self._reset_editor_fields()
            return

        self.editor = tk.Toplevel(self.root)
        self.editor.title("VocabBuddy Editor")
        self.editor.geometry("840x640")
        self.editor.minsize(840, 640)
        self.editor.resizable(True, True)
        self.editor.configure(bg=THEME["bg"])
        self.apply_app_icon()

        header = tk.Frame(self.editor, bg=THEME["bg"])
        header.pack(fill="x", padx=24, pady=(16, 6))
        header.columnconfigure(0, weight=1)
        header.columnconfigure(1, weight=0)

        tk.Label(
            header,
            text="Editor",
            font=("Segoe UI", 18, "bold"),
            bg=THEME["bg"],
            fg=THEME["text"],
        ).grid(row=0, column=0, sticky="w")

        settings = tk.Frame(header, bg=THEME["bg"])
        settings.grid(row=0, column=1, sticky="e")

        tk.Label(settings, text="Interval (sec):", bg=THEME["bg"], fg=THEME["text"]).grid(
            row=0, column=0, sticky="e", padx=6, pady=2
        )
        self.interval_entry = tk.Entry(
            settings,
            width=8,
            bg=THEME["card"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
        )
        self.interval_entry.insert(0, str(self.interval_seconds))
        self.interval_entry.grid(row=0, column=1, padx=6, pady=2)

        tk.Label(settings, text="Words/day:", bg=THEME["bg"], fg=THEME["text"]).grid(
            row=1, column=0, sticky="e", padx=6, pady=2
        )
        self.words_per_day_entry = tk.Entry(
            settings,
            width=8,
            bg=THEME["card"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
        )
        self.words_per_day_entry.insert(0, str(self.words_per_day))
        self.words_per_day_entry.grid(row=1, column=1, padx=6, pady=2)

        card_canvas = tk.Canvas(self.editor, bg=THEME["bg"], highlightthickness=0, width=720, height=440)
        card_canvas.pack(padx=24, pady=(8, 16))
        card = self._card_container(card_canvas, width=720, height=440, radius=16)
        card_canvas.create_window((0, 0), window=card, anchor="nw")

        form = tk.Frame(card, bg=THEME["card"])
        form.pack(pady=16, padx=18, fill="x")
        form.columnconfigure(0, weight=1)

        tk.Label(form, text="New word", bg=THEME["card"], fg=THEME["text"]).grid(
            row=0, column=0, sticky="w", padx=6, pady=(0, 6)
        )
        self.word_entry = tk.Entry(
            form,
            bg=THEME["card"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
        )
        self.word_entry.grid(row=1, column=0, padx=6, pady=(0, 12), sticky="we")

        tk.Label(form, text="Meaning / Example", bg=THEME["card"], fg=THEME["text"]).grid(
            row=2, column=0, sticky="w", padx=6, pady=(0, 6)
        )
        self.meaning_entry = tk.Text(
            form,
            height=5,
            bg=THEME["card"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            wrap="word",
        )
        self.meaning_entry.grid(row=3, column=0, padx=6, pady=(0, 12), sticky="we")

        tk.Label(form, text="Image (PNG)", bg=THEME["card"], fg=THEME["text"]).grid(
            row=4, column=0, sticky="w", padx=6, pady=(0, 6)
        )
        image_row = tk.Frame(form, bg=THEME["card"])
        image_row.grid(row=5, column=0, padx=6, pady=(0, 12), sticky="we")
        image_row.columnconfigure(0, weight=1)

        self.image_entry = tk.Entry(
            image_row,
            bg=THEME["card"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            state="readonly",
        )
        self.image_entry.grid(row=0, column=0, sticky="we", padx=(0, 8))
        self.image_browse_btn = self.make_button(image_row, text="Upload", command=self.browse_image, kind="secondary")
        self.image_browse_btn.configure(padx=10, pady=4, font=("Segoe UI", 9), radius=12)
        self.image_browse_btn.grid(row=0, column=1, sticky="e", padx=(8,0))

        actions = tk.Frame(card, bg=THEME["card"])
        actions.pack(padx=18, pady=(4, 16), fill="x")
        actions.columnconfigure(0, weight=1, uniform="actions")
        actions.columnconfigure(1, weight=1, uniform="actions")

        self.add_btn = self.make_button(actions, text="Add word", command=self.add_word, kind="primary")
        self.add_btn.grid(row=0, column=0, padx=8, pady=4, sticky="ew")

        self.apply_all_btn = self.make_button(actions, text="Apply changes", command=self.apply_all_changes, kind="danger")
        self.apply_all_btn.grid(row=0, column=1, padx=8, pady=4, sticky="ew")

        # Save all button removed per request

        self.editor.protocol("WM_DELETE_WINDOW", self.editor.withdraw)
        self._apply_button_icons()
        self._init_text_undo_bindings()
        self._reset_editor_fields()

    def _build_menu_bar(self) -> None:
        menu_bar = tk.Menu(self.manager)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Import .vb", command=self.import_words_vb)
        file_menu.add_command(label="Export .vb", command=self.export_words_vb)
        file_menu.add_separator()
        file_menu.add_command(label="Create Backup", command=self.create_backup)
        file_menu.add_command(label="Revert to Backup", command=self.revert_to_backup)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_manager_close)
        menu_bar.add_cascade(label="File", menu=file_menu)
        edit_menu = tk.Menu(menu_bar, tearoff=0)
        edit_menu.add_command(label="Undo", command=self.undo_last_action, accelerator="Ctrl+Z")
        menu_bar.add_cascade(label="Edit", menu=edit_menu)
        view_menu = tk.Menu(menu_bar, tearoff=0)
        view_menu.add_command(label="Full Screen", command=self.toggle_fullscreen)
        view_menu.add_separator()
        view_menu.add_command(label="Zoom In", command=self.zoom_in)
        view_menu.add_command(label="Zoom Out", command=self.zoom_out)
        view_menu.add_command(label="Reset Zoom", command=self.reset_zoom)
        menu_bar.add_cascade(label="View", menu=view_menu)
        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="Guide", command=self.open_guide, accelerator="F1")
        help_menu.add_command(label="Support VocabBuddy", command=self.open_support)
        help_menu.add_command(label="About VocabBuddy", command=self.open_about)
        help_menu.add_separator()
        help_menu.add_command(label="Check for updates", command=lambda: self.check_for_updates(manual=True))
        menu_bar.add_cascade(label="Help", menu=help_menu)
        self.manager.config(menu=menu_bar)
        self.manager.bind_all("<Control-z>", self._handle_undo_event)
        self.manager.bind_all("<Control-Z>", self._handle_undo_event)
        self.manager.bind_all("<Control-MouseWheel>", self._handle_ctrl_mousewheel)
        self.manager.bind_all("<Control-0>", self._handle_reset_zoom_event)
        self.manager.bind_all("<Control-Key-0>", self._handle_reset_zoom_event)
        self.manager.bind_all("<F1>", lambda _e: self.open_guide())

    def apply_interval(self) -> None:
        raw = self.interval_entry.get().strip()
        try:
            seconds = int(raw)
            if seconds <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Invalid interval", "Enter a whole number greater than 0.")
            return

        self.interval_seconds = seconds
        self.remaining_seconds = self.interval_seconds
        self.notify_duration_seconds = NOTIFY_DURATION_SECONDS
        save_config(DB_PATH, {"interval_seconds": self.interval_seconds})
        self.update_ui()
        self.update_countdown_label(running=self.timer_id is not None)
        if self.timer_id is not None:
            self.root.after_cancel(self.timer_id)
            self.timer_id = self.root.after(self.interval_seconds * 1000, self._tick)
            self.start_countdown()

    def apply_words_per_day(self) -> None:
        if not hasattr(self, "words_per_day_entry"):
            return
        raw = self.words_per_day_entry.get().strip()
        try:
            count = int(raw)
            if count <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Invalid number", "Enter a whole number greater than 0.")
            return
        self.words_per_day = count
        save_config(DB_PATH, {"words_per_day": self.words_per_day})
        self.index = 0
        self.active_date = None
        self.active_words = []
        self.refresh_listbox()
        active = self.get_active_words()
        if active:
            self.word_var.set(active[self.index]["word"])
            self.meaning_var.set(active[self.index].get("meaning", ""))
        self.update_ui()

    def edit_selected(self) -> None:
        sel = self.words_list.curselection()
        if not sel:
            return
        idx = sel[0]
        active = self.get_filtered_words()
        if not active:
            return
        item = active[idx]
        try:
            real_index = self.words.index(item)
        except ValueError:
            word = item.get("word", "")
            meaning = item.get("meaning", "")
            image = item.get("image", "")
            real_index = next(
                (i for i, w in enumerate(self.words)
                 if w.get("word", "") == word
                 and w.get("meaning", "") == meaning
                 and w.get("image", "") == image),
                None,
            )
        if real_index is None:
            return
        self.edit_index = real_index
        self.open_editor(reset_fields=False)
        self.word_entry.delete(0, tk.END)
        self.word_entry.insert(0, self.words[real_index].get("word", ""))
        self.meaning_entry.delete("1.0", tk.END)
        self.meaning_entry.insert("1.0", self.words[real_index].get("meaning", ""))
        self.image_entry.configure(state="normal")
        self.image_entry.delete(0, tk.END)
        self.image_entry.insert(0, self.words[real_index].get("image", ""))
        self.image_entry.configure(state="readonly")
        self.add_btn.configure(text="Update word")
        self._reset_text_undo_stacks()

    def _reset_editor_fields(self) -> None:
        if not hasattr(self, "word_entry"):
            return
        self.edit_index = None
        self.word_entry.delete(0, tk.END)
        self.meaning_entry.delete("1.0", tk.END)
        self.image_entry.configure(state="normal")
        self.image_entry.delete(0, tk.END)
        self.image_entry.configure(state="readonly")
        if hasattr(self, "add_btn"):
            self.add_btn.configure(text="Add word")
        if self._text_undo_initialized:
            self._reset_text_undo_stacks()

    def _ensure_today_includes_new_word(self, new_index: int) -> None:
        today = date.today().isoformat()
        if self.active_date != today:
            self.active_words = self.get_active_words()
        if not self.active_words:
            self.active_words = []
            self.active_date = today
        if new_index < 0 or new_index >= len(self.words):
            return
        new_word = self.words[new_index]
        if new_word in self.active_words:
            return
        if len(self.active_words) < self.words_per_day:
            self.active_words.append(new_word)
        else:
            self.active_words[-1] = new_word
        self.active_date = today

    def _persist_today_words(self) -> None:
        if not self.active_words:
            save_config(DB_PATH, {"cycle_order": "", "cycle_index": 0, "daily_set": "", "cycle_date": ""})
            return
        indices = [self.words.index(w) for w in self.active_words if w in self.words]
        cfg = load_config(DB_PATH)
        order_raw = cfg.get("cycle_order", "")
        try:
            order = json.loads(order_raw) if order_raw else []
        except json.JSONDecodeError:
            order = []
        if not order:
            order = list(range(len(self.words)))
            random.shuffle(order)
        else:
            for i in range(len(self.words)):
                if i not in order:
                    order.append(i)
        save_config(
            DB_PATH,
            {
                "daily_set": json.dumps(indices),
                "cycle_date": self.active_date or date.today().isoformat(),
                "cycle_order": json.dumps(order),
            },
        )

    def apply_all_changes(self) -> None:
        self.apply_interval()
        self.apply_words_per_day()
        self.update_ui()

    def save_all(self) -> None:
        save_words(WORDS_FILE, self.words)
        save_config(DB_PATH, {"words_per_day": self.words_per_day})
        self.update_ui()
        messagebox.showinfo("Saved", "All changes have been saved.")

    def refresh_listbox(self) -> None:
        self.words_list.delete(0, tk.END)
        filtered = self.get_filtered_words()
        for item in filtered:
            self.words_list.insert(tk.END, item.get("word", ""))
        if hasattr(self, "count_var"):
            self.count_var.set(f"{len(filtered)} words")

    def get_active_words(self) -> list[dict[str, str]]:
        if not self.words:
            return []
        today = date.today().isoformat()
        if self.active_date != today or not self.active_words:
            cfg = load_config(DB_PATH)
            if cfg.get("cycle_date") == today and cfg.get("daily_set"):
                try:
                    picked = json.loads(cfg.get("daily_set", "[]"))
                except json.JSONDecodeError:
                    picked = []
                if picked:
                    self.active_words = [self.words[i] for i in picked if i < len(self.words)]
                    self.active_date = today
                    return self.active_words
            self.update_daily_words(today)
        return self.active_words

    def get_filtered_words(self) -> list[dict[str, str]]:
        active = self.get_active_words()
        if not hasattr(self, "search_var"):
            return active
        q = self.search_var.get().strip().lower()
        if q == "search words...":
            q = ""
        if not q:
            return active
        return [w for w in active if q in w.get("word", "").lower()]

    def update_daily_words(self, today: str) -> None:
        cfg = load_config(DB_PATH)
        order_raw = cfg.get("cycle_order", "")
        index_raw = cfg.get("cycle_index", "0")
        try:
            order = json.loads(order_raw) if order_raw else []
        except json.JSONDecodeError:
            order = []
        try:
            idx = int(index_raw)
        except ValueError:
            idx = 0

        total = len(self.words)
        if not order or sorted(order) != list(range(total)):
            order = list(range(total))
            random.shuffle(order)
            idx = 0

        n = max(1, int(self.words_per_day))
        if idx >= total:
            order = list(range(total))
            random.shuffle(order)
            idx = 0

        end = min(idx + n, total)
        picked = order[idx:end]
        idx = end

        self.active_words = [self.words[i] for i in picked]
        self.active_date = today
        save_config(
            DB_PATH,
            {
                "cycle_order": json.dumps(order),
                "cycle_index": idx,
                "cycle_date": today,
                "daily_set": json.dumps(picked),
            },
        )

    def remove_selected(self) -> None:
        sel = self.words_list.curselection()
        if not sel:
            return
        idx = sel[0]
        active = self.get_filtered_words()
        if not active:
            return
        real_item = active[idx]
        self._push_undo_state()
        if real_item in self.words:
            self.words.remove(real_item)
        else:
            word = real_item.get("word", "")
            meaning = real_item.get("meaning", "")
            image = real_item.get("image", "")
            self.words = [
                w
                for w in self.words
                if not (
                    w.get("word", "") == word
                    and w.get("meaning", "") == meaning
                    and w.get("image", "") == image
                )
            ]
        save_words(WORDS_FILE, self.words)
        save_config(DB_PATH, {"cycle_order": "", "cycle_index": 0, "daily_set": "", "cycle_date": ""})
        self.active_date = None
        self.active_words = []
        self.update_ui()
        if self.index >= len(self.get_active_words()):
            self.index = 0
        self.refresh_listbox()
        active = self.get_active_words()
        if active:
            self.word_var.set(active[self.index]["word"])
            self.meaning_var.set(active[self.index].get("meaning", ""))
            self.page_title_var.set(active[self.index]["word"])
        else:
            self.word_var.set("No words found")
            self.meaning_var.set("")
            self.page_title_var.set("No words found")
        self.status_var.set(
            f"Interval: {self.interval_seconds} seconds | Total words: {len(self.words)} | Per day: {self.words_per_day}"
        )

    def clear_all(self) -> None:
        if not self.words:
            return
        if messagebox.askyesno("Clear all", "Remove all words?"):
            self._push_undo_state()
            self.words.clear()
            self.index = 0
            self.refresh_listbox()
            self.word_var.set("No words found")
            self.meaning_var.set("")
            self.page_title_var.set("No words found")
            self.status_var.set(
                f"Interval: {self.interval_seconds} seconds | Total words: {len(self.words)} | Per day: {self.words_per_day}"
            )
            save_words(WORDS_FILE, self.words)
            save_config(DB_PATH, {"cycle_order": "", "cycle_index": 0, "daily_set": "", "cycle_date": ""})
            self.active_date = None
            self.active_words = []

    def show_notification(self, word: str, meaning: str, image_path: str, image_blob: bytes | None = None) -> None:
        for w in self.notify.winfo_children():
            w.destroy()
        frame = tk.Frame(self.notify, bg=THEME["card"], bd=1, relief="solid")
        frame.pack(fill="both", expand=True)

        text_frame = tk.Frame(frame, bg=THEME["card"])
        text_frame.pack(fill="x", padx=14, pady=(8, 4))

        tk.Label(
            text_frame,
            text=word,
            font=("Segoe UI", 20, "bold"),
            bg=THEME["card"],
            fg=THEME["text"],
        ).pack()

        tk.Message(
            text_frame,
            text=meaning or "",
            font=("Segoe UI", 11),
            bg=THEME["card"],
            fg=THEME["muted"],
            width=NOTIFY_MAX_WIDTH - 30,
            justify="center",
        ).pack(pady=(4, 0))

        image_frame = tk.Frame(frame, bg=THEME["card"])
        image_frame.pack(fill="both", expand=True, padx=10, pady=(4, 8))

        self.notify_image_ref = None
        image_label = tk.Label(image_frame, bg=THEME["card"])
        image_label.pack()

        if image_blob or (image_path and os.path.exists(image_path)):
            try:
                self.notify.update_idletasks()
                text_h = text_frame.winfo_reqheight()
                available_h = NOTIFY_MAX_HEIGHT - text_h - 40
                max_h = max(80, min(available_h, 200))
                max_w = 180
                max_h = min(max_h, 180)
                if image_blob:
                    img = Image.open(io.BytesIO(image_blob)).convert("RGBA")
                else:
                    img = Image.open(image_path).convert("RGBA")
                img.thumbnail((max_w, max_h), Image.LANCZOS)
                self.notify_image_ref = ImageTk.PhotoImage(img)
                image_label.configure(image=self.notify_image_ref)
            except Exception:
                image_label.configure(
                    text="(Image not supported)",
                    font=("Segoe UI", 9),
                    fg=THEME["muted"],
                )
        elif image_path:
            image_label.configure(
                text="(Image not found)",
                font=("Segoe UI", 9),
                fg=THEME["muted"],
            )

        self.notify.update_idletasks()
        req_w = frame.winfo_reqwidth()
        req_h = frame.winfo_reqheight()
        w = max(NOTIFY_MIN_WIDTH, min(req_w, NOTIFY_MAX_WIDTH))
        h = max(NOTIFY_MIN_HEIGHT, min(req_h, NOTIFY_MAX_HEIGHT))
        x = self.notify.winfo_screenwidth() - w - 20
        y = self.notify.winfo_screenheight() - h - 60
        self.notify.geometry(f"{w}x{h}+{x}+{y}")
        self.notify.deiconify()

        if self.notify_hide_id is not None:
            self.root.after_cancel(self.notify_hide_id)
        self.notify_hide_id = self.root.after(self.notify_duration_seconds * 1000, self.notify.withdraw)

    def start_countdown(self) -> None:
        if self.countdown_id is not None:
            self.root.after_cancel(self.countdown_id)
        if self.remaining_seconds <= 0:
            self.remaining_seconds = self.interval_seconds
        self.update_countdown_label(running=True)
        self.countdown_id = self.root.after(1000, self.countdown_tick)

    def countdown_tick(self) -> None:
        if self.timer_id is None:
            return
        if self.remaining_seconds > 0:
            self.remaining_seconds -= 1
        self.update_countdown_label(running=True)
        self.countdown_id = self.root.after(1000, self.countdown_tick)

    def update_countdown_label(self, running: bool) -> None:
        if not running:
            self.next_word_var.set("Next word: stopped")
            return
        self.next_word_var.set(f"Next word: {self.remaining_seconds}s")

    def on_manager_close(self) -> None:
        self.stop()
        if self.backup_timer_id is not None:
            try:
                self.root.after_cancel(self.backup_timer_id)
            except Exception:
                pass
            self.backup_timer_id = None
        release_single_instance()
        self.root.quit()


def main() -> None:
    if not ensure_single_instance():
        return
    atexit.register(release_single_instance)
    words = load_words(WORDS_FILE)
    root = tk.Tk()
    app = WordRotatorApp(root, words)
    root.mainloop()
    release_single_instance()


if __name__ == "__main__":
    main()

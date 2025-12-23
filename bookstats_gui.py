
import json
import os
import random
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Default Library column widths (used for reset + initial layout)
DEFAULT_LIBRARY_COLUMN_WIDTHS = {
    "Title": 300,
    "Author": 170,
    "Year": 70,
    "Pages": 70,
    "Rating": 70,
    "Format": 150,
    "Tags": 170,
    "Collections": 220,
    "Genres": 220,
}

# Optional stats charts (recommended)
try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    _HAS_MPL = True
except Exception:
    Figure = None
    FigureCanvasTkAgg = None
    _HAS_MPL = False




# Optional cover support (recommended)
# Requires: pip install pillow requests
try:
    from PIL import Image, ImageTk  # type: ignore
    _HAS_PIL = True
except Exception:
    Image = None
    ImageTk = None
    _HAS_PIL = False

try:
    # cover_cache.py should be in the same folder as this script
    from cover_cache import CoverCache, choose_best_isbn, normalize_isbn  # type: ignore
    _HAS_COVER_CACHE = True
except Exception:
    CoverCache = None  # type: ignore
    choose_best_isbn = None  # type: ignore
    normalize_isbn = None  # type: ignore
    _HAS_COVER_CACHE = False


# ----------------------------
# Settings (persisted to JSON)
# ----------------------------

APP_NAME = "BookStats"
APP_AUTHOR = "KyleCarroll"

_DEFAULT_SETTINGS = {
    "covers_enabled": True,
    "auto_fill_summary": True,
    "auto_load_last_file": True,
    "last_opened_path": "",
    # Library columns: visibility + order (Treeview displaycolumns)
    # If empty, defaults to all columns in the app's default order.
    "library_visible_columns": [],
    "library_column_order": [],
}

def _settings_path(settings_dir: str) -> str:
    return os.path.join(settings_dir, "bookstats_settings.json")
def load_settings(settings_dir: str) -> Dict[str, Any]:
    path = _settings_path(settings_dir)
    if not os.path.exists(path):
        return dict(_DEFAULT_SETTINGS)
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return dict(_DEFAULT_SETTINGS)
        out = dict(_DEFAULT_SETTINGS)
        out.update(data)
        return out
    except Exception:
        return dict(_DEFAULT_SETTINGS)

def save_settings(settings_dir: str, settings: Dict[str, Any]) -> None:
    path = _settings_path(settings_dir)
    tmp = path + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)
        os.replace(tmp, path)
    except Exception:
        # Best-effort; avoid crashing the app if disk is read-only, etc.
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass

# ----------------------------
# Data model + parsing helpers
# ----------------------------

def get_app_dirs(app_dir: str) -> tuple[str, str]:
    """
    Return (data_dir, cache_dir) in a location that is writable and persistent.

    - data_dir: user-specific persistent data (settings, last-opened file)
    - cache_dir: user-specific cache (cover images)

    Prefers platformdirs when available. If not, uses OS-specific user folders.
    If a legacy settings file or legacy .cache folder exists next to the script,
    continue using those to avoid breaking existing installs.
    """
    # Legacy (older builds): store next to the script folder.
    legacy_settings = os.path.join(app_dir, "bookstats_settings.json")
    legacy_cache = os.path.join(app_dir, ".cache")

    try:
        from platformdirs import user_data_dir, user_cache_dir  # type: ignore
        data_dir = user_data_dir(APP_NAME, APP_AUTHOR)
        cache_dir = user_cache_dir(APP_NAME, APP_AUTHOR)
    except Exception:
        # If we already have a legacy install, keep using it.
        if os.path.exists(legacy_settings) or os.path.isdir(legacy_cache):
            data_dir = app_dir
            cache_dir = legacy_cache
        else:
            home = os.path.expanduser("~")

            if sys.platform.startswith("win"):
                appdata = os.environ.get("APPDATA") or os.environ.get("LOCALAPPDATA") or home
                localappdata = os.environ.get("LOCALAPPDATA") or appdata
                data_dir = os.path.join(appdata, APP_NAME)
                cache_dir = os.path.join(localappdata, APP_NAME, "Cache")
            elif sys.platform == "darwin":
                data_dir = os.path.join(home, "Library", "Application Support", APP_NAME)
                cache_dir = os.path.join(home, "Library", "Caches", APP_NAME)
            else:
                xdg_data = os.environ.get("XDG_DATA_HOME") or os.path.join(home, ".local", "share")
                xdg_cache = os.environ.get("XDG_CACHE_HOME") or os.path.join(home, ".cache")
                data_dir = os.path.join(xdg_data, APP_NAME)
                cache_dir = os.path.join(xdg_cache, APP_NAME)

    os.makedirs(data_dir, exist_ok=True)
    os.makedirs(cache_dir, exist_ok=True)
    return data_dir, cache_dir


def _as_list(v: Any) -> List[Any]:
    if v is None:
        return []
    if isinstance(v, list):
        return v
    if isinstance(v, dict):
        # LibraryThing sometimes uses dicts like {"0": "...", "2": "..."} for ISBN etc.
        items = list(v.items())
        items.sort(key=lambda x: str(x[0]))
        return [x[1] for x in items]
    return [v]


def _safe_float(v: Any) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        return float(v)
    except Exception:
        return None


def _safe_bool(v: Any) -> bool:
    """Best-effort boolean parse for JSON fields."""
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    if isinstance(v, (int, float)):
        return v != 0
    s = str(v).strip().lower()
    if s in {"1", "true", "yes", "y", "t"}:
        return True
    if s in {"0", "false", "no", "n", "f", ""}:
        return False
    # Fallback: non-empty strings count as True
    return True


def _digits_to_int(v: Any) -> Optional[int]:
    if v is None:
        return None
    s = str(v)
    m = re.search(r"(\d+)", s)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def _parse_year(v: Any) -> Optional[int]:
    if v is None:
        return None
    s = str(v).strip()
    m = re.search(r"(\d{4})", s)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def _extract_best_isbn(raw: Dict[str, Any]) -> str:
    """
    Extract a usable ISBN (prefer ISBN-13) from LibraryThing export fields.

    LibraryThing commonly stores ISBN in `isbn` (and sometimes `originalisbn`) as:
      - a dict like {"isbn10": "...", "isbn13": "..."} OR {"0": "...", "2": "..."}
      - a list of strings
      - a single string
    """
    candidates: List[str] = []

    for field in ("isbn", "originalisbn"):
        v = raw.get(field)
        if v is None:
            continue
        if isinstance(v, dict):
            candidates.extend([str(x) for x in v.values() if x is not None])
        else:
            candidates.extend([str(x) for x in _as_list(v) if x is not None])

    candidates = [c for c in (str(x).strip() for x in candidates) if c]
    if not candidates:
        return ""

    if choose_best_isbn:
        try:
            return choose_best_isbn(*candidates)  # type: ignore
        except Exception:
            pass

    # Minimal fallback: prefer 13-digit ISBNs starting with 978/979
    cleaned = []
    for c in candidates:
        s = re.sub(r"[^0-9Xx]", "", c).upper()
        if len(s) in (10, 13):
            cleaned.append(s)

    for s in cleaned:
        if len(s) == 13 and (s.startswith("978") or s.startswith("979")):
            return s
    return cleaned[0] if cleaned else ""


def _parse_date_yyyy_mm_dd(v: Any) -> Optional[datetime]:
    if not v:
        return None
    s = str(v).strip()
    try:
        return datetime.strptime(s, "%Y-%m-%d")
    except Exception:
        return None


@dataclass
class Book:
    books_id: str
    title: str
    primaryauthor: str
    authors: List[str] = field(default_factory=list)
    collections: List[str] = field(default_factory=list)
    tags: List[str] = field(default_factory=list)
    formats: List[str] = field(default_factory=list)  # from LibraryThing "format" -> [{"text": "..."}]
    genre: List[str] = field(default_factory=list)
    series: List[str] = field(default_factory=list)
    rating: Optional[float] = None
    pages: Optional[int] = None
    year: Optional[int] = None
    dateread: Optional[str] = None
    entrydate: Optional[str] = None
    isbn: str = ""
    publication: str = ""
    summary: str = ""
    summary_checked: bool = False

    @property
    def display_author(self) -> str:
        # Prefer primaryauthor; if missing, fall back to first author in authors list
        if self.primaryauthor:
            return self.primaryauthor
        if self.authors:
            return self.authors[0]
        return ""

    @property
    def author_last(self) -> str:
        a = self.display_author.strip()
        if not a:
            return ""
        # Handle "Last, First"
        if "," in a:
            return a.split(",", 1)[0].strip().lower()
        # Otherwise "First Last"
        parts = a.split()
        return parts[-1].strip().lower() if parts else a.lower()

    @property
    def is_read(self) -> bool:
        return "Read" in self.collections or (self.dateread is not None and str(self.dateread).strip() != "")

    @property
    def is_unread(self) -> bool:
        return ("Unread" in self.collections) and not self.is_read

    @property
    def is_owned(self) -> bool:
        return "Owned" in self.collections

    @property
    def is_to_read(self) -> bool:
        # Use explicit LibraryThing collection (case-insensitive), e.g. "To read".
        # User requested this be driven by the JSON collection rather than derived.
        return any(str(c).strip().lower() in {"to read", "to-read", "toread"} for c in self.collections)

    def collections_str(self) -> str:
        return ", ".join(self.collections)

    def genre_str(self) -> str:
        return ", ".join(self.genre)

    def authors_str(self) -> str:
        return ", ".join(self.authors)

    def tags_str(self) -> str:
        return ", ".join(self.tags)

    def formats_str(self) -> str:
        return ", ".join(self.formats)

    def primary_format(self) -> str:
        return self.formats[0] if self.formats else ""


def parse_books(data: Dict[str, Any]) -> List[Book]:
    """
    Parse the LibraryThing JSON export (a dict keyed by books_id) into Book objects.

    Notes on known fields in your export:
      - collections: ["Owned","Unread"] etc.
      - tags: ["Good","Like New"] etc.
      - format: [{"text": "Hardcover"}] etc.
      - dateread: "YYYY-MM-DD"
      - entrydate: "YYYY-MM-DD"
    """
    books: List[Book] = []
    for key, raw in data.items():
        if not isinstance(raw, dict):
            continue

        books_id = str(raw.get("books_id") or key)
        title = str(raw.get("title") or "").strip()
        primaryauthor = str(raw.get("primaryauthor") or "").strip()

        author_objs = _as_list(raw.get("authors"))
        authors: List[str] = []
        for a in author_objs:
            if isinstance(a, dict):
                nm = (a.get("fl") or a.get("lf") or "").strip()
                if nm:
                    authors.append(nm)
            elif isinstance(a, str) and a.strip():
                authors.append(a.strip())

        collections = [str(x).strip() for x in _as_list(raw.get("collections")) if str(x).strip()]
        tags = [str(x).strip() for x in _as_list(raw.get("tags")) if str(x).strip()]
        genre = [str(x).strip() for x in _as_list(raw.get("genre")) if str(x).strip()]
        series = [str(x).strip() for x in _as_list(raw.get("series")) if str(x).strip()]

        # format looks like: [{"code": "...", "text": "Hardcover"}]
        formats: List[str] = []
        for fe in _as_list(raw.get("format")):
            if isinstance(fe, dict):
                t = str(fe.get("text") or "").strip()
                if t:
                    formats.append(t)
            elif isinstance(fe, str) and fe.strip():
                formats.append(fe.strip())
        # de-dupe while preserving order
        seen = set()
        formats = [x for x in formats if not (x.lower() in seen or seen.add(x.lower()))]

        rating = _safe_float(raw.get("rating"))
        pages = _digits_to_int(raw.get("pages"))
        year = _parse_year(raw.get("date"))

        isbn = _extract_best_isbn(raw)

        book = Book(
            books_id=books_id,
            title=title,
            primaryauthor=primaryauthor,
            authors=authors,
            collections=collections,
            tags=tags,
            formats=formats,
            genre=genre,
            series=series,
            rating=rating,
            pages=pages,
            year=year,
            isbn=isbn,
            dateread=str(raw.get("dateread") or "").strip() or None,
            entrydate=str(raw.get("entrydate") or "").strip() or None,
            publication=str(raw.get("publication") or "").strip(),
            summary=str(raw.get("summary") or "").strip(),
            summary_checked=_safe_bool(raw.get("summary_checked")),
        )
        books.append(book)

    # Default sort: author last name, then title
    books.sort(key=lambda b: (b.author_last, b.display_author.lower(), b.title.lower()))
    return books


# ----------------------------
# GUI App
# ----------------------------

class _TopTable:
    def __init__(self, frame: ttk.Frame, title: str):
        self.frame = ttk.LabelFrame(frame, text=title, padding=8)
        self.tree = ttk.Treeview(self.frame, columns=("Item", "Count"), show="headings", height=12)
        self.tree.heading("Item", text="Item")
        self.tree.heading("Count", text="Count")
        self.tree.column("Item", width=180)
        self.tree.column("Count", width=60, anchor="e")
        self.tree.pack(fill="both", expand=True)

    def grid(self, **kwargs):
        self.frame.grid(**kwargs)


class BookStatsApp(tk.Tk):
    def __init__(self, initial_path: Optional[str] = None):
        super().__init__()
        self.title("Book Collection Stats")
        self.geometry("1220x760")
        self.minsize(980, 620)

        # App directory + writable data/cache directories (packaging-friendly)
        self.app_dir = os.path.dirname(os.path.abspath(__file__))
        self.data_dir, self.cache_dir = get_app_dirs(self.app_dir)
        self.settings: Dict[str, Any] = load_settings(self.data_dir)

        # Cover cache (Open Library by ISBN via cover_cache.py)
        self.covers_available = bool(_HAS_PIL and _HAS_COVER_CACHE)
        self.covers_enabled = bool(self.settings.get("covers_enabled", True)) and self.covers_available
        self.cover_cache = None
        if self.covers_enabled and CoverCache:
            cache_dir = os.path.join(self.cache_dir, "covers")
            try:
                self.cover_cache = CoverCache(cache_dir=cache_dir)
            except Exception:
                self.cover_cache = None
                self.covers_enabled = False

        # Keep PhotoImage references alive (prevents images disappearing)
        self._img_refs: Dict[str, Any] = {}

        # Column drag state
        self._col_drag_name: Optional[str] = None
        self._col_drag_start_x: int = 0
        self._col_drag_start_index: int = -1
        self._col_dragging: bool = False

        self.books: List[Book] = []
        self.filtered_books: List[Book] = []
        self.current_path: Optional[str] = None

        # Sorting state (Library tab)
        self.sort_col = "Author"
        self.sort_desc = False

        self._build_menu()
        self._build_layout()

        if (not initial_path) and self.settings.get("auto_load_last_file", True):
            lp = str(self.settings.get("last_opened_path") or "").strip()
            if lp and os.path.exists(lp):
                initial_path = lp

        if initial_path:
            self.load_json(initial_path)

    # ----- Menu -----


    # ----- Menu -----

    def _build_menu(self):
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Open JSON…", command=self.open_json_dialog)
        file_menu.add_command(label="Settings…", command=self.open_settings_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=False)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Cover images help", command=self.show_cover_help)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.config(menu=menubar)



    # ----- Settings -----

    def open_settings_dialog(self):
        if not hasattr(self, "all_columns"):
            messagebox.showinfo("Settings", "Settings are available after the UI is initialized.")
            return

        win = tk.Toplevel(self)
        win.title("Settings")
        win.geometry("520x520")
        win.resizable(True, True)

        container = ttk.Frame(win, padding=12)
        container.pack(fill="both", expand=True)

        # Columns section
        cols_frame = ttk.LabelFrame(container, text="Library columns", padding=10)
        cols_frame.pack(fill="both", expand=False)

        # Current order comes from persisted settings (or defaults). Hidden columns still appear here.
        current_order = list(self.settings.get("library_column_order") or list(self.all_columns))
        current_order = [c for c in current_order if c in self.all_columns]
        for c in self.all_columns:
            if c not in current_order:
                current_order.append(c)

        col_vars: Dict[str, tk.BooleanVar] = {}
        visible_now = set(current_order)

        # Use settings visibility if present
        vis_setting = self.settings.get("library_visible_columns") or []
        if isinstance(vis_setting, list) and vis_setting:
            visible_now = set([c for c in vis_setting if c in self.all_columns])

        # show checkboxes in current column order
        for c in current_order:
            v = tk.BooleanVar(value=(c in visible_now))
            col_vars[c] = v
            cb = ttk.Checkbutton(cols_frame, text=c, variable=v)
            cb.pack(anchor="w")

        help_lbl = ttk.Label(
            cols_frame,
            text="Tip: You can also drag column headers in the Library table to reorder them.",
            foreground="#555",
            wraplength=460,
        )
        help_lbl.pack(anchor="w", pady=(8, 0))

        # Other settings
        other = ttk.LabelFrame(container, text="Other", padding=10)
        other.pack(fill="x", pady=(12, 0))

        covers_var = tk.BooleanVar(value=bool(self.settings.get("covers_enabled", True)))
        auto_load_var = tk.BooleanVar(value=bool(self.settings.get("auto_load_last_file", True)))

        covers_cb = ttk.Checkbutton(other, text="Enable cover images (requires Pillow + requests)", variable=covers_var)
        covers_cb.pack(anchor="w")
        if not self.covers_available:
            covers_cb.state(["disabled"])
            ttk.Label(other, text="Cover support unavailable (install Pillow + requests and keep cover_cache.py beside this file).",
                      foreground="#a00", wraplength=460).pack(anchor="w", pady=(4, 0))

        ttk.Checkbutton(other, text="Auto-load last opened JSON on startup", variable=auto_load_var).pack(anchor="w", pady=(6, 0))

        # Actions
        actions = ttk.Frame(container)
        actions.pack(fill="x", pady=(14, 0))

        def do_reset_columns():
            # Reset to defaults (all columns visible in default order)
            # Update local dialog state
            current_order[:] = list(self.all_columns)
            for c in self.all_columns:
                if c in col_vars:
                    col_vars[c].set(True)

            # Live-preview the reset in the Library table (without persisting until Save)
            try:
                if hasattr(self, "tree") and self.tree is not None:
                    self.tree.configure(displaycolumns=list(self.all_columns))
                    self._reset_library_column_widths(live_only=True)
                    self._populate_tree(self.filtered_books)
                    self._update_library_footer()
            except Exception:
                pass

        def do_clear_cover_cache():
            cache_dir = os.path.join(self.cache_dir, "covers")
            try:
                if os.path.isdir(cache_dir):
                    for fn in os.listdir(cache_dir):
                        fp = os.path.join(cache_dir, fn)
                        if os.path.isfile(fp):
                            os.remove(fp)
                messagebox.showinfo("Cover cache", "Cover cache cleared.")
            except Exception as e:
                messagebox.showerror("Cover cache", f"Could not clear cache:\n{e}")

        ttk.Button(actions, text="Reset columns", command=do_reset_columns).pack(side="left")
        ttk.Button(actions, text="Clear cover cache", command=do_clear_cover_cache).pack(side="left", padx=(8, 0))

        btns = ttk.Frame(container)
        btns.pack(fill="x", pady=(14, 0))

        def do_cancel():
            win.destroy()

        def do_save():
            # Column visibility
            visible = [c for c in current_order if col_vars.get(c) and col_vars[c].get()]
            if "Title" not in visible:
                visible.insert(0, "Title")  # Always keep Title visible
            if not visible:
                visible = ["Title"]

            # Column order is whatever the user currently has in the Library
            # (current_order), but we persist it explicitly.
            self.settings["library_column_order"] = list(current_order)
            self.settings["library_visible_columns"] = list(visible)
            # Persist current column widths
            self.settings["library_column_widths"] = self._get_current_column_widths()

            self.settings["covers_enabled"] = bool(covers_var.get())
            self.settings["auto_load_last_file"] = bool(auto_load_var.get())

            save_settings(self.data_dir, self.settings)
            self.apply_library_column_settings()

            # Refresh cover state
            self._apply_cover_settings()

            win.destroy()

        ttk.Button(btns, text="Cancel", command=do_cancel).pack(side="right")
        ttk.Button(btns, text="Save", command=do_save).pack(side="right", padx=(0, 8))

    def _apply_cover_settings(self):
        self.covers_enabled = bool(self.settings.get("covers_enabled", True)) and self.covers_available
        if self.covers_enabled and (self.cover_cache is None) and CoverCache:
            try:
                cache_dir = os.path.join(self.cache_dir, "covers")
                self.cover_cache = CoverCache(cache_dir=cache_dir)
            except Exception:
                self.cover_cache = None
                self.covers_enabled = False
        if (not self.covers_enabled):
            self.cover_cache = None

    def _get_displaycolumns(self) -> Tuple[str, ...]:
        if hasattr(self, "tree") and self.tree is not None:
            dc = self.tree.cget("displaycolumns")
            if isinstance(dc, (tuple, list)) and dc:
                return tuple(dc)
            if dc == "#all":
                return tuple(self.tree.cget("columns"))
        # fall back to settings / defaults
        order = self.settings.get("library_column_order") or []
        if isinstance(order, list) and order:
            return tuple([c for c in order if c in getattr(self, "all_columns", order)])
        return tuple(getattr(self, "all_columns", ()))

    def apply_library_column_settings(self):
        if not hasattr(self, "tree"):
            return

        order = self.settings.get("library_column_order") or []
        if not (isinstance(order, list) and order):
            order = list(self.all_columns)

        # Ensure order contains only known columns and includes all columns at least once
        order = [c for c in order if c in self.all_columns]
        for c in self.all_columns:
            if c not in order:
                order.append(c)

        visible = self.settings.get("library_visible_columns") or []
        if not (isinstance(visible, list) and visible):
            visible = list(self.all_columns)
        visible = [c for c in visible if c in self.all_columns]
        if "Title" not in visible:
            visible.insert(0, "Title")

        display = [c for c in order if c in visible]
        if not display:
            display = ["Title"]

        self.tree.configure(displaycolumns=display)
        self._apply_persisted_column_widths()
        # Re-populate to reflect displaycolumns order/visibility
        self._populate_tree(self.filtered_books)


    def _build_layout(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        self.tab_library = ttk.Frame(self.notebook, padding=10)
        self.tab_random = ttk.Frame(self.notebook, padding=10)
        self.tab_stats = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab_library, text="Library")
        self.notebook.add(self.tab_random, text="Random Picker")
        self.notebook.add(self.tab_stats, text="Statistics")

        self._build_library_tab()
        self._build_random_tab()
        self._build_stats_tab()

        self.status_var = tk.StringVar(value="Open a JSON export to begin.")
        status = ttk.Label(self, textvariable=self.status_var, anchor="w", padding=(8, 4))
        status.pack(fill="x", side="bottom")

    # ----------------------------
    # Library tab
    # ----------------------------


    def _build_library_tab(self):
        top = ttk.Frame(self.tab_library)
        top.pack(fill="x")

        ttk.Label(top, text="Search:").pack(side="left")
        self.search_var = tk.StringVar(value="")
        search_entry = ttk.Entry(top, textvariable=self.search_var, width=34)
        search_entry.pack(side="left", padx=(8, 12))
        search_entry.bind("<KeyRelease>", lambda e: self.apply_filters())

        ttk.Label(top, text="Quick filter:").pack(side="left")
        self.quick_filter_var = tk.StringVar(value="All")
        self.quick_filter = ttk.Combobox(
            top,
            textvariable=self.quick_filter_var,
            values=["All", "Read", "Unread", "To Read", "Owned", "Unowned"],
            state="readonly",
            width=12,
        )
        self.quick_filter.pack(side="left", padx=(8, 12))
        self.quick_filter.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        ttk.Label(top, text="Tag:").pack(side="left")
        self.tag_filter_var = tk.StringVar(value="Any")
        self.tag_filter = ttk.Combobox(
            top,
            textvariable=self.tag_filter_var,
            values=["Any"],
            state="readonly",
            width=16,
        )
        self.tag_filter.pack(side="left", padx=(8, 12))
        self.tag_filter.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        ttk.Label(top, text="Media type:").pack(side="left")
        self.media_filter_var = tk.StringVar(value="Any")
        self.media_filter = ttk.Combobox(
            top,
            textvariable=self.media_filter_var,
            values=["Any"],
            state="readonly",
            width=20,
        )
        self.media_filter.pack(side="left", padx=(8, 12))
        self.media_filter.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        ttk.Button(top, text="Clear", command=self._clear_search).pack(side="left", padx=8)

        # Treeview + scrollbars
        table_frame = ttk.Frame(self.tab_library)
        table_frame.pack(fill="both", expand=True, pady=(10, 0))

        # Keep a stable list of all columns; visibility/order is controlled by displaycolumns + settings.
        self.all_columns = ("Title", "Author", "Year", "Pages", "Rating", "Format", "Tags", "Collections", "Genres")

        self.tree = ttk.Treeview(table_frame, columns=self.all_columns, show="headings", height=22)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Headings (sorting handled by our header-click handler so we can also support drag-reorder)
        for c in self.all_columns:
            self.tree.heading(c, text=c)

        # Default widths (can be overridden by persisted settings)
        self._library_default_widths = dict(DEFAULT_LIBRARY_COLUMN_WIDTHS)
        self.tree.column("Title", width=self._library_default_widths["Title"], anchor="w")
        self.tree.column("Author", width=self._library_default_widths["Author"], anchor="w")
        self.tree.column("Year", width=self._library_default_widths["Year"], anchor="e")
        self.tree.column("Pages", width=self._library_default_widths["Pages"], anchor="e")
        self.tree.column("Rating", width=self._library_default_widths["Rating"], anchor="e")
        self.tree.column("Format", width=self._library_default_widths["Format"], anchor="w")
        self.tree.column("Tags", width=self._library_default_widths["Tags"], anchor="w")
        self.tree.column("Collections", width=self._library_default_widths["Collections"], anchor="w")
        self.tree.column("Genres", width=self._library_default_widths["Genres"], anchor="w")

        # Apply persisted widths (if any)
        self._apply_persisted_column_widths()

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)


        # Library footer (shown count for current filters)
        footer = ttk.Frame(self.tab_library)
        footer.pack(fill="x", pady=(6, 0))
        self.library_footer_var = tk.StringVar(value="Showing 0 of 0")
        ttk.Label(footer, textvariable=self.library_footer_var, anchor="w").pack(side="left")

        self.tree.bind("<Double-1>", self._on_double_click_book)

        # Header drag-to-reorder + click-to-sort
        self.tree.bind("<ButtonPress-1>", self._on_tree_button_press, add="+")
        self.tree.bind("<B1-Motion>", self._on_tree_mouse_drag, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_tree_button_release, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_tree_any_button_release, add="+")

        # Apply persisted column order/visibility
        self.apply_library_column_settings()

    # ----- Column drag + header click sorting -----

    def _identify_display_column(self, x: int) -> Optional[str]:
        if not hasattr(self, "tree") or self.tree is None:
            return None
        col_token = self.tree.identify_column(x)  # '#1', '#2', ...
        if not col_token or not col_token.startswith("#"):
            return None
        try:
            idx = int(col_token[1:]) - 1
        except Exception:
            return None
        displaycols = list(self._get_displaycolumns())
        if idx < 0 or idx >= len(displaycols):
            return None
        return displaycols[idx]

    def _on_tree_button_press(self, event):
        # Ignore clicks that aren't on the header
        region = self.tree.identify_region(event.x, event.y)
        if region != "heading":
            self._col_drag_name = None
            return

        col_name = self._identify_display_column(event.x)
        if not col_name:
            self._col_drag_name = None
            return

        self._col_drag_name = col_name
        self._col_drag_start_x = event.x
        self._col_dragging = False
        displaycols = list(self._get_displaycolumns())
        try:
            self._col_drag_start_index = displaycols.index(col_name)
        except ValueError:
            self._col_drag_start_index = -1

    def _on_tree_mouse_drag(self, event):
        if not self._col_drag_name:
            return
        # If the mouse moved enough, treat as drag
        if abs(event.x - self._col_drag_start_x) > 6:
            self._col_dragging = True

    def _on_tree_button_release(self, event):
        if not self._col_drag_name:
            return

        region = self.tree.identify_region(event.x, event.y)
        if region != "heading":
            self._reset_col_drag()
            return

        col_name = self._col_drag_name
        displaycols = list(self._get_displaycolumns())

        if not self._col_dragging:
            # Simple click: sort by this column
            self._on_sort_by(col_name)
            self._reset_col_drag()
            return

        # Drag: reorder columns
        target = self._identify_display_column(event.x)
        if not target or target not in displaycols:
            self._reset_col_drag()
            return

        if col_name not in displaycols:
            self._reset_col_drag()
            return

        src_idx = displaycols.index(col_name)
        dst_idx = displaycols.index(target)
        if src_idx == dst_idx:
            self._reset_col_drag()
            return

        displaycols.pop(src_idx)
        # After pop, if dragging to the right and src before dst, dst shifts left
        if src_idx < dst_idx:
            dst_idx -= 1
        displaycols.insert(dst_idx, col_name)

        # Update Treeview and persist settings
        self.tree.configure(displaycolumns=displaycols)

        # Persist full order list (include hidden columns after visible ones so they remain known)
        full_order = list(displaycols)
        for c in self.all_columns:
            if c not in full_order:
                full_order.append(c)

        self.settings["library_column_order"] = full_order
        # Keep existing visibility list
        if not (isinstance(self.settings.get("library_visible_columns"), list) and self.settings.get("library_visible_columns")):
            self.settings["library_visible_columns"] = list(displaycols)

        save_settings(self.data_dir, self.settings)

        # Re-populate to match new order
        self._populate_tree(self.filtered_books)

        self._reset_col_drag()

    def _reset_col_drag(self):
        self._col_drag_name = None
        self._col_drag_start_x = 0
        self._col_drag_start_index = -1
        self._col_dragging = False


    def _clear_search(self):
        self.search_var.set("")
        self.quick_filter_var.set("All")
        self.tag_filter_var.set("Any")
        self.media_filter_var.set("Any")
        self.apply_filters()

    def _on_double_click_book(self, event):
        item = self.tree.selection()
        if not item:
            return
        tags = self.tree.item(item[0], "tags")
        if not tags:
            return
        books_id = tags[0]
        b = next((x for x in self.books if x.books_id == books_id), None)
        if b:
            self.show_book_details(b)

    def _on_sort_by(self, col: str):
        if self.sort_col == col:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_col = col
            self.sort_desc = False

        self.filtered_books.sort(key=self._sort_key_for(self.sort_col), reverse=self.sort_desc)
        self._populate_tree(self.filtered_books)

        arrow = "▼" if self.sort_desc else "▲"
        for c in self.tree["columns"]:
            self.tree.heading(c, text=c, command=lambda x=c: self._on_sort_by(x))
        self.tree.heading(col, text=f"{col} {arrow}", command=lambda x=col: self._on_sort_by(x))

    def apply_filters(self):
        q = self.search_var.get().strip().lower()
        quick = self.quick_filter_var.get()
        tag_sel = self.tag_filter_var.get()
        media_sel = self.media_filter_var.get()

        def matches_quick(b: Book) -> bool:
            if quick == "All":
                return True
            if quick == "Read":
                return b.is_read
            if quick == "Unread":
                return b.is_unread
            if quick == "To Read":
                return b.is_to_read
            if quick == "Owned":
                return b.is_owned
            if quick == "Unowned":
                return not b.is_owned
            return True

        def matches_tag(b: Book) -> bool:
            if tag_sel in ("", "Any"):
                return True
            return any(t.lower() == tag_sel.lower() for t in b.tags)

        def matches_media(b: Book) -> bool:
            if media_sel in ("", "Any"):
                return True
            return any(f.lower() == media_sel.lower() for f in b.formats)

        def matches_query(b: Book) -> bool:
            if not q:
                return True
            hay = " | ".join([
                b.title,
                b.display_author,
                b.publication,
                b.collections_str(),
                b.genre_str(),
                b.tags_str(),
                b.formats_str(),
                " ".join(b.series),
            ]).lower()
            return q in hay

        self.filtered_books = [
            b for b in self.books
            if matches_quick(b) and matches_tag(b) and matches_media(b) and matches_query(b)
        ]

        self.filtered_books.sort(key=self._sort_key_for(self.sort_col), reverse=self.sort_desc)
        self._populate_tree(self.filtered_books)

        self.status_var.set(
            f"{len(self.filtered_books):,} shown / {len(self.books):,} total"
            + (f" — {os.path.basename(self.current_path)}" if self.current_path else "")
        )

    def _sort_key_for(self, col: str):
        if col == "Title":
            return lambda b: (b.title.lower(), b.author_last, b.display_author.lower())
        if col == "Author":
            return lambda b: (b.author_last, b.display_author.lower(), b.title.lower())
        if col == "Year":
            return lambda b: (b.year is None, b.year if b.year is not None else 0, b.title.lower())
        if col == "Pages":
            return lambda b: (b.pages is None, b.pages if b.pages is not None else 0, b.title.lower())
        if col == "Rating":
            return lambda b: (b.rating is None, b.rating if b.rating is not None else 0.0, b.title.lower())
        if col == "Format":
            return lambda b: (b.primary_format().lower(), b.title.lower())
        if col == "Tags":
            return lambda b: (b.tags_str().lower(), b.title.lower())
        if col == "Collections":
            return lambda b: (b.collections_str().lower(), b.title.lower())
        if col == "Genres":
            return lambda b: (b.genre_str().lower(), b.title.lower())
        return lambda b: (b.author_last, b.title.lower())

    def _populate_tree(self, books: List[Book]):
        self.tree.delete(*self.tree.get_children())
        for b in books:
            year = b.year if b.year is not None else ""
            pages = b.pages if b.pages is not None else ""
            rating = "" if b.rating is None else f"{b.rating:g}"
            vals = (
                b.title,
                b.display_author,
                year,
                pages,
                rating,
                b.primary_format(),
                b.tags_str(),
                b.collections_str(),
                b.genre_str(),
            )
            self.tree.insert("", "end", values=vals, tags=(b.books_id,))

        # Reset selected-count footer after repopulating
        self._update_library_footer()

    def _update_library_footer(self):
        """Update the Library-tab footer showing how many rows are shown under current filters."""
        if not hasattr(self, "library_footer_var"):
            return
        shown = len(getattr(self, "filtered_books", []) or [])
        total = len(getattr(self, "books", []) or [])
        suffix = f" — {os.path.basename(self.current_path)}" if getattr(self, "current_path", None) else ""
        self.library_footer_var.set(f"Showing {shown:,} of {total:,}{suffix}")

    def _get_current_column_widths(self) -> Dict[str, int]:
        widths: Dict[str, int] = {}
        if not hasattr(self, "tree") or self.tree is None:
            return widths
        for c in getattr(self, "all_columns", ()):
            try:
                w = int(self.tree.column(c, "width"))
                widths[c] = w
            except Exception:
                continue
        return widths

    def _apply_persisted_column_widths(self):
        """Apply saved widths (if present) after Treeview is created."""
        if not hasattr(self, "tree") or self.tree is None:
            return
        saved = self.settings.get("library_column_widths") or {}
        if not isinstance(saved, dict):
            return
        for c in getattr(self, "all_columns", ()):
            if c in saved:
                try:
                    w = int(saved[c])
                    if w >= 40:
                        self.tree.column(c, width=w)
                except Exception:
                    pass

    def _reset_library_column_widths(self, live_only: bool = False):
        """Reset column widths to defaults. If live_only=True, do not persist until user saves settings."""
        if not hasattr(self, "tree") or self.tree is None:
            return
        defaults = getattr(self, "_library_default_widths", None) or DEFAULT_LIBRARY_COLUMN_WIDTHS
        for c in getattr(self, "all_columns", ()):
            w = defaults.get(c)
            if w:
                try:
                    self.tree.column(c, width=int(w))
                except Exception:
                    pass
        if not live_only:
            self.settings["library_column_widths"] = self._get_current_column_widths()
            save_settings(self.data_dir, self.settings)

    def _on_tree_any_button_release(self, _event):
        """Catch column resize events and persist widths with a small debounce."""
        if not hasattr(self, "tree") or self.tree is None:
            return
        new_widths = self._get_current_column_widths()
        last = getattr(self, "_last_saved_widths", None)
        if last == new_widths:
            return
        self._last_saved_widths = dict(new_widths)

        # Update settings in memory
        self.settings["library_column_widths"] = dict(new_widths)

        # Debounce disk writes
        after_id = getattr(self, "_width_save_after_id", None)
        if after_id:
            try:
                self.after_cancel(after_id)
            except Exception:
                pass

        def commit():
            save_settings(self.data_dir, self.settings)

        self._width_save_after_id = self.after(600, commit)


    # ----------------------------
    # Random selection
    # ----------------------------

    def _build_random_tab(self):
        left = ttk.Frame(self.tab_random)
        left.pack(side="left", fill="y", padx=(0, 12))

        ttk.Label(left, text="Collections (pick one or more):").pack(anchor="w")
        self.collections_list = tk.Listbox(left, selectmode="extended", height=16, exportselection=False)
        self.collections_list.pack(fill="y", expand=False, pady=(6, 10))

        filters = ttk.LabelFrame(left, text="Filters", padding=10)
        filters.pack(fill="x", pady=(8, 0))

        self.rp_status_var = tk.StringVar(value="Any")
        ttk.Label(filters, text="Read status:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            filters,
            textvariable=self.rp_status_var,
            values=["Any", "Read", "Unread"],
            state="readonly",
            width=10
        ).grid(row=0, column=1, sticky="w", padx=(8, 0))

        self.rp_owned_only = tk.BooleanVar(value=False)
        ttk.Checkbutton(filters, text="Owned only", variable=self.rp_owned_only).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(6, 0)
        )

        self.rp_min_rating = tk.DoubleVar(value=0.0)
        ttk.Label(filters, text="Min rating:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.rp_rating_scale = ttk.Scale(filters, from_=0.0, to=5.0, variable=self.rp_min_rating)
        self.rp_rating_scale.grid(row=2, column=1, sticky="we", padx=(8, 0), pady=(8, 0))
        filters.columnconfigure(1, weight=1)

        self.rp_genre_var = tk.StringVar(value="")
        ttk.Label(filters, text="Genre contains:").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(filters, textvariable=self.rp_genre_var, width=14).grid(
            row=3, column=1, sticky="we", padx=(8, 0), pady=(8, 0)
        )

        ttk.Button(left, text="Pick Random Book", command=self.pick_random).pack(fill="x", pady=(12, 6))

        right = ttk.Frame(self.tab_random)
        right.pack(side="left", fill="both", expand=True)

        cover_and_text = ttk.Frame(right)
        cover_and_text.pack(fill="both", expand=True)

        self.random_cover_label = ttk.Label(cover_and_text, text="", anchor="center")
        self.random_cover_label.pack(anchor="n", pady=(0, 8))

        self.random_result = tk.Text(cover_and_text, height=18, wrap="word")
        self.random_result.pack(fill="both", expand=True)

        self._random_current_book: Optional[Book] = None
        self._random_last_pool_size: Optional[int] = None
        self.random_result.configure(state="disabled")

    def _selected_collections(self) -> List[str]:
        sel = self.collections_list.curselection()
        items = [self.collections_list.get(i) for i in sel]
        return items

    def _random_pool(self) -> List[Book]:
        if not self.books:
            return []

        selected = set(self._selected_collections())
        if selected:
            pool = [b for b in self.books if any(c in selected for c in b.collections)]
        else:
            pool = list(self.books)

        status = self.rp_status_var.get()
        if status == "Read":
            pool = [b for b in pool if b.is_read]
        elif status == "Unread":
            pool = [b for b in pool if b.is_unread]

        if self.rp_owned_only.get():
            pool = [b for b in pool if b.is_owned]

        min_rating = float(self.rp_min_rating.get())
        if min_rating > 0:
            pool = [b for b in pool if (b.rating is not None and b.rating >= min_rating)]

        genre_q = self.rp_genre_var.get().strip().lower()
        if genre_q:
            pool = [b for b in pool if genre_q in b.genre_str().lower()]

        return pool

    def pick_random(self):
        pool = self._random_pool()
        if not pool:
            self._random_current_book = None
            self._random_last_pool_size = None
            self._clear_cover_label(getattr(self, "random_cover_label", None), key="random")
            self._set_random_text("No matching books for your filters.")
            return
        b = random.choice(pool)
        self._random_current_book = b
        self._random_last_pool_size = len(pool)
        self._set_cover_label_for_book(getattr(self, "random_cover_label", None), b, key="random", details_text=getattr(self, "random_result", None))
        self._set_random_text(self._book_text(b, pool_size=len(pool)))

    def _set_random_text(self, s: str):
        self.random_result.configure(state="normal")
        self.random_result.delete("1.0", "end")
        self.random_result.insert("1.0", s)
        self.random_result.configure(state="disabled")

    # ----------------------------
    # Stats tab
    # ----------------------------

    def _build_stats_tab(self):
        # Sub-tabs for nicer organization
        self.stats_nb = ttk.Notebook(self.tab_stats)
        self.stats_nb.pack(fill="both", expand=True)

        self.stats_overview = ttk.Frame(self.stats_nb, padding=10)
        self.stats_trends = ttk.Frame(self.stats_nb, padding=10)
        self.stats_dist = ttk.Frame(self.stats_nb, padding=10)

        self.stats_nb.add(self.stats_overview, text="Overview")
        self.stats_nb.add(self.stats_trends, text="Trends")
        self.stats_nb.add(self.stats_dist, text="Distributions")

        # --- Overview ---
        self.stats_summary = tk.Text(self.stats_overview, height=12, wrap="word")
        self.stats_summary.pack(fill="x", expand=False, pady=(6, 10))
        self.stats_summary.configure(state="disabled")

        grids = ttk.Frame(self.stats_overview)
        grids.pack(fill="both", expand=True)

        self.top_authors = _TopTable(grids, "Top Authors (by count)")
        self.top_genres = _TopTable(grids, "Top Genres (by count)")
        self.top_collections = _TopTable(grids, "Top Collections (by count)")
        self.top_formats = _TopTable(grids, "Top Formats (by count)")
        self.top_tags = _TopTable(grids, "Top Tags (by count)")

        self.top_authors.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))
        self.top_genres.grid(row=0, column=1, sticky="nsew", padx=(0, 8), pady=(0, 8))
        self.top_collections.grid(row=0, column=2, sticky="nsew", pady=(0, 8))

        self.top_formats.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        self.top_tags.grid(row=1, column=1, sticky="nsew", padx=(0, 8))
        
        # --- Books read by year (table) ---
        year_overview_frame = ttk.LabelFrame(grids, text="Books read by year", padding=8)
        year_overview_frame.grid(row=1, column=2, sticky="nsew")

        self.overview_year_tree = ttk.Treeview(
            year_overview_frame,
            columns=("Year", "Books"),
            show="headings",
            height=10
        )
        self.overview_year_tree.heading("Year", text="Year")
        self.overview_year_tree.heading("Books", text="Books")
        self.overview_year_tree.column("Year", width=70, anchor="e")
        self.overview_year_tree.column("Books", width=80, anchor="e")
        self.overview_year_tree.pack(fill="both", expand=True)
# Leave (row=1, col=2) empty for breathing room / future tables

        for c in range(3):
            grids.columnconfigure(c, weight=1)
        grids.rowconfigure(0, weight=1)
        grids.rowconfigure(1, weight=1)

        # --- Trends ---
        if _HAS_MPL:
            trends_grid = ttk.Frame(self.stats_trends)
            trends_grid.pack(fill="both", expand=True)

            self._fig_reads = Figure(figsize=(5, 3), dpi=100)
            self._ax_reads = self._fig_reads.add_subplot(111)
            self._canvas_reads = FigureCanvasTkAgg(self._fig_reads, master=trends_grid)
            self._canvas_reads.get_tk_widget().grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 10))

            self._fig_pages = Figure(figsize=(5, 3), dpi=100)
            self._ax_pages = self._fig_pages.add_subplot(111)
            self._canvas_pages = FigureCanvasTkAgg(self._fig_pages, master=trends_grid)
            self._canvas_pages.get_tk_widget().grid(row=0, column=1, sticky="nsew", pady=(0, 10))

            self._fig_added = Figure(figsize=(5, 3), dpi=100)
            self._ax_added = self._fig_added.add_subplot(111)
            self._canvas_added = FigureCanvasTkAgg(self._fig_added, master=trends_grid)
            self._canvas_added.get_tk_widget().grid(row=1, column=0, sticky="nsew", padx=(0, 10))

            year_table_frame = ttk.LabelFrame(trends_grid, text="Reads by year", padding=8)
            year_table_frame.grid(row=1, column=1, sticky="nsew")

            self.year_tree = ttk.Treeview(
                year_table_frame,
                columns=("Year", "Books Read", "Pages Read", "Avg Rating"),
                show="headings",
                height=10
            )
            for c, w, a in [
                ("Year", 60, "e"),
                ("Books Read", 90, "e"),
                ("Pages Read", 90, "e"),
                ("Avg Rating", 90, "e"),
            ]:
                self.year_tree.heading(c, text=c)
                self.year_tree.column(c, width=w, anchor=a)
            self.year_tree.pack(fill="both", expand=True)

            trends_grid.columnconfigure(0, weight=1)
            trends_grid.columnconfigure(1, weight=1)
            trends_grid.rowconfigure(0, weight=1)
            trends_grid.rowconfigure(1, weight=1)
        else:
            ttk.Label(
                self.stats_trends,
                text="Matplotlib not available — install it to see charts (pip install matplotlib)."
            ).pack(anchor="w")

        # --- Distributions ---
        if _HAS_MPL:
            dist_grid = ttk.Frame(self.stats_dist)
            dist_grid.pack(fill="both", expand=True)

            self._fig_format = Figure(figsize=(5, 3), dpi=100)
            self._ax_format = self._fig_format.add_subplot(111)
            self._canvas_format = FigureCanvasTkAgg(self._fig_format, master=dist_grid)
            self._canvas_format.get_tk_widget().grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 10))

            self._fig_rating = Figure(figsize=(5, 3), dpi=100)
            self._ax_rating = self._fig_rating.add_subplot(111)
            self._canvas_rating = FigureCanvasTkAgg(self._fig_rating, master=dist_grid)
            self._canvas_rating.get_tk_widget().grid(row=0, column=1, sticky="nsew", pady=(0, 10))

            self._fig_tags = Figure(figsize=(5, 3), dpi=100)
            self._ax_tags = self._fig_tags.add_subplot(111)
            self._canvas_tags = FigureCanvasTkAgg(self._fig_tags, master=dist_grid)
            self._canvas_tags.get_tk_widget().grid(row=1, column=0, sticky="nsew", padx=(0, 10))

            self._fig_len = Figure(figsize=(5, 3), dpi=100)
            self._ax_len = self._fig_len.add_subplot(111)
            self._canvas_len = FigureCanvasTkAgg(self._fig_len, master=dist_grid)
            self._canvas_len.get_tk_widget().grid(row=1, column=1, sticky="nsew")

            dist_grid.columnconfigure(0, weight=1)
            dist_grid.columnconfigure(1, weight=1)
            dist_grid.rowconfigure(0, weight=1)
            dist_grid.rowconfigure(1, weight=1)
        else:
            ttk.Label(
                self.stats_dist,
                text="Matplotlib not available — install it to see charts (pip install matplotlib)."
            ).pack(anchor="w")

    def _fill_top_table(self, tree: ttk.Treeview, rows: List[Tuple[str, int]]):
        tree.delete(*tree.get_children())
        for item, count in rows:
            tree.insert("", "end", values=(item, f"{count:,}"))

    def refresh_stats(self):
        if not self.books:
            self._set_stats_summary("No data loaded.")
            self._fill_top_table(self.top_authors.tree, [])
            self._fill_top_table(self.top_genres.tree, [])
            self._fill_top_table(self.top_collections.tree, [])
            self._fill_top_table(self.top_formats.tree, [])
            self._fill_top_table(self.top_tags.tree, [])
            if hasattr(self, "overview_year_tree"):
                self.overview_year_tree.delete(*self.overview_year_tree.get_children())
            if _HAS_MPL:
                self._clear_axes()
            return

        total = len(self.books)
        read = sum(1 for b in self.books if b.is_read)
        unread = sum(1 for b in self.books if b.is_unread)
        to_read = sum(1 for b in self.books if b.is_to_read)
        owned = sum(1 for b in self.books if b.is_owned)
        unowned = total - owned

        ratings_all = [b.rating for b in self.books if b.rating is not None]
        avg_rating_all = (sum(ratings_all) / len(ratings_all)) if ratings_all else None

        ratings_read = [b.rating for b in self.books if b.is_read and b.rating is not None]
        avg_rating_read = (sum(ratings_read) / len(ratings_read)) if ratings_read else None

        pages_total = sum(b.pages for b in self.books if b.pages is not None)
        pages_read = sum(b.pages for b in self.books if b.is_read and b.pages is not None)

        authors = Counter()
        genres = Counter()
        collections = Counter()
        formats = Counter()
        tags = Counter()

        for b in self.books:
            if b.display_author:
                authors[b.display_author] += 1
            for g in b.genre:
                genres[g] += 1
            for c in b.collections:
                collections[c] += 1
            if b.primary_format():
                formats[b.primary_format()] += 1
            for t in b.tags:
                tags[t] += 1

        unique_authors = len({b.display_author for b in self.books if b.display_author})
        unique_collections = len(collections)
        unique_formats = len(formats)

        # Reading timeline
        read_dates = [d for d in (_parse_date_yyyy_mm_dd(b.dateread) for b in self.books if b.is_read) if d is not None]
        first_read = min(read_dates).date().isoformat() if read_dates else None
        last_read = max(read_dates).date().isoformat() if read_dates else None

        # Yearly stats
        yr_books = Counter()
        yr_pages = Counter()
        yr_ratings_sum = defaultdict(float)
        yr_ratings_n = Counter()

        for b in self.books:
            if not b.is_read:
                continue
            d = _parse_date_yyyy_mm_dd(b.dateread)
            if not d:
                continue
            y = d.year
            yr_books[y] += 1
            if b.pages is not None:
                yr_pages[y] += b.pages
            if b.rating is not None:
                yr_ratings_sum[y] += b.rating
                yr_ratings_n[y] += 1

        busiest_year = None
        if yr_books:
            busiest_year, busiest_count = max(yr_books.items(), key=lambda kv: kv[1])
        else:
            busiest_count = 0

        # Reading streak (consecutive years with >=1 read)
        best_streak = 0
        best_streak_range: Optional[Tuple[int, int]] = None
        if yr_books:
            years = sorted(yr_books.keys())
            cur_start = years[0]
            cur_prev = years[0]
            cur_len = 1
            for y in years[1:]:
                if y == cur_prev + 1:
                    cur_prev = y
                    cur_len += 1
                else:
                    if cur_len > best_streak:
                        best_streak = cur_len
                        best_streak_range = (cur_start, cur_prev)
                    cur_start = y
                    cur_prev = y
                    cur_len = 1
            if cur_len > best_streak:
                best_streak = cur_len
                best_streak_range = (cur_start, cur_prev)

        summary_lines = [
            f"Total books: {total:,}",
            f"Owned: {owned:,}   |   Unowned: {unowned:,}",
            f"Read: {read:,}   |   Unread: {unread:,}   |   To Read (Owned+Unread): {to_read:,}",
            f"Unique authors: {unique_authors:,}",
            f"Unique collections: {unique_collections:,}",
            f"Unique formats: {unique_formats:,}",
            f"Total pages (where known): {pages_total:,}",
            f"Pages read (where known): {pages_read:,}",
        ]

        if avg_rating_all is None:
            summary_lines.append("Average rating (rated only): (no ratings found)")
        else:
            summary_lines.append(f"Average rating (rated only): {avg_rating_all:.2f}  (n={len(ratings_all):,})")

        if avg_rating_read is not None:
            summary_lines.append(f"Average rating (read + rated): {avg_rating_read:.2f}  (n={len(ratings_read):,})")

        if first_read and last_read:
            summary_lines.append(f"Reading history (from Date Read): {first_read} → {last_read}")
        if busiest_year is not None:
            summary_lines.append(f"Busiest year: {busiest_year} ({busiest_count} books)")
        if best_streak_range and best_streak > 1:
            summary_lines.append(f"Longest yearly streak: {best_streak} years ({best_streak_range[0]}–{best_streak_range[1]})")

        self._set_stats_summary("\n".join(summary_lines))

        self._fill_top_table(self.top_authors.tree, authors.most_common(20))
        self._fill_top_table(self.top_genres.tree, genres.most_common(20))
        self._fill_top_table(self.top_collections.tree, collections.most_common(20))
        self._fill_top_table(self.top_formats.tree, formats.most_common(20))
        self._fill_top_table(self.top_tags.tree, tags.most_common(20))

        # Populate the overview "Books read by year" table (most recent first)
        if hasattr(self, "overview_year_tree"):
            self.overview_year_tree.delete(*self.overview_year_tree.get_children())
            for y in sorted(yr_books.keys(), reverse=True):
                self.overview_year_tree.insert("", "end", values=(y, f"{yr_books[y]:,}"))

        if _HAS_MPL:
            self._update_trends_charts(yr_books, yr_pages)
            self._update_distributions_charts(formats, ratings_all, tags)
            self._update_added_chart()

    def _set_stats_summary(self, s: str):
        self.stats_summary.configure(state="normal")
        self.stats_summary.delete("1.0", "end")
        self.stats_summary.insert("1.0", s)
        self.stats_summary.configure(state="disabled")

    def _clear_axes(self):
        # trends
        for ax in [getattr(self, "_ax_reads", None), getattr(self, "_ax_pages", None), getattr(self, "_ax_added", None),
                   getattr(self, "_ax_format", None), getattr(self, "_ax_rating", None), getattr(self, "_ax_tags", None),
                   getattr(self, "_ax_len", None)]:
            if ax is not None:
                ax.clear()
        for canvas in [getattr(self, "_canvas_reads", None), getattr(self, "_canvas_pages", None), getattr(self, "_canvas_added", None),
                       getattr(self, "_canvas_format", None), getattr(self, "_canvas_rating", None), getattr(self, "_canvas_tags", None),
                       getattr(self, "_canvas_len", None)]:
            if canvas is not None:
                canvas.draw_idle()

        if hasattr(self, "year_tree"):
            self.year_tree.delete(*self.year_tree.get_children())

    def _update_trends_charts(self, yr_books: Counter, yr_pages: Counter):
        # Books read per year
        self._ax_reads.clear()
        if yr_books:
            years = sorted(yr_books.keys())
            vals = [yr_books[y] for y in years]
            self._ax_reads.bar(years, vals)
            self._ax_reads.set_title("Books read per year")
            self._ax_reads.set_xlabel("Year")
            self._ax_reads.set_ylabel("Books")
        else:
            self._ax_reads.set_title("Books read per year (no Date Read data)")
        self._fig_reads.tight_layout()
        self._canvas_reads.draw_idle()

        # Pages read per year
        self._ax_pages.clear()
        if yr_pages:
            years = sorted(yr_pages.keys())
            vals = [yr_pages[y] for y in years]
            self._ax_pages.bar(years, vals)
            self._ax_pages.set_title("Pages read per year (known pages only)")
            self._ax_pages.set_xlabel("Year")
            self._ax_pages.set_ylabel("Pages")
        else:
            self._ax_pages.set_title("Pages read per year (no pages+Date Read data)")
        self._fig_pages.tight_layout()
        self._canvas_pages.draw_idle()

        # Year table
        if hasattr(self, "year_tree"):
            self.year_tree.delete(*self.year_tree.get_children())
            years = sorted(yr_books.keys()) if yr_books else []
            for y in years:
                books = yr_books.get(y, 0)
                pages = yr_pages.get(y, 0)
                # Avg rating per year (read + rated)
                r_sum = 0.0
                r_n = 0
                for b in self.books:
                    if not b.is_read or b.rating is None:
                        continue
                    d = _parse_date_yyyy_mm_dd(b.dateread)
                    if d and d.year == y:
                        r_sum += b.rating
                        r_n += 1
                avg = (r_sum / r_n) if r_n else None
                self.year_tree.insert("", "end", values=(y, books, f"{pages:,}" if pages else "", f"{avg:.2f}" if avg is not None else ""))

    def _update_added_chart(self):
        # Books added per year (from entrydate)
        self._ax_added.clear()
        yr_added = Counter()
        for b in self.books:
            d = _parse_date_yyyy_mm_dd(b.entrydate)
            if d:
                yr_added[d.year] += 1
        if yr_added:
            years = sorted(yr_added.keys())
            vals = [yr_added[y] for y in years]
            self._ax_added.bar(years, vals)
            self._ax_added.set_title("Books added per year (Entry Date)")
            self._ax_added.set_xlabel("Year")
            self._ax_added.set_ylabel("Books")
        else:
            self._ax_added.set_title("Books added per year (no Entry Date data)")
        self._fig_added.tight_layout()
        self._canvas_added.draw_idle()

    def _update_distributions_charts(self, formats: Counter, ratings_all: List[float], tags: Counter):
        # Format distribution
        self._ax_format.clear()
        if formats:
            items = formats.most_common(10)
            labels = [i[0] for i in items]
            vals = [i[1] for i in items]
            self._ax_format.bar(range(len(labels)), vals)
            self._ax_format.set_title("Top formats")
            self._ax_format.set_ylabel("Books")
            self._ax_format.set_xticks(range(len(labels)))
            self._ax_format.set_xticklabels(labels, rotation=30, ha="right")
        else:
            self._ax_format.set_title("Top formats (no format data)")
        self._fig_format.tight_layout()
        self._canvas_format.draw_idle()

        # Rating histogram (all rated)
        self._ax_rating.clear()
        if ratings_all:
            # Histogram bins for half-star steps
            bins = [x / 2 for x in range(0, 11)]  # 0..5 in 0.5
            self._ax_rating.hist(ratings_all, bins=bins, edgecolor="black")
            self._ax_rating.set_title("Ratings distribution (rated only)")
            self._ax_rating.set_xlabel("Rating")
            self._ax_rating.set_ylabel("Count")
        else:
            self._ax_rating.set_title("Ratings distribution (no ratings)")
        self._fig_rating.tight_layout()
        self._canvas_rating.draw_idle()

        # Top tags
        self._ax_tags.clear()
        if tags:
            items = tags.most_common(12)
            labels = [i[0] for i in items]
            vals = [i[1] for i in items]
            self._ax_tags.bar(range(len(labels)), vals)
            self._ax_tags.set_title("Top tags")
            self._ax_tags.set_ylabel("Books")
            self._ax_tags.set_xticks(range(len(labels)))
            self._ax_tags.set_xticklabels(labels, rotation=30, ha="right")
        else:
            self._ax_tags.set_title("Top tags (no tags)")
        self._fig_tags.tight_layout()
        self._canvas_tags.draw_idle()

        # Page-length distribution (known pages)
        self._ax_len.clear()
        pages = [b.pages for b in self.books if b.pages is not None]
        if pages:
            # coarse bins in 100-page increments
            maxp = max(pages)
            step = 100
            bins = list(range(0, (maxp // step + 2) * step, step))
            self._ax_len.hist(pages, bins=bins, edgecolor="black")
            self._ax_len.set_title("Page count distribution (known pages)")
            self._ax_len.set_xlabel("Pages")
            self._ax_len.set_ylabel("Books")
        else:
            self._ax_len.set_title("Page count distribution (no pages)")
        self._fig_len.tight_layout()
        self._canvas_len.draw_idle()

    # ----------------------------
    # File operations + refresh
    # ----------------------------

    def open_json_dialog(self):
        path = filedialog.askopenfilename(
            title="Open book JSON",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if path:
            self.load_json(path)

    def load_json(self, path: str):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                raise ValueError("Top-level JSON must be an object (dict).")
            self.books = parse_books(data)
            self.current_path = path

            # Persist last opened file
            self.settings["last_opened_path"] = path
            save_settings(self.data_dir, self.settings)

            # Default sort state (Author last name)
            self.sort_col = "Author"
            self.sort_desc = False

            self.status_var.set(f"Loaded {len(self.books):,} books from: {os.path.basename(path)}")
            self._refresh_all()
        except Exception as e:
            messagebox.showerror("Load failed", f"Could not load JSON:\n{e}")

    def _refresh_all(self):
        self._refresh_collections_list()
        self._refresh_library_filter_dropdowns()
        self.apply_filters()
        self.refresh_stats()

    def _refresh_collections_list(self):
        self.collections_list.delete(0, "end")
        if not self.books:
            return
        all_cols = sorted({c for b in self.books for c in b.collections if c})
        for c in all_cols:
            self.collections_list.insert("end", c)

    def _refresh_library_filter_dropdowns(self):
        # Populate Tag + Media dropdown options based on the loaded library
        if not self.books:
            self.tag_filter.configure(values=["Any"])
            self.media_filter.configure(values=["Any"])
            self.tag_filter_var.set("Any")
            self.media_filter_var.set("Any")
            return

        all_tags = sorted({t for b in self.books for t in b.tags if t})
        all_media = sorted({f for b in self.books for f in b.formats if f})

        tag_values = ["Any"] + all_tags
        media_values = ["Any"] + all_media

        self.tag_filter.configure(values=tag_values)
        self.media_filter.configure(values=media_values)

        # Keep current selection if still valid
        if self.tag_filter_var.get() not in tag_values:
            self.tag_filter_var.set("Any")
        if self.media_filter_var.get() not in media_values:
            self.media_filter_var.set("Any")

    # ----------------------------
    # Book details
    # ----------------------------


    def show_book_details(self, b: Book):
        win = tk.Toplevel(self)
        win.title(b.title or "Book Details")
        win.geometry("860x560")
        win.minsize(740, 520)

        outer = ttk.Frame(win, padding=10)
        outer.pack(fill="both", expand=True)

        left = ttk.Frame(outer)
        left.pack(side="left", fill="y")

        cover_lbl = ttk.Label(left, text="", anchor="center")
        cover_lbl.pack(anchor="n", pady=(0, 8))

        right = ttk.Frame(outer)
        right.pack(side="left", fill="both", expand=True, padx=(12, 0))

        txt = tk.Text(right, wrap="word")
        txt.pack(fill="both", expand=True)

        txt.insert("1.0", self._book_text(b))
        txt.configure(state="disabled")

        self._set_cover_label_for_book(cover_lbl, b, key=f"details:{b.books_id}:{id(win)}", details_text=txt)

    # ----- Covers -----

    def _clear_cover_label(self, label: Optional[ttk.Label], key: str):
        if label is None:
            return
        try:
            label.configure(image="", text="")
            if hasattr(label, "image"):
                label.image = None
        except Exception:
            pass
        if key in self._img_refs:
            self._img_refs.pop(key, None)

    def _maybe_persist_summary(self, book: Book, new_summary: str) -> None:
        """
        If enabled and the book's summary is empty, save a fetched summary back
        into the currently loaded JSON file (and update the in-memory Book).

        Creates a one-time .bak backup alongside the JSON before first write.
        """
        try:
            if not self.settings.get("auto_fill_summary", True):
                return
            if not new_summary:
                return
            if (book.summary or "").strip():
                return
            if not self.current_path:
                return

            book.summary = new_summary.strip()
            book.summary_checked = True

            path = self.current_path
            backup = path + ".bak"
            if (not os.path.exists(backup)) and os.path.exists(path):
                try:
                    import shutil
                    shutil.copy2(path, backup)
                except Exception:
                    pass

            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return

            key = str(book.books_id)
            raw = data.get(key)
            if isinstance(raw, dict) and (not str(raw.get("summary") or "").strip()):
                raw["summary"] = book.summary
                raw["summary_checked"] = True
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            return

    def _maybe_mark_summary_checked(self, book: Book) -> None:
        """Persist a 'summary_checked' marker so we don't repeatedly ping the server
        when no summary is available for a given book."""
        try:
            if not self.settings.get("auto_fill_summary", True):
                return
            if (book.summary or "").strip():
                # If summary exists, _maybe_persist_summary handles persistence.
                return
            if getattr(book, "summary_checked", False):
                return
            if not self.current_path:
                return

            book.summary_checked = True
            path = self.current_path
            backup = path + ".bak"
            if (not os.path.exists(backup)) and os.path.exists(path):
                try:
                    import shutil
                    shutil.copy2(path, backup)
                except Exception:
                    pass

            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return

            key = str(book.books_id)
            raw = data.get(key)
            if isinstance(raw, dict):
                raw["summary_checked"] = True
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            return


    def _set_cover_label_for_book(self, label: Optional[ttk.Label], b: Book, key: str, details_text: Optional[tk.Text] = None):
        if label is None:
            return
        if not self.covers_enabled or not self.cover_cache:
            label.configure(text="")
            return
        isbn = (b.isbn or "").strip()
        # Even if ISBN is missing, cover_cache can fall back to title/author search.
        self._set_cover_label_for_isbn(
            label,
            isbn,
            key=key,
            title=b.title,
            author=b.display_author,
            book=b,
            details_text=details_text,
        )


    def _set_cover_label_for_isbn(
        self,
        label: ttk.Label,
        isbn: str,
        key: str,
        title: str = "",
        author: str = "",
        book: Optional[Book] = None,
        details_text: Optional[tk.Text] = None,
    ):
        if not self.covers_enabled or not self.cover_cache:
            label.configure(text="")
            return

        # show a quick loading hint
        label.configure(text="Loading cover…", image="")
        if hasattr(label, "image"):
            label.image = None

        def on_done_extras(extras: Any):
            # extras can be either a dict {"path": ..., "summary": ...} or a raw path
            path: Optional[str] = None
            summary: str = ""
            if isinstance(extras, dict):
                path = extras.get("path")
                summary = str(extras.get("summary") or "")
            else:
                path = extras

            def apply_on_ui_thread():
                # Persist summary even if there is no cover image or Pillow isn't installed.
                try:
                    if book is not None and summary:
                        self._maybe_persist_summary(book, summary)
                except Exception:
                    pass

                # If we asked for a summary but the source had none, mark it as checked
                # so we don't keep pinging the server on every open.
                try:
                    if book is not None and want_summary and (not str(summary).strip()):
                        self._maybe_mark_summary_checked(book)
                except Exception:
                    pass

                # If this cover load is tied to an open Book Details window, refresh its text
                # so the newly saved summary appears immediately.
                try:
                    if book is not None and details_text is not None and details_text.winfo_exists() and want_summary:
                        y = details_text.yview()
                        details_text.configure(state="normal")
                        details_text.delete("1.0", "end")
                        if key == "random" and getattr(self, "_random_last_pool_size", None) is not None:
                            details_text.insert("1.0", self._book_text(book, pool_size=getattr(self, "_random_last_pool_size")))
                        else:
                            details_text.insert("1.0", self._book_text(book))
                        details_text.configure(state="disabled")
                        if isinstance(y, tuple) and y:
                            details_text.yview_moveto(y[0])
                except Exception:
                    pass

                if not label.winfo_exists():
                    return

                if not path or not os.path.exists(path):
                    label.configure(text="(No cover found)", image="")
                    return
                if not _HAS_PIL or Image is None or ImageTk is None:
                    label.configure(text="(Install Pillow for covers)", image="")
                    return
                try:
                    img = Image.open(path)
                    # Fit within a reasonable preview box
                    img.thumbnail((220, 320))
                    photo = ImageTk.PhotoImage(img)
                    label.configure(image=photo, text="")
                    label.image = photo  # keep alive
                    self._img_refs[key] = photo
                except Exception:
                    label.configure(text="(Could not load cover)", image="")
                    return

            self.after(0, apply_on_ui_thread)

        try:
            want_summary = bool(
                book is not None
                and (not (book.summary or "").strip())
                and (not getattr(book, "summary_checked", False))
                and self.settings.get("auto_fill_summary", True)
            )

            # Prefer the richer API if available (supports ISBN + title/author fallback + summary)
            if hasattr(self.cover_cache, "fetch_async_extras"):
                self.cover_cache.fetch_async_extras(  # type: ignore
                    isbn=isbn,
                    size="L",
                    title=title,
                    author=author,
                    want_summary=want_summary,
                    on_done=on_done_extras,
                )
            else:
                # Back-compat
                try:
                    self.cover_cache.fetch_async(isbn, "L", on_done_extras, title=title, author=author)  # type: ignore
                except TypeError:
                    self.cover_cache.fetch_async(isbn, "L", on_done_extras)  # type: ignore
        except Exception:
            label.configure(text="(Cover fetch failed)", image="")


    def show_cover_help(self):
        msg = (
            "Cover images are fetched by ISBN and cached locally in a '.covers_cache' folder.\n\n"
            "Source: Open Library Covers API.\n\n"
            "Requirements:\n"
            "  - pip install pillow requests\n"
            "  - keep cover_cache.py next to this app file\n\n"
            "If a book has no ISBN or Open Library doesn't have a cover, you'll see '(No cover found)'."
        )
        messagebox.showinfo("Cover images", msg)


    def _book_text(self, b: Book, pool_size: Optional[int] = None) -> str:
        lines: List[str] = []
        if pool_size is not None:
            lines.append(f"Random pick from {pool_size:,} matching books\n")

        lines.append(f"Title: {b.title}")
        if b.display_author:
            lines.append(f"Author: {b.display_author}")
        if b.year:
            lines.append(f"Year: {b.year}")
        if b.pages:
            lines.append(f"Pages: {b.pages}")
        if getattr(b, "isbn", ""):
            lines.append(f"ISBN: {b.isbn}")
        if b.primary_format():
            lines.append(f"Format: {b.primary_format()}")
        if b.tags:
            lines.append(f"Tags: {', '.join(b.tags)}")
        if b.rating is not None:
            lines.append(f"Rating: {b.rating:g}")
        if b.collections:
            lines.append(f"Collections: {', '.join(b.collections)}")
        if b.genre:
            lines.append(f"Genres: {', '.join(b.genre)}")
        if b.series:
            lines.append(f"Series: {', '.join(b.series)}")
        if b.dateread:
            lines.append(f"Date read: {b.dateread}")
        if b.entrydate:
            lines.append(f"Entry date: {b.entrydate}")
        if b.publication:
            lines.append(f"Publication: {b.publication}")

        if b.summary:
            lines.append("\nSummary:\n" + b.summary)

        lines.append(f"\nbooks_id: {b.books_id}")
        return "\n".join(lines)

    # ----------------------------
    # About
    # ----------------------------

    def show_about(self):
        msg = (
            "Book Collection Stats\n\n"
            "Loads a LibraryThing JSON export and provides:\n"
            "• Sortable library view\n"
            "• Random picker with collection filters\n"
            "• Statistics + charts (if matplotlib installed)\n\n"
            "Tip: use Library → column headers to sort.\n\n"
            "Version 1.0.1"
        )
        messagebox.showinfo("About", msg)


def main():
    initial_path = None
    if len(sys.argv) > 1:
        initial_path = sys.argv[1]
    app = BookStatsApp(initial_path=initial_path)
    app.mainloop()


if __name__ == "__main__":
    main()
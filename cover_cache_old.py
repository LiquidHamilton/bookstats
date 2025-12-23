"""
cover_cache.py

Simple cover-fetching + disk cache.

Primary source: Open Library Covers API:
  https://covers.openlibrary.org/b/isbn/{ISBN}-{S|M|L}.jpg?default=false

Notes:
- Uses Pillow (PIL) for image decoding in the main GUI (recommended).
- Uses requests for HTTP.
"""

from __future__ import annotations

import os
import re
import threading
from dataclasses import dataclass
from typing import Callable, Optional

import requests


def normalize_isbn(isbn: str) -> str:
    """
    Normalize an ISBN string by stripping non ISBN chars (keeps X),
    uppercasing, and removing hyphens/spaces.
    """
    if not isbn:
        return ""
    s = isbn.strip().upper()
    # keep digits and X
    s = re.sub(r"[^0-9X]", "", s)
    return s


def choose_best_isbn(*candidates: str) -> str:
    """
    Prefer ISBN-13 (13 digits starting with 978/979), else ISBN-10.
    """
    cleaned = [normalize_isbn(c) for c in candidates if c]
    cleaned = [c for c in cleaned if len(c) in (10, 13)]
    if not cleaned:
        return ""
    for c in cleaned:
        if len(c) == 13 and (c.startswith("978") or c.startswith("979")):
            return c
    # fall back to first valid
    return cleaned[0]


@dataclass
class CoverCache:
    cache_dir: str
    timeout_s: float = 10.0
    user_agent: str = "BookStatsGUI/1.0 (+https://example.local)"
    # You can extend with more sources later.

    def __post_init__(self):
        os.makedirs(self.cache_dir, exist_ok=True)

    def cache_path(self, isbn: str, size: str) -> str:
        safe = normalize_isbn(isbn)
        size = (size or "L").upper()
        return os.path.join(self.cache_dir, f"{safe}_{size}.jpg")

    def openlibrary_url(self, isbn: str, size: str) -> str:
        size = (size or "L").upper()
        safe = normalize_isbn(isbn)
        return f"https://covers.openlibrary.org/b/isbn/{safe}-{size}.jpg?default=false"

    def get_cover_path(self, isbn: str, size: str = "L", force_refresh: bool = False) -> Optional[str]:
        """
        Returns a local file path for the cached cover, downloading if needed.
        Returns None if not found / download failed / ISBN missing.
        """
        safe = normalize_isbn(isbn)
        if not safe:
            return None

        out_path = self.cache_path(safe, size)
        if (not force_refresh) and os.path.exists(out_path) and os.path.getsize(out_path) > 0:
            return out_path

        url = self.openlibrary_url(safe, size)
        try:
            r = requests.get(
                url,
                timeout=self.timeout_s,
                headers={"User-Agent": self.user_agent},
            )
            if r.status_code == 200 and r.content:
                with open(out_path, "wb") as f:
                    f.write(r.content)
                return out_path
            # Open Library returns 404 if default=false and no cover is present.
            return None
        except Exception:
            return None

    def fetch_async(
        self,
        isbn: str,
        size: str,
        on_done: Callable[[Optional[str]], None],
    ) -> None:
        """
        Run get_cover_path in a background thread and call on_done(path_or_none).
        """
        def worker():
            path = self.get_cover_path(isbn, size=size)
            on_done(path)

        t = threading.Thread(target=worker, daemon=True)
        t.start()

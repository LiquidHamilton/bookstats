"""
cover_cache.py

Cover fetching + disk cache with ISBN-first lookup and a title/author fallback.

Primary source: Open Library Covers API:
  - By ISBN: https://covers.openlibrary.org/b/isbn/{ISBN}-{S|M|L}.jpg?default=false
  - By cover id: https://covers.openlibrary.org/b/id/{COVER_ID}-{S|M|L}.jpg?default=false

Fallback (when ISBN missing OR no ISBN cover is found):
  - Open Library Search API: https://openlibrary.org/search.json?title=...&author=...
    We request specific fields to avoid relying on defaults.

Notes:
- Uses requests for HTTP.
- The GUI can optionally use Pillow (PIL) to load/display the downloaded images.
"""

from __future__ import annotations

import hashlib
import os
import re
import threading
from dataclasses import dataclass
from typing import Callable, Dict, Optional, Tuple

import requests


def normalize_isbn(isbn: str) -> str:
    """
    Normalize an ISBN string by stripping non ISBN chars (keeps X),
    uppercasing, and removing hyphens/spaces.
    """
    if not isbn:
        return ""
    s = isbn.strip().upper()
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
    return cleaned[0]


def _stable_hash_key(*parts: str) -> str:
    """
    Produce a stable short key for caching non-ISBN lookups (title/author).
    """
    joined = "||".join((p or "").strip().lower() for p in parts)
    h = hashlib.sha1(joined.encode("utf-8")).hexdigest()
    return h[:16]


@dataclass
class CoverCache:
    cache_dir: str
    timeout_s: float = 10.0
    user_agent: str = "BookStatsGUI/1.0 (+local)"
    base_search_url: str = "https://openlibrary.org/search.json"

    def __post_init__(self):
        os.makedirs(self.cache_dir, exist_ok=True)

    # -----------------------------
    # Cache paths
    # -----------------------------

    def cache_path_isbn(self, isbn: str, size: str) -> str:
        safe = normalize_isbn(isbn)
        size = (size or "L").upper()
        return os.path.join(self.cache_dir, f"isbn_{safe}_{size}.jpg")

    def cache_path_coverid(self, cover_id: int, size: str) -> str:
        size = (size or "L").upper()
        return os.path.join(self.cache_dir, f"olid_{cover_id}_{size}.jpg")

    def cache_path_query(self, title: str, author: str, size: str) -> str:
        size = (size or "L").upper()
        key = _stable_hash_key(title, author)
        return os.path.join(self.cache_dir, f"q_{key}_{size}.jpg")

    # -----------------------------
    # Open Library URL helpers
    # -----------------------------

    def openlibrary_url_isbn(self, isbn: str, size: str) -> str:
        size = (size or "L").upper()
        safe = normalize_isbn(isbn)
        return f"https://covers.openlibrary.org/b/isbn/{safe}-{size}.jpg?default=false"

    def openlibrary_url_coverid(self, cover_id: int, size: str) -> str:
        size = (size or "L").upper()
        return f"https://covers.openlibrary.org/b/id/{cover_id}-{size}.jpg?default=false"

    # -----------------------------
    # Public API
    # -----------------------------

    def get_cover_path(
        self,
        isbn: str,
        *,
        size: str = "L",
        force_refresh: bool = False,
        title: Optional[str] = None,
        author: Optional[str] = None,
    ) -> Optional[str]:
        """
        Return a local file path for the cached cover, downloading if needed.

        Strategy:
          1) Try ISBN -> Covers API
          2) If ISBN is missing OR no cover found:
             search by (title, author) using /search.json
             - If we find a cover_i, fetch via Covers API by id
             - Else if we find an ISBN, try ISBN fetch again

        Returns None if nothing is found.
        """
        # 1) Try ISBN first if present
        safe_isbn = normalize_isbn(isbn)
        if safe_isbn:
            cached = self.cache_path_isbn(safe_isbn, size)
            if (not force_refresh) and os.path.exists(cached) and os.path.getsize(cached) > 0:
                return cached

            url = self.openlibrary_url_isbn(safe_isbn, size)
            path = self._download_to(url, cached, allow_404=True)
            if path:
                return path

        # 2) Fallback: search by title/author (requires at least a title)
        if not title and not author:
            return None

        found = self._search_openlibrary_best(title=title or "", author=author or "")
        if not found:
            return None

        cover_id, found_isbn = found

        if cover_id is not None:
            # Cache by cover id (most reliable fallback)
            cached = self.cache_path_coverid(cover_id, size)
            if (not force_refresh) and os.path.exists(cached) and os.path.getsize(cached) > 0:
                return cached

            url = self.openlibrary_url_coverid(cover_id, size)
            return self._download_to(url, cached, allow_404=True)

        if found_isbn:
            # Avoid infinite recursion by NOT passing title/author again here.
            return self.get_cover_path(found_isbn, size=size, force_refresh=force_refresh)

        return None

    def fetch_async(
        self,
        isbn: str,
        size: str,
        on_done: Callable[[Optional[str]], None],
        *,
        title: Optional[str] = None,
        author: Optional[str] = None,
    ) -> None:
        """
        Run get_cover_path in a background thread and call on_done(path_or_none).
        """
        def worker():
            path = self.get_cover_path(isbn, size=size, title=title, author=author)
            on_done(path)

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    # -----------------------------
    # Internals
    # -----------------------------

    def _download_to(self, url: str, out_path: str, *, allow_404: bool) -> Optional[str]:
        try:
            r = requests.get(
                url,
                timeout=self.timeout_s,
                headers={"User-Agent": self.user_agent},
            )
            if allow_404 and r.status_code == 404:
                return None
            if r.status_code == 200 and r.content:
                with open(out_path, "wb") as f:
                    f.write(r.content)
                return out_path
            return None
        except Exception:
            return None

    def _search_openlibrary_best(self, *, title: str, author: str) -> Optional[Tuple[Optional[int], str]]:
        """
        Query Open Library search and return a best match:
          - (cover_id, "") if a cover_i exists
          - (None, isbn) if only ISBN exists
          - None if nothing useful found

        We request fields explicitly to avoid relying on Open Library's default fields.
        """
        q_title = (title or "").strip()
        q_author = (author or "").strip()
        if not q_title and not q_author:
            return None

        params: Dict[str, str] = {
            "title": q_title,
            "author": q_author,
            "limit": "10",
            "fields": "cover_i,isbn,title,author_name,first_publish_year",
        }

        try:
            r = requests.get(
                self.base_search_url,
                params=params,
                timeout=self.timeout_s,
                headers={"User-Agent": self.user_agent},
            )
            if r.status_code != 200:
                return None
            data = r.json()
            docs = data.get("docs") or []
            if not isinstance(docs, list):
                return None

            # Prefer docs with cover_i
            for d in docs:
                if not isinstance(d, dict):
                    continue
                cover_i = d.get("cover_i")
                if isinstance(cover_i, int) and cover_i > 0:
                    return (cover_i, "")

            # Otherwise, try ISBNs in results
            for d in docs:
                if not isinstance(d, dict):
                    continue
                isbns = d.get("isbn")
                if isinstance(isbns, list) and isbns:
                    best = choose_best_isbn(*[str(x) for x in isbns])
                    if best:
                        return (None, best)

            return None
        except Exception:
            return None

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

    def _download_to(self, url: str, out_path: str, *, allow_404: bool = False, force_refresh: bool = False) -> Optional[str]:
        try:
            if (not force_refresh) and os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                return out_path

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
    def _search_openlibrary_best_extras(self, *, title: str, author: str) -> Optional[Dict[str, Any]]:
        """
        Like _search_openlibrary_best, but returns the best-matching doc so we can
        extract cover + work keys for summary lookups.
        """
        q_title = (title or "").strip()
        q_author = (author or "").strip()
        if not q_title and not q_author:
            return None

        params: Dict[str, str] = {
            "limit": "10",
            "fields": "title,author_name,cover_i,isbn,key,edition_key,first_publish_year",
        }
        if q_title:
            params["title"] = q_title
        if q_author:
            params["author"] = q_author

        try:
            r = requests.get(
                "https://openlibrary.org/search.json",
                params=params,
                timeout=self.timeout_s,
                headers={"User-Agent": self.user_agent},
            )
            if r.status_code != 200:
                return None
            payload = r.json()
            docs = payload.get("docs")
            if not isinstance(docs, list) or not docs:
                return None

            # Prefer: has cover_i and matches roughly
            for d in docs:
                if not isinstance(d, dict):
                    continue
                if d.get("cover_i"):
                    return d

            # Otherwise return the first doc that has ISBNs
            for d in docs:
                if not isinstance(d, dict):
                    continue
                isbns = d.get("isbn")
                if isinstance(isbns, list) and isbns:
                    return d

            return None
        except Exception:
            return None

    def _extract_description_from_work_json(self, work_json: Any) -> str:
        desc = work_json.get("description") if isinstance(work_json, dict) else None
        if isinstance(desc, str):
            return desc.strip()
        if isinstance(desc, dict):
            val = desc.get("value")
            if isinstance(val, str):
                return val.strip()
        return ""

    def _fetch_work_description(self, work_key: str) -> str:
        if not work_key:
            return ""
        try:
            key = str(work_key).strip()
            # Open Library keys are usually like "/works/OL123W" but sometimes show up as "OL123W" or "works/OL123W".
            if key.startswith("works/"):
                key = "/" + key
            if not key.startswith("/"):
                # Assume bare work id
                key = "/works/" + key
            # If a full URL was accidentally passed, keep only the path
            if key.startswith("http://") or key.startswith("https://"):
                # very defensive; shouldn't normally happen
                key = "/" + key.split("openlibrary.org", 1)[-1].lstrip("/")
            # work_key like '/works/OL123W'
            url = "https://openlibrary.org" + key + ".json"
            r = requests.get(url, timeout=self.timeout_s, headers={"User-Agent": self.user_agent})
            if r.status_code != 200:
                return ""
            return self._extract_description_from_work_json(r.json())
        except Exception:
            return ""

    def _fetch_edition_description(self, edition_key: str) -> str:
        if not edition_key:
            return ""
        try:
            # edition_key like 'OL123M' (sometimes), or sometimes '/books/OL123M'
            key = str(edition_key).strip()
            if key.startswith("books/"):
                key = key[len("books/"):]
            if key.startswith("/books/"):
                key = key[len("/books/"):]
            if key.startswith("/"):
                key = key.lstrip("/")
            url = f"https://openlibrary.org/books/{key}.json"
            r = requests.get(url, timeout=self.timeout_s, headers={"User-Agent": self.user_agent})
            if r.status_code != 200:
                return ""
            data = r.json()
            # editions can include description directly
            desc = data.get("description")
            if isinstance(desc, str):
                return desc.strip()
            if isinstance(desc, dict) and isinstance(desc.get("value"), str):
                return str(desc.get("value")).strip()
            return ""
        except Exception:
            return ""

    def _fetch_description_by_isbn(self, isbn: str) -> str:
        safe = normalize_isbn(isbn)
        if not safe:
            return ""
        try:
            # edition lookup by ISBN
            url = f"https://openlibrary.org/isbn/{safe}.json"
            r = requests.get(url, timeout=self.timeout_s, headers={"User-Agent": self.user_agent})
            if r.status_code != 200:
                return ""
            data = r.json()
            # direct description
            desc = data.get("description")
            if isinstance(desc, str):
                return desc.strip()
            if isinstance(desc, dict) and isinstance(desc.get("value"), str):
                return str(desc.get("value")).strip()
            works = data.get("works")
            if isinstance(works, list) and works:
                wk = works[0].get("key") if isinstance(works[0], dict) else None
                if isinstance(wk, str):
                    return self._fetch_work_description(wk)
            return ""
        except Exception:
            return ""

    def get_cover_and_summary(
        self,
        isbn: str = "",
        size: str = "L",
        *,
        title: str = "",
        author: str = "",
        force_refresh: bool = False,
        want_summary: bool = True,
    ) -> Dict[str, Any]:
        """
        Returns {'path': <local path or None>, 'summary': <string>}.
        - Tries ISBN cover first (if provided)
        - Falls back to searching by title/author when needed
        - Optionally fetches a description/summary from Open Library
        """
        safe_isbn = normalize_isbn(isbn)
        isbn_summary = ""

        # 1) ISBN attempt
        if safe_isbn:
            out_path = self.cache_path_isbn(safe_isbn, size)
            if (not force_refresh) and os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                summary = self._fetch_description_by_isbn(safe_isbn) if want_summary else ""
                return {"path": out_path, "summary": summary}

            url = self.openlibrary_url_isbn(safe_isbn, size)
            path = self._download_to(url, out_path, force_refresh=force_refresh)
            if path:
                summary = self._fetch_description_by_isbn(safe_isbn) if want_summary else ""
                return {"path": path, "summary": summary}

            # No cover found via ISBN; still try to fetch a description by ISBN.
            if want_summary:
                isbn_summary = self._fetch_description_by_isbn(safe_isbn)

        # 2) Fallback search by title/author
        doc = self._search_openlibrary_best_extras(title=title, author=author)
        if not doc:
            return {"path": None, "summary": isbn_summary}

        summary = isbn_summary
        if want_summary:
            wk = doc.get("key") if isinstance(doc.get("key"), str) else ""
            edk = ""
            ek = doc.get("edition_key")
            if isinstance(ek, list) and ek:
                edk = str(ek[0])
            elif isinstance(ek, str):
                edk = ek
            summary = self._fetch_work_description(wk) or self._fetch_edition_description(edk)

        # Try cover_i
        cover_i = doc.get("cover_i")
        if isinstance(cover_i, int):
            out_path = self.cache_path_coverid(cover_i, size)
            if (not force_refresh) and os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                return {"path": out_path, "summary": summary}
            url = self.openlibrary_url_coverid(cover_i, size)
            path = self._download_to(url, out_path, force_refresh=force_refresh)
            if path:
                return {"path": path, "summary": summary}

        # Try ISBNs from doc
        isbns = doc.get("isbn")
        if isinstance(isbns, list) and isbns:
            best = choose_best_isbn(*[str(x) for x in isbns])
            if best:
                out_path = self.cache_path_isbn(best, size)
                if (not force_refresh) and os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                    return {"path": out_path, "summary": summary}
                url = self.openlibrary_url_isbn(best, size)
                path = self._download_to(url, out_path, force_refresh=force_refresh)
                if path:
                    return {"path": path, "summary": summary}

        return {"path": None, "summary": summary}

    def fetch_async_extras(
        self,
        *,
        isbn: str = "",
        size: str = "L",
        title: str = "",
        author: str = "",
        want_summary: bool = False,
        force_refresh: bool = False,
        on_done: Optional[Callable[[Any], None]] = None,
    ) -> None:
        """
        Async version of get_cover_and_summary.
        Calls on_done(extras_dict) on completion.
        """
        def worker():
            result = self.get_cover_and_summary(
                isbn=isbn,
                size=size,
                title=title,
                author=author,
                want_summary=want_summary,
                force_refresh=force_refresh,
            )
            if on_done:
                on_done(result)

        t = threading.Thread(target=worker, daemon=True)
        t.start()

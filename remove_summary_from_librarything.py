#!/usr/bin/env python3
"""
Empty out all "summary" fields in a LibraryThing-style JSON export.

Input format (top-level object):
{
  "243750757": { "title": "...", "summary": "...", ... },
  "243750758": { ... },
  ...
}

By default, writes to "<input>.summary_emptied.json".
Use --in-place to overwrite the input file.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict


def empty_summaries(data: Any) -> int:
    """
    Sets data[book_id]["summary"] = "" for each top-level book record dict.
    Returns the number of records touched.
    """
    if not isinstance(data, dict):
        raise ValueError("Expected the JSON root to be an object/dict.")

    touched = 0
    for _, book in data.items():
        if isinstance(book, dict):
            book["summary"] = ""
            touched += 1
    return touched


def main() -> int:
    parser = argparse.ArgumentParser(description='Make all "summary" fields empty in a JSON file.')
    parser.add_argument("input", help="Path to input JSON file")
    parser.add_argument(
        "-o",
        "--output",
        help="Path to output JSON file (default: <input>.summary_emptied.json)",
        default=None,
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Overwrite the input file (creates a .bak backup first)",
    )
    args = parser.parse_args()

    in_path = Path(args.input)

    if not in_path.exists():
        raise FileNotFoundError(f"Input file not found: {in_path}")

    with in_path.open("r", encoding="utf-8") as f:
        data: Dict[str, Any] = json.load(f)

    touched = empty_summaries(data)

    if args.in_place:
        backup_path = in_path.with_suffix(in_path.suffix + ".bak")
        backup_path.write_text(in_path.read_text(encoding="utf-8"), encoding="utf-8")
        out_path = in_path
    else:
        out_path = Path(args.output) if args.output else in_path.with_suffix(in_path.suffix + ".summary_emptied.json")

    with out_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"Updated {touched} book records. Wrote: {out_path}")
    if args.in_place:
        print(f"Backup created: {backup_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


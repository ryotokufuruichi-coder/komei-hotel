#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Refresh the <lastmod> dates in sitemap.xml.

Default behavior: rewrite every <lastmod> to today's date (UTC). This is the
"monthly freshness ping" — Google rewards sites that publish new lastmod
values as a small but consistent signal of activity.

Why this matters:
  Google re-crawls more aggressively when lastmod values appear to advance.
  Even when the underlying page content hasn't changed, a refreshed lastmod
  invites Googlebot to re-verify and tends to keep the cached title /
  description fresh in SERP.

Important nuance:
  This script intentionally bumps ALL entries unconditionally. Some teams
  prefer per-file mtime comparison so that lastmod reflects "real" edits,
  but for a small static site like this, the overhead of accurate
  per-page tracking outweighs the benefit. Monthly blanket refresh is a
  documented, low-risk pattern.

Usage:
    python scripts/refresh_sitemap.py             # apply today's date
    python scripts/refresh_sitemap.py --check     # exit 1 if any change needed
    python scripts/refresh_sitemap.py --dry-run   # print diff, don't write
    python scripts/refresh_sitemap.py --date YYYY-MM-DD   # use a specific date

Designed to be invoked from GitHub Actions on a monthly schedule (see
.github/workflows/refresh-sitemap.yml).
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import date, datetime, timezone
from pathlib import Path

SITE_ROOT = Path(__file__).resolve().parent.parent
SITEMAP_PATH = SITE_ROOT / "sitemap.xml"

LASTMOD_RE = re.compile(r"(<lastmod>)([^<]+)(</lastmod>)")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def replace_lastmods(xml: str, new_date: str) -> tuple[str, int]:
    """Replace every <lastmod>…</lastmod> with the given date. Returns (new_xml, changes)."""
    changes = 0

    def _sub(m: re.Match) -> str:
        nonlocal changes
        if m.group(2) != new_date:
            changes += 1
        return f"{m.group(1)}{new_date}{m.group(3)}"

    new_xml = LASTMOD_RE.sub(_sub, xml)
    return new_xml, changes


def emit_diff(old: str, new: str) -> None:
    import difflib
    diff = difflib.unified_diff(
        old.splitlines(keepends=True),
        new.splitlines(keepends=True),
        fromfile="sitemap.xml (current)",
        tofile="sitemap.xml (after refresh_sitemap.py)",
        n=2,
    )
    sys.stdout.writelines(diff)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--date", default=None,
                        help="Override target date (YYYY-MM-DD). Defaults to today UTC.")
    parser.add_argument("--dry-run", action="store_true",
                        help="Print the diff and exit; don't write the file.")
    parser.add_argument("--check", action="store_true",
                        help="Exit 1 if any change is required, else 0. Does not write.")
    args = parser.parse_args()

    if args.date:
        if not DATE_RE.fullmatch(args.date):
            raise SystemExit(f"--date must be YYYY-MM-DD, got {args.date!r}")
        target_date = args.date
    else:
        target_date = datetime.now(timezone.utc).date().isoformat()

    if not SITEMAP_PATH.exists():
        raise SystemExit(f"sitemap.xml not found at {SITEMAP_PATH}")

    xml = SITEMAP_PATH.read_text(encoding="utf-8")
    new_xml, changes = replace_lastmods(xml, target_date)

    if changes == 0:
        print(f"sitemap.xml already at lastmod {target_date} — nothing to do.",
              file=sys.stderr)
        return 0

    if args.check:
        print(f"sitemap.xml has {changes} entries needing refresh to {target_date}.",
              file=sys.stderr)
        return 1

    if args.dry_run:
        emit_diff(xml, new_xml)
        return 0

    SITEMAP_PATH.write_text(new_xml, encoding="utf-8")
    print(f"✓ refreshed {changes} <lastmod> entries → {target_date}",
          file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())

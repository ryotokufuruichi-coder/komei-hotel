#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Inject AggregateRating + Review into the LodgingBusiness JSON-LD on index.html
from the public_reviews GAS API.

Idempotent. Safe to run repeatedly — no-op when nothing changed.
Skips injection when the API reports count == 0 (Schema.org rejects
AggregateRating with reviewCount=0 / no ratingValue).

Usage:
    python scripts/inject_reviews.py                # update index.html in place
    python scripts/inject_reviews.py --dry-run      # show diff, don't write
    python scripts/inject_reviews.py --check        # exit 1 if update is needed
    python scripts/inject_reviews.py --soft-fail    # exit 0 on network errors

Designed to be invoked from GitHub Actions on a daily schedule.
See .github/workflows/update-reviews.yml.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import urllib.error
import urllib.request
from pathlib import Path

API_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbx6JO_l8Hhb5pQCIZkBTgEkChylOcD4e32JuCcCG-XtvuaQ6I5qacd7UuyeHQ8fnOZR"
    "/exec?action=public_reviews"
)

SITE_ROOT = Path(__file__).resolve().parent.parent
INDEX_PATH = SITE_ROOT / "index.html"

# Cap how much we inline so the JSON-LD block stays a reasonable size for crawlers
MAX_REVIEWS = 5
MAX_REVIEW_TEXT = 500

# Match the JSON-LD <script> that contains LodgingBusiness. The regex is
# anchored on the @type to avoid matching the Organization / FAQPage / Breadcrumb
# blocks below it.
LODGING_RE = re.compile(
    r'(<script type="application/ld\+json">\s*)'
    r'(\{[\s\S]+?"@type":\s*"LodgingBusiness"[\s\S]+?\n\})'
    r'(\s*</script>)'
)


def fetch_reviews(timeout: int = 30) -> dict:
    req = urllib.request.Request(
        API_URL, headers={"User-Agent": "komei-seo-injector/1.0"}
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def build_payload(data: dict) -> tuple[dict | None, list[dict] | None]:
    """Return (aggregateRating, review[]) or (None, None) when not injectable."""
    if not data.get("ok"):
        return None, None
    count = int(data.get("count") or 0)
    if count <= 0:
        return None, None

    avg = data.get("avg") or {}
    overall = float(avg.get("overall") or 0)
    if overall <= 0:
        return None, None

    aggregate = {
        "@type": "AggregateRating",
        "ratingValue": round(overall, 1),
        "reviewCount": count,
        "bestRating": 5,
        "worstRating": 1,
    }

    reviews: list[dict] = []
    for r in (data.get("reviews") or [])[:MAX_REVIEWS]:
        review = {
            "@type": "Review",
            "reviewRating": {
                "@type": "Rating",
                "ratingValue": int(r.get("overall") or 5),
                "bestRating": 5,
                "worstRating": 1,
            },
            "author": {
                "@type": "Person",
                "name": (r.get("rep_name") or "Guest").strip() or "Guest",
            },
        }
        comment = (r.get("comment") or "").strip()
        if comment:
            if len(comment) > MAX_REVIEW_TEXT:
                comment = comment[: MAX_REVIEW_TEXT - 1].rstrip() + "…"
            review["reviewBody"] = comment
        created = r.get("created_at")
        if created:
            # ISO date prefix only (YYYY-MM-DD)
            review["datePublished"] = str(created)[:10]
        reviews.append(review)

    return aggregate, (reviews or None)


def patch_html(
    html: str,
    aggregate: dict | None,
    reviews: list[dict] | None,
) -> tuple[str, bool]:
    """Return (new_html, changed)."""
    m = LODGING_RE.search(html)
    if not m:
        raise SystemExit(
            "ERROR: LodgingBusiness JSON-LD block not found in index.html. "
            "Aborting to avoid corrupting the file."
        )

    try:
        obj = json.loads(m.group(2))
    except json.JSONDecodeError as e:
        raise SystemExit(
            f"ERROR: existing LodgingBusiness JSON-LD is not valid JSON: {e}"
        )

    if aggregate is None:
        # No usable review data — strip any pre-existing keys so we never
        # display a stale rating after reviews are deleted/reset.
        changed_local = False
        for key in ("aggregateRating", "review"):
            if key in obj:
                obj.pop(key)
                changed_local = True
        if not changed_local:
            return html, False
    else:
        obj["aggregateRating"] = aggregate
        if reviews:
            obj["review"] = reviews
        elif "review" in obj:
            obj.pop("review")

    new_json = json.dumps(obj, ensure_ascii=False, indent=2)
    new_block = m.group(1) + new_json + m.group(3)
    new_html = html[: m.start()] + new_block + html[m.end():]
    return new_html, new_html != html


def emit_diff(old: str, new: str) -> None:
    import difflib

    diff = difflib.unified_diff(
        old.splitlines(keepends=True),
        new.splitlines(keepends=True),
        fromfile="index.html (current)",
        tofile="index.html (after inject_reviews.py)",
        n=3,
    )
    sys.stdout.writelines(diff)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--dry-run", action="store_true", help="Print the unified diff and exit"
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Exit 1 if index.html would change, else 0. Does not write.",
    )
    parser.add_argument(
        "--soft-fail",
        action="store_true",
        help="Exit 0 on network/API errors (useful for unattended CI runs)",
    )
    args = parser.parse_args()

    try:
        print("Fetching reviews from GAS API…", file=sys.stderr)
        data = fetch_reviews()
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError) as e:
        msg = f"network error fetching review API: {e}"
        if args.soft_fail:
            print(f"WARN: {msg} (--soft-fail, exit 0)", file=sys.stderr)
            return 0
        print(f"ERROR: {msg}", file=sys.stderr)
        return 2

    ok = data.get("ok")
    count = data.get("count")
    print(f"  api response: ok={ok} count={count}", file=sys.stderr)

    aggregate, reviews = build_payload(data)
    if aggregate is None:
        print(
            "  → no usable rating yet (count=0 or overall=0); will strip any "
            "stale aggregateRating/review keys",
            file=sys.stderr,
        )
    else:
        print(
            f"  → ratingValue={aggregate['ratingValue']} "
            f"reviewCount={aggregate['reviewCount']} "
            f"reviews_inlined={len(reviews) if reviews else 0}",
            file=sys.stderr,
        )

    html = INDEX_PATH.read_text(encoding="utf-8")
    new_html, changed = patch_html(html, aggregate, reviews)

    if not changed:
        print("index.html already up to date — nothing to do.", file=sys.stderr)
        return 0

    if args.check:
        print("index.html would change (run without --check to apply).", file=sys.stderr)
        return 1

    if args.dry_run:
        emit_diff(html, new_html)
        return 0

    INDEX_PATH.write_text(new_html, encoding="utf-8")
    print(f"✓ Updated {INDEX_PATH.relative_to(SITE_ROOT)}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())

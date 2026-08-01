#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate language-subdirectory versions of a multilingual page.

The site authors pages with inline <span class="lang-xx"> blocks and a
client-side switcher. For SEO, each language needs its own indexable URL.
This script emits /<lang>/<page> from a root source page by:
  - setting <html lang="..">
  - self-canonical + hreflang alternates (subdirectory URLs)
  - og:url / og:locale for the language
  - rewriting RELATIVE asset + page references (in attributes AND JS
    strings) to root-absolute (/css, /images, /booking.html, ...) so they
    resolve correctly from the /<lang>/ path
  - turning the language <select> into a navigation to /<lang>/<page>

Additive + idempotent: rerunning overwrites the generated files only.
Usage:  python scripts/gen_lang.py <source.html> <slug>
        slug '' means the index page  (=> /<lang>/)
        e.g. python scripts/gen_lang.py index.html ""
"""
from __future__ import annotations
import re, sys
from pathlib import Path

SITE = Path(__file__).resolve().parent.parent
BASE = "https://komei.yoshinarcorp.com"
LANGS = {
    "en": {"html": "en", "hreflang": "en",      "locale": "en_US"},
    "ja": {"html": "ja", "hreflang": "ja",      "locale": "ja_JP"},
    "zh": {"html": "zh", "hreflang": "zh-Hans", "locale": "zh_CN"},
    "ko": {"html": "ko", "hreflang": "ko",      "locale": "ko_KR"},
}
ORDER = ["ja", "en", "zh", "ko"]


def hreflang_block(slug: str) -> str:
    lines = [
        f'<link rel="alternate" hreflang="{LANGS[c]["hreflang"]}" href="{BASE}/{c}/{slug}">'
        for c in ORDER
    ]
    lines.append(f'<link rel="alternate" hreflang="x-default" href="{BASE}/en/{slug}">')
    return "\n".join(lines)


def _abs_attr(m: re.Match) -> str:
    attr, val = m.group(1), m.group(2)
    if re.match(r"https?:|//|#|mailto:|tel:|data:|javascript:|/", val):
        return m.group(0)
    return f'{attr}="/{val}"'


def generate(src: str, lang: str, slug: str) -> str:
    info = LANGS[lang]
    s = Path(SITE / src).read_text(encoding="utf-8")
    url = f"{BASE}/{lang}/{slug}"

    n = 0
    s, n = re.subn(r'<html lang="en">', f'<html lang="{info["html"]}">', s, count=1); assert n == 1, "html lang"
    s, n = re.subn(r'<link rel="canonical" href="[^"]*">', f'<link rel="canonical" href="{url}">', s, count=1); assert n == 1, "canonical"
    s, n = re.subn(r'(?:<!-- hreflang.*?-->\s*)?(?:<link rel="alternate"[^>]*>\s*){5}', hreflang_block(slug) + "\n", s, count=1, flags=re.S); assert n == 1, "hreflang"
    s, n = re.subn(r'(<meta property="og:url" content=")[^"]*(">)', rf"\g<1>{url}\g<2>", s, count=1); assert n == 1, "og:url"
    s = s.replace('<meta property="og:locale" content="en_US">', f'<meta property="og:locale" content="{info["locale"]}">', 1)

    # rewrite relative asset/page refs in attributes, then in JS strings
    s = re.sub(r'(href|src)="([^"]+)"', _abs_attr, s)
    s = re.sub(r"""(['"])([a-z][a-z0-9_\-]*\.html)""", r"\1/\2", s)

    # language switcher -> navigate to /<lang>/<slug>
    target = "'/'+this.value+'/'" if slug == "" else "'/'+this.value+'/" + slug + "'"
    s, n = re.subn(r'onchange="setLang\(this\.value\)"', f'onchange="location.href={target}"', s); assert n >= 1, "switcher"
    s = s.replace('value="en" selected', 'value="en"')
    s, n = re.subn(rf'(<option value="{lang}")>', r"\1 selected>", s, count=1); assert n == 1, "selected option"
    return s


def main() -> int:
    src = sys.argv[1] if len(sys.argv) > 1 else "index.html"
    slug = sys.argv[2] if len(sys.argv) > 2 else ""
    for lang in LANGS:
        out = generate(src, lang, slug)
        d = SITE / lang
        d.mkdir(exist_ok=True)
        (d / (slug or "index.html")).write_text(out, encoding="utf-8")
        print(f"  wrote {lang}/{slug or 'index.html'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

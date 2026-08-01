#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate a new long-form travel guide page from a JSON config file.

Workflow:
    1. Copy _templates/example_guide_config.json to a new file, e.g.
       _templates/configs/2026-06-rainy-asakusa.json
    2. Edit the JSON (slug, title, sections, etc.)
    3. python scripts/new_guide.py _templates/configs/2026-06-rainy-asakusa.json
    4. Review the generated <slug>.html in the site root
    5. git add / commit / push / open PR

The script also:
- Adds a <url> entry to sitemap.xml (idempotent — won't duplicate)
- Prints suggested internal-link snippets you can paste into index.html
  and other related guides

Usage:
    python scripts/new_guide.py <config.json>
    python scripts/new_guide.py <config.json> --no-sitemap     # skip sitemap update
    python scripts/new_guide.py <config.json> --dry-run        # don't write files
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date
from html import escape
from pathlib import Path

SITE_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = SITE_ROOT / "_templates" / "guide_template.html"
SITEMAP_PATH = SITE_ROOT / "sitemap.xml"

# ─────────────────────────── Color presets ───────────────────────────
# Pick a preset that fits the topic. Pages with different topics get
# visually distinct treatment so they don't blur together.
COLOR_PRESETS = {
    # Warm/cultural — for: temple, food, neighborhood walks
    "amber": {
        "accent_color": "#f59e0b",
        "accent_color_dark": "#b45309",
        "accent_light": "#fde68a",
        "accent_bg_soft": "#fef3c7",
        "accent_border": "#fbcfe8",
        "hero_gradient": "linear-gradient(135deg, #f59e0b 0%, #ea580c 100%)",
    },
    # Festive/dramatic — for: fireworks, samba, night events
    "indigo-pink": {
        "accent_color": "#db2777",
        "accent_color_dark": "#be185d",
        "accent_light": "#fde68a",
        "accent_bg_soft": "#fdf2f8",
        "accent_border": "#fbcfe8",
        "hero_gradient": "linear-gradient(135deg, #1e3a8a 0%, #6d28d9 50%, #be185d 100%)",
    },
    # Cool/rainy/winter — for: rainy day guides, winter, indoor
    "teal": {
        "accent_color": "#0d9488",
        "accent_color_dark": "#0f766e",
        "accent_light": "#a7f3d0",
        "accent_bg_soft": "#ccfbf1",
        "accent_border": "#5eead4",
        "hero_gradient": "linear-gradient(135deg, #134e4a 0%, #0d9488 50%, #06b6d4 100%)",
    },
    # Spring/floral — for: sakura, blossoms, gentle springtime
    "rose": {
        "accent_color": "#e11d48",
        "accent_color_dark": "#9f1239",
        "accent_light": "#fecdd3",
        "accent_bg_soft": "#fff1f2",
        "accent_border": "#fda4af",
        "hero_gradient": "linear-gradient(135deg, #fda4af 0%, #f472b6 50%, #be185d 100%)",
    },
    # Autumn — for: foliage, autumn festivals
    "amber-rust": {
        "accent_color": "#c2410c",
        "accent_color_dark": "#7c2d12",
        "accent_light": "#fed7aa",
        "accent_bg_soft": "#ffedd5",
        "accent_border": "#fdba74",
        "hero_gradient": "linear-gradient(135deg, #7c2d12 0%, #c2410c 50%, #ea580c 100%)",
    },
}


# ───────────────────────── HTML rendering ─────────────────────────

def render_hero_meta(meta_items: list[str]) -> str:
    """Render the hero meta chip row (3-4 quick facts)."""
    if not meta_items:
        return ""
    spans = "\n      ".join(f'<span>{escape(m)}</span>' for m in meta_items)
    return f'<div class="meta">\n      {spans}\n    </div>'


def render_toc(sections: list[dict]) -> str:
    """Render the TOC block. Skipped if config sets toc=false."""
    items = []
    for i, s in enumerate(sections):
        sid = s.get("id") or f"section-{i+1}"
        toc_label = s.get("toc_label") or s.get("heading")
        items.append(f'    <li><a href="#{sid}">{escape(toc_label)}</a></li>')
    return (
        '<div class="toc">\n'
        '  <h2>What\'s in this guide</h2>\n'
        '  <ol>\n'
        + "\n".join(items)
        + "\n  </ol>\n"
        + "</div>"
    )


def _render_block(block: dict) -> str:
    """Render a single sub-element inside a section: paragraph, list, callout, h3."""
    btype = block.get("type", "p")
    if btype == "p":
        return f'  <p>{block["text"]}</p>'
    if btype == "h3":
        return f'  <h3>{escape(block["text"])}</h3>'
    if btype == "ul" or btype == "ol":
        items = "\n".join(f'    <li>{item}</li>' for item in block["items"])
        return f'  <{btype}>\n{items}\n  </{btype}>'
    if btype == "callout":
        return f'  <div class="callout">\n    <p>{block["text"]}</p>\n  </div>'
    raise ValueError(f"Unknown block type: {btype!r}")


def render_section(section: dict, idx: int) -> str:
    """Render one <section> with heading + accent + content blocks."""
    sid = section.get("id") or f"section-{idx+1}"
    parts = [f'<section class="section" id="{sid}">']
    if section.get("walk_time"):
        parts.append(f'  <span class="walk-time">{escape(section["walk_time"])}</span>')
    parts.append(f'  <h2>{escape(section["heading"])}</h2>')
    parts.append('  <div class="accent"></div>')
    for block in section.get("blocks", []):
        parts.append(_render_block(block))
    parts.append("</section>")
    return "\n".join(parts)


# ───────────────────────── JSON-LD rendering ─────────────────────────

def build_breadcrumb_jsonld(config: dict) -> dict:
    items = [{"@type": "ListItem", "position": 1, "name": "Home", "item": "https://komei.yoshinarcorp.com/"}]
    pos = 2
    for crumb in config.get("breadcrumb", []):
        items.append({
            "@type": "ListItem",
            "position": pos,
            "name": crumb["name"],
            "item": crumb["url"],
        })
        pos += 1
    items.append({
        "@type": "ListItem",
        "position": pos,
        "name": config.get("breadcrumb_self") or config["hero_h1"],
        "item": f"https://komei.yoshinarcorp.com/{config['slug']}.html",
    })
    return {
        "@context": "https://schema.org",
        "@type": "BreadcrumbList",
        "itemListElement": items,
    }


def build_travelguide_jsonld(config: dict) -> dict:
    return {
        "@context": "https://schema.org",
        "@type": "TravelGuide",
        "headline": config.get("schema_headline") or config["hero_h1"],
        "description": config["description"],
        "inLanguage": "en",
        "datePublished": config.get("date_published", date.today().isoformat()),
        "author": {"@type": "Organization", "name": "Komei Hotel 光明荘"},
        "publisher": {
            "@type": "Organization",
            "name": "Komei Hotel 光明荘",
            "url": "https://komei.yoshinarcorp.com/",
        },
        "mainEntityOfPage": f"https://komei.yoshinarcorp.com/{config['slug']}.html",
    }


def build_event_jsonld(config: dict) -> dict | None:
    e = config.get("event")
    if not e:
        return None
    block = {
        "@context": "https://schema.org",
        "@type": "Event",
        "name": e["name"],
        "description": e.get("description", config["description"]),
        "startDate": e["start_date"],
        "eventStatus": "https://schema.org/EventScheduled",
        "eventAttendanceMode": "https://schema.org/OfflineEventAttendanceMode",
        "url": f"https://komei.yoshinarcorp.com/{config['slug']}.html",
    }
    if e.get("end_date"):
        block["endDate"] = e["end_date"]
    if e.get("location"):
        block["location"] = e["location"]
    if e.get("organizer"):
        block["organizer"] = {"@type": "Organization", "name": e["organizer"]}
    if e.get("image"):
        block["image"] = e["image"]
    if e.get("alternate_name"):
        block["alternateName"] = e["alternate_name"]
    return block


def render_jsonld_blocks(config: dict) -> str:
    blocks = []
    ev = build_event_jsonld(config)
    if ev:
        blocks.append(ev)
    blocks.append(build_travelguide_jsonld(config))
    blocks.append(build_breadcrumb_jsonld(config))
    rendered = []
    for b in blocks:
        rendered.append(
            '<script type="application/ld+json">\n'
            + json.dumps(b, ensure_ascii=False, indent=2)
            + "\n</script>"
        )
    return "\n".join(rendered)


# ───────────────────────── Template assembly ─────────────────────────

def render_template(config: dict) -> str:
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    preset_name = config.get("color_preset", "amber")
    if preset_name not in COLOR_PRESETS:
        raise SystemExit(
            f"Unknown color_preset {preset_name!r}. "
            f"Choices: {', '.join(COLOR_PRESETS)}"
        )
    colors = COLOR_PRESETS[preset_name]

    show_toc = config.get("show_toc", True)
    sections_html = "\n\n".join(
        render_section(s, i) for i, s in enumerate(config["sections"])
    )

    subs = {
        "{{slug}}": config["slug"],
        "{{title}}": escape(config["title"]),
        "{{description}}": escape(config["description"]),
        "{{keywords}}": escape(config["keywords"]),
        "{{og_title}}": escape(config.get("og_title") or config["title"]),
        "{{og_description}}": escape(config.get("og_description") or config["description"]),
        "{{og_image}}": config.get("og_image", "p45.jpg"),
        "{{jsonld_blocks}}": render_jsonld_blocks(config),
        "{{hero_eyebrow}}": escape(config.get("hero_eyebrow", "")),
        "{{hero_h1}}": escape(config["hero_h1"]),
        "{{hero_lead}}": config.get("hero_lead", ""),
        "{{hero_meta_html}}": render_hero_meta(config.get("hero_meta", [])),
        "{{toc_block}}": render_toc(config["sections"]) if show_toc else "",
        "{{sections_html}}": sections_html,
        "{{cta_h2}}": escape(config.get("cta_h2", "Stay walkable. Book the whole house.")),
        "{{cta_lead}}": config.get("cta_lead",
            "Komei Hotel sleeps up to 10, has a private rooftop, and "
            "puts everything in this guide within walking distance. "
            "Book direct — save up to 10% vs Airbnb."),
        "{{cta_btn_label}}": escape(config.get("cta_btn_label", "Check availability →")),
        **{f"{{{{{k}}}}}": v for k, v in colors.items()},
    }

    result = template
    for k, v in subs.items():
        result = result.replace(k, v)
    return result


# ───────────────────────── Sitemap update ─────────────────────────

def update_sitemap(slug: str, dry_run: bool = False) -> tuple[bool, str]:
    """Append a <url> entry for the new slug to sitemap.xml.
    Idempotent: returns (False, reason) if the slug is already in the sitemap.
    """
    sitemap = SITEMAP_PATH.read_text(encoding="utf-8")
    target_url = f"https://komei.yoshinarcorp.com/{slug}.html"
    if target_url in sitemap:
        return False, "already in sitemap"

    today = date.today().isoformat()
    new_entry = (
        "  <url>\n"
        f"    <loc>{target_url}</loc>\n"
        f"    <lastmod>{today}</lastmod>\n"
        "    <changefreq>monthly</changefreq>\n"
        "    <priority>0.8</priority>\n"
        "  </url>\n"
    )

    # Insert before </urlset>
    new_sitemap = sitemap.replace("</urlset>", new_entry + "</urlset>")
    if dry_run:
        return True, "would insert (dry-run)"
    SITEMAP_PATH.write_text(new_sitemap, encoding="utf-8")
    return True, "inserted"


# ───────────────────────── Suggestions ─────────────────────────

def print_link_suggestions(config: dict) -> None:
    slug = config["slug"]
    sys.stderr.write("\n" + "=" * 70 + "\n")
    sys.stderr.write("Suggested internal-link snippets — paste these manually:\n")
    sys.stderr.write("=" * 70 + "\n\n")

    sys.stderr.write("1) Add a card to index.html's Nearby/Guide grid:\n\n")
    sys.stderr.write(f'''      <a href="{slug}.html" class="block bg-amber-50 border border-amber-200 rounded-2xl p-6 hover:shadow-lg hover:border-amber-300 transition">
        <div class="text-3xl mb-2">{config.get("emoji", "📖")}</div>
        <h3 class="font-bold text-lg mb-1">
          <span class="lang-ja">{config.get("card_title_ja", config["hero_h1"])}</span>
          <span class="lang-en">{config.get("card_title_en", config["hero_h1"])}</span>
        </h3>
        <p class="text-sm text-slate-600 mb-2">
          <span class="lang-en">{config.get("card_lead", config["description"][:120] + "...")}</span>
        </p>
      </a>
''')

    sys.stderr.write("\n2) Update footer of related guides to include this one:\n")
    sys.stderr.write(f'   · <a href="{slug}.html">{config.get("footer_label", config["hero_h1"])}</a>\n')

    sys.stderr.write("\n3) Don't forget to:\n")
    sys.stderr.write("   - Run Rich Results Test:\n")
    sys.stderr.write(f"     https://search.google.com/test/rich-results?url=https://komei.yoshinarcorp.com/{slug}.html\n")
    sys.stderr.write("   - Submit to GSC: URL Inspection → Request Indexing\n")
    sys.stderr.write("   - Resubmit sitemap.xml in GSC\n\n")


# ───────────────────────── CLI ─────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("config", help="Path to JSON config file")
    parser.add_argument("--no-sitemap", action="store_true",
                        help="Skip the sitemap.xml update")
    parser.add_argument("--dry-run", action="store_true",
                        help="Generate to stdout, don't write any files")
    parser.add_argument("--force", action="store_true",
                        help="Overwrite the output HTML if it exists")
    args = parser.parse_args()

    config_path = Path(args.config)
    if not config_path.exists():
        raise SystemExit(f"Config file not found: {config_path}")
    config = json.loads(config_path.read_text(encoding="utf-8"))

    # Validate required fields
    required = ["slug", "title", "description", "keywords", "hero_h1", "sections"]
    missing = [k for k in required if k not in config]
    if missing:
        raise SystemExit(f"Config missing required fields: {missing}")

    if not re.fullmatch(r"[a-z0-9][a-z0-9-]*", config["slug"]):
        raise SystemExit(
            f"slug {config['slug']!r} must be kebab-case "
            "(lowercase letters, digits, hyphens)"
        )

    html_out = render_template(config)

    if args.dry_run:
        sys.stdout.write(html_out)
        sys.stderr.write("\n(dry-run; nothing written)\n")
        return 0

    out_path = SITE_ROOT / f"{config['slug']}.html"
    if out_path.exists() and not args.force:
        raise SystemExit(
            f"{out_path.name} already exists. Use --force to overwrite."
        )
    out_path.write_text(html_out, encoding="utf-8")
    sys.stderr.write(f"✓ wrote {out_path.relative_to(SITE_ROOT)}\n")

    if not args.no_sitemap:
        ok, reason = update_sitemap(config["slug"])
        if ok:
            sys.stderr.write(f"✓ added to sitemap.xml ({reason})\n")
        else:
            sys.stderr.write(f"⊘ sitemap.xml: {reason}\n")

    print_link_suggestions(config)
    return 0


if __name__ == "__main__":
    sys.exit(main())

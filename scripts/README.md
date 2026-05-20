# scripts/

Automation scripts that keep the site healthy and the publishing cadence light.

| Script | What it does | When it runs |
|---|---|---|
| `inject_reviews.py` | Pulls reviews from the GAS API and writes `AggregateRating` + `Review` into the LodgingBusiness JSON-LD on index.html | Daily cron (`.github/workflows/update-reviews.yml`) |
| `new_guide.py` | Generates a new long-form travel guide page (HTML + JSON-LD) from a JSON config, and adds the entry to sitemap.xml | On demand — run when publishing a monthly guide |

## Publishing a new monthly guide

```bash
# 1. Copy the example config
cp _templates/example_guide_config.json _templates/configs/2026-08-asakusa-samba.json

# 2. Edit the new file (slug, title, sections, etc.)
#    See _templates/example_guide_config.json for full field reference.

# 3. Generate the page
python scripts/new_guide.py _templates/configs/2026-08-asakusa-samba.json

# 4. The script writes:
#    - {slug}.html in the site root
#    - a <url> entry in sitemap.xml
#    - prints suggested internal-link snippets to paste into index.html

# 5. Manually add the CTA card on index.html (snippet printed by the script)
# 6. Add a footer link to related guides (asakusa-guide.html, sumida-fireworks.html)

# 7. Commit, push, open PR, merge → GitHub Pages auto-deploys.

# 8. After deploy, in Google Search Console:
#    - URL Inspection → Request Indexing on the new page
#    - Resubmit sitemap.xml
```

### Color presets in `new_guide.py`

Pick the one that fits the topic so guides feel visually distinct from each other:

| Preset | Hero gradient | Best for |
|---|---|---|
| `amber` | warm amber→orange | temples, food, daytime neighborhood walks |
| `indigo-pink` | indigo→purple→pink | festivals, night events, fireworks |
| `teal` | teal→cyan | rainy days, indoor, winter |
| `rose` | pink→rose | sakura, blossoms, spring |
| `amber-rust` | rust→amber | autumn foliage, fall festivals |

### Required JSON config fields

- `slug` (kebab-case; becomes `{slug}.html`)
- `title` (full HTML `<title>`, can include both English and Japanese)
- `description` (meta description, 150-160 chars ideal)
- `keywords` (comma-separated; mix English and Japanese)
- `hero_h1` (the on-page H1)
- `sections` (list of section objects with `heading` + `blocks`)

### Optional but recommended

- `color_preset` (default: `amber`)
- `hero_eyebrow`, `hero_lead`, `hero_meta` (chip-row of 3-4 facts)
- `breadcrumb` (for BreadcrumbList JSON-LD)
- `event` (for Event JSON-LD — use when the page is about a dated event)
- `cta_h2`, `cta_lead`, `cta_btn_label` (overrides CTA defaults)

### Section block types

```jsonc
{
  "heading": "...",
  "walk_time": "10-min walk · 800m",   // optional pill
  "blocks": [
    {"type": "p",  "text": "Paragraph text. HTML allowed."},
    {"type": "h3", "text": "Sub-heading"},
    {"type": "ul", "items": ["item one", "item two"]},
    {"type": "ol", "items": ["step 1", "step 2"]},
    {"type": "callout", "text": "Highlighted box, HTML allowed."}
  ]
}
```

### Validation gotchas

- `slug` must be kebab-case (`a-z0-9-`) — script will reject anything else
- `description` should not exceed ~160 chars (Google truncates beyond that)
- For `event`, `start_date` must be ISO-8601 with timezone (e.g. `2026-08-29T13:00:00+09:00`)

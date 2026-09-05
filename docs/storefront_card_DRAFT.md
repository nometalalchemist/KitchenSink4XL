# Storefront aisle card: KitchenSink4XL

**DRAFT FOR AUTHOR.** Nothing here has been applied anywhere. This file is staged in the
xlsx repo so that release day is a paste rather than a writing session. The storefront
repo (`MCP Servers/storefront/index.html`) is LIVE and has not been touched.

Its own header comment carries the standing rule: *"Update an aisle card only when that
product's release is actually public."* So this card goes in when KS4XL ships, not before.

Every string below is English only. The storefront carries the same seven-language
dictionary as the product pages, so at paste time the six translations have to be added
for each new key (`x.tag`, `x.badge`, `x.pitch`, `x.counts`, `x.show`), matching the
`w.*` and `p.*` entries already in that file.

---

## 1. The card itself

Aisle letter: **X**. Department: **ELECTRICAL**. Fixture emoji: **🔌**.
Card class: `x` (the existing cards use `w` and `p`, so the accent color for `.aisle.x`
needs adding to the storefront stylesheet; the product page uses brass `#7d5c14` with
`#a3781c` as its bright variant and `#e8a31e` for the live/energized state).

```html
<div class="aisle x"><a class="stretch" href="https://nometalalchemist.github.io/KitchenSink4XL/" aria-label="KitchenSink4XL showroom"></a>
  <span class="aisle-tag" data-i18n="x.tag">AISLE X · ELECTRICAL</span>
  <span class="fixture">🔌</span>
  <h3>KitchenSink<span class="four">4</span>XL<span class="badge" data-i18n="x.badge">NOW IN BETA</span></h3>
  <p class="pitch" data-i18n="x.pitch">Everything plus the kitchen sink for Excel (.xlsx), headlined by the thing every other server gets wrong: a spreadsheet remembers its answers, not its thinking, so a stale number looks exactly like a live one. This one labels every value it hands back, names the cells that went stale, and asks Excel itself to recalculate when it matters. Structural edits carry their formulas with them. Tiered loading from day one.</p>
  <p class="counts" data-i18n="x.counts"><b>69</b> tools · <b>640</b> tests · <b>{CORRUPTIONS}</b> corruptions · <b>4</b> tiers</p>
  <div class="links">
    <a href="https://nometalalchemist.github.io/KitchenSink4XL/" data-i18n="x.show">SHOWROOM →</a>
    <a href="https://github.com/nometalalchemist/KitchenSink4XL">GITHUB</a>
    <a href="https://pypi.org/project/kitchensink4xl/">PYPI</a>
  </div>
</div>
```

**Author decision needed on the counts line.** The Word and PPT cards read
`operations · tools · tests · corruptions`. This repo has no script that derives an
operations figure the way those two do (`scripts/measure_surface.py` reports tools and
tokens only), so the draft above swaps the first slot for tiers rather than publishing a
hand-counted number. Two other options if the four-stat symmetry matters more: add an
operations counter to the measuring script before ship, or run the card as
`69 tools · 640 tests · {CORRUPTIONS} corruptions · 10.8k lite tokens`, which is the one
number no competing Excel server publishes at all.

## 2. Dictionary entries (English; six translations still needed at paste time)

```js
"x.tag": "AISLE X · ELECTRICAL",
"x.badge": "NOW IN BETA",
"x.pitch": "Everything plus the kitchen sink for Excel (.xlsx), headlined by the thing every other server gets wrong: a spreadsheet remembers its answers, not its thinking, so a stale number looks exactly like a live one. This one labels every value it hands back, names the cells that went stale, and asks Excel itself to recalculate when it matters. Structural edits carry their formulas with them. Tiered loading from day one.",
"x.counts": "<b>69</b> tools · <b>640</b> tests · <b>{CORRUPTIONS}</b> corruptions · <b>4</b> tiers",
"x.show": "SHOWROOM →",
```

## 3. Switchbar line (add after the PPT entry)

```html
<a class="x" href="https://nometalalchemist.github.io/KitchenSink4XL/">🔌 KitchenSink4XL</a>
```

The same line has to be added to the switchbar on the Word and PPT product pages, which
currently list two aisles and a storefront. The KS4XL page already carries all three.

## 4. Empty shelf

The storefront currently shows shelves reading THIS SHELF INTENTIONALLY LEFT EMPTY, with
the line *"The next fixture earns its department the same way the first two did: by
shipping."* One of those shelves comes down when this card goes up. That sentence stays
exactly as it is; it is the reason the card was worth waiting for.

## 5. Meta description (whole-page, storefront `<head>`)

The current one names two servers. Proposed replacement, same shape:

> The KitchenSink family of MCP servers for Office documents: KitchenSink4Word (219 operations across 108 tools for .docx), KitchenSink4PPT (138 tools for .pptx, native editable diagrams from SVG with connectors that stay glued), and KitchenSink4XL (69 tools for .xlsx, every value labeled calculated, cached, or absent). All start small and load capability packs on demand. Engineered not to corrupt. Refuses rather than guesses.

## 6. Beta wording

The product page carries a BETA chip in its masthead and a full panel explaining why, in
the author's own words. The storefront card uses the shorter `NOW IN BETA` badge in the
same slot where the PPT card said `NEW STOCK`. If the author would rather the storefront
not carry a beta signal at all, drop the `<span class="badge">` and the `x.badge` key; the
product page keeps the honest label either way.

---

## Placeholders to fill before pasting

| Placeholder | Where | How to fill |
|---|---|---|
| `{CORRUPTIONS}` | counts line, twice | corruption count across the fixture corpus; the build log records zero, but the figure is not script-derived yet |

Tools (69) and tests (640) above are script-derived and current as of 2026-09-04. Re-run
`scripts/measure_surface.py` and `pytest tests/unit --co -q` at paste time; the storefront
rule is that its numbers track the shipped release.

# Storefront aisle card: KitchenSink4XL

**APPLIED 2026-09-06, LOCAL ONLY, NOT PUSHED.** Everything below has been placed in
`MCP Servers/storefront/index.html` and committed there (205a9a3). That repo is LIVE but
the commit has NOT been pushed, so nothing is public until the KitchenSink4XL release is.
This file is kept as the record of what was placed and which open questions were settled.

Settled at placement: the counts line runs four stats, the operations figure now exists
(`scripts/count_operations.py`), `{CORRUPTIONS}` resolves to 0, `NOW IN BETA` is gone
because the release is a full 1.0.0, and the NEW STOCK badge moved here from the PPT card.
Still open: the storefront has three aisles but only two service desks, and the KS4XL
switchbar line is not yet on the Word and PPT product pages (section 3 below).

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
  <h3>KitchenSink<span class="four">4</span>XL<span class="badge" data-i18n="x.badge">NEW STOCK</span></h3>
  <p class="pitch" data-i18n="x.pitch">Everything plus the kitchen sink for Excel (.xlsx), headlined by the thing every other server gets wrong: a spreadsheet remembers its answers, not its thinking, so a stale number looks exactly like a live one. This one labels every value it hands back, names the cells that went stale, and asks Excel itself to recalculate when it matters. Structural edits carry their formulas with them. Tiered loading from day one.</p>
  <p class="counts" data-i18n="x.counts"><b>129</b> operations · <b>69</b> tools · <b>1,041</b> tests · <b>0</b> corruptions</p>
  <div class="links">
    <a href="https://nometalalchemist.github.io/KitchenSink4XL/" data-i18n="x.show">SHOWROOM →</a>
    <a href="https://github.com/nometalalchemist/KitchenSink4XL">GITHUB</a>
    <a href="https://pypi.org/project/kitchensink4xl/">PYPI</a>
  </div>
</div>
```

**RESOLVED at placement.** The counts line runs the same four stats as the Word and PPT
cards. `scripts/count_operations.py` landed with the release collateral and derives the
operations figure from committed source, so the first slot is a measured 129 rather than a
hand-counted number or a substitute stat.

## 2. Dictionary entries (English below; all six translations were written at placement)

```js
"x.tag": "AISLE X · ELECTRICAL",
"x.badge": "NEW STOCK",
"x.pitch": "Everything plus the kitchen sink for Excel (.xlsx), headlined by the thing every other server gets wrong: a spreadsheet remembers its answers, not its thinking, so a stale number looks exactly like a live one. This one labels every value it hands back, names the cells that went stale, and asks Excel itself to recalculate when it matters. Structural edits carry their formulas with them. Tiered loading from day one.",
"x.counts": "<b>129</b> operations · <b>69</b> tools · <b>1,041</b> tests · <b>0</b> corruptions",
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

## 6. The badge

**RESOLVED at placement.** KitchenSink4XL ships as a full 1.0.0, so the maturity badge this
section used to propose is gone from both the card and the product page. The slot carries
NEW STOCK instead, moved off the PPT card, which is no longer the newest arrival.

---

## Placeholders to fill before pasting

None remain. `{CORRUPTIONS}` resolved to 0, the figure the build log and the fixture
corpus both record.

Every number on the card was re-measured on 2026-09-06: operations 129
(`scripts/count_operations.py`), tools 69 (`scripts/measure_surface.py`), tests 1,041
(`pytest tests/unit --co -q`). Re-run all three before any future edit to this card; the
storefront rule is that its numbers track the shipped release.

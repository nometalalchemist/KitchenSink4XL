<!-- mcp-name: io.github.nometalalchemist/kitchensink4xl -->

# KitchenSink4XL

Everything plus the kitchen sink for Microsoft Excel files: a round-trip-safe
`.xlsx` MCP server with an honest calculation story and tiered loading.

> Status: PRE-RELEASE (Phase 0 scaffold). The infrastructure is ported and
> wired; the Excel engine and the tool surface land in later build phases.
> Counts are published only once `scripts/measure_surface.py` measures the
> real surface, so no numbers appear here yet.

## What it is

A kitchen-sink MCP server for reading and editing Microsoft Excel workbooks.
It pairs a cross-platform file-based tier with an environment-gated Excel
application (COM) tier, and it is built around a safety core: a round-trip
hazard scan, backup-before-mutation, and verify-after-write, so rich workbooks
are not silently damaged on edit.

- Cross-platform file-based tier for cells, ranges, formulas, formatting,
  styles, conditional formatting, data validation, tables, named ranges, sort
  and filter, images, and charts.
- An honest calculation story: reads label every value as cached, computed,
  formula, or absent, because no pure-Python engine computes formulas.
- Tiered loading: sessions start with a small lite core; optional packs load
  on request.
- On Windows with Excel installed, an optional COM tier adds real pivot
  tables, recalculation, goal seek, PDF and image export, format conversion,
  and more.

Placeholder for the measured figures (published once the surface is built):

- Operations: _[measured by scripts/measure_surface.py]_
- Tools: _[measured]_
- Lite startup surface: _[measured]_ tokens
- Full surface: _[measured]_ tokens

## Install

_[Publishing details are added at the ship phase.]_

## Trademarks and affiliation

Not affiliated with or endorsed by Microsoft Corporation. Microsoft and Excel
are trademarks of Microsoft Corporation. KitchenSink4XL works with Microsoft
Excel files; the trademarks are used nominatively to describe that
compatibility, and no Microsoft logos or trade dress are used.

## License

AGPL-3.0-only.

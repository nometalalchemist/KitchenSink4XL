<!-- mcp-name: io.github.nometalalchemist/kitchensink4xl -->

# KitchenSink4XL

Everything plus the kitchen sink for Microsoft Excel files: a round-trip-safe
`.xlsx` MCP server with an honest calculation story and tiered loading.

> Status: PRE-RELEASE (file tier COMPLETE: safety core, cells and ranges,
> query, formulas with write and audit, formatting, tables, names,
> conditional formatting, data validation, sort/filter, comments,
> hyperlinks, import/export, images and charts, protection, page layout,
> inspectors, backups, document properties, sheet views, a multiplex
> validate battery, and workflow recipes). The COM application tier
> (recalculation, real pivots, and the rest) lands in a later build phase.
> Counts are published only once `scripts/measure_surface.py` measures the
> final surface.

## What it is

A kitchen-sink MCP server for reading and editing Microsoft Excel workbooks.
It pairs a cross-platform file-based tier with an environment-gated Excel
application (COM) tier, and it is built around a safety core: a round-trip
hazard scan, backup-before-mutation, and verify-after-write, so rich workbooks
are not silently damaged on edit.

- Cross-platform file-based tier for cells, ranges, formulas, server-side
  query and aggregation, formatting, styles, conditional formatting, data
  validation, tables, named ranges, sort and filter, comments, hyperlinks,
  and CSV/TSV/JSON interop (images and charts join with the gated
  families).
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

## Environment variables

| Variable | Effect |
|---|---|
| `KS4XL_MODE` | startup surface: `lite` (default), `full`, or a pack list |
| `KS4XL_PACK_POLICY` | `auto` (default) or `locked` (enable_tools refuses; surface fixed at startup) |
| `KS4XL_ALLOWED_ROOTS` | opt-in path sandbox; tools refuse to touch files outside these roots |
| `KS4XL_NO_UPDATE_CHECK` | `1` or `true` turns the update check off completely: no network call, no cache file |

The server checks PyPI, the package index it was installed from, at most once
every 14 days to see whether a newer version exists; the check sends nothing
but a standard HTTP request for that package's public JSON, and setting
`KS4XL_NO_UPDATE_CHECK=1` turns it off entirely. It runs on a background
thread at startup, so it never delays a call, and it fails silently: a timeout
or an offline machine leaves no error anywhere. When a newer release exists,
`get_server_info` adds one line saying so. That is the only place it ever
appears, and the server never downloads or installs anything on its own.

## Trademarks and affiliation

Not affiliated with or endorsed by Microsoft Corporation. Microsoft and Excel
are trademarks of Microsoft Corporation. KitchenSink4XL works with Microsoft
Excel files; the trademarks are used nominatively to describe that
compatibility, and no Microsoft logos or trade dress are used.

## License

AGPL-3.0-only.

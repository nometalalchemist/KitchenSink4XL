<!-- mcp-name: io.github.nometalalchemist/kitchensink4xl -->

# 🔌 KitchenSink4XL

[![Tests](https://github.com/nometalalchemist/KitchenSink4XL/actions/workflows/tests.yml/badge.svg)](https://github.com/nometalalchemist/KitchenSink4XL/actions/workflows/tests.yml)
[![PyPI](https://img.shields.io/pypi/v/kitchensink4xl)](https://pypi.org/project/kitchensink4xl/)
[![License: AGPL-3.0](https://img.shields.io/badge/license-AGPL--3.0-blue)](LICENSE)

[Landing page](https://nometalalchemist.github.io/KitchenSink4XL/) · [llms.txt](https://nometalalchemist.github.io/KitchenSink4XL/llms.txt) (machine-readable capability manifest for agents and LLM crawlers)

**Everything plus the kitchen sink for Microsoft Excel: the .xlsx MCP server
that never hands your AI a number it cannot back up.** 129 workbook
operations across 67 tools, a lite core that opens at about 21,200 tokens,
and a safety core that backs up before every change and verifies every save.

## The number that looks right and isn't

Here is the dirty secret of every tool that reads Excel files: a spreadsheet
remembers its answers, not its thinking. When something changes upstream, the
old answers just sit there, looking exactly like real numbers. Most AI tools
will read one anyway and confidently report it, and you will make a decision
based on it. This server refuses to bluff. Every number it hands your AI comes
with the truth attached: freshly calculated, remembered from the last save, or
missing, with the affected cells named. And when it really matters, it asks
Excel itself to recalculate and prove it.

## What happens to column F when you insert column D

Insert one column in a real workbook and you find out how much of a
spreadsheet lives to the right of where you clicked. Formulas, totals, charts,
that table on sheet three that quietly feeds the summary: all of them have
opinions about where their cells just went.

Most tools do the cheap version: move the cells, save the file, hope. The
damage never shows up when the edit runs. It shows up three days later, in a
meeting, in the one total that quietly stopped including its last row.

This server moves everything that should move, and checks the result against
what Excel itself would have done. When an edit cannot be done safely, it says
no out loud instead of guessing quietly. A refusal costs you a minute. A quiet
guess costs you the deal.

## Two numbers that matter

- **129 workbook operations, 67 tools.** Many tools here are action
  multiplexers, so the tool count undersells the surface: `manage_worksheet`
  alone performs seven distinct operations, `manage_table` nine, `validate`
  runs nine correctness batteries. The operations figure comes from
  `scripts/count_operations.py`, which reads the dispatch values each tool
  actually validates, out of committed source, and never from hand-math. Its
  docstring carries the counting definition and what is deliberately excluded.
  A committed snapshot (`scripts/operations_snapshot.json`) plus a guard test
  make a drifting figure a test failure rather than a marketing decision.
- **Tiered loading: starts at about 21.2k tokens, scales to everything.** A
  fresh session loads the 40-tool lite core and turns on capability packs only
  when a task needs them, with one `enable_tools` call. Load every pack and
  the full surface measures about 32,700 tokens. All figures come from
  `scripts/measure_surface.py`; see [Context cost](#context-cost-measured).

## The packs

The stock is arranged in capability packs, the way a good shop groups what
belongs together. A session opens on the lite core and switches on the pack a
job needs with one call. Numbers below come straight from
`scripts/measure_surface.py`, never hand-counted.

| Pack | Tools | Approx tokens | What it carries |
|---|---:|---:|---|
| **lite** (startup) | 40 | ~21.2k | The everyday bench: an anchored grid view of the workbook, labeled reads of cells and ranges, server-side query and aggregation, formula write and audit, structural row and column edits that carry their references, sort and filter with Excel's own ranking, tables, formatting, import and export, backups, diagnostics, and the `enable_tools` switchboard |
| design | 9 | ~6.1k | Named cell styles, a format painter, a style-bloat audit, conditional formatting, data validation, images, charts, the full table lifecycle (columns, totals, resize, banding), and named ranges including LAMBDA definitions and a cleanup pass |
| io | 9 | ~2.8k | Page layout and print setup, headers and footers, advisory protection, legacy comments, multi-sheet export, and the read-side inspectors: external links, VBA, existing pivot tables, and data connections |
| com | 11 | ~2.6k | Drives a private hidden Excel instance, never your open session: real recalculation, real pivot tables, goal seek, PDF export, sheet render to image, format conversion, real encryption, sparklines, true autofit, an opens-clean check, and an honest status report |
| **Full surface** | **69** | **~32.7k** | Everything (67 workbook tools plus `enable_tools` / `disable_tools`) |

## Quickstart: start lite, enable what you need

A session begins with the lite core. When a task needs more, the agent turns
on the pack by name:

```
enable_tools(["design"])        # styles, conditional formatting, charts, names
enable_tools(["com"])           # real recalculation, real pivots, PDF export
```

The lite core carries no degraded stand-ins, so the lazy path is a dead end on
purpose: a refusal for out-of-scope work names the exact pack and the exact
call that unlocks it. Power users who want everything loaded from the start
can pin it with `KS4XL_MODE=full` in the server environment, or a
comma-separated pack list. Administrators can lock the selection with
`KS4XL_PACK_POLICY=locked`.

## Against the rest of the aisle

Feature grids flatter everybody. Here it is instead as the requests an agent
actually gets:

| "Hey, could you…" | 🔌 KitchenSink4XL | haris-musa (openpyxl), 4,155★† | The rest of the aisle* |
|---|---|---|---|
| "Insert a column in the middle, keep every formula pointing where it should." | ✔ formulas, tables, named ranges, charts and cross-sheet links all follow, checked against Excel's own recalculation | ✘ nothing in the docs says otherwise, and openpyxl, the library underneath, states plainly that it does not maintain formulas, tables or charts when rows and columns move | Excel does the shifting itself, on Windows, with the workbook closed and the application running |
| "Is this total current, or left over from the last save?" | ✔ every value labeled calculated, cached or absent, and the stale cells named by address | ✘ reads return the formula string; there is no way to get even the cached value | the COM server recalculates with Excel's own engine, but the value comes back with no freshness marking of any kind |
| "Sort it the way Excel would sort it, blanks and all." | ✔ numbers, then text case-insensitively, then FALSE, TRUE, errors, blanks last, both directions, checked cell for cell against Excel | ✘ no sort tool | Excel's own sort, with the application running |
| "Write a LAMBDA the file can still open afterward." | ✔ modern functions and their declared parameters prefixed correctly; 80 formulas across every documented family opened with no repair prompt | n/d, and no Excel MCP server surveyed documents the modern functions either way, as support or as a limitation | n/d |
| "Read the big workbook without torching my context window." | ✔ compact reads, a grid view, and server-side query and aggregation so the rows stay on the server | ✘ no pagination and no token budgeting anywhere in the docs; with no end cell given it reads the whole used range | the Go server pages reads by cell count, 4,000 cells at a time by default |
| "Recalculate it for real and prove it." | ✔ a private hidden Excel does the arithmetic and the result comes back labeled calculated, never mixed in with cached values | ✘ no recalculation tool, and the documentation never raises the subject | real recalculation on the COM server, which asks for Windows, an installed Excel, an interactive desktop and the workbook closed first |

Capability survey as of 2026-09-04, compiled from public repositories,
documentation, and issue trackers. n/d: not documented in the sources
surveyed, which is not a claimed absence. The most popular Excel MCP server is
haris-musa/excel-mcp-server, 25 tools on openpyxl, last released 2026-04-12;
its own open issues #94 and #118 ask for computed values instead of formula
text. openpyxl's documentation states that it does not manage dependencies
such as formulae, tables and charts when rows or columns are inserted or
deleted. \* The rest of the aisle: negokaz/excel-mcp-server (Go, seven tools,
paginated reads, live editing on Windows), sbroenne/mcp-server-excel (C#, 31
tools over 326 operations driving the real Excel application, the capability
ceiling, but Windows with Excel and exclusive access to the file, and no
file-based tier). Two newer servers, PSU3D0/agent-spreadsheet and
logisky/logisheets-mcp, embed real calculation engines and are worth watching.
Microsoft ships no Excel MCP server, and Claude for Excel is an Office add-in
rather than an MCP server. † Star counts read from the GitHub API on
2026-09-04, shown for context: stars say how long a shelf has been in view,
not what is on it. Corrections welcome:
[open an issue](https://github.com/nometalalchemist/KitchenSink4XL/issues).

## Requirements

- Python 3.12+ (developed on 3.14)
- Most of this server needs no Excel installed at all and runs on any
  computer. The parts that ask Excel to do the work want Windows with Excel on
  it, and they open their own private copy, so the workbook you have on screen
  is never touched.

## Install

Pick the line that describes you. Most people are the first one.

First launch through uvx downloads and builds the environment and can take 20
to 30 seconds before the server answers; every launch after that starts in
about two. If a client reports a timeout on first install, launch once from a
terminal and try again.

### Using Claude Desktop? One double-click.

Download `kitchensink4xl.mcpb` from the
[latest release](https://github.com/nometalalchemist/KitchenSink4XL/releases/latest)
and double-click it, or drag it into the Claude Desktop window. Desktop adds
it as an extension and the sink is connected. Nothing to type, nothing to
configure. One-time requirement: [uv](https://docs.astral.sh/uv/) on your
PATH (`pip install uv`), which the bundle uses to start the server. If Desktop
does not pick the file up on a double-click, use Settings > Extensions >
Advanced settings > Install extension.

The extension's settings page offers two switches. **Verify every save with
Excel** turns on the deep check, where a real hidden Excel has to open the
saved file cleanly or the backup is restored; it is off by default because it
costs a round trip through Excel on every write. **Limit the server to one
folder** confines every path the server touches, reads included, to a
directory you pick.

### Using Claude Code? Paste this.

```
claude mcp add xl -s user -- uvx kitchensink4xl
```

One line in a terminal and you are done. It fetches and runs the server for
you, so there is nothing to install first.

<details>
<summary><b>For developers: pip, uvx, source, other MCP clients</b></summary>

Install the package and point any MCP client at the executable:

```
pip install kitchensink4xl
```

```json
{"mcpServers": {"xl": {"command": "kitchensink4xl", "env": {"KS4XL_MODE": "lite"}}}}
```

The mode is `"lite"` (the default), `"full"`, or a comma-separated pack list.
The `xl-mcp` executable is an equivalent entry point. The installed package is
named `xlsx_mcp`, so a client that wants an interpreter and a module instead
of a console script can run `python -m xlsx_mcp`.
`python -m xlsx_mcp.server` starts the same server and is what releases
before 1.1 support. Running from a clone
works the same way; point the command at the `xl-mcp` executable in the
clone's virtual environment:

Windows:

```
git clone https://github.com/nometalalchemist/KitchenSink4XL
cd KitchenSink4XL
python -m venv .venv
.venv\Scripts\pip install -e ".[com]"
claude mcp add xl -s user -- <absolute-path>\.venv\Scripts\xl-mcp.exe
```

macOS and Linux:

```
git clone https://github.com/nometalalchemist/KitchenSink4XL
cd KitchenSink4XL
python3 -m venv .venv
.venv/bin/pip install -e ".[com]"
claude mcp add xl -s user -- <absolute-path>/.venv/bin/xl-mcp
```

The COM pack is an optional extra, `pip install kitchensink4xl[com]`, and it
is a no-op off Windows. On Windows the COM dependency often arrives
transitively with the base install; installing with `[com]` is the guaranteed
route either way, and harmless to repeat. With no install at all:
`uvx kitchensink4xl`.

Environment variables the server reads:

| Variable | What it does |
|---|---|
| `KS4XL_MODE` | Startup surface: `lite` (default), `full`, or a comma-separated pack list |
| `KS4XL_PACK_POLICY` | `auto` (default) or `locked`, which fixes the surface at startup |
| `KS4XL_ALLOWED_ROOTS` | Path sandbox: an `os.pathsep`-separated list of directories |
| `KS4XL_VERIFY_COM` | `1` makes the deep Excel verification the default for every save |
| `KS4XL_VALIDATE_COM` | `1` routes `validate`'s structure check through Excel's own verdict |
| `KS4XL_COM_TIMEOUT` | Bounds how long a COM call may take |
| `KS4XL_NO_UPDATE_CHECK` | `1` or `true` turns the update check off completely: no network call, no cache file |

The server checks PyPI, the package index it was installed from, at most once
every 14 days to see whether a newer version exists; the check sends nothing
but a standard HTTP request for that package's public JSON, and setting
`KS4XL_NO_UPDATE_CHECK=1` turns it off entirely. It runs on a background
thread at startup, so it never delays a call, and it fails silently: a timeout
or an offline machine leaves no error anywhere. When a newer release exists,
`get_server_info` adds one line saying so. That is the only place it ever
appears, and the server never downloads or installs anything on its own.

</details>

## Context cost (measured)

Most MCP servers move into your AI's context like a hoarder: everything, up
front, whether the job needs it or not. This one starts light and only unpacks
a shelf when the work calls for it. Here is the bill, measured by
`scripts/measure_surface.py`:

| On the meter | Tools | Tokens | When it draws |
|---|---:|---:|---|
| Lite core | 40 | ~21.2k | From the first message of every session |
| Design pack | 9 | ~6.1k | Only after `enable_tools` |
| Layout plus inspection pack (`io`) | 9 | ~2.8k | Only after `enable_tools` |
| COM pack | 11 | ~2.6k | Only after `enable_tools`, and only on Windows with Excel |
| Everything switched on | 69 | ~32.7k | `KS4XL_MODE=full`, if you want it all up front |

No other Excel MCP server surveyed publishes what its own tool definitions
cost to load. The figures above come from a measuring script that ships with
the source, so you can check them yourself. Clients that defer tool schemas
until first use pay close to zero until a tool is actually called.

## Safety model

- Automatic timestamped backup before every mutation, in two rotating slots in
  a hidden `.ks4xl-backups/` folder next to the workbook. `manage_backups`
  lists, restores, and prunes them. Exclude that folder from cloud sync tools:
  the slots churn on every edit and sync clients can hold locks that slow
  saves down.
  {MAIN_THREAD_COPY:V1-16-readme} <!-- facts: the slots hold only what this
  server itself changed. An edit made in Excel, or by any other program, is
  never captured, so `prev` restores the state before the last SERVER
  mutation, not the state before the last edit to the file. -->
- Filter-hidden rows: {MAIN_THREAD_COPY:V1-12-readme} <!-- facts:
  query_range, export_range, and the aggregates read every row in the range,
  including rows an autofilter is hiding. Excel's own SUBTOTAL ignores hidden
  rows; these do not. Filter first with `where` if hidden rows should be out
  of the answer. -->
- Saves are atomic and validated; a failed operation leaves the original
  byte-identical.
- A round-trip hazard scan runs before a mutating save. Parts the writer
  cannot preserve are named to the caller, and a loss is refused unless the
  caller explicitly allows the specific classes the gate named.
- Verify-after-write, with restore-from-backup on failure. `verify_com=true`
  on any mutating call adds the deep check: the produced file has to open in a
  real hidden Excel with no repair prompt, or the backup goes back and the
  save refuses. `KS4XL_VERIFY_COM=1` makes that the default for every save.
- Structural edits are checked against an independent expectation of Excel's
  own behavior for formulas, tables, named ranges, charts, and cross-sheet
  references.
- Formula text is never inferred from a leading `=` character; cell type
  decides, so text that merely looks like a formula is not re-armed as one on
  copy, move, sort, or reference rewrite.
- Ambiguous targets are refused with the candidates listed, never guessed.
- The COM tier opens its own hidden instance, journals process IDs, and leaves
  no orphan Excel processes behind.

### Sandboxing (opt-in)

Off by default: with nothing configured, the server reads and writes wherever
you point it. Set `KS4XL_ALLOWED_ROOTS` to a list of directories separated by
the OS path separator (`;` on Windows, `:` elsewhere) and every path the
server touches must resolve inside one of them. Reads are gated as well as
writes, since a read outside the sandbox exfiltrates workbook content just as
surely as a write plants it. The containment check runs on canonicalized
paths, so `..\` traversal, symlink and junction escapes, short names, case
tricks, and lookalike sibling directories are all caught. A blocked call
refuses with a typed error naming the offending path and the allowed roots
before any file is opened.

## Testing

1,046 tests in `tests/unit`, plus a separate local COM gate battery that drives a
real Excel. On top of the suite, this release went through the family's
gauntlet:

- **Adversarial rounds through the raw MCP transport.** Roughly 600 tool calls
  across six waves against scratch workbooks, every call journaled. Thirteen
  findings confirmed and fixed, each with a regression test.
- **A live COM stress round.** Nine phases, roughly 600 executor operations
  including a 200-operation endurance soak, more than 40 distinct Excel
  processes, all of it unattended with nobody available to dismiss a dialog.
  Zero critical, zero high, six lower findings, all fixed. It verified 87
  properties that held, among them serialization under twelve concurrent
  submits and zero lost updates under eight concurrent writers on one
  workbook.
- **A numbers-safety gate.** 61 checks against a real Excel covering insert and
  delete of rows and columns, sort, move, merge, and an insert-plus-delete
  composition, each compared against Excel's own recalculation.
- **A formula-fidelity gate.** 80 formulas across every documented function
  family, written through all four writers, opened with zero repair prompts
  and 80 of 80 correct values.
- **A discoverability gate.** Six fresh agents starting from the 40-tool lite
  surface found the right pack six times out of six.
- **The author's own field testing**, running the tools in insane mode against
  the installed build from a separate window, which is where the last round of
  fixes came from.

Zero corruptions across the fixture corpus. That corpus is committed in-tree
because the hazard, fidelity and adversarial tests need real Excel-authored
workbooks (pivots, slicers, shapes, x14 conditional formatting) that openpyxl
cannot author, and every fixture is scrubbed of authoring identity by a
test-guarded pass. The corruption figure is the build record's, not a script's
output; the script-derived numbers here are the test count and the surface
measurements.

## Young, and tested like it isn't

This is the newest sink in the family, and it shipped through the hardest
gauntlet we have ever run: adversarial rounds through the raw protocol,
ground-truth checks against Excel's own arithmetic, interrupted-save torture,
and a final round built around one question: can any sequence of operations
ruin the only copy of a workbook a business depends on. Every finding was
fixed before this release. What it has not had yet is a long life in
strangers' spreadsheets, and numbers earn trust in the field. The safety net
is structural: a backup before every change, saves that verify before they
replace your file, and honest labels on every number it hands you. If Excel is
where you live, bring your ugliest workbook to the service counter and tell us
what broke.

- 🔧 [Found a dead circuit?](https://github.com/nometalalchemist/KitchenSink4XL/issues/new?template=bug_report.yml)
  Something glitched, refused, or came back with the wrong number. Never
  attach a private workbook; rebuild the structure with placeholder data.
- 🔌 [Missing an outlet?](https://github.com/nometalalchemist/KitchenSink4XL/issues/new?template=feature_request.yml)
  An Excel capability the sink should also have. Describe the real task behind
  it and it goes on the workbench.

### Engineered not to corrupt

Atomic saves. Automatic backups before every change. And one extra wire,
because this is a spreadsheet: before a release ships, every kind of edit is
checked against Excel's own math. The wiring gets inspected before the power
goes on.

Provided as-is, without warranty of any kind, per the license; the engineering
above is simply how seriously your workbooks are taken. Keep backups. (It
makes them for you.)

## Known limits

- **`com_render_sheet` needs a real desktop.** Rendering a sheet to an image
  goes through Excel's clipboard, and a hidden Excel started in a
  non-interactive session has no window station to copy through, so the call
  fails there. It works from a normal signed-in desktop session, which is
  where anyone actually uses it. Every other COM tool runs fine unattended.
- **One user, one machine.** This is a desktop tool, not a shared service.
  The server runs on your computer, over stdio, under your own account, and
  the COM tier drives an Excel that belongs to your Windows session. Two
  people cannot point at one installation, and it is not something to stand up
  on a server for a team.
- **Deleting or renaming a sheet does not rewrite references.** That is
  Excel's own behavior, and it is documented on the tool: formulas and defined
  names pointing at a deleted sheet break to `#REF!` when Excel opens the
  file. Audit references first when in doubt.
- **Cached values go stale after a file-tier write.** Writing through the file
  tier cannot recompute anything, so a cell downstream of your edit keeps its
  old cached number until Excel recalculates. That is exactly what the labels
  are for, and `recalculate` in the COM pack is the fix.
- **Some MCP clients drop a tool's schema when a pack is disabled** and do not
  pick it back up on re-enable, even though the server announces the change
  both ways. If a re-enabled tool comes back as "no such tool", refresh the
  tool list on the client side.

## License

**AGPL-3.0.** Free for individuals and personal use, and it stays that way.

Companies building it into their own products need a commercial license, with
terms worked out case by case.
[Open an issue](https://github.com/nometalalchemist/KitchenSink4XL/issues/new?template=commercial_license.yml)
and we will talk it through.

---

*Not affiliated with or endorsed by Microsoft Corporation. Microsoft and Excel
are trademarks of Microsoft Corporation. KitchenSink4XL works with Microsoft
Excel files; the trademarks are used nominatively to describe that
compatibility, and no Microsoft logos or trade dress are used.*

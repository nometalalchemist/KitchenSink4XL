"""Public-copy guards: the published numbers must be the scripts' numbers.

The Word repo shipped stale marketing figures twice in launch week (an
llms.txt claiming the wrong version, a site claiming a retired tool count),
which is why this guard exists here from the first release instead of after
the first mistake. It makes the failure mode mechanical: if the surface
changes and the copy does not, the suite goes red.

Guarded files: README.md, docs/llms.txt (and docs/index.html once the site
branch merges; the list below picks it up automatically when it exists).

The operations snapshot guard is the one to understand. count_operations.py
derives the headline operations figure from committed source, and
scripts/operations_snapshot.json records the per-tool breakdown that produced
the published number. Adding a dispatch value to any multiplexer moves the
figure, and that is a publishing decision, not a silent one. When this test
fails legitimately:

    .venv/Scripts/python.exe -X utf8 scripts/count_operations.py \\
        --json scripts/operations_snapshot.json

then update every published figure, then update this test's expectations. All
three in the same commit.
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
SNAPSHOT = ROOT / "scripts" / "operations_snapshot.json"
STRESS_SNAPSHOT = ROOT / "scripts" / "stress_calls_snapshot.json"

#: The figures the public copy currently claims. Change these only together
#: with the copy itself.
PUBLISHED_OPERATIONS = 129
PUBLISHED_TOOLS_TOTAL = 69      # 67 workbook tools + enable_tools/disable_tools
PUBLISHED_TOOLS_LITE = 40
PUBLISHED_STRESS_CALLS = 882    # journaled adversarial calls; a floor, not a total


def _public_files() -> list[Path]:
    candidates = [
        ROOT / "README.md",
        ROOT / "docs" / "llms.txt",
        ROOT / "docs" / "index.html",
    ]
    return [p for p in candidates if p.exists()]


def _run(script: str, *args: str) -> str:
    out = subprocess.run(
        [sys.executable, "-X", "utf8", str(ROOT / "scripts" / script), *args],
        capture_output=True, text=True, encoding="utf-8", cwd=ROOT,
    )
    assert out.returncode == 0, f"{script} failed: {out.stderr}"
    return out.stdout


def _measured() -> tuple[int, int, int]:
    """(lite tools, full tools, operations), straight from the scripts."""
    surf = _run("measure_surface.py")
    ops = _run("count_operations.py")
    lite = int(re.search(r"lite startup surface: (\d+) tools", surf).group(1))
    full = int(re.search(r"full surface .*?: (\d+) tools", surf).group(1))
    n_ops = int(re.search(r"TOTAL DISTINCT OPERATIONS:\s+(\d+)", ops).group(1))
    return lite, full, n_ops


def test_operations_snapshot_matches_source():
    """Every tool's operation count is the one the snapshot recorded."""
    stored = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
    out = _run("count_operations.py", "--check", str(SNAPSHOT))
    assert "snapshot MATCHES" in out, (
        "count_operations.py disagrees with scripts/operations_snapshot.json. "
        "Re-run it with --json to re-record, then update the published "
        "figures and PUBLISHED_OPERATIONS in this file.\n" + out
    )
    assert stored["total_operations"] == PUBLISHED_OPERATIONS


def test_stress_call_figure_is_the_journal_figure():
    """The site's stress-call headline is the one the journal counts.

    The journal itself (integration/adversarial_log.md) carries real machine
    paths and is gitignored, so the count travels as a committed snapshot and
    this guard checks the published copy against that. On the build machine,
    where the journal does exist, count_stress_calls.py --check re-derives it
    and this test catches a snapshot that drifted from the log.
    """
    stored = json.loads(STRESS_SNAPSHOT.read_text(encoding="utf-8"))
    assert stored["stress_calls"] == PUBLISHED_STRESS_CALLS
    out = _run("count_stress_calls.py", "--check")
    assert "MATCHES" in out or "journal absent" in out, out
    page = ROOT / "docs" / "index.html"
    if page.exists():
        assert str(PUBLISHED_STRESS_CALLS) in page.read_text(
            encoding="utf-8"), "index.html is missing the stress-call figure"


def test_published_numbers_match_scripts():
    """The headline figures in the public copy are the scripts' figures."""
    lite, full, n_ops = _measured()
    assert (lite, full, n_ops) == (
        PUBLISHED_TOOLS_LITE, PUBLISHED_TOOLS_TOTAL, PUBLISHED_OPERATIONS
    ), (
        f"scripts now report lite={lite} full={full} ops={n_ops}; update the "
        "public copy AND this test together (that is the whole point)."
    )
    workbook_tools = full - 2  # excluding enable_tools/disable_tools
    for path in _public_files():
        text = path.read_text(encoding="utf-8")
        assert str(n_ops) in text, f"{path.name} is missing the operations count"

    # Which tool total a file states is a per-file doctrine, and the guard
    # follows it rather than overriding it. README and llms.txt spell out the
    # relationship ("129 operations across 67 tools, plus two pack toggles"),
    # so they carry the workbook figure. index.html states 69 and only 69, by
    # a doctrine written into its own header comment: a reader who does the
    # arithmetic on the pack table must not arrive at a second answer. Both
    # figures are the measured ones either way.
    for name, expected in (("README.md", workbook_tools),
                           ("llms.txt", workbook_tools),
                           ("index.html", full)):
        matches = [p for p in _public_files() if p.name == name]
        if not matches:
            continue
        text = matches[0].read_text(encoding="utf-8")
        assert str(expected) in text, (
            f"{name} is missing its tool total ({expected})"
        )


def test_token_headlines_match_measurement():
    """The lite and full token figures in every public file are measured."""
    surf = _run("measure_surface.py")
    lite_k = re.search(r"lite startup surface: \d+ tools, ~([\d.]+)k",
                       surf).group(1)
    full_k = re.search(r"full surface .*?: \d+ tools, ~([\d.]+)k",
                       surf).group(1)

    def variants(k: str) -> list[str]:
        n = int(round(float(k) * 1000))
        comma = f"{n:,}"
        return [f"{k}k", comma] + [comma.replace(",", s)
                                   for s in (".", " ", " ", " ")]

    for path in _public_files():
        text = path.read_text(encoding="utf-8")
        for label, k in (("lite", lite_k), ("full", full_k)):
            assert any(v in text for v in variants(k)), (
                f"{path.name}: measured {label} figure ~{k}k appears in no "
                f"accepted format {variants(k)}"
            )


def test_pack_tables_match_measurement():
    """Every pack row in the README carries the measured tools and tokens."""
    surf = _run("measure_surface.py")
    rows = re.findall(r"^([a-z][\w-]*)\s+(\d+)\s+(~?[\d.]+k)\s*$", surf, re.M)
    assert len(rows) >= 4, f"measure_surface pack table not parsed: {rows!r}"
    lines = (ROOT / "README.md").read_text(encoding="utf-8").splitlines()
    for pack, tools, toks in rows:
        cands = [ln for ln in lines
                 if ln.startswith(f"| {pack} ") or ln.startswith(f"| **{pack}**")]
        assert cands, f"README pack table has no row for {pack!r}"
        ln = cands[0]
        assert f" {tools} " in ln and toks.lstrip("~") in ln, (
            f"README row for {pack!r} does not carry measured "
            f"{tools} tools / {toks}: {ln!r}"
        )


def test_llms_txt_pack_costs_match_measurement():
    """The agent-facing doc's per-pack costs are measured too.

    The headline lite/full figures were guarded from the first release; the
    pack line under them was not, and it drifted: it published `design 9
    tools / 6.1k` against a measured 4.1k until the 2026-09-08 trim pass
    caught it by hand. A figure nobody measures is a figure that goes stale,
    so this one is measured.
    """
    surf = _run("measure_surface.py")
    rows = dict((p, t) for p, _n, t in
                re.findall(r"^([a-z][\w-]*)\s+(\d+)\s+~?([\d.]+k)\s*$",
                           surf, re.M))
    text = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    line = next((ln for ln in text.splitlines() if "Pack costs" in ln), None)
    assert line, "llms.txt no longer carries a pack-costs line"
    for pack in ("design", "io", "com"):
        assert rows[pack] in line, (
            f"llms.txt pack costs claim something other than the measured "
            f"{pack} {rows[pack]}: {line!r}")


def test_calc_labels_in_llms_txt_are_the_wire_values():
    """The agent-facing doc uses the labels the code actually emits.

    The lay-facing copy and the wire used to disagree: the website said
    "calculated" where core/calc.py emitted `computed`, and an agent branching
    on the marketing word got nothing. The pre-release rename settled it on
    `calculated` in both places. This guard keeps llms.txt reading the
    constants rather than a copy of them, so the next divergence goes red.
    """
    sys.path.insert(0, str(ROOT / "src"))
    from xlsx_mcp.core import calc

    text = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    for label in (calc.LABEL_VALUE, calc.LABEL_CACHED, calc.LABEL_CALCULATED,
                  calc.LABEL_ABSENT, calc.LABEL_FORMULA):
        assert f"`{label}`" in text, (
            f"llms.txt does not document the {label!r} calc label"
        )


def test_closed_error_codes_are_all_documented():
    """Every code an agent can receive appears in llms.txt Section 5."""
    sys.path.insert(0, str(ROOT / "src"))
    from xlsx_mcp import envelope

    text = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    missing = sorted(c for c in envelope.CLOSED_CODES if f"`{c}`" not in text)
    assert not missing, f"llms.txt does not document error codes: {missing}"


def test_pack_membership_matches_llms_txt():
    """Every registered tool is listed under its pack in llms.txt."""
    sys.path.insert(0, str(ROOT / "src"))
    from xlsx_mcp import packs, server  # noqa: F401  (server registers tools)

    text = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    listed = set(re.findall(r"`([a-z_0-9]+)`", text))
    for pack, members in packs.tool_names().items():
        missing = [m for m in members if m not in listed]
        assert not missing, f"llms.txt pack {pack!r} is missing {missing}"


@pytest.mark.parametrize("path", _public_files(), ids=lambda p: p.name)
def test_no_em_dashes_in_public_copy(path: Path):
    assert "—" not in path.read_text(encoding="utf-8"), (
        f"{path.name} contains an em dash"
    )


def test_documented_module_entry_point_resolves():
    """The README's `python -m` line must name a module that can be run.

    The install gauntlet found the guessable names all wrong: the
    distribution is kitchensink4xl, the package is xlsx_mcp, and
    `import kitchensink4xl` fails. README now points at
    `python -m xlsx_mcp.server`, so that module has to stay runnable, and it
    has to reach the same entry point the console scripts do.

    1.1 adds the shorter `python -m xlsx_mcp` route to the same two
    documents, which is why both routes are asserted here rather than only
    the `.server` one: the package `__main__` shipped in 1.0.0's tree but
    not in its release, so until now the docs deliberately named only the
    route the published wheel could honor.
    """
    import importlib.util

    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    llms = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    for route in ("python -m xlsx_mcp.server", "python -m xlsx_mcp`"):
        assert route in readme, (
            f"README no longer documents the module entry point {route!r}")
    assert "python -m xlsx_mcp`" in llms and "python -m xlsx_mcp.server" in llms, (
        "docs/llms.txt no longer documents both module entry points")
    assert importlib.util.find_spec("xlsx_mcp.server") is not None
    assert importlib.util.find_spec("xlsx_mcp.__main__") is not None
    from xlsx_mcp import __main__ as pkg_main
    from xlsx_mcp.server import main as server_main

    assert pkg_main.main is server_main


def test_mcp_name_marker_survives():
    first = (ROOT / "README.md").read_text(encoding="utf-8").splitlines()[0]
    assert first == (
        "<!-- mcp-name: io.github.nometalalchemist/kitchensink4xl -->"
    ), "README line 1 mcp-name marker was lost"


def test_non_affiliation_disclaimer_present():
    for path in _public_files():
        assert "Not affiliated with or endorsed by Microsoft" in (
            path.read_text(encoding="utf-8")
        ), f"{path.name} is missing the non-affiliation disclaimer"


def test_version_is_consistent_across_manifests():
    """pyproject, server.json, the bundle manifest and the package agree."""
    import tomllib

    pyproject = tomllib.loads(
        (ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    version = pyproject["project"]["version"]

    server_json = json.loads((ROOT / "server.json").read_text(encoding="utf-8"))
    assert server_json["version"] == version
    assert server_json["packages"][0]["version"] == version

    manifest = json.loads(
        (ROOT / "bundle" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["version"] == version
    assert manifest["server"]["mcp_config"]["args"] == [
        f"kitchensink4xl[com]=={version}"
    ], "the bundle's uvx pin does not match the package version"

    sys.path.insert(0, str(ROOT / "src"))
    import xlsx_mcp
    if xlsx_mcp.__version__ == "0.0.0":
        pytest.skip(
            "package version not stamped yet; stamping "
            "src/xlsx_mcp/__init__.py at ship turns this guard on"
        )
    assert xlsx_mcp.__version__ == version, (
        f"src/xlsx_mcp/__init__.py says {xlsx_mcp.__version__}, "
        f"pyproject says {version}"
    )


def test_no_beta_language_in_public_copy():
    """Excel ships as a full 1.0 (author ruling). No beta labels anywhere."""
    for path in _public_files():
        text = path.read_text(encoding="utf-8").lower()
        assert "beta" not in text, f"{path.name} still carries a beta label"

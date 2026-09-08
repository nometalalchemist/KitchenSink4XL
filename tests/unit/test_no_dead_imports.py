"""An import nothing uses is a claim nothing backs.

Six of them had accumulated in src/ by September 2026 and CI could not have
caught any: the workflow runs pytest and nothing else (fat audit 2026-09-08,
finding 9). The obvious fix is a linter, which means a new dependency, a
network install, and a second gate to keep green. This is the same catch in
forty lines of the standard library, inside the suite that already runs.

Deliberately narrow. It flags a module-level import whose bound name appears
nowhere else in the file, and nothing else: no style, no complexity, no
opinions. Feature-probe imports (inside a try), re-exports (named in
__all__), __future__, star imports, and anything marked noqa are all
legitimate and are skipped.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

SRC = Path(__file__).resolve().parents[2] / "src"


def _module_files() -> list[Path]:
    return sorted(p for p in SRC.rglob("*.py")
                  if "__pycache__" not in p.parts)


def _bound_names(node: ast.Import | ast.ImportFrom):
    """The names an import statement actually binds, with the source text
    that would have to be deleted to remove each."""
    for alias in node.names:
        if alias.name == "*":
            continue
        bound = alias.asname or alias.name.split(".")[0]
        yield bound, (alias.asname or alias.name)


def _unused_imports(path: Path) -> list[str]:
    text = path.read_text(encoding="utf-8")
    tree = ast.parse(text, filename=str(path))
    lines = text.splitlines()

    # Imports nested inside a try are feature probes: the binding may be
    # unused because the IMPORT ITSELF is the test (envelope.py's lxml).
    probed: set[int] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Try):
            for child in ast.walk(node):
                if isinstance(child, ast.Import | ast.ImportFrom):
                    probed.add(id(child))

    imports: dict[str, tuple[int, str]] = {}
    for node in tree.body:
        if not isinstance(node, ast.Import | ast.ImportFrom):
            continue
        if id(node) in probed:
            continue
        if isinstance(node, ast.ImportFrom) and node.module == "__future__":
            continue
        line = lines[node.lineno - 1]
        for bound, shown in _bound_names(node):
            src_line = lines[min(node.end_lineno, len(lines)) - 1]
            if "noqa" in line or "noqa" in src_line:
                continue
            imports[bound] = (node.lineno, shown)

    used: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Name):
            used.add(node.id)
        elif isinstance(node, ast.Attribute):
            pass  # the base of an attribute chain is a Name, caught above
        elif isinstance(node, ast.Constant) and isinstance(node.value, str):
            # __all__ entries and string annotations keep a re-export alive
            used.add(node.value)

    try:
        where = path.relative_to(SRC)
    except ValueError:
        where = path.name
    return [f"{where}:{lineno} {shown}"
            for name, (lineno, shown) in sorted(imports.items())
            if name not in used]


@pytest.mark.parametrize("path", _module_files(),
                         ids=lambda p: str(p.relative_to(SRC)))
def test_module_has_no_unused_imports(path: Path):
    dead = _unused_imports(path)
    assert not dead, "imports nothing uses: " + ", ".join(dead)


def test_the_check_can_actually_see_a_dead_import(tmp_path):
    """Red-first, permanently: the detector is itself pinned, so it cannot
    quietly degrade into a test that passes because it stopped looking."""
    live = tmp_path / "live.py"
    live.write_text("import os\n\nprint(os.sep)\n", encoding="utf-8")
    dead = tmp_path / "dead.py"
    dead.write_text("import os\nimport re\n\nprint(os.sep)\n",
                    encoding="utf-8")
    probe = tmp_path / "probe.py"
    probe.write_text("try:\n    import lxml\nexcept ImportError:\n    pass\n",
                     encoding="utf-8")
    excused = tmp_path / "excused.py"
    excused.write_text("import re  # noqa: F401\n", encoding="utf-8")

    assert _unused_imports(live) == []
    assert len(_unused_imports(dead)) == 1
    assert "re" in _unused_imports(dead)[0]
    assert _unused_imports(probe) == []
    assert _unused_imports(excused) == []

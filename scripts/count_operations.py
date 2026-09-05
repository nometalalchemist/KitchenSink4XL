"""Count distinct workbook OPERATIONS across the KitchenSink4XL surface.

Ported from KitchenSink4Word/scripts/count_operations.py and adapted to this
codebase. The reason the port exists: a tool count undersells a consolidated
surface. Many tools here are action-dispatching multiplexers
(manage_worksheet action=add|delete|rename|copy|reorder|hide|unhide,
modify_grid_structure action=..., validate checks=..., manage_name
action=..., and so on), and one multiplexer performs several distinct
operations. The published HUMAN-facing figure is the number of distinct
operations, and this script derives it from committed source so it is never
hand-math.

THE 187-VS-219 LESSON, which is why this is dispatch-aware and not a
signature count. On the Word tree the first counter read only each tool's own
function body and reported 182. A re-audit found dispatch values validated one
layer down, inside the ops module the tool delegates to, and the figure moved
to 187. A second audit found dispatch values that ride INSIDE a caller-supplied
argument (a batch of edit dicts where the ops layer reads `op = edit["op"]`
and validates it against a committed table), and the figure moved to 219. All
three runs were "the count"; only the last one was the surface. So this script
looks in all three places:

  1. THE TOOL'S OWN BODY. The distinct values of the tool's primary dispatch
     parameter (action / operation / op / check / direction / type / mode /
     kind), read out of the validation the tool itself performs: a membership
     test against a tuple/set/dict, an equality dispatch chain, or a table
     lookup.
  2. ONE LAYER DOWN, THROUGH THE DELEGATE. Nearly every tool in server.py is a
     thin docstring-carrying wrapper that hands its arguments to an ops module
     function (manage_worksheet -> ops.lifecycle.manage_worksheet, which is
     where `if action not in WS_ACTIONS` actually lives). The scan follows the
     call, matches the passed parameter to the callee's parameter name, and
     reads the validation there. Module-level constants (WS_ACTIONS and its
     siblings) resolve by importing the ops module.
  3. INSIDE THE ARGUMENT. When the dispatch value rides in caller-supplied
     data rather than beside it (`op = edit["op"]` inside a batch loop), the
     scan reads that validation too, under three provenance conditions: the
     local must be named for the key it reads, the dict must trace back to
     caller-supplied data, and the value must be checked against a committed
     value set. Values the code assigns to itself, or discovers in the
     workbook, are not caller-dispatched operations and do not count.

A single-purpose tool contributes exactly 1. There is no second detector on
this tree: KitchenSink4XL has no v1 predecessor and therefore no migration
capability map, so the Word counter's max(scan, migration) reduces to the
scan alone. The result is a defensible lower bound drawn entirely from
committed source.

What is deliberately NOT an operation, so the figure never inflates:
boolean feature flags (backup, allow_loss, verify_com), alternative argument
FORMS for one operation (a location object given as a cell, an R1C1 ref, a
name, or a grid anchor), tuning parameters (the COM timeout), verbosity or
values-shape selectors that change how one read is rendered rather than what
it does, and guidance content (get_workflows recipes). Several of those are
real capability. They are not distinct dispatched operations under this
definition, and counting them would change the yardstick instead of measuring
the surface.

The two pack-toggle tools (enable_tools, disable_tools) are surface plumbing,
not workbook operations, and are excluded. They are registered by direct call
rather than by decorator, so they fall out of the scan anyway; the exclusion
set keeps that from being an accident.

Run:  .venv/Scripts/python.exe -X utf8 scripts/count_operations.py
      ... --json scripts/operations_snapshot.json   (write per-tool snapshot)
      ... --check scripts/operations_snapshot.json  (diff against snapshot)
"""

from __future__ import annotations

import argparse
import ast
import importlib
import json
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# Set by configure(); the module under measurement supplies the constants
# that Name comparators resolve against.
SRC: Path
SERVER_PY: Path
PKG = "xlsx_mcp"
server_mod = None


def configure(root: Path):
    """Point the counter at a checkout."""
    global SRC, SERVER_PY, server_mod
    SRC = root / "src"
    SERVER_PY = SRC / PKG / "server.py"
    sys.path.insert(0, str(SRC))
    server_mod = importlib.import_module(f"{PKG}.server")
    return server_mod


# Parameter names that name a PRIMARY dispatch dimension. The first one of
# these that a tool actually validates against a value-set is the dimension
# whose distinct values are counted as operations.
DISPATCH_NAMES = (
    "action", "operation", "op", "check", "checks", "direction", "mode",
    "kind", "type",
)

# Plumbing, not workbook operations.
EXCLUDE = {"enable_tools", "disable_tools"}


def _decorated_with_tool(node: ast.FunctionDef, decorator: str = "_tool") -> bool:
    """True when the function carries the checkout's tool decorator."""
    attr = decorator.split(".")[-1]
    for dec in node.decorator_list:
        target = dec.func if isinstance(dec, ast.Call) else dec
        if isinstance(target, ast.Name) and target.id == attr:
            return True
        if isinstance(target, ast.Attribute) and target.attr == attr:
            return True
    return False


def _literal_values(node: ast.AST):
    """String values from a tuple/list/set literal or a dict literal's keys."""
    if isinstance(node, (ast.Tuple, ast.List, ast.Set)):
        vals = [el.value for el in node.elts
                if isinstance(el, ast.Constant) and isinstance(el.value, str)]
        return vals if vals and len(vals) == len(node.elts) else None
    if isinstance(node, ast.Dict):
        vals = [k.value for k in node.keys
                if isinstance(k, ast.Constant) and isinstance(k.value, str)]
        return vals if vals and len(vals) == len(node.keys) else None
    return None


def _local_map(fn: ast.FunctionDef):
    """Map local names assigned a tuple/list/set/dict of string constants."""
    out: dict[str, list[str]] = {}
    for node in ast.walk(fn):
        if isinstance(node, ast.Assign):
            vals = _literal_values(node.value)
            if vals:
                for tgt in node.targets:
                    if isinstance(tgt, ast.Name):
                        out[tgt.id] = vals
    return out


def _scan_param_values(fn: ast.FunctionDef, wanted: tuple[str, ...],
                       module_obj=None) -> dict[str, set[str]]:
    """Union every validation signal for each wanted parameter name inside
    one function:
      - `X in/not in <set|Name>`            (membership tests)
      - `X == "const"` / `"const" == X`     (if/elif dispatch chains)
      - `DICT.get(X)`, `DICT[X]`, `X in DICT` (table-driven dispatch)
    Name comparators resolve against local assignments first, then the given
    module's constants."""
    if module_obj is None:
        module_obj = server_mod
    local = _local_map(fn)
    found: dict[str, set[str]] = {}

    def module_values(name: str):
        val = getattr(module_obj, name, None)
        if isinstance(val, dict):
            return [k for k in val.keys() if isinstance(k, str)]
        if isinstance(val, (tuple, list, set, frozenset)):
            return [v for v in val if isinstance(v, str)]
        return None

    def resolve(node: ast.AST):
        lit = _literal_values(node)
        if lit is not None:
            return lit
        if isinstance(node, ast.Name):
            return local.get(node.id) or module_values(node.id)
        return None

    def add(param: str, vals):
        if vals:
            found.setdefault(param, set()).update(vals)

    for sub in ast.walk(fn):
        # Membership: X in/not in <set|Name|DICT>
        if isinstance(sub, ast.Compare) and len(sub.ops) == 1:
            op = sub.ops[0]
            left = sub.left
            comp = sub.comparators[0]
            if isinstance(op, (ast.In, ast.NotIn)) and isinstance(left, ast.Name) \
                    and left.id in wanted:
                add(left.id, resolve(comp))
            # Equality dispatch: X == "const"  /  "const" == X
            if isinstance(op, (ast.Eq,)):
                for a, b in ((left, comp), (comp, left)):
                    if isinstance(a, ast.Name) and a.id in wanted \
                            and isinstance(b, ast.Constant) \
                            and isinstance(b.value, str):
                        add(a.id, [b.value])
        # Table-driven: DICT.get(X) or DICT[X]
        if isinstance(sub, ast.Call) and isinstance(sub.func, ast.Attribute) \
                and sub.func.attr == "get" and sub.args:
            arg = sub.args[0]
            base = sub.func.value
            if isinstance(arg, ast.Name) and arg.id in wanted \
                    and isinstance(base, ast.Name):
                add(arg.id, resolve(base))
        if isinstance(sub, ast.Subscript) and isinstance(sub.value, ast.Name):
            idx = sub.slice
            if isinstance(idx, ast.Name) and idx.id in wanted:
                add(idx.id, resolve(sub.value))

    return found


def _ops_index():
    """(module_name -> (ast func map, imported module)) for xlsx_mcp.ops.*,
    plus the alias map server.py imports them under (lifecycle as _lifecycle,
    and the rest)."""
    tree = ast.parse(SERVER_PY.read_text(encoding="utf-8"))
    aliases: dict[str, str] = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module in ("ops", ".ops"):
            for a in node.names:
                aliases[a.asname or a.name] = a.name
    modules: dict[str, tuple[dict[str, ast.FunctionDef], object]] = {}
    for modname in set(aliases.values()):
        path = SRC / PKG / "ops" / f"{modname}.py"
        if not path.exists():
            continue
        mtree = ast.parse(path.read_text(encoding="utf-8"))
        fns = {n.name: n for n in ast.walk(mtree)
               if isinstance(n, ast.FunctionDef)}
        modules[modname] = (fns, importlib.import_module(
            f"{PKG}.ops.{modname}"))
    return aliases, modules


def _delegate_values(fn: ast.FunctionDef, aliases, modules):
    """Follow calls that pass a dispatch-named parameter through to an
    ops-module function (positionally or as a keyword) and scan the callee's
    validation for that parameter. Returns {tool_param: set[values]}."""
    out: dict[str, set[str]] = defaultdict(set)
    for sub in ast.walk(fn):
        if not (isinstance(sub, ast.Call) and isinstance(sub.func, ast.Attribute)
                and isinstance(sub.func.value, ast.Name)
                and sub.func.value.id in aliases):
            continue
        modname = aliases[sub.func.value.id]
        if modname not in modules:
            continue
        fns, mod_obj = modules[modname]
        callee = fns.get(sub.func.attr)
        if callee is None:
            continue
        pos = [a.arg for a in callee.args.args]
        pairs = []  # (callee_param_name, passed node)
        for i, arg in enumerate(sub.args):
            if i < len(pos):
                pairs.append((pos[i], arg))
        for kw in sub.keywords:
            if kw.arg:
                pairs.append((kw.arg, kw.value))
        for callee_param, node in pairs:
            if isinstance(node, ast.Name) and node.id in DISPATCH_NAMES:
                vals = _scan_param_values(
                    callee, (callee_param,), mod_obj).get(callee_param)
                if vals:
                    out[node.id] |= vals
                elem = _element_values(callee, callee_param, mod_obj)
                if elem:
                    out[node.id] |= elem
    return out


def _element_values(fn: ast.FunctionDef, param: str, module_obj):
    """Dispatch values validated for the ELEMENTS of a caller-supplied list.

    The battery pattern: `validate(checks=[...])` takes a list and the ops
    layer does `for check in checks: if check not in _CHECKS: refuse`. The
    caller dispatches one operation per element, and the evidence is the same
    membership test against a committed table the parameter scan already
    accepts; only the container differs (a list instead of a bare argument or
    a dict field).

    Two conditions keep it from inflating. The loop variable has to trace back
    to the caller's own parameter through iteration, and the loop variable has
    to itself be dispatch-named (`check` out of `checks`), which is the same
    naming test the dict-field scan applies. A loop over rows, cells, or
    anything else the code assembled never qualifies.
    """
    tracked = {param}
    elements: set[str] = set()

    def from_tracked(node: ast.AST) -> bool:
        if isinstance(node, ast.Name):
            return node.id in tracked
        if isinstance(node, ast.Subscript):
            return from_tracked(node.value)
        if isinstance(node, (ast.IfExp,)):
            return from_tracked(node.body) or from_tracked(node.orelse)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) \
                and node.func.id in ("enumerate", "list", "reversed", "iter",
                                     "sorted", "set", "tuple") and node.args:
            return from_tracked(node.args[0])
        return False

    def bind(target: ast.AST) -> None:
        if isinstance(target, ast.Name):
            if target.id in DISPATCH_NAMES:
                elements.add(target.id)
        elif isinstance(target, (ast.Tuple, ast.List)):
            for el in target.elts:
                bind(el)

    for node in ast.walk(fn):
        if isinstance(node, ast.For) and from_tracked(node.iter):
            bind(node.target)
        elif isinstance(node, ast.comprehension) and from_tracked(node.iter):
            bind(node.target)
    if not elements:
        return set()
    found = _scan_param_values(fn, tuple(sorted(elements)), module_obj)
    out: set[str] = set()
    for vals in found.values():
        out |= vals
    return out


def _caller_data_names(fn: ast.FunctionDef, seeded: set[str]) -> set[str]:
    """Names inside `fn` that hold CALLER-SUPPLIED data.

    Seeded with the parameters the caller's data arrived in, then grown
    through the ways a caller's collection gets unpacked: `for x in
    <tracked>`, `for i, x in enumerate(<tracked>)`, `x = <tracked>`, and
    `x = <tracked>[...]`. Names built from anything else (a plan the code
    assembled, a value the code discovered in the workbook) never enter the
    set. This is the guard that keeps the nested scan honest: only a value the
    CALLER chooses is a dispatched operation."""
    tracked = set(seeded)

    def from_tracked(node: ast.AST) -> bool:
        if isinstance(node, ast.Name):
            return node.id in tracked
        if isinstance(node, ast.Subscript):
            return from_tracked(node.value)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) \
                and node.func.id in ("enumerate", "list", "reversed", "iter") \
                and node.args:
            return from_tracked(node.args[0])
        # tracked.get("key") reaches caller data exactly as tracked["key"] does
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) \
                and node.func.attr in ("get", "pop"):
            return from_tracked(node.func.value)
        return False

    def bind(target: ast.AST) -> None:
        if isinstance(target, ast.Name):
            tracked.add(target.id)
        elif isinstance(target, (ast.Tuple, ast.List)):
            for el in target.elts:
                bind(el)

    for _ in range(3):  # small fixed point; assignment order is not guaranteed
        before = len(tracked)
        for node in ast.walk(fn):
            if isinstance(node, ast.For) and from_tracked(node.iter):
                bind(node.target)
            elif isinstance(node, ast.comprehension) and from_tracked(node.iter):
                bind(node.target)
            elif isinstance(node, ast.Assign) and from_tracked(node.value):
                for tgt in node.targets:
                    bind(tgt)
        if len(tracked) == before:
            break
    return tracked


def _nested_key_values(callee: ast.FunctionDef, fns, mod_obj, depth: int = 2,
                       seeded: set[str] | None = None):
    """Dispatch values validated for a key read OUT of a caller-supplied dict.

    Some multiplexers carry their dispatch value inside the argument instead
    of beside it: a batch tool takes a list of edit dicts and the ops layer
    does `op = edit["op"]`, then validates `op` against a committed table. The
    evidence is the same kind the parameter scan already accepts (a membership
    test against a committed value set); only the container differs, so the
    scan reads it the same way.

    Three conditions keep it from inflating. The local must be named for the
    key it reads, the dict it reads from must trace back to caller-supplied
    data (see _caller_data_names), and the value must be validated against a
    committed value set. Internal plan fields the code assigns to itself, and
    kinds the code discovers in the workbook, fail the provenance test and are
    not operations. Same-module helpers are followed `depth` levels with
    provenance carried across the call, since the entry point usually
    delegates validation to a private validator."""
    out: dict[str, set[str]] = defaultdict(set)
    seen: set[str] = set()
    seed = seeded if seeded is not None else {
        a.arg for a in callee.args.args + callee.args.kwonlyargs}
    stack = [(callee, depth, seed)]
    while stack:
        fn, level, fn_seed = stack.pop()
        if fn.name in seen:
            continue
        seen.add(fn.name)
        tracked = _caller_data_names(fn, fn_seed)
        for node in ast.walk(fn):
            if not (isinstance(node, ast.Assign) and len(node.targets) == 1
                    and isinstance(node.targets[0], ast.Name)):
                continue
            name = node.targets[0].id
            if name not in DISPATCH_NAMES:
                continue
            val = node.value
            key = base = None
            if isinstance(val, ast.Subscript) and isinstance(
                    val.slice, ast.Constant) and isinstance(
                        val.slice.value, str):
                key, base = val.slice.value, val.value
            elif (isinstance(val, ast.Call)
                  and isinstance(val.func, ast.Attribute)
                  and val.func.attr == "get" and val.args
                  and isinstance(val.args[0], ast.Constant)
                  and isinstance(val.args[0].value, str)):
                key, base = val.args[0].value, val.func.value
            if key != name or not isinstance(base, ast.Name) \
                    or base.id not in tracked:
                continue
            vals = _scan_param_values(fn, (name,), mod_obj).get(name)
            if vals:
                out[name] |= vals
        if level > 1:
            for node in ast.walk(fn):
                if not (isinstance(node, ast.Call)
                        and isinstance(node.func, ast.Name)):
                    continue
                helper = fns.get(node.func.id)
                if helper is None:
                    continue
                pos = [a.arg for a in helper.args.args]
                passed = set()
                for i, arg in enumerate(node.args):
                    if i < len(pos) and isinstance(arg, ast.Name) \
                            and arg.id in tracked:
                        passed.add(pos[i])
                for kw in node.keywords:
                    if kw.arg and isinstance(kw.value, ast.Name) \
                            and kw.value.id in tracked:
                        passed.add(kw.arg)
                if passed:
                    stack.append((helper, level - 1, passed))
    return out


def _nested_values(fn: ast.FunctionDef, aliases, modules):
    """Union of _nested_key_values across every ops function this tool calls."""
    out: dict[str, set[str]] = defaultdict(set)
    for sub in ast.walk(fn):
        if not (isinstance(sub, ast.Call) and isinstance(sub.func, ast.Attribute)
                and isinstance(sub.func.value, ast.Name)
                and sub.func.value.id in aliases):
            continue
        modname = aliases[sub.func.value.id]
        if modname not in modules:
            continue
        fns, mod_obj = modules[modname]
        callee = fns.get(sub.func.attr)
        if callee is None:
            continue
        for param, vals in _nested_key_values(callee, fns, mod_obj).items():
            out[param] |= vals
    return out


def _dispatch_values(fn: ast.FunctionDef, aliases, modules):
    """Return (param_name, sorted[values]) for the primary dispatch dimension,
    or None. Unions the tool function's own validation with any validation
    found by following pass-through calls into ops modules. Picks the dispatch
    param with the largest value union; ties break by DISPATCH_NAMES priority
    so `action` beats a minor secondary `type`."""
    found = _scan_param_values(fn, DISPATCH_NAMES)
    for param in (a.arg for a in fn.args.args + fn.args.kwonlyargs):
        if param not in DISPATCH_NAMES:
            continue
        elem = _element_values(fn, param, server_mod)
        if elem:
            found.setdefault(param, set()).update(elem)
    for param, vals in _delegate_values(fn, aliases, modules).items():
        found.setdefault(param, set()).update(vals)
    for param, vals in _nested_values(fn, aliases, modules).items():
        found.setdefault(param, set()).update(vals)
    if not found:
        return None
    best = max(found, key=lambda p: (len(found[p]),
                                     -DISPATCH_NAMES.index(p)))
    return best, sorted(found[best])


def measure(decorator: str = "_tool") -> dict:
    tree = ast.parse(SERVER_PY.read_text(encoding="utf-8"))
    tool_fns = {
        n.name: n
        for n in ast.walk(tree)
        if isinstance(n, ast.FunctionDef) and _decorated_with_tool(n, decorator)
    }
    aliases, modules = _ops_index()

    total = 0
    multiplex = 0
    single = 0
    rows = []
    per_tool: dict[str, int] = {}
    for name in sorted(tool_fns):
        if name in EXCLUDE:
            continue
        disp = _dispatch_values(tool_fns[name], aliases, modules)
        ops = len(disp[1]) if disp else 1
        detail = ""
        if disp:
            detail = f"{disp[0]}: {', '.join(disp[1])}"
        if ops > 1:
            multiplex += 1
        else:
            single += 1
        rows.append((name, ops, detail))
        per_tool[name] = ops
        total += ops

    return {
        "tools_counted": len(tool_fns) - len(EXCLUDE & set(tool_fns)),
        "single": single,
        "multiplex": multiplex,
        "total_operations": total,
        "rows": rows,
        "per_tool": per_tool,
    }


def _report(res: dict, label: str) -> None:
    print(f"{'tool':<28} {'ops':>4}  dispatch values")
    print("-" * 90)
    for name, ops, detail in res["rows"]:
        d = detail if len(detail) <= 52 else detail[:49] + "..."
        print(f"{name:<28} {ops:>4}  {d}")
    print("-" * 90)
    print(f"{label}")
    print(f"tools counted (excl. enable/disable_tools): {res['tools_counted']}")
    print(f"  single-operation tools:                   {res['single']}")
    print(f"  multi-operation tools:                    {res['multiplex']}")
    print(f"TOTAL DISTINCT OPERATIONS:                  {res['total_operations']}")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--root", default=str(ROOT),
                    help="checkout to measure (default: this repo)")
    ap.add_argument("--decorator", default="_tool",
                    help="tool decorator to look for")
    ap.add_argument("--label", default="")
    ap.add_argument("--json", dest="json_out",
                    help="write the per-tool snapshot to this path")
    ap.add_argument("--check", dest="check",
                    help="diff the measurement against a stored snapshot")
    args = ap.parse_args(argv)

    configure(Path(args.root).resolve())
    res = measure(decorator=args.decorator)
    _report(res, args.label or f"root: {args.root}")

    if args.json_out:
        payload = {
            "tree": args.label or Path(args.root).resolve().name,
            "decorator": args.decorator,
            "tools_counted": res["tools_counted"],
            "total_operations": res["total_operations"],
            "per_tool": res["per_tool"],
        }
        Path(args.json_out).write_text(
            json.dumps(payload, indent=2, sort_keys=True) + "\n",
            encoding="utf-8")
        print(f"snapshot written: {args.json_out}")

    if args.check:
        stored = json.loads(Path(args.check).read_text(encoding="utf-8"))
        old, new = stored["per_tool"], res["per_tool"]
        diffs = []
        for name in sorted(set(old) | set(new)):
            if old.get(name) != new.get(name):
                diffs.append((name, old.get(name), new.get(name)))
        print("-" * 90)
        if not diffs:
            print(f"snapshot MATCHES {args.check} "
                  f"({stored['total_operations']} operations)")
        else:
            print(f"snapshot DIFFERS from {args.check}: "
                  f"{stored['total_operations']} -> {res['total_operations']}")
            for name, o, n in diffs:
                print(f"  {name:<28} {o} -> {n}")
        return 0
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

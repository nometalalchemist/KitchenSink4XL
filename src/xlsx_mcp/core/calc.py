"""core/calc.py: the formula / calc layer (honest calc story, DESIGN Section 4).

The load-bearing fact (library_gap_analysis fact 2): openpyxl writes formula
strings and computes NOTHING. A file it saves after a formula edit has no
cached value for those cells, so every non-Excel consumer reads them empty
until Excel opens and recalculates. This module owns:

  - the _xlfn / _xlfn._xlws / _xlpm normalization shim so modern functions do
    not land as #NAME?,
  - cached-value labeling (cached | computed | formula | absent) so no read
    ever silently returns an empty cell where a formula lives,
  - fullCalcOnLoad injection (the no-Excel flag that forces recalc on open),
  - the COM recalc round-trip (the fidelity path, com_ground_truth exp 5),
  - the `formulas`-library pure-Python fallback with a coverage report.

No pure-Python engine computes at Excel fidelity; COM is the fidelity path and
`formulas` is a best-effort convenience with a stated coverage disclaimer.
pywin32 and `formulas` are imported lazily so this module imports on any
platform.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .errors import FormulaRejected

# ---------------------------------------------------------- the _xlfn shim

# Functions Excel stores with a _xlfn. prefix (post-2007 "future functions").
# Writing the bare name via openpyxl lands #NAME? in Excel (or, for dynamic
# arrays, a repair-only file: Phase 1 spike, bare FILTER). The pre-audit list
# held ~34 names; the re-audit expanded it to the full documented future-
# function set (the 2010 statistical dot-family, the 2013 math/engineering
# wave, 2016 IFS/TEXTJOIN, and the 365 dynamic-array and lambda-helper era),
# because a missing entry means a silently broken formula write. Sources: the
# MS-XLSX future-function storage rules and Microsoft's _xlfn documentation;
# XLOOKUP was verified end to end in Excel by the Phase 1 spike.
XLFN_FUNCS = frozenset({
    # 2010: statistical renames and friends
    "AGGREGATE", "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.INV",
    "CEILING.PRECISE", "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV",
    "CHISQ.INV.RT", "CHISQ.TEST", "CONFIDENCE.NORM", "CONFIDENCE.T",
    "COVARIANCE.P", "COVARIANCE.S", "ERF.PRECISE", "ERFC.PRECISE",
    "EXPON.DIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST",
    "FLOOR.PRECISE", "GAMMA.DIST", "GAMMA.INV", "GAMMALN.PRECISE",
    "HYPGEOM.DIST", "ISO.CEILING", "LOGNORM.DIST", "LOGNORM.INV",
    "MODE.MULT", "MODE.SNGL", "NEGBINOM.DIST", "NETWORKDAYS.INTL",
    "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "PERCENTILE.EXC",
    "PERCENTILE.INC", "PERCENTRANK.EXC", "PERCENTRANK.INC", "POISSON.DIST",
    "QUARTILE.EXC", "QUARTILE.INC", "RANK.AVG", "RANK.EQ", "STDEV.P",
    "STDEV.S", "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T",
    "T.TEST", "VAR.P", "VAR.S", "WEIBULL.DIST", "WORKDAY.INTL", "Z.TEST",
    # 2013: math / engineering / info wave
    "ACOT", "ACOTH", "ARABIC", "BASE", "BINOM.DIST.RANGE", "BITAND",
    "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "CEILING.MATH", "COMBINA",
    "COT", "COTH", "CSC", "CSCH", "DAYS", "DECIMAL", "ENCODEURL",
    "FILTERXML", "FLOOR.MATH", "FORMULATEXT", "GAMMA", "GAUSS", "IFNA",
    "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMSEC", "IMSECH", "IMSINH",
    "IMTAN", "ISFORMULA", "ISOWEEKNUM", "MUNIT", "NUMBERVALUE", "PDURATION",
    "PERMUTATIONA", "PHI", "RRI", "SEC", "SECH", "SHEET", "SHEETS", "SKEW.P",
    "UNICHAR", "UNICODE", "WEBSERVICE", "XOR",
    # 2016: forecasting and aggregation
    "FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT", "FORECAST.LINEAR", "CONCAT", "IFS", "MAXIFS",
    "MINIFS", "SWITCH", "TEXTJOIN",
    # 365: lookups, LET/LAMBDA and helpers, dynamic-array builders, text
    "XLOOKUP", "XMATCH", "LET", "LAMBDA", "SINGLE", "ANCHORARRAY",
    "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY",
    "MAP", "REDUCE", "SCAN", "MAKEARRAY", "BYROW", "BYCOL", "ISOMITTED",
    "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT", "VSTACK", "HSTACK", "TOCOL",
    "TOROW", "WRAPROWS", "WRAPCOLS", "TAKE", "DROP", "EXPAND", "CHOOSECOLS",
    "CHOOSEROWS", "ARRAYTOTEXT", "VALUETOTEXT", "STOCKHISTORY", "IMAGE",
    "GROUPBY", "PIVOTBY", "PERCENTOF", "REGEXTEST", "REGEXEXTRACT",
    "REGEXREPLACE", "TRANSLATE", "DETECTLANGUAGE", "TRIMRANGE",
})

# Functions stored with the _xlfn._xlws. worksheet-scope prefix. Documented
# storage uses this ONLY for FILTER and SORT; the Phase 1 spike verified
# _xlfn._xlws.FILTER end to end in Excel. The pre-audit list wrongly put
# SORTBY / UNIQUE / SEQUENCE / RANDARRAY / ANCHORARRAY / SINGLE here, which
# risks the same repair-only failure the spike proved for a mis-prefixed
# dynamic array; they are plain _xlfn. names (moved to XLFN_FUNCS above).
# COM-tier task: verify each 365-era prefix empirically in Excel.
XLFN_XLWS_FUNCS = frozenset({"FILTER", "SORT"})

_CALL = re.compile(r"(?<![A-Za-z0-9_.])([A-Za-z][A-Za-z0-9_.]*)\s*\(")

# Excel string literals: double-quoted, "" as the escape. Function names
# inside them must never be prefixed.
_STRING_RE = re.compile(r'"(?:[^"]|"")*"')

# ----------------------------------------------------- the _xlpm shim
#
# LET and LAMBDA declare NAMES, and Excel stores every occurrence of such a
# name with an _xlpm. (parameter) prefix. Writing them bare does not merely
# produce a #NAME?: the numbers-safety gate found that Excel REFUSES TO OPEN
# the workbook at all ("Open method of Workbooks class failed") for every
# LET / LAMBDA / MAP / REDUCE / SCAN / BYROW / BYCOL / MAKEARRAY formula.
# So the shim prefixes the declared names too.

_DECLARER = re.compile(
    r"(?<![A-Za-z0-9_.])(?:_xlfn\.)?(LAMBDA|LET)\s*\(", re.IGNORECASE)
_NAME_TOKEN = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.?\\]*$")

#: A LAMBDA OPTIONAL parameter declaration: the name in square brackets,
#: =LAMBDA(x,[y],...). Excel stores it WITHOUT the brackets and under a third
#: prefix, _xlop. (see _OPT_PREFIX below).
_OPT_TOKEN = re.compile(r"^\[\s*([A-Za-z_\\][A-Za-z0-9_.?\\]*)\s*\]$")

#: The optional-parameter prefix Excel writes at a LAMBDA DECLARATION site.
#: Use sites in the body still take _xlpm. (COM ground truth, edge audit
#: 2026-09-04 follow-up: Excel authored =LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))
#: and stored
#:   _xlfn.LAMBDA(_xlpm.x,_xlop.y,IF(_xlfn.ISOMITTED(_xlpm.y),
#:                                   _xlpm.x,_xlpm.x+_xlpm.y))
#: -- brackets gone, _xlop. at the declaration, _xlpm. everywhere else,
#: including at a call site when the optional parameter is lambda-valued
#: (_xlpm.f(_xlpm.x)). A bare [y] was dropped by the name regex before this,
#: so every y landed unprefixed and the workbook would not OPEN.)
_OPT_PREFIX = "_xlop."

#: Every name a call site could mean instead of a declared parameter. Excel's
#: rule at a call site is BUILTIN WINS: =LET(mod,2,mod+MOD(7,3)) is stored
#: _xlfn.LET(_xlpm.mod,2,_xlpm.mod+MOD(7,3)) and evaluates 3 (COM ground
#: truth, edge audit 2026-09-04). Enforcing that rule needs the WHOLE builtin
#: catalog, not a sample: the hand-written 30-name list below left MOD, TRIM,
#: TODAY, COUNTIF, SUMIF, CONCATENATE, ROUNDUP and several hundred others
#: outside the guard, so any of them could be captured into _xlpm.MOD(7,3).
#: openpyxl.utils.FORMULAE is the ECMA-376 classic function catalog (355
#: names) and ships with a hard dependency, so it is used as the catalog; the
#: hand list stays as a floor in case that export ever moves. The classic set
#: is CLOSED by construction -- every function Excel has added since 2007 is
#: stored with an _xlfn. prefix and lives in XLFN_FUNCS / XLFN_XLWS_FUNCS --
#: so the union below is complete, not a subset that has to keep growing.
_CLASSIC_SEED = frozenset({
    "SUM", "IF", "AND", "OR", "NOT", "MIN", "MAX", "AVERAGE", "COUNT",
    "COUNTA", "INDEX", "MATCH", "VLOOKUP", "HLOOKUP", "LOOKUP", "TEXT",
    "LEN", "LEFT", "RIGHT", "MID", "ROUND", "ABS", "NA", "ROW", "COLUMN",
    "ROWS", "COLUMNS", "IFERROR", "SUMPRODUCT", "OFFSET", "INDIRECT",
})


def _classic_catalog() -> frozenset[str]:
    try:
        from openpyxl.utils import FORMULAE  # type: ignore
        return frozenset(str(n).upper() for n in FORMULAE)
    except Exception:  # noqa: BLE001 - the seed still guards the common names
        return frozenset()


_CLASSIC_FUNCS = _CLASSIC_SEED | _classic_catalog()

#: Every name that is a function call, not a parameter use, when followed by
#: "(" -- the call-site guard's whole vocabulary.
_BUILTIN_FUNCS = _CLASSIC_FUNCS | XLFN_FUNCS | XLFN_XLWS_FUNCS

#: A single-quoted sheet-name span ('Sales x'!A1), with '' as the escape.
#: Excel never writes a parameter prefix inside one: =LET(x,1,x+'Sales x'!A1)
#: is stored with 'Sales x'!A1 untouched (COM ground truth, edge audit
#: 2026-09-04). The "!" lookahead does not cover this, because the occurrence
#: inside the quotes is followed by an apostrophe, not a "!".
_SHEET_QUOTE_RE = re.compile(r"'(?:[^']|'')*'")


def literal_spans(body: str) -> list[tuple[int, int, str]]:
    """Every INERT span of a formula, left to right: double-quoted string
    literals and single-quoted sheet-name spans, each with '' / "" as its own
    escape. Returns [(start, end, kind)] with kind 'string' or 'sheet'; an
    unterminated quote runs to the end of the text.

    ONE scan handling BOTH quote characters is the load-bearing part. Scanning
    for double quotes alone (what this module did before) misreads two legal
    shapes, and the insane round proved both produce a workbook Excel REFUSES
    TO OPEN:

      - a '"' inside a single-quoted sheet name starts a phantom string that
        swallows the rest of the argument list, hiding later declarations, so
        =LET(x,'a"b'!A1,y,"c",x&y) lost its `y` and wrote it bare. Excel's own
        storage for that formula is
        _xlfn.LET(_xlpm.x,'a"b'!A1,_xlpm.y,"c",_xlpm.x&_xlpm.y)
        (COM ground truth, fix wave 2, 2026-09-05): the double quote inside
        the sheet name is ordinary text, not a string delimiter.
      - an unbalanced '(' inside a single-quoted sheet name (legal: Excel bans
        only : \\ / ? * [ ]) threw the paren depth counter off, so the whole
        _xlpm pass silently did nothing.
    """
    spans: list[tuple[int, int, str]] = []
    i, n = 0, len(body)
    while i < n:
        ch = body[i]
        if ch not in ('"', "'"):
            i += 1
            continue
        kind = "string" if ch == '"' else "sheet"
        j = i + 1
        while j < n:
            if body[j] == ch:
                if j + 1 < n and body[j + 1] == ch:
                    j += 2          # the doubled-quote escape
                    continue
                j += 1
                break
            j += 1
        else:
            j = n                   # unterminated: runs to the end
        spans.append((i, min(j, n), kind))
        i = j
    return spans


def _blank_literals(body: str) -> str:
    """The formula with every inert span (string literal AND single-quoted
    sheet name) replaced by same-length filler, so a structural scan can use
    the offsets of the original text."""
    out = list(body)
    for start, end, _kind in literal_spans(body):
        for i in range(start, end):
            out[i] = " "
    return "".join(out)


#: An identifier immediately followed by whitespace and then '(' -- a shape
#: Excel's formula bar tolerates on INPUT but its FILE PARSER refuses. COM
#: ground truth (fix wave 2, 2026-09-05): a worksheet part carrying
#: <f>SUM (1,2)</f> gives "Open method of Workbooks class failed", and so does
#: the tab and newline form; Excel's own Formula2 setter rejects the same
#: string with 0x800A03EC. Nothing Excel authors ever contains it, so a
#: formula reaching the write path with one is refused rather than stored.
_SPACE_BEFORE_PAREN = re.compile(r"(?<![A-Za-z0-9_.])([A-Za-z_\\]"
                                 r"[A-Za-z0-9_.?\\]*)[ \t\r\n]+\(")


def _refuse_space_before_paren(body: str) -> None:
    blanked = _blank_literals(body)
    m = _SPACE_BEFORE_PAREN.search(blanked)
    if m is None:
        return
    raise FormulaRejected(
        f"the formula puts whitespace between {m.group(1)!r} and its opening "
        "parenthesis. Excel's formula bar accepts that while you type, but "
        "Excel's FILE parser does not: a workbook stored with "
        f"'{m.group(1)} (' cannot be opened at all ('Open method of Workbooks "
        "class failed', verified against Excel). Remove the space so the call "
        f"reads '{m.group(1)}('.")


def _top_level_args(text: str, open_idx: int) -> tuple[list[tuple[str, int]], int]:
    """Split the argument list of a call whose '(' is at open_idx. Returns
    ([(arg_text, start_offset), ...], index_of_matching_close) or ([], -1)
    when unbalanced. The offsets are what lets an optional-parameter
    declaration be rewritten in place. `text` must already have its inert
    spans blanked (_blank_literals), or a paren inside a quoted sheet name
    throws the depth count off."""
    depth = 0
    args: list[tuple[str, int]] = []
    start = open_idx + 1
    i = open_idx
    while i < len(text):
        ch = text[i]
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                args.append((text[start:i], start))
                return args, i
        elif ch == "," and depth == 1:
            args.append((text[start:i], start))
            start = i + 1
        i += 1
    return [], -1


@dataclass
class Declaration:
    """One LET / LAMBDA call: the names it declares and the span they are
    prefixed inside."""
    start: int          # index of the 'L' of LET/LAMBDA in the body
    end: int            # index just past the matching ')'
    names: list[str]


def _declarations(body: str) -> tuple[list[Declaration], list[tuple[int, int, str]]]:
    """Every LET / LAMBDA in the formula as a SCOPED declaration, plus the
    spans of the OPTIONAL LAMBDA declarations ([y]) that need _xlop.

    SCOPE IS REAL, and the previous global substitution got it wrong. The old
    comment here claimed "Excel prefixes every occurrence of a declared name
    regardless of scope"; Excel disagrees for occurrences OUTSIDE the
    declaring call. COM ground truth (fix wave 2, 2026-09-05), authored by
    Excel with workbook-scoped defined names Rate / Factor / a present:

        =LET(Rate,1,Rate)+Rate
          -> _xlfn.LET(_xlpm.Rate,1,_xlpm.Rate)+Rate
        =LAMBDA(Factor,Factor*2)(3)+Factor
          -> _xlfn.LAMBDA(_xlpm.Factor,_xlpm.Factor*2)(3)+Factor
        =LET(a,1,a)+LET(b,2,b)+a
          -> _xlfn.LET(_xlpm.a,1,_xlpm.a)+_xlfn.LET(_xlpm.b,2,_xlpm.b)+a

    The trailing bare name is the DEFINED NAME, and prefixing it produced a
    file that opens and shows #NAME?. So each declaration carries its own
    span and the prefix pass only rewrites inside it.

    Raises FormulaRejected when a LET/LAMBDA's argument list cannot be parsed.
    Skipping it (the old behavior) is what let bare names reach the file, and
    a bare LET name is not a #NAME?: Excel REFUSES TO OPEN the workbook."""
    blanked = _blank_literals(body)
    decls: list[Declaration] = []
    optional: list[tuple[int, int, str]] = []
    for m in _DECLARER.finditer(blanked):
        kind = m.group(1).upper()
        args, close = _top_level_args(blanked, m.end() - 1)
        if close < 0:
            raise FormulaRejected(
                f"the {kind} at character {m.start() + 1} has an unbalanced "
                "argument list, so its declared parameter names cannot be "
                "resolved. Writing them unprefixed produces a workbook Excel "
                "REFUSES TO OPEN, so this write is refused instead. Check the "
                "parentheses and quotes in the formula.")
        if len(args) < 2:
            raise FormulaRejected(
                f"the {kind} at character {m.start() + 1} declares no "
                f"parameter ({len(args)} argument(s)); {kind} needs at least "
                "a name and a body.")
        if kind == "LAMBDA":
            declared = args[:-1]
        else:  # LET: name, value, name, value, ..., calculation
            declared = [a for i, a in enumerate(args[:-1]) if i % 2 == 0]
        names: list[str] = []
        for raw, at in declared:
            tok = raw.strip()
            if not tok:
                continue
            if _NAME_TOKEN.match(tok):
                names.append(tok)
                continue
            # Only LAMBDA takes optional parameters; LET names cannot be
            # bracketed.
            om = _OPT_TOKEN.match(tok) if kind == "LAMBDA" else None
            if om:
                lead = len(raw) - len(raw.lstrip())
                names.append(om.group(1))
                optional.append((at + lead, at + lead + len(tok), om.group(1)))
        if names:
            decls.append(Declaration(m.start(), close + 1, names))
    return decls, optional


def _mark_optional_declarations(
        body: str, spans: list[tuple[int, int, str]]) -> str:
    """Rewrite each `[y]` LAMBDA declaration to `_xlop.y` in place. Done
    BEFORE the _xlpm pass, which then leaves the result alone (the name is
    preceded by a '.', which the lookbehind excludes) and prefixes only the
    body's use sites."""
    out: list[str] = []
    pos = 0
    for start, end, name in sorted(spans):
        if start < pos:
            continue
        out.append(body[pos:start])
        out.append(_OPT_PREFIX + name)
        pos = end
    out.append(body[pos:])
    return "".join(out)


#: An identifier as Excel's formula grammar spells one. Maximal munch, and
#: '.' is a continuation character, so an ALREADY-prefixed name (_xlpm.x) is
#: one token that can never collide with the bare name -- which is what makes
#: the whole pass idempotent.
_IDENT = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.?\\]*")

#: Characters that, immediately BEFORE an identifier, mean it is not a
#: parameter use: '$' (an absolute address) and '!' (a sheet-qualified name).
_NOT_BEFORE = "$!"


def _prefix_params_scoped(body: str, decls: list[Declaration]) -> str:
    """Insert _xlpm. before every USE of a declared name, inside the declaring
    LET/LAMBDA's span only.

    One left-to-right tokenizer pass over the whole formula, which is what
    makes the three rules Excel actually follows expressible at once:

      1. INERT SPANS are skipped whole. String literals and single-quoted
         sheet names never take a prefix ('Sales x'!A1 stays bare in Excel's
         own storage), and the one scan means a quote of either kind inside
         the other cannot desynchronize the parse.
      2. BUILTIN WINS at a call site: =LET(mod,2,mod+MOD(7,3)) stores
         _xlfn.LET(_xlpm.mod,2,_xlpm.mod+MOD(7,3)) and evaluates 3. The tail
         test skips whitespace before the '(' so the guard matches Excel's own
         tolerance (a spaced call is separately refused by
         _refuse_space_before_paren, because Excel's file parser rejects it).
         A NON-builtin name followed by '(' is a lambda call site and DOES get
         the prefix (_xlpm.f(_xlpm.x)).
      3. SCOPE: a name is prefixed only within the span of the LET/LAMBDA that
         declared it (see _declarations for the COM ground truth).

    Maximal-munch tokenizing also retires the old sorted-by-length regex
    loop, which needed lookarounds to keep 'sum' out of 'summary'."""
    if not decls:
        return body
    inert = literal_spans(body)
    pieces: list[str] = []
    pos = 0
    for m in _IDENT.finditer(body):
        start, end = m.start(), m.end()
        if start < pos:
            continue
        if any(s <= start < e for s, e, _k in inert):
            continue                                  # rule 1
        if start and body[start - 1] in _NOT_BEFORE:
            continue
        token = m.group(0)
        tail = body[end:end + 1]
        if tail == "!":
            continue                                  # a sheet qualifier
        scope_names = {n.lower() for d in decls
                       if d.start <= start < d.end for n in d.names}
        if token.lower() not in scope_names:          # rule 3
            continue
        rest = body[end:]
        stripped = rest.lstrip(" \t\r\n")
        if stripped[:1] == "(" and token.upper() in _BUILTIN_FUNCS:
            continue                                  # rule 2
        pieces.append(body[pos:start])
        pieces.append("_xlpm." + token)
        pos = end
    pieces.append(body[pos:])
    return "".join(pieces)


def _bare_declared_names(body: str, decls: list[Declaration]) -> list[str]:
    """Declared names still sitting bare in a NORMALIZED formula.

    The post-pass assertion behind the C-1 promise: correct prefixing or a
    loud refusal, never a bare LET/LAMBDA name reaching the file. Uses the
    same three rules as the prefix pass, so a name legitimately left bare (a
    builtin call site) is not reported."""
    inert = literal_spans(body)
    bare: list[str] = []
    for m in _IDENT.finditer(body):
        start, end = m.start(), m.end()
        if any(s <= start < e for s, e, _k in inert):
            continue
        if start and body[start - 1] in _NOT_BEFORE:
            continue
        token = m.group(0)
        if body[end:end + 1] == "!":
            continue
        scope_names = {n.lower() for d in decls
                       if d.start <= start < d.end for n in d.names}
        if token.lower() not in scope_names:
            continue
        if body[end:].lstrip(" \t\r\n")[:1] == "(" and \
                token.upper() in _BUILTIN_FUNCS:
            continue
        bare.append(token)
    return bare


def _outside_strings(body: str, fn) -> str:
    """Apply fn to every stretch of the formula that is not an inert span.

    Inert means BOTH kinds: a double-quoted string literal and a
    single-quoted sheet name. Excel prefixes nothing inside either, and
    treating only the double-quoted kind as inert is how a sheet named
    'a"b' used to desynchronize the whole pass."""
    pieces: list[str] = []
    pos = 0
    for start, end, _kind in literal_spans(body):
        if start > pos:
            pieces.append(fn(body[pos:start]))
        pieces.append(body[start:end])   # inert span, verbatim
        pos = end
    if pos < len(body):
        pieces.append(fn(body[pos:]))
    return "".join(pieces)


def normalize_formula(formula: str) -> tuple[str, list[str]]:
    """Prefix modern function names, and LET / LAMBDA parameter names, so an
    openpyxl-written formula lands in Excel exactly as Excel would store it.
    Returns (normalized, list_of_functions_prefixed).

    Prefixes call-site names outside an existing _xlfn context. Quoted string
    literals and quoted sheet names are skipped verbatim.

    Raises FormulaRejected rather than storing something Excel cannot open:
    whitespace between a name and its '(' , an unbalanced LET/LAMBDA argument
    list, or (as a backstop) a declared name that could not be prefixed.

    NORMALIZATION IS NOT VERBATIM ROUND-TRIPPING, by design. Two cosmetic
    differences are expected and are not drift to be fixed (insane round,
    L-4): an optional declaration written =LAMBDA(x, [ y ], x+y) comes back
    as [y], because Excel stores the name without the brackets or the spaces
    and the brackets are re-derived on read; and a caller who supplies an
    ALREADY-prefixed formula, =LET(_xlpm.x,1,_xlpm.x+1), denormalizes to
    =LET(x,1,x+1), because the display form of a stored prefix is the bare
    name. Both re-normalize to the identical stored string, which is the
    property that actually matters."""
    prefixed: list[str] = []

    def sub(m: re.Match) -> str:
        name = m.group(1)
        upper = name.upper()
        if name.startswith("_xlfn"):
            return m.group(0)
        if upper in XLFN_XLWS_FUNCS:
            prefixed.append(upper)
            return f"_xlfn._xlws.{name}("
        if upper in XLFN_FUNCS:
            prefixed.append(upper)
            return f"_xlfn.{name}("
        return m.group(0)

    body = formula[1:] if formula.startswith("=") else formula
    _refuse_space_before_paren(body)
    out = _outside_strings(body, lambda seg: _CALL.sub(sub, seg))
    _decls, optional = _declarations(out)
    if optional:
        # Rewriting [y] to _xlop.y moves every offset after it, so the spans
        # are recomputed against the marked body rather than shifted by hand.
        out = _mark_optional_declarations(out, optional)
        _decls, _ = _declarations(out)
    # An already-prefixed declaration keeps the pass idempotent: the scope
    # still owns the BARE name (so body use sites get _xlpm.) while the
    # prefixed token itself can never match it.
    decls = [Declaration(d.start, d.end,
                         [_strip_param_prefix(n) for n in d.names])
             for d in _decls]
    out = _prefix_params_scoped(out, decls)
    bare = _bare_declared_names(out, decls)
    if bare:
        # Unreachable by construction; kept because the cost of being wrong
        # here is a workbook that will not open at all, not a #NAME?.
        raise FormulaRejected(
            "the LET/LAMBDA parameter name(s) "
            + ", ".join(sorted(set(bare)))
            + " could not be given their _xlpm. storage prefix. Excel REFUSES "
              "TO OPEN a workbook that stores them bare, so this write is "
              "refused rather than producing an unopenable file.")
    return ("=" + out if formula.startswith("=") else out), prefixed


def _strip_param_prefix(name: str) -> str:
    for p in ("_xlpm.", _OPT_PREFIX):
        if name.startswith(p):
            return name[len(p):]
    return name


#: An optional-parameter declaration as Excel stores it, for the read side.
_XLOP_RE = re.compile(re.escape(_OPT_PREFIX) + r"([A-Za-z0-9_.?\\]+)")


def denormalize_formula(formula: str) -> str:
    """Strip the storage prefixes for a human-facing display of a formula read
    back from a workbook.

    All three of Excel's prefixes are handled: _xlfn._xlws. and _xlfn. for
    future functions, _xlpm. for LET/LAMBDA parameter uses, and _xlop. for a
    LAMBDA OPTIONAL parameter declaration. The last one is restored to its
    bracketed source form (_xlop.y -> [y]) rather than merely stripped: that
    is what Excel's own formula bar shows, it is the only form that says the
    parameter is optional, and it round-trips back through normalize_formula
    to the same stored string. Before this, every read of a real Excel
    workbook using optional lambda parameters displayed a raw _xlop.y.

    LITERAL-AWARE, like normalize_formula's own passes. Three blind
    str.replace calls over the WHOLE formula edited the user's data: Excel
    stores =CONCATENATE("_xlpm.","total") verbatim (COM ground truth, fix
    wave 2, 2026-09-05 -- the cell displays "_xlpm.total"), and reading it
    back returned =CONCATENATE("","total"). A read followed by a write then
    made that corruption permanent. Prefixes inside a string literal or a
    quoted sheet name are the user's TEXT, never storage syntax, so this
    strips only outside them."""
    return _outside_strings(formula, _denormalize_span)


def _denormalize_span(segment: str) -> str:
    out = segment.replace("_xlfn._xlws.", "").replace("_xlfn.", "") \
        .replace("_xlpm.", "")
    return _XLOP_RE.sub(lambda m: "[" + m.group(1) + "]", out)


# ------------------------------------------------ what IS a formula cell

def is_formula_cell(cell) -> bool:
    """True only for a cell that actually holds a formula.

    The whole server used to ask ``value.startswith("=")``, which is a lie in
    one important direction: import_data's injection lint deliberately stores
    text like ``=cmd|'/c calc'!A1`` as TEXT (data_type 's') so a spreadsheet
    never executes it, and every read surface then labelled that text as a
    live formula with a cached result. Excel's own answer is the cell type, so
    that is what this asks: data_type 'f' (or an ArrayFormula /
    DataTableFormula object, which openpyxl types 'f' as well).

    Only meaningful on a data_only=False load; a data_only=True load types
    every cell by its CACHED value, so a formula cell reads as 'n'/'s' there.
    """
    if cell is None:
        return False
    if getattr(cell, "data_type", None) == "f":
        return True
    text = getattr(getattr(cell, "value", None), "text", None)
    return isinstance(text, str) and text.startswith("=")


def formula_text_of(cell) -> str | None:
    """The formula string of a real formula cell, else None. Array and data-
    table formulas carry their text on ``.text``."""
    if not is_formula_cell(cell):
        return None
    v = getattr(cell, "value", None)
    if isinstance(v, str) and v.startswith("="):
        return v
    text = getattr(v, "text", None)
    return text if isinstance(text, str) else None


def looks_like_formula_text(cell) -> bool:
    """True for a TEXT cell whose string starts with '=' (neutralized
    injection text). Such a cell must never be re-armed into a live formula
    by a copy, move, sort, or reference rewrite."""
    if cell is None or is_formula_cell(cell):
        return False
    v = getattr(cell, "value", None)
    return isinstance(v, str) and v.startswith("=")


# ---------------------------------------------------- cached-value labeling

LABEL_CACHED = "cached"      # formula present, a prior real calc left a value
LABEL_COMPUTED = "computed"  # this session recalculated it
LABEL_FORMULA = "formula"    # returning the formula string, not a value
LABEL_ABSENT = "absent"      # formula present, NO cached value (openpyxl write)
LABEL_VALUE = "value"        # a plain literal, no formula


def label_cell(formula: str | None, cached, *, computed: bool = False) -> str:
    """Classify what a returned cell value actually is, so a read tool never
    passes off an empty formula cell as a blank cell."""
    if formula is None:
        return LABEL_VALUE
    if computed:
        return LABEL_COMPUTED
    if cached is None:
        return LABEL_ABSENT
    return LABEL_CACHED


# ------------------------------------------------------ fullCalcOnLoad flag

def part_encoding(data: bytes) -> str:
    """The codec an OOXML part is stored in, from its byte-order mark.

    Every part this server has ever seen is UTF-8, but a UTF-16 part is legal
    XML and openpyxl loads one without complaint (verified: a hand-built
    UTF-16 worksheet part round-trips through load_workbook with its formula
    intact). The raw-zip surgery below used to assume UTF-8 and raised
    UnicodeDecodeError on the first byte of such a part, so a legal workbook
    crashed the save rather than merely evading a scan (edge audit 2026-09-04,
    S3)."""
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return "utf-16"
    if data[:3] == b"\xef\xbb\xbf":
        return "utf-8-sig"
    return "utf-8"


def inject_full_calc_on_load(path: str) -> bool:
    """Set <calcPr fullCalcOnLoad="1"/> in xl/workbook.xml so the next Excel /
    LibreOffice open recalculates rather than showing an empty/stale cache.
    Raw-zip surgery (no openpyxl model load); returns True if it changed the
    flag. The no-Excel answer to the stale-cache problem (DESIGN 4.2)."""
    import shutil
    import tempfile
    import zipfile
    from pathlib import Path

    src = Path(path)
    with zipfile.ZipFile(src) as zf:
        names = zf.namelist()
        raw = zf.read("xl/workbook.xml")
    enc = part_encoding(raw)
    wb = raw.decode(enc)

    new_wb, changed = _set_full_calc(wb)
    if not changed:
        return False

    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=str(src.parent))
    import os
    os.close(fd)
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = new_wb.encode(enc)
            zout.writestr(item, data)
    shutil.move(tmp, src)
    return True


# openpyxl writes "<f>B1+1</f><v></v>" for a formula it has no cached value
# for: an EMPTY value element rather than no element at all. Excel reads that
# empty number as an error, which a normal recalculation overwrites, so it
# went unnoticed -- except under ITERATIVE calculation, where the iteration
# seeds from the cell's current value: the numbers-safety gate found a
# circular formula stuck at #VALUE! forever instead of converging (Excel
# converges the identical formula it authored itself). Excel omits the
# element; so do we.
_EMPTY_CACHED_VALUE = re.compile(r"(</f>)\s*<v\s*(?:/>|>\s*</v>)")


def strip_empty_cached_values(path: str) -> int:
    """Remove openpyxl's empty <v></v> from formula cells in every worksheet
    part. Raw-zip surgery, no model load; returns the number of parts changed.

    Safe by construction: openpyxl never carries a cached value through a
    load, so an empty <v> in a package it just wrote is always this artifact
    and never a formula whose real cached result was the empty string."""
    import os
    import shutil
    import tempfile
    import zipfile
    from pathlib import Path

    src = Path(path)
    with zipfile.ZipFile(src) as zf:
        targets = {}
        for name in zf.namelist():
            if not name.startswith("xl/worksheets/") or \
                    not name.endswith(".xml"):
                continue
            raw = zf.read(name)
            enc = part_encoding(raw)
            try:
                data = raw.decode(enc)
            except UnicodeDecodeError:
                continue  # an encoding nobody declared: leave the part alone
            new, n = _EMPTY_CACHED_VALUE.subn(r"\1", data)
            if n:
                targets[name] = (new, enc)
    if not targets:
        return 0

    fd, tmp = tempfile.mkstemp(suffix=src.suffix, dir=str(src.parent))
    os.close(fd)
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in targets:
                text, enc = targets[item.filename]
                data = text.encode(enc)
            zout.writestr(item, data)
    shutil.move(tmp, src)
    return len(targets)


# ------------------------------------ the always-calculate cache (ca="1")

#: A formula cell as Excel stores it, with its optional attributes and its
#: optional cached value. Used only by restore_always_calc_cache below.
_CELL_RE = re.compile(
    r'<c r="([A-Z]+[0-9]+)"(?P<cattrs>[^>]*)>'
    r'(?P<body>.*?)</c>', re.S)
_F_RE = re.compile(r"<f(?P<fattrs>[^>]*)>(?P<text>.*?)</f>", re.S)
_V_RE = re.compile(r"<v[^>]*>.*?</v>|<v[^>]*/>", re.S)


def _sheet_parts(zf) -> dict[str, str]:
    """sheet title -> worksheet part name, resolved through workbook.xml and
    its rels (openpyxl renumbers part names, so position is not identity)."""
    try:
        wb_xml = zf.read("xl/workbook.xml").decode(
            part_encoding(zf.read("xl/workbook.xml")), "replace")
        rels_xml = zf.read("xl/_rels/workbook.xml.rels").decode(
            part_encoding(zf.read("xl/_rels/workbook.xml.rels")), "replace")
    except Exception:  # noqa: BLE001
        return {}
    targets = {}
    for m in re.finditer(r'<Relationship\b[^>]*?Id="([^"]+)"[^>]*?'
                         r'Target="([^"]+)"', rels_xml):
        targets[m.group(1)] = m.group(2)
    for m in re.finditer(r'<Relationship\b[^>]*?Target="([^"]+)"[^>]*?'
                         r'Id="([^"]+)"', rels_xml):
        targets.setdefault(m.group(2), m.group(1))
    out: dict[str, str] = {}
    for m in re.finditer(r'<sheet\b[^>]*>', wb_xml):
        tag = m.group(0)
        name = re.search(r'name="([^"]*)"', tag)
        rid = re.search(r'r:id="([^"]+)"', tag) or \
            re.search(r'\bid="(rId[^"]+)"', tag)
        if not name or not rid:
            continue
        target = targets.get(rid.group(1))
        if not target:
            continue
        out[name.group(1)] = _resolve_part(target)
    return out


def _resolve_part(target: str) -> str:
    """A workbook-rels Target as a package part name. Targets come in three
    spellings depending on who wrote the file: absolute ("/xl/worksheets/
    sheet1.xml"), relative to xl/ ("worksheets/sheet1.xml"), and occasionally
    already prefixed."""
    part = target.replace("\\", "/").replace("/./", "/")
    if part.startswith("/"):
        return part.lstrip("/")
    if part.startswith("xl/"):
        return part
    return "xl/" + part


def _always_calc_cells(zf, part: str) -> dict[str, tuple[str, str]]:
    """addr -> (formula_text, cached_value_element) for the ca="1" formula
    cells of one worksheet part."""
    try:
        raw = zf.read(part)
    except KeyError:
        return {}
    try:
        xml = raw.decode(part_encoding(raw))
    except UnicodeDecodeError:
        return {}
    out: dict[str, tuple[str, str]] = {}
    for m in _CELL_RE.finditer(xml):
        body = m.group("body")
        fm = _F_RE.search(body)
        if fm is None or 'ca="1"' not in fm.group("fattrs"):
            continue
        vm = _V_RE.search(body[fm.end():])
        out[m.group(1)] = (fm.group("text"), vm.group(0) if vm else "")
    return out


def restore_always_calc_cache(original: str, produced: str) -> int:
    """Put back the ca="1" flag and the cached value on formula cells the
    edit did not touch. Returns the number of cells restored.

    WHY THIS EXISTS. openpyxl does not model the always-calculate flag or any
    cached value, so a file-tier save of an Excel-authored workbook rewrites

        <c r="B1"><f ca="1">B1+A1</f><v>2.5</v></c>   as   <f>B1+A1</f>

    Under normal calculation that is invisible: Excel recalculates on open
    (the fullCalcOnLoad flag guarantees it) and fills the value back in.
    Under ITERATIVE calculation it is not, because the iteration SEEDS from
    the cell's current value: a circular formula with no seed lands on
    #VALUE! and stays there, while Excel converges the identical formula it
    authored itself. That is the exact failure the numbers-safety gate closed
    once already via strip_empty_cached_values, re-created by the adjacent
    case where a REAL cached value is dropped (insane round, M-4).

    Conservative by construction: a cell is restored only when the produced
    package still holds the SAME formula text at the SAME address on the same
    sheet, so an edited formula never gets a stale value pinned to it, and a
    cell the edit removed is simply not restored."""
    import os
    import shutil
    import tempfile
    import zipfile
    from pathlib import Path

    src = Path(produced)
    patches: dict[str, str] = {}
    try:
        with zipfile.ZipFile(original) as zin, zipfile.ZipFile(src) as zout:
            src_parts = _sheet_parts(zin)
            out_parts = _sheet_parts(zout)
            if not src_parts or not out_parts:
                return 0
            for title, part in src_parts.items():
                target = out_parts.get(title)
                if target is None:
                    continue
                wanted = _always_calc_cells(zin, part)
                if not wanted:
                    continue
                raw = zout.read(target)
                enc = part_encoding(raw)
                try:
                    xml = raw.decode(enc)
                except UnicodeDecodeError:
                    continue
                new_xml, n = _reapply_always_calc(xml, wanted)
                if n:
                    patches[target] = new_xml
    except Exception:  # noqa: BLE001 - never fail a save over a cache repair
        return 0
    if not patches:
        return 0

    restored = 0
    fd, tmp = tempfile.mkstemp(suffix=src.suffix, dir=str(src.parent))
    os.close(fd)
    try:
        with zipfile.ZipFile(src) as zin, \
                zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in patches:
                    raw = data
                    data = patches[item.filename].encode(part_encoding(raw))
                    restored += 1
                zw.writestr(item, data)
        shutil.move(tmp, src)
    except Exception:  # noqa: BLE001
        try:
            os.remove(tmp)
        except OSError:
            pass
        return 0
    return restored


def _reapply_always_calc(xml: str, wanted: dict[str, tuple[str, str]]
                         ) -> tuple[str, int]:
    n = 0

    def fix(m: re.Match) -> str:
        addr = m.group(1)
        entry = wanted.get(addr)
        if entry is None:
            return m.group(0)
        text, cached = entry
        body = m.group("body")
        fm = _F_RE.search(body)
        if fm is None or fm.group("text") != text:
            return m.group(0)          # the edit changed this formula
        attrs = fm.group("fattrs")
        if 'ca="1"' not in attrs:
            attrs = attrs + ' ca="1"'
        rebuilt = f"<f{attrs}>{text}</f>"
        tail = body[fm.end():]
        if cached and not _V_RE.search(tail):
            tail = cached + tail
        nonlocal n
        n += 1
        return (f'<c r="{addr}"{m.group("cattrs")}>'
                f"{body[:fm.start()]}{rebuilt}{tail}</c>")

    return _CELL_RE.sub(fix, xml), n


def _set_full_calc(workbook_xml: str) -> tuple[str, bool]:
    if 'fullCalcOnLoad="1"' in workbook_xml:
        return workbook_xml, False
    if "<calcPr" in workbook_xml:
        new = re.sub(r"<calcPr\b([^/>]*)/>",
                     lambda m: f"<calcPr{m.group(1)} fullCalcOnLoad=\"1\"/>",
                     workbook_xml, count=1)
        if new != workbook_xml:
            return new, True
        new = re.sub(r"<calcPr\b([^>]*)>",
                     lambda m: f"<calcPr{m.group(1)} fullCalcOnLoad=\"1\">",
                     workbook_xml, count=1)
        return new, new != workbook_xml
    # No calcPr element: insert one before </workbook>.
    if "</workbook>" in workbook_xml:
        return workbook_xml.replace(
            "</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>'), True
    return workbook_xml, False


# ------------------------------------------------------- COM recalc (fidelity)

@dataclass
class RecalcResult:
    engine: str                       # "com" | "formulas" | "flag"
    ok: bool
    values: dict = field(default_factory=dict)   # addr -> value (COM readback)
    unsupported: list[str] = field(default_factory=list)  # formulas coverage
    note: str = ""


def recalc_via_com(path: str, manager, *, calculate: bool = True,
                   full: bool = True) -> RecalcResult:
    """The fidelity recalc (com_ground_truth exp 5): open in a worker, force a
    calculation even under manual mode, save so the cache is populated, close.
    `manager` is an ExcelInstanceManager. Returns the readback of any formula
    cells it can see on the active sheet is left to the caller; this just
    guarantees the file's cache is populated on disk."""
    import os
    w = manager.acquire()
    wb = None
    try:
        app = w.app
        prior_calc = None
        try:
            prior_calc = app.Calculation
        except Exception:
            prior_calc = None
        wb = app.Workbooks.Open(os.path.abspath(path))
        if calculate:
            # Explicit calc fires even in manual mode (the exp-5 caveat cover).
            if full:
                app.CalculateFull()
            else:
                app.Calculate()
        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None
        if prior_calc is not None:
            try:
                app.Calculation = prior_calc
            except Exception:
                pass
        return RecalcResult(engine="com", ok=True,
                            note="cache populated by Excel")
    except Exception as exc:  # noqa: BLE001
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        return RecalcResult(engine="com", ok=False, note=f"{type(exc).__name__}: {exc}")
    finally:
        manager.release_to_pool(w)


# ------------------------------------------------- formulas-lib fallback

def recalc_via_formulas(path: str, out_path: str) -> RecalcResult:
    """Pure-Python best-effort recalc via the `formulas` library. Computes the
    dependency graph without Excel and writes cached values back. Reports any
    functions outside its coverage. Import is lazy so `formulas` stays an
    optional extra; a clean ImportError becomes an honest not-available note."""
    try:
        import formulas  # type: ignore
    except Exception as exc:  # noqa: BLE001
        return RecalcResult(engine="formulas", ok=False,
                            note=f"formulas unavailable: {exc}")
    try:
        xl_model = formulas.ExcelModel().loads(path).finish()
        solution = xl_model.calculate()
        # Write results back next to out_path (the formulas lib names its own
        # output files inside dirpath; the caller reads them from there).
        xl_model.write(dirpath=_dir_of(out_path))
        return RecalcResult(engine="formulas", ok=True,
                            values={k: _clean(v) for k, v in
                                    list(solution.items())[:50]},
                            note="best-effort pure-Python recalc")
    except Exception as exc:  # noqa: BLE001
        return RecalcResult(engine="formulas", ok=False,
                            note=f"{type(exc).__name__}: {exc}")


def _dir_of(p: str) -> str:
    import os
    return os.path.dirname(os.path.abspath(p)) or "."


def _clean(v):
    try:
        return v.value[0, 0] if hasattr(v, "value") else v
    except Exception:
        return str(v)


__all__ = [
    "normalize_formula", "denormalize_formula", "label_cell",
    "is_formula_cell", "formula_text_of", "looks_like_formula_text",
    "inject_full_calc_on_load", "strip_empty_cached_values",
    "recalc_via_com", "recalc_via_formulas", "literal_spans",
    "RecalcResult", "XLFN_FUNCS", "XLFN_XLWS_FUNCS",
    "LABEL_CACHED", "LABEL_COMPUTED", "LABEL_FORMULA", "LABEL_ABSENT",
    "LABEL_VALUE",
]

"""core/limits.py: the STATIC REFUSE-CLASS checks (the cheap layer of the
"Excel will refuse this" gate).

THE ARCHITECTURAL GAP THIS CLOSES. verify-after-write (core.verify) is an
honest check on three things: the produced package re-opens as valid OOXML,
no part was lost, and the cells read back as intended. It has no notion of
"Excel will refuse this," and it never did: openpyxl will happily load a
package that Excel throws out. The insane round (2026-09-05) found ten single
tool calls that returned ``ok: true, saved: true, verified: true`` and
produced a file answering "Open method of Workbooks class failed."

The gate is closed in three layers, and this module is layer (a):

  (a) STATIC REFUSAL CHECKS, here. Every refuse-class the round proved gets a
      cheap, Excel-ground-truthed check BEFORE the write, so the common routes
      refuse loudly with an actionable message and never cost a COM call.
  (b) HONEST REPORTING. ops/validation.py no longer hard-codes
      ``structure.opens_clean: true``; the field says "not checked" unless a
      real Excel check ran.
  (c) THE AUTHORITATIVE CHECK, OPT-IN. ``com_validate_opens_clean`` gets every
      one of these right and is wired as a post-write verification
      (``verify_com``, or KS4XL_VERIFY_COM=1 to make it the default).

Layer (a) can never be complete on its own -- that is exactly why layer (c)
exists -- so nothing here claims a file is safe. Each threshold below is a
REFUSAL BOUNDARY MEASURED AGAINST THIS MACHINE'S EXCEL (fix wave 2,
2026-09-05): a package was built at the boundary, handed to Excel through
``Workbooks.Open``, and the exact value where the open flips from success to
"Open method of Workbooks class failed" is recorded next to the constant.

No new refusal vocabulary: these raise ExcelWouldRefuse (a plain XlMcpError,
so the envelope's closed BAD_PARAMS code covers it) or FormulaRejected, both
of which already exist.
"""

from __future__ import annotations

import re

from .errors import ExcelWouldRefuse

# ---------------------------------------------------------------- thresholds

#: Comment text. MEASURED: 32,766 and 32,767 characters open; 32,768 does not.
#: (The insane round's 60,000-character comment is the same refusal.)
MAX_COMMENT_CHARS = 32_767

#: A data-validation formula1/formula2 string. MEASURED by bisection: 4,097
#: characters open, 4,098 does not. Excel's DOCUMENTED authoring limit is 255
#: and the tool has always warned above it; this is the separate, harder line
#: where the FILE stops opening at all, which is what the round's
#: 10,000-item inline list crossed.
MAX_DV_FORMULA_CHARS = 4_097
DV_FORMULA_SOFT_CHARS = 255

#: A header or footer, as ONE assembled string including its &L / &C / &R
#: section codes. MEASURED: a single left section of 253 characters opens
#: ("&L" + 253 = 255) and 254 does not; three sections of 80 open (246 total)
#: and three of 84 do not (258). The limit is on the whole string, not the
#: section, which is why a "255 per section" check would still ship a file
#: Excel refuses.
MAX_HEADER_FOOTER_CHARS = 255

#: Picture display extents. MEASURED: width/height 0 and 10 open; -5 does not
#: (a negative EMU extent is not representable in the drawing XML).
MIN_IMAGE_EXTENT = 0

#: Cell text. Excel's documented ceiling, and check_cell_text below is what
#: consults it. It used to be a bare constant that nothing on the set_cells
#: path read, so a 40,000-character cell paid a full write-and-verify cycle
#: and then refused through the WRONG LAYER: "1 written cell did not read
#: back as intended", a generic verify failure, instead of "Excel stores at
#: most 32,767 characters" (live COM stress, L-1).
MAX_CELL_CHARS = 32_767


# ------------------------------------------------------------ illegal text

#: Characters openpyxl itself refuses to serialize (and Excel cannot store):
#: the C0 controls except tab, newline and carriage return. Without this check
#: openpyxl's IllegalCharacterError escaped the Section 7 envelope entirely
#: and echoed the raw control bytes back to the caller (insane round, M-2).
_ILLEGAL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def check_text_storable(value, *, what: str = "this value"):
    """Refuse text Excel cannot store, BEFORE openpyxl raises its own
    un-enveloped error. Two populations:

      1. C0 control characters. openpyxl raises IllegalCharacterError, which
         escaped the envelope as a bare tool error with a Rich traceback on
         stderr and the raw control bytes echoed back.
      2. LONE SURROGATES. A request carrying "\\ud800" is valid JSON and the
         server accepted it, but the value cannot be encoded to UTF-8 on the
         way out, so the response was never written and the caller blocked
         until its own timeout with the server still healthy (insane round,
         M-3). A PAIRED surrogate (a real emoji) is fine and is left alone.

    Returns the value unchanged when it is storable."""
    if not isinstance(value, str):
        return value
    m = _ILLEGAL_CHARS.search(value)
    if m is not None:
        point = value.index(m.group(0))
        raise ExcelWouldRefuse(
            f"{what} contains the control character U+{ord(m.group(0)):04X} "
            f"at position {point}, which Excel cannot store in a worksheet. "
            "Strip the control characters (tab, newline and carriage return "
            "are fine) and retry.")
    for i, ch in enumerate(value):
        if 0xD800 <= ord(ch) <= 0xDFFF:
            raise ExcelWouldRefuse(
                f"{what} contains an unpaired surrogate code point "
                f"U+{ord(ch):04X} at position {i}. It is not encodable as "
                "UTF-8, so neither the workbook nor this server's own reply "
                "can carry it. Send the character as a proper pair, or drop "
                "it.")
    return value


# --------------------------------------------------------------- the checks


def check_cell_text(value, *, what: str = "this cell"):
    """The storable-text check plus Excel's 32,767-character cell ceiling.
    Non-strings pass straight through (a number has no length to check).

    openpyxl writes an oversize string happily and the produced package is
    valid OOXML, so verify-after-write is what used to catch it, one layer too
    late and with the wrong words. Refusing here costs nothing and says what
    the limit is."""
    check_text_storable(value, what=what)
    if isinstance(value, str) and len(value) > MAX_CELL_CHARS:
        raise ExcelWouldRefuse(
            f"{what} is {len(value):,} characters; Excel stores at most "
            f"{MAX_CELL_CHARS:,} in a cell and truncates or refuses beyond "
            "that. Shorten the text, or split it across cells.")
    return value


def check_comment_text(text: str) -> str:
    check_text_storable(text, what="the comment text")
    if len(text) > MAX_COMMENT_CHARS:
        raise ExcelWouldRefuse(
            f"the comment is {len(text):,} characters; Excel stores at most "
            f"{MAX_COMMENT_CHARS:,} and REFUSES TO OPEN a workbook that "
            "carries a longer one (measured against Excel: 32,767 opens, "
            "32,768 does not). Shorten the note, or put the long text in a "
            "cell.")
    return text


def check_dv_formula(formula, *, field: str = "formula1") -> str | None:
    if formula is None:
        return None
    text = str(formula)
    check_text_storable(text, what=f"the data-validation {field}")
    if len(text) > MAX_DV_FORMULA_CHARS:
        raise ExcelWouldRefuse(
            f"the data-validation {field} is {len(text):,} characters; Excel "
            f"REFUSES TO OPEN a workbook whose {field} is longer than "
            f"{MAX_DV_FORMULA_CHARS:,} (measured). An inline list this long "
            "belongs in a range: put the items on a sheet and point the "
            "validation at that range instead.")
    return text


def check_header_footer(sections: dict, *, which: str = "header") -> None:
    """Refuse a header/footer whose ASSEMBLED string crosses Excel's limit.
    ``sections`` maps 'left'/'center'/'right' to the ENCODED text (the & codes
    already expanded), which is what actually lands in the file."""
    codes = {"left": "&L", "center": "&C", "right": "&R"}
    total = 0
    for name, text in sections.items():
        if not text:
            continue
        check_text_storable(text, what=f"the {which} {name} section")
        total += len(codes.get(name, "")) + len(str(text))
    if total > MAX_HEADER_FOOTER_CHARS:
        raise ExcelWouldRefuse(
            f"the assembled {which} is {total:,} characters including its "
            f"&L/&C/&R section codes; Excel REFUSES TO OPEN a workbook whose "
            f"{which} string is longer than {MAX_HEADER_FOOTER_CHARS} "
            "(measured: 255 opens, 256 does not). Note the limit is on the "
            "WHOLE string, so three sections share it. Shorten the text or "
            "drop a section.")


#: Characters legal in the AUTHORITY (host) component of a URI: RFC 3986
#: reg-name plus the userinfo and port punctuation and the IPv6 brackets.
_AUTHORITY_OK = re.compile(r"^[A-Za-z0-9\-._~%!$&'()*+,;=:@\[\]]*$")

#: Schemes this server refuses to plant in a workbook. Excel OPENS all of
#: them (measured), so this is not a corruption guard: it is the same posture
#: the server already takes on formula injection, where =cmd|'/c calc'!A1 is
#: deliberately stored as TEXT. An agent-planted hyperlink target is that
#: threat model with a click instead of a recalculation (insane round, M-7).
UNSAFE_SCHEMES = ("javascript:", "vbscript:", "data:", "ms-msdt:",
                  "ms-search:", "ms-officecmd:", "search-ms:")

#: Schemes that reach the local machine or the network. Legitimate in a real
#: workbook (a link to a file share is ordinary), so these WARN rather than
#: refuse; the caller sees what was written.
SENSITIVE_PREFIXES = ("file:", "\\\\")


def check_hyperlink_target(target: str) -> list[str]:
    """Validate a hyperlink target and return advisory warnings.

    Refuses two classes:

      1. URI-INVALID targets, which make the workbook unopenable. MEASURED:
         a quote, a tab, or a space in the AUTHORITY (``http://x"y``,
         ``http://x y``) gives "Open method of Workbooks class failed", while
         the identical characters in the PATH (``http://example.com/x"y``)
         open fine. The insane round's ``http://x"/><evil a="`` is the first
         case: not an XML injection (openpyxl escaped it correctly), just a
         target Excel cannot parse.
      2. UNSAFE SCHEMES, per the server's standing injection posture.
    """
    if not isinstance(target, str):
        raise ExcelWouldRefuse("the hyperlink target must be a string")
    check_text_storable(target, what="the hyperlink target")
    low = target.strip().lower()
    for scheme in UNSAFE_SCHEMES:
        if low.startswith(scheme):
            raise ExcelWouldRefuse(
                f"refusing to write a {scheme} hyperlink target. Excel opens "
                "the workbook and the link fires on a click, which is the "
                "same threat model this server already blocks on the formula "
                "side (=cmd|'/c calc'!A1 is stored as TEXT, never armed). "
                "Use an http(s) target, or an in-workbook 'Sheet!A1' "
                "reference.")
    m = re.match(r"^([A-Za-z][A-Za-z0-9+.\-]*):\/\/([^/?#]*)", target)
    if m is not None:
        authority = m.group(2)
        if not _AUTHORITY_OK.match(authority):
            bad = sorted({c for c in authority if not _AUTHORITY_OK.match(c)})
            shown = ", ".join(
                repr(c) if c.isprintable() else f"U+{ord(c):04X}" for c in bad)
            raise ExcelWouldRefuse(
                f"the hyperlink target {target!r} has {shown} in its host "
                f"({authority!r}), which is not a valid URI authority. Excel "
                "REFUSES TO OPEN a workbook carrying such a target (measured; "
                "the same characters later in the PATH are fine). "
                "Percent-encode them, or fix the host.")
    warnings: list[str] = []
    for prefix in SENSITIVE_PREFIXES:
        if low.startswith(prefix.lower()):
            warnings.append(
                f"the hyperlink target {target!r} points at the local machine "
                "or a network share; a UNC target leaks credentials to "
                "whoever owns that share when a reader clicks it")
            break
    return warnings


def check_image_extents(width, height) -> None:
    for label, value in (("width", width), ("height", height)):
        if value is None:
            continue
        try:
            n = int(value)
        except (TypeError, ValueError):
            raise ExcelWouldRefuse(
                f"image {label} must be a whole number of pixels, got "
                f"{value!r}") from None
        if n < MIN_IMAGE_EXTENT:
            raise ExcelWouldRefuse(
                f"image {label} is {n}; a negative display extent cannot be "
                "represented in the drawing XML and Excel REFUSES TO OPEN the "
                "workbook (measured: 0 and 10 open, -5 does not). Pass a "
                "non-negative size, or omit it to keep the image's natural "
                "size.")


def check_sheet_title(title: str) -> str:
    """Excel's sheet-name grammar, for completeness of the refuse-class
    table: 1-31 characters, none of : \\ / ? * [ ], and not bracketed by
    apostrophes."""
    if not isinstance(title, str) or not title:
        raise ExcelWouldRefuse("a sheet name must be a non-empty string")
    check_text_storable(title, what="the sheet name")
    if len(title) > 31:
        raise ExcelWouldRefuse(
            f"the sheet name is {len(title)} characters; Excel's limit is 31")
    bad = sorted(set(title) & set(":\\/?*[]"))
    if bad:
        raise ExcelWouldRefuse(
            f"the sheet name contains {', '.join(repr(c) for c in bad)}; "
            "Excel bans : \\ / ? * [ ] in sheet names")
    if title.startswith("'") or title.endswith("'"):
        raise ExcelWouldRefuse(
            "a sheet name may not start or end with an apostrophe")
    return title


def describe() -> dict:
    """The refuse-class table, for the docs and the safety gate."""
    return {
        "comment_text_chars": MAX_COMMENT_CHARS,
        "data_validation_formula_chars": MAX_DV_FORMULA_CHARS,
        "header_footer_string_chars": MAX_HEADER_FOOTER_CHARS,
        "image_extent_minimum": MIN_IMAGE_EXTENT,
        "cell_text_chars": MAX_CELL_CHARS,
        "hyperlink": "valid URI authority; unsafe schemes refused",
        "formula": "no whitespace before '('; no bare LET/LAMBDA names",
        "measured_against": "Excel, this machine, fix wave 2 2026-09-05",
    }


__all__ = [
    "MAX_COMMENT_CHARS", "MAX_DV_FORMULA_CHARS", "DV_FORMULA_SOFT_CHARS",
    "MAX_HEADER_FOOTER_CHARS", "MIN_IMAGE_EXTENT", "MAX_CELL_CHARS",
    "UNSAFE_SCHEMES", "SENSITIVE_PREFIXES",
    "check_text_storable", "check_cell_text", "check_comment_text",
    "check_dv_formula",
    "check_header_footer", "check_hyperlink_target", "check_image_extents",
    "check_sheet_title", "describe",
]

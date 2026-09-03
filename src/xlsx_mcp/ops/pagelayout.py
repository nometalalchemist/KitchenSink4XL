"""ops/pagelayout.py: page setup and headers/footers (DESIGN Section 11, io).

set_page_layout covers the print-shaped sheet settings: orientation, paper
size, margins, scaling and fit-to, print area, print titles (the whole-row
and whole-column spans core.locate/refs already understand), and the
gridlines/headings print flags. One call, every parameter optional, at
least one required; unset parameters leave the current values alone.

set_header_footer writes the three-section (left/center/right) headers and
footers. Excel stores them in a & code mini-language, so by default the
tool treats section text as LITERAL: every & is escaped to && and the
placeholders {page} {pages} {date} {time} {file} {sheet} {path} expand to
their codes (&P &N &D &T &F &A &Z). raw=true passes the text through as
codes for callers who speak the mini-language. The even/first variants set
the matching differentOddEven/differentFirst flags automatically.

Both tools mutate through WorkbookPackage (hazard gate, backup, atomic
verified save); tests round-trip every field through a reload.
"""

from __future__ import annotations

import re

from ..core.errors import TargetNotFound, XlMcpError
from ..core.package import WorkbookPackage

ORIENTATIONS = ("portrait", "landscape")

#: common paper-size names -> ECMA-376 paperSize codes (int codes also pass)
PAPER_SIZES: dict[str, int] = {
    "letter": 1, "tabloid": 3, "ledger": 4, "legal": 5, "executive": 7,
    "a3": 8, "a4": 9, "a5": 11, "b4": 12, "b5": 13,
}

MARGIN_KEYS = ("left", "right", "top", "bottom", "header", "footer")

HF_TARGETS = ("odd", "even", "first")
_HF_SECTIONS = ("left", "center", "right")
_PLACEHOLDERS = {
    "{page}": "&P", "{pages}": "&N", "{date}": "&D", "{time}": "&T",
    "{file}": "&F", "{sheet}": "&A", "{path}": "&Z",
}

_ROWS_SPAN = re.compile(r"^\$?\d+:\$?\d+$")
_COLS_SPAN = re.compile(r"^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$")


def _require_sheet(pkg: WorkbookPackage, sheet: str | None):
    wb = pkg.workbook
    if sheet is None:
        return wb.active
    if sheet not in wb.sheetnames:
        raise TargetNotFound(
            f"no sheet named {sheet!r}; sheets: {wb.sheetnames}")
    return wb[sheet]


def set_page_layout(path: str, sheet: str | None = None,
                    orientation: str | None = None,
                    paper_size: str | int | None = None,
                    margins: dict | None = None,
                    scale: int | None = None,
                    fit_to_width: int | None = None,
                    fit_to_height: int | None = None,
                    print_area: str | None = None,
                    print_title_rows: str | None = None,
                    print_title_cols: str | None = None,
                    gridlines: bool | None = None,
                    headings: bool | None = None,
                    allow_loss: bool = False, backup: bool = True) -> dict:
    """Set the print-shaped page settings. Backup + verify."""
    supplied = [v for v in (orientation, paper_size, margins, scale,
                            fit_to_width, fit_to_height, print_area,
                            print_title_rows, print_title_cols, gridlines,
                            headings) if v is not None]
    if not supplied:
        raise XlMcpError(
            "nothing to set: pass at least one page-layout parameter")
    if scale is not None and (fit_to_width is not None
                              or fit_to_height is not None):
        raise XlMcpError(
            "scale and fit-to are mutually exclusive (Excel ignores scale "
            "when fit-to-page is on); pass one or the other")

    pkg = WorkbookPackage.open(path)
    ws = _require_sheet(pkg, sheet)
    changed: dict = {"sheet": ws.title}

    if orientation is not None:
        if orientation not in ORIENTATIONS:
            raise XlMcpError(
                f"orientation must be one of {ORIENTATIONS}, got "
                f"{orientation!r}")
        ws.page_setup.orientation = orientation
        changed["orientation"] = orientation
    if paper_size is not None:
        if isinstance(paper_size, str):
            code = PAPER_SIZES.get(paper_size.strip().lower())
            if code is None:
                raise XlMcpError(
                    f"unknown paper size {paper_size!r}; names: "
                    f"{sorted(PAPER_SIZES)} (or pass the numeric ECMA-376 "
                    "code)")
        else:
            code = int(paper_size)
            if not 1 <= code <= 118:
                raise XlMcpError(
                    f"paper size code {code} is outside the ECMA-376 range "
                    "1-118")
        ws.page_setup.paperSize = code
        changed["paper_size"] = code
    if margins is not None:
        unknown = sorted(set(margins) - set(MARGIN_KEYS))
        if unknown:
            raise XlMcpError(
                f"unknown margin key(s) {unknown}; valid: {MARGIN_KEYS} "
                "(inches)")
        for key, val in margins.items():
            val = float(val)
            if not 0 <= val <= 10:
                raise XlMcpError(
                    f"margin {key}={val} is outside 0-10 inches")
            setattr(ws.page_margins, key, val)
        changed["margins"] = {k: float(v) for k, v in margins.items()}
    if scale is not None:
        if not 10 <= int(scale) <= 400:
            raise XlMcpError(f"scale must be 10-400 percent, got {scale}")
        ws.page_setup.scale = int(scale)
        ws.sheet_properties.pageSetUpPr.fitToPage = False
        changed["scale"] = int(scale)
    if fit_to_width is not None or fit_to_height is not None:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        if fit_to_width is not None:
            ws.page_setup.fitToWidth = int(fit_to_width)
            changed["fit_to_width"] = int(fit_to_width)
        if fit_to_height is not None:
            ws.page_setup.fitToHeight = int(fit_to_height)
            changed["fit_to_height"] = int(fit_to_height)
    if print_area is not None:
        if print_area.strip().lower() == "clear":
            ws.print_area = None
            changed["print_area"] = "cleared"
        else:
            grid = pkg.resolve({"range": print_area},
                               default_sheet=ws.title)
            if grid.sheet != ws.title:
                raise XlMcpError(
                    "print_area must be a range on the target sheet")
            ws.print_area = grid.a1
            changed["print_area"] = grid.a1
    if print_title_rows is not None:
        if print_title_rows.strip().lower() == "clear":
            ws.print_title_rows = None
            changed["print_title_rows"] = "cleared"
        elif not _ROWS_SPAN.match(print_title_rows.strip()):
            raise XlMcpError(
                f"print_title_rows must be a whole-row span like '1:2', "
                f"got {print_title_rows!r}")
        else:
            ws.print_title_rows = print_title_rows.strip()
            changed["print_title_rows"] = print_title_rows.strip()
    if print_title_cols is not None:
        if print_title_cols.strip().lower() == "clear":
            ws.print_title_cols = None
            changed["print_title_cols"] = "cleared"
        elif not _COLS_SPAN.match(print_title_cols.strip()):
            raise XlMcpError(
                f"print_title_cols must be a whole-column span like 'A:B', "
                f"got {print_title_cols!r}")
        else:
            ws.print_title_cols = print_title_cols.strip()
            changed["print_title_cols"] = print_title_cols.strip()
    if gridlines is not None:
        ws.print_options.gridLines = bool(gridlines)
        changed["gridlines"] = bool(gridlines)
    if headings is not None:
        ws.print_options.headings = bool(headings)
        changed["headings"] = bool(headings)

    pkg._changed["page_layout"] = changed
    return pkg.save(allow_loss=allow_loss, backup=backup)


def _encode_section(text: str, raw: bool) -> str:
    if raw:
        return text
    # escape literal ampersands first, then expand placeholders to codes
    out = text.replace("&", "&&")
    for ph, code in _PLACEHOLDERS.items():
        out = out.replace(ph, code)
    return out


def set_header_footer(path: str, sheet: str | None = None,
                      header: dict | None = None,
                      footer: dict | None = None,
                      apply_to: str = "odd", raw: bool = False,
                      allow_loss: bool = False, backup: bool = True) -> dict:
    """Write three-section headers/footers with & code escaping.
    Backup + verify."""
    if header is None and footer is None:
        raise XlMcpError(
            "nothing to set: pass header and/or footer as "
            "{left, center, right} (an empty string clears a section)")
    if apply_to not in HF_TARGETS:
        raise XlMcpError(
            f"apply_to must be one of {HF_TARGETS}, got {apply_to!r}")
    for name, block in (("header", header), ("footer", footer)):
        if block is None:
            continue
        unknown = sorted(set(block) - set(_HF_SECTIONS))
        if unknown:
            raise XlMcpError(
                f"unknown {name} section(s) {unknown}; valid: "
                f"{_HF_SECTIONS}")

    pkg = WorkbookPackage.open(path)
    ws = _require_sheet(pkg, sheet)
    targets = {
        "odd": (ws.oddHeader, ws.oddFooter),
        "even": (ws.evenHeader, ws.evenFooter),
        "first": (ws.firstHeader, ws.firstFooter),
    }
    hdr_obj, ftr_obj = targets[apply_to]
    if apply_to == "even":
        ws.HeaderFooter.differentOddEven = True
    elif apply_to == "first":
        ws.HeaderFooter.differentFirst = True

    changed: dict = {"sheet": ws.title, "apply_to": apply_to}
    for name, block, obj in (("header", header, hdr_obj),
                             ("footer", footer, ftr_obj)):
        if block is None:
            continue
        done = {}
        for section in _HF_SECTIONS:
            if section not in block:
                continue
            text = block[section]
            if text is None or text == "":
                getattr(obj, section).text = None
                done[section] = "cleared"
            else:
                encoded = _encode_section(str(text), raw)
                getattr(obj, section).text = encoded
                done[section] = encoded
        changed[name] = done

    pkg._changed["header_footer"] = changed
    return pkg.save(allow_loss=allow_loss, backup=backup)


__all__ = ["set_page_layout", "set_header_footer", "PAPER_SIZES",
           "ORIENTATIONS", "HF_TARGETS"]

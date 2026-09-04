"""Exception types for xlsx-mcp (KitchenSink4XL).

Ported from KitchenSink4Word core/errors.py and extended for the grid
domain (DESIGN Section 7). Every tool call maps these to actionable
messages; envelope.py turns them into the closed-code refusal shape.

The base class is XlMcpError. The Word taxonomy is carried over (renamed
Workbook/sheet where the Word wording was document-specific) plus the
grid-specific additions the DESIGN names: HazardRefused, FormulaRejected,
CalcUnavailable.
"""

from __future__ import annotations


class XlMcpError(Exception):
    """Base class; message text is user-facing."""


# ------------------------------------------------------- file / package

class WorkbookNotFound(XlMcpError):
    pass


class WorkbookLocked(XlMcpError):
    """File is open in Excel (or another process holds a lock).

    The empirically confirmed case (com_ground_truth exp 4): a book held
    open in Excel raises PermissionError on the openpyxl write path and
    Excel drops a ~$name.xlsx owner lockfile. Maps to WORKBOOK_LOCKED; the
    hint names the live/COM route."""


class WorkbookCorrupt(XlMcpError):
    """File is not a valid .xlsx/.xlsm package (BadZipFile and kin)."""


class WorkbookProtected(XlMcpError):
    """File is encrypted / password-protected (read needs decryption)."""


# ------------------------------------------------------------- location

class TargetNotFound(XlMcpError):
    """A cell, range, name, table, sheet, or object the tool was told to
    act on does not exist."""


class AmbiguousTarget(XlMcpError):
    """More than one match for a search or name selector; the caller must
    disambiguate. Carries every match on .matches for the refusal shape
    (DESIGN Section 5.2)."""


class RangeOutOfBounds(XlMcpError):
    """A range is inverted or exceeds the grid limits (1,048,576 rows x
    16,384 columns) or the 32,767-char cell ceiling. Maps to
    RANGE_OUT_OF_BOUNDS; the message names the valid bounds."""


class StaleAnchor(XlMcpError):
    """A get_grid_view anchor no longer resolves; the sheet changed since
    the view was taken. Maps to STALE_ANCHOR (Phase 4 view/batch layer)."""


class UnsupportedStructure(XlMcpError):
    """A grid topology or XML shape we refuse to guess about (conservative
    mode). Maps to UNSUPPORTED_CONTENT."""


# ------------------------------------------------------ safety / verify

class ValidationFailed(XlMcpError):
    """Verify-after-write caught a problem; the original file was NOT
    modified (the mutation wrote to a temp path and was not promoted).
    Maps to VALIDATION_FAILED (DESIGN Section 3.3)."""


class ExcelWouldRefuse(XlMcpError):
    """A STATIC refuse-class check caught a value Excel itself will not
    accept: a comment past 32,767 characters, a data-validation formula past
    the length where the file stops opening, a header/footer string past 255,
    an invalid hyperlink authority, a negative image extent, a control
    character or lone surrogate in cell text.

    These all used to sail through: verify-after-write proves the package is
    valid OOXML and reads back as intended, which a workbook Excel refuses to
    open can be. Every threshold is measured against Excel (core.limits).
    Maps to the existing BAD_PARAMS code -- a refuse-class value is a bad
    parameter, and the closed vocabulary does not grow for this."""


class HazardRefused(XlMcpError):
    """The round-trip hazard scan found unpreservable parts and no safe
    route, and no allow_loss override was passed. Carries the named parts
    and the available routes on .detail. Maps to HAZARD_REFUSED (DESIGN
    Section 3.1). This is the loud refusal that never repeats the
    incumbent's silent loss."""


class FormulaRejected(XlMcpError):
    """The formula-injection lint or the unsafe-function policy blocked a
    write; the message names the function and the override. Maps to
    FORMULA_REJECTED (DESIGN Section 10)."""


class CalcUnavailable(XlMcpError):
    """A fidelity recalc was requested, COM is absent, and the workbook
    uses functions outside the fallback engine's coverage. Maps to
    CALC_UNAVAILABLE (DESIGN Section 4.2)."""


# --------------------------------------------------------- COM / live

class ExcelNotRunning(XlMcpError):
    """No attachable Excel instance (com_ tools need one)."""


class WorkbookNotOpenInExcel(XlMcpError):
    """A live/COM tool targeted a workbook that is not open in the running
    Excel."""


class ExcelBusy(XlMcpError):
    """Excel rejected the call (modal dialog, or a running command)."""


class ExcelBlocked(XlMcpError):
    """Excel is not answering at all (long synchronous operation)."""


class ExcelDisconnected(XlMcpError):
    """Excel or the workbook closed mid-call; the edit may be partially
    applied."""

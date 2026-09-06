"""core/hazard.py: the round-trip hazard scan and per-tool routing decision.

The single most important new module (PLAN reuse ledger 1.2, DESIGN Section
3.1). A cheap zip central-directory read detects the fragile parts openpyxl
would drop on save (slicers, timelines, drawings/shapes, ActiveX, threaded
comments, Power Query DataMashup, connections, dynamic-array metadata, pivot
caches, VBA) WITHOUT an object-model load, then routes each mutation to
openpyxl / raw OOXML surgery / COM / HAZARD_REFUSED.

The knowledge table in HAZARD_SPECS encodes what openpyxl preserves vs drops.
The Phase 1 fidelity harness (tests/fixtures + scripts) VALIDATES every
`survives_openpyxl` claim in this table empirically against real fixtures,
and the re-audit's part-injection round-trips validated the additions
(embeddings, persons, richData, namedSheetViews, queryTables, data model).
Any file whose parts openpyxl would drop MUST be flagged here; a lossy file
this scan calls clean is the exact incumbent failure the brand exists to
prevent, so the table errs toward flagging (a false positive costs a needless
raw/COM route; a false negative silently destroys user content).

Two deliberate NON-flags, verified empirically (re-audit round-trips):
- xl/externalLinks/ is PRESERVED by openpyxl (modeled; re-serialized but
  semantically intact), so it carries no spec. Only externalBook links were
  tested; ddeLink/oleLink variants are untested and would surface in
  verify-after-write's part-loss diff if they ever dropped.
- xl/printerSettings/ IS dropped by openpyxl, but flagging it would refuse
  mutations on nearly every workbook that ever printed, for a loss that is
  cosmetic (page-setup device settings; the pageSetup element itself
  survives). It stays unflagged by policy and is treated as a routine drop.

IN-PART HAZARDS (the former known limit, now closed). Some fragile content
does not live in a part of its own: x14 conditional formatting (every modern
data bar and icon set), sparkline groups, worksheet slicer lists and the rest
of the worksheet-level extLst extensions sit INSIDE xl/worksheets/sheetN.xml,
which survives the save at full size. The fidelity gate measured what that
cost: an Excel-authored data bar plus icon set (tests/fixtures corpus
x14_condformat.xlsx) added no part at all, the namelist scan reported CLEAN,
the mutation applied without a murmur, and the x14 rules were gone from the
saved file. openpyxl's worksheet reader warns "<name> extension is not
supported and will be removed" for every ext it meets and writes none of them
back, so the whole block is DROP-class, not degrade-class.

The scan therefore reads the worksheet parts and looks at the top-level
extLst itself (EXT_DROP_LABELS / the "in_part_extensions" spec below). That
costs a decompress-and-parse of each worksheet, so scan_path is no longer
namelist-only for workbooks with sheets; the parse is a streaming expat walk
with no tree built, but the read is linear in sheet size and that is the price
of not destroying rules in silence. scan_names WITHOUT a part reader cannot
see in-part content and reports only the name-detectable hazards; core/
package.py keeps openpyxl's own load-time warnings as a second, independent
detector and refuses on those too, so a URI this table has never heard of is
still caught.

The author's package policy governs the outcome (ratified): drop-risk refuses
without allow_loss, degrade-risk warns and proceeds. In-part extensions are
drop-risk, so a mutation that would lose them now raises HAZARD_REFUSED naming
the rules at risk and the ways out. Under allow_loss the loss is still
announced item by item. This replaces the previous warn-and-proceed behavior,
which announced the loss but let it happen.

The refusal used to name a COM route (see refusal_routes below). It does not
any more: the com pack holds no tool that writes a cell, a format, or a row,
so for a refused file-tier mutation that route was a dead end the caller had
to walk before finding out. Preservation is the real fix and it is not in
this release; until it lands the refusal states the exact losses and the two
options that exist.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from typing import Callable, Iterable

# ------------------------------------------- legacy OLE containers (.xls)
#
# The geriatric round's H-1: a genuine .xls handed to any tool escaped the
# envelope entirely (openpyxl's InvalidFileException is a bare Exception
# subclass), or -- worse -- half-parsed. Modern Excel's .xls files embed a
# zip FRAGMENT (the theme package), so Python's zipfile "successfully" opens
# a BIFF file, finds [Content_Types].xml with no workbook part, and the
# caller of a perfectly healthy 2003 workbook was told the file was corrupt.
# Four different behaviors for one user situation, two out of envelope, and
# the one leaked message recommended xlrd, a Python library an MCP caller
# cannot invoke, instead of the server's own com_convert_format.
#
# The fix is an 8-byte signature sniff at every file entry point. An OLE
# compound file (D0CF11E0A1B11AE1) is never a readable OOXML package, but it
# IS one of two very different user situations, told apart by the container's
# top-level stream names: a legacy BIFF workbook carries a "Workbook" (or
# BIFF5 "Book") stream, while Excel's real encryption wraps an OOXML package
# in streams named "EncryptionInfo"/"EncryptedPackage". Diagnosing an
# encrypted heirloom as "an old .xls" (or vice versa) sends the owner down
# the wrong remedy, so the sniff reads the first directory sector and names
# the situation it actually found.

_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

OLE_KIND_BIFF = "biff"            # legacy Excel 97-2003 (or older) workbook
OLE_KIND_ENCRYPTED = "encrypted"  # Excel's real password encryption
OLE_KIND_UNKNOWN = "ole"          # an OLE container this sniff cannot place


def _ole_stream_names(fh, header: bytes) -> set[str]:
    """Lower-cased top-level stream names from the FIRST directory sector.
    The root entry and the main streams (Workbook / Book / EncryptionInfo /
    EncryptedPackage / WordDocument) live there in every real file; a parse
    failure returns what was read (the caller stays conservative)."""
    import struct
    names: set[str] = set()
    try:
        sector_shift = struct.unpack_from("<H", header, 30)[0]
        if not 6 <= sector_shift <= 20:
            return names
        sector = 1 << sector_shift
        first_dir = struct.unpack_from("<i", header, 48)[0]
        if first_dir < 0:
            return names
        fh.seek((first_dir + 1) * sector)
        data = fh.read(sector)
    except (OSError, struct.error):
        return names
    for off in range(0, len(data) - 127, 128):
        nlen = int.from_bytes(data[off + 64:off + 66], "little")
        if not 2 <= nlen <= 64:
            continue
        try:
            names.add(
                data[off:off + nlen - 2].decode("utf-16-le").strip().lower())
        except UnicodeDecodeError:
            continue
    return names


def ole_container_kind(path: str) -> str | None:
    """None when the file is not an OLE compound container; otherwise one of
    OLE_KIND_BIFF / OLE_KIND_ENCRYPTED / OLE_KIND_UNKNOWN."""
    try:
        with open(path, "rb") as fh:
            header = fh.read(512)
            if header[:8] != _OLE_MAGIC:
                return None
            if len(header) < 512:
                return OLE_KIND_UNKNOWN
            names = _ole_stream_names(fh, header)
    except OSError:
        return None  # unreadable here: the normal open path diagnoses locks
    if {"workbook", "book"} & names:
        return OLE_KIND_BIFF
    if {"encryptedpackage", "encryptioninfo"} & names:
        return OLE_KIND_ENCRYPTED
    return OLE_KIND_UNKNOWN


def ole_refusal_text(kind: str, path: str) -> str:
    """The one honest message per container kind, shared by every entry
    point so a .xls meets the same words whatever tool it hits first."""
    from pathlib import Path as _P
    name = _P(path).name
    if kind == OLE_KIND_BIFF:
        return (
            f"{name} is a legacy Excel 97-2003 workbook (BIFF format in an "
            "OLE compound container), not an OOXML .xlsx/.xlsm package -- "
            "whatever its extension says. The file-based tools cannot read "
            "BIFF. Remedy: com_convert_format(path=..., output='...xlsx') "
            "has Excel itself write a faithful .xlsx copy; then run this "
            "tool on the converted file.")
    if kind == OLE_KIND_ENCRYPTED:
        return (
            f"{name} is password-protected (Excel's real encryption wraps "
            "the whole package in an OLE container); the file-based tools "
            "cannot decrypt it. The com tools accept password=... "
            "(com_validate_opens_clean verifies it opens; "
            "com_save_with_password with password='' removes the "
            "encryption so the file tier can work on it).")
    return (
        f"{name} is an OLE compound file, not an OOXML package: most likely "
        "a legacy Office file (.xls/.doc/.ppt) or an encrypted workbook "
        "this scan could not classify. If it is a legacy Excel workbook, "
        "com_convert_format(path=..., output='...xlsx') converts it via "
        "Excel; if it is password-protected, the com tools accept "
        "password=... .")


def refuse_ole_container(path: str) -> None:
    """Raise the in-envelope refusal when `path` is an OLE container.
    UNSUPPORTED_CONTENT for BIFF/unknown (UnsupportedStructure),
    WorkbookProtected for the encrypted kind; both carry hint_tools so the
    envelope names the com pack when it is disabled."""
    kind = ole_container_kind(path)
    if kind is None:
        return
    from .errors import UnsupportedStructure, WorkbookProtected
    if kind == OLE_KIND_ENCRYPTED:
        exc: Exception = WorkbookProtected(ole_refusal_text(kind, path))
        exc.hint_tools = ("com_validate_opens_clean",
                          "com_save_with_password")
    else:
        exc = UnsupportedStructure(ole_refusal_text(kind, path))
        exc.hint_tools = ("com_convert_format",)
    raise exc


# ------------------------------------------------------------ routing verbs

ROUTE_OPENPYXL = "openpyxl"       # clean workbook, model rewrite is safe
ROUTE_RAW_OOXML = "raw_ooxml"     # hazard present, surgical edit, byte-preserve
ROUTE_COM = "com"                 # hazard present, structural edit, Excel saves
ROUTE_REFUSE = "refuse"           # hazard present, no safe route, no override

# ------------------------------------------------------------ severity

SEV_DROPS = "drops"        # openpyxl silently loses this part on save
SEV_DEGRADES = "degrades"  # openpyxl round-trips it through a lossy model
SEV_CONDITIONAL = "conditional"  # preserved only under a special load flag


@dataclass(frozen=True)
class HazardSpec:
    """One fragile-part class: how to detect it in the zip and what openpyxl
    does to it. `matcher` is a predicate over a single archive member name."""

    key: str
    label: str
    severity: str
    survives_openpyxl: bool
    note: str
    matcher: "callable"


def _prefix(*prefixes: str):
    def m(name: str) -> bool:
        low = name.lower()
        return any(low.startswith(p) for p in prefixes)
    return m


def _exact(*names: str):
    wanted = {n.lower() for n in names}
    def m(name: str) -> bool:
        return name.lower() in wanted
    return m


def _basename(*names: str):
    wanted = {n.lower() for n in names}
    def m(name: str) -> bool:
        return name.rsplit("/", 1)[-1].lower() in wanted
    return m


def _basename_or_prefix(basenames: tuple[str, ...], prefixes: tuple[str, ...]):
    wanted = {n.lower() for n in basenames}
    def m(name: str) -> bool:
        low = name.lower()
        if low.rsplit("/", 1)[-1] in wanted:
            return True
        return any(low.startswith(p) for p in prefixes)
    return m


def _never(_name: str) -> bool:
    """Matcher for a hazard that has no part of its own. The in-part
    extension hazard is detected by READING worksheet content, never by a
    member name, and this matters beyond the scan: verify.part_loss_check
    walks HAZARD_SPECS matchers over the parts that VANISHED, and the
    worksheet part does not vanish (its extLst does). A never-matching
    matcher keeps that check honest instead of teaching it to excuse a lost
    worksheet."""
    return False


def _drawings_matcher(name: str) -> bool:
    """Match drawing parts under xl/drawings/, EXCEPT legacy-comment VML
    anchors named commentsDrawing*.vml (openpyxl's own output naming).
    openpyxl models legacy notes and regenerates their VML anchor on the
    round-trip (empirically confirmed), so a comment anchor is not shape
    loss and must not trip the SEV_DROPS drawings hazard.

    EXCEL'S naming is different: real Excel writes comment anchors as
    vmlDrawing*.vml -- the same basename form controls use -- so a NAME
    cannot tell a sticky note from a checkbox (this docstring used to claim
    the opposite, and every real old file with a comment refused plain edits
    on it: geriatric round, M-1). vmlDrawing*.vml therefore stays matched
    here, and _refine_comment_vml then READS each one: an anchor whose every
    shape is a ClientData ObjectType=\"Note\" is demoted from the hazard
    (content-verified comment-only), while anything carrying form controls
    or drawn shapes -- or that cannot be read or parsed -- stays flagged."""
    low = name.lower()
    if not low.startswith("xl/drawings/"):
        return False
    base = low.rsplit("/", 1)[-1]
    if base.startswith("commentsdrawing") and base.endswith(".vml"):
        return False
    return True


#: VML element local names that mean the part carries a DRAWN object of its
#: own (not a comment note): any of these keeps the part flagged.
_VML_DRAWN_LOCALS = frozenset({
    "oval", "rect", "roundrect", "line", "polyline", "arc", "curve",
    "image", "group",
})


def _vml_local(tag) -> str:
    return tag.rsplit("}", 1)[-1].lower() if isinstance(tag, str) else ""


def _vml_is_comment_only(data: bytes) -> bool | None:
    """True when a VML part holds nothing but legacy-comment note anchors:
    at least one shape, every shape carrying exactly one
    x:ClientData ObjectType="Note", and no drawn VML object anywhere.
    False for form controls (Checkbox/Button/Drop/... ClientData), drawn
    shapes, or an empty part; None when the XML cannot be parsed (the
    caller stays conservative and keeps the part flagged)."""
    try:
        from xml.etree.ElementTree import fromstring
        root = fromstring(data)
    except Exception:  # noqa: BLE001
        return None
    shapes = 0
    for el in root.iter():
        local = _vml_local(el.tag)
        if local in _VML_DRAWN_LOCALS:
            return False
        if local != "shape":
            continue
        shapes += 1
        client_data = [c for c in el.iter()
                       if _vml_local(c.tag) == "clientdata"]
        if len(client_data) != 1 or \
                client_data[0].get("ObjectType") != "Note":
            return False
    return shapes > 0


#: A legacy comments part, either spelling (Excel's xl/comments1.xml or
#: openpyxl's xl/comments/comment1.xml). The comment-anchor demotion requires
#: one: openpyxl regenerates a VML anchor only for notes it actually models,
#: so an ORPHANED Note-anchor part (no comments anywhere) stays flagged.
_COMMENTS_PART_RE = re.compile(
    r"^xl/(comments\d+\.xml|comments/comment\d+\.xml)$")


def _refine_comment_vml(
    found: dict[str, "Hazard"],
    nameset: set[str],
    reader: Callable[[str], bytes] | None,
) -> list[str]:
    """Demote content-verified comment-only vmlDrawing*.vml anchors from the
    SEV_DROPS drawings hazard, returning the demoted part names.

    Excel names comment anchors vmlDrawing*.vml (see _drawings_matcher), so
    the single most common geriatric scar -- a sticky note on an old business
    file -- refused every plain edit while the note in fact round-trips
    (openpyxl regenerates the anchor under its own commentsDrawing name,
    text, author, and position intact; empirically confirmed on the
    geriatric corpus). A namelist alone cannot make this call, because the
    same basename carries form controls; the demotion happens ONLY after
    reading the part and finding nothing but Note anchors. Without a reader
    (scan_names namelist-only) everything stays flagged, the conservative
    direction."""
    hz = found.get("drawings")
    if hz is None or reader is None:
        return []
    if not any(_COMMENTS_PART_RE.match(_normalized(n).lower())
               for n in nameset):
        return []  # no legacy comments to anchor: nothing to demote
    demoted: list[str] = []
    for part in hz.parts:
        base = part.lower().rsplit("/", 1)[-1]
        if not (base.startswith("vmldrawing") and base.endswith(".vml")):
            continue
        try:
            data = reader(part)
        except Exception:  # noqa: BLE001
            continue
        if _vml_is_comment_only(data) is True:
            demoted.append(part)
    if demoted:
        kept = [p for p in hz.parts if p not in demoted]
        if kept:
            hz.parts = kept
        else:
            del found["drawings"]
    return demoted


# ------------------------------------------------- in-part extension blocks

#: Hazard key for worksheet-level extLst content. It has no part of its own,
#: so scan_names detects it by reading the worksheet rather than by matching a
#: member name (see _never and _detect_in_part_extensions).
IN_PART_EXT_KEY = "in_part_extensions"

#: Worksheet extLst URIs and what each one carries, mirroring
#: openpyxl.xml.constants.EXT_TYPES so these labels line up with the
#: "... extension is not supported and will be removed" warnings the package
#: surfaces from the load. openpyxl drops EVERY worksheet-level ext, listed or
#: not, so a URI missing from this table is still reported (as an
#: unrecognized extension) rather than passed over.
EXT_DROP_LABELS: dict[str, str] = {
    "{78C0D931-6437-407D-A8EE-F0AAD7539E65}":
        "x14 conditional formatting (data bars, icon sets)",
    "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}": "x14 data validation",
    "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}": "sparkline groups",
    "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}": "worksheet slicer list",
    "{3A4CF648-6AED-40F4-86FF-DC5316D8AED3}": "worksheet slicer list",
    "{FC87AEE6-9EDD-4A0A-B7FB-166176984837}": "protected ranges",
    "{01252117-D84E-4E92-8308-4BE1C098FCBB}": "ignored-error markers",
    "{F7C9EE02-42E1-4005-9D12-6889AFFD525C}": "web extensions",
    "{7E03D99C-DC04-49D9-9315-930204A7B6E9}": "timeline references",
}


def _top_level_ext_uris(data: bytes) -> list[str]:
    """URIs of the <ext> elements directly under the worksheet's top-level
    <extLst>, in document order, de-duplicated.

    Depth matters. A cfRule carries its OWN nested extLst holding the x14 rule
    id ({B025F937-...}), and that one is not a dropped extension block, it is
    a pointer into the block. Matching URIs with a regex over the whole part
    would report it as a second hazard and misdescribe what is at risk, so the
    walk only accepts worksheet > extLst > ext.

    Streaming expat, no tree: a worksheet can be hundreds of megabytes and
    this runs on every open. A malformed part yields whatever was read before
    the error, which is the conservative direction (report, do not swallow)."""
    import xml.parsers.expat

    uris: list[str] = []
    stack: list[str] = []

    def start(name: str, attrs: dict) -> None:
        local = name.rsplit(":", 1)[-1]
        stack.append(local)
        if len(stack) == 3 and stack[1] == "extLst" and local == "ext":
            uri = attrs.get("uri")
            if uri and uri not in uris:
                uris.append(uri)

    def end(_name: str) -> None:
        if stack:
            stack.pop()

    parser = xml.parsers.expat.ParserCreate()
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    try:
        parser.Parse(data, True)
    except Exception:  # noqa: BLE001
        return uris
    return uris


def _is_worksheet_part(name: str) -> bool:
    low = _normalized(name).lower()
    return (low.startswith("xl/worksheets/") and low.endswith(".xml")
            and "/_rels/" not in low)


def _detect_in_part_extensions(
    found: dict[str, "Hazard"], names: Iterable[str],
    part_reader: Callable[[str], bytes] | None,
) -> None:
    """Add the in-part extension hazard when a worksheet carries a top-level
    extLst. Needs part content, so it is a no-op on the namelist-only path."""
    if part_reader is None:
        return
    labels: list[str] = []
    parts: list[str] = []
    for name in names:
        if not _is_worksheet_part(name):
            continue
        try:
            data = part_reader(name)
        except Exception:  # noqa: BLE001
            continue
        # Cheap reject before the parse. It is a UTF-8 byte scan, so a part
        # with a UTF-16 byte-order mark (legal XML, and openpyxl loads one)
        # skips the reject and goes straight to expat, which reads the
        # declared encoding itself (edge audit 2026-09-04, S3).
        if not data[:2] in (b"\xff\xfe", b"\xfe\xff") and b"extLst" not in data:
            continue
        uris = _top_level_ext_uris(data)
        if not uris:
            continue
        parts.append(name)
        for uri in uris:
            label = EXT_DROP_LABELS.get(
                uri.upper(), f"unrecognized extension {uri}")
            if label not in labels:
                labels.append(label)
    if not parts:
        return
    spec = _SPEC_BY_KEY[IN_PART_EXT_KEY]
    found[IN_PART_EXT_KEY] = Hazard(
        spec.key, f"{spec.label}: {', '.join(labels)}", spec.severity,
        spec.survives_openpyxl, spec.note, parts,
    )


# The knowledge table. survives_openpyxl is the CLAIM the fidelity harness
# checks. Ordered most-to-least destructive for readable reports.
HAZARD_SPECS: tuple[HazardSpec, ...] = (
    HazardSpec(
        "slicers", "slicers", SEV_DROPS, False,
        "openpyxl has no slicer model; slicers are removed on save.",
        _prefix("xl/slicers/", "xl/slicercaches/"),
    ),
    HazardSpec(
        "timelines", "timelines", SEV_DROPS, False,
        "openpyxl has no timeline model; timelines are removed on save.",
        _prefix("xl/timelines/", "xl/timelinecaches/"),
    ),
    HazardSpec(
        "activex", "ActiveX controls", SEV_DROPS, False,
        "ActiveX controls are dropped; even with keep_vba they break.",
        _prefix("xl/activex/"),
    ),
    HazardSpec(
        "ctrlprops", "form/ActiveX control properties", SEV_DROPS, False,
        "control property parts are dropped on the openpyxl round-trip.",
        _prefix("xl/ctrlprops/"),
    ),
    HazardSpec(
        "drawings", "drawings / shapes", SEV_DROPS, False,
        "openpyxl tutorial states shapes are lost from existing files; only "
        "charts are re-modeled, textboxes and shapes vanish. Legacy-comment "
        "VML anchors are excluded (openpyxl preserves legacy notes).",
        _drawings_matcher,
    ),
    HazardSpec(
        "media", "embedded media / images", SEV_DROPS, False,
        "images embedded in an existing file are not carried through the "
        "openpyxl load+save of that file.",
        _prefix("xl/media/"),
    ),
    HazardSpec(
        "threaded_comments", "threaded comments", SEV_DROPS, False,
        "openpyxl models legacy notes only; threadedComments parts and the "
        "xl/persons/ author registry they reference are lost.",
        _prefix("xl/threadedcomments/", "xl/persons/"),
    ),
    HazardSpec(
        "embeddings", "embedded OLE objects", SEV_DROPS, False,
        "xl/embeddings/ holds embedded documents (Word, PDF, other OLE "
        "payloads); openpyxl does not model them and drops them on save.",
        _prefix("xl/embeddings/"),
    ),
    HazardSpec(
        "power_query", "customXml (incl. Power Query DataMashup)", SEV_DROPS,
        False,
        "customXml parts are not modeled by openpyxl and are dropped. Power "
        "Query M lives here (the DataMashup item), but so do add-in stores "
        "and SharePoint property sets; a namelist scan cannot tell them "
        "apart, so every customXml part is flagged.",
        _prefix("customxml/"),
    ),
    HazardSpec(
        "connections", "data connections", SEV_DROPS, False,
        "xl/connections.xml is not modeled by openpyxl and is dropped.",
        _exact("xl/connections.xml"),
    ),
    HazardSpec(
        "query_tables", "legacy query tables", SEV_DROPS, False,
        "xl/queryTables/ (legacy web/database query definitions) is not "
        "modeled by openpyxl and is dropped.",
        _prefix("xl/querytables/"),
    ),
    HazardSpec(
        "data_model", "Power Pivot data model", SEV_DROPS, False,
        "xl/model/ holds the Power Pivot / data-model payload; openpyxl does "
        "not model it and drops it on save.",
        _prefix("xl/model/"),
    ),
    HazardSpec(
        "named_sheet_views", "named sheet views", SEV_DROPS, False,
        "xl/namedSheetViews/ (saved temporary filter/sort views) is not "
        "modeled by openpyxl and is dropped.",
        _prefix("xl/namedsheetviews/"),
    ),
    HazardSpec(
        "rich_metadata", "dynamic-array / rich-value metadata", SEV_DROPS,
        False,
        "xl/metadata.xml carries dynamic-array spill and rich-value metadata, "
        "and xl/richData/ holds the rich-value payloads (stock/geo types, "
        "images-in-cells); openpyxl models neither, so spill ranges and rich "
        "values lose their backing.",
        _basename_or_prefix(("metadata.xml", "richdata.xml"),
                            ("xl/richdata/",)),
    ),
    HazardSpec(
        IN_PART_EXT_KEY, "in-sheet extension blocks", SEV_DROPS, False,
        "worksheet-level extLst blocks (x14 conditional formatting, sparkline "
        "groups, slicer lists, protected ranges) live inside the worksheet "
        "part, which survives the save at full size. openpyxl parses them, "
        "warns that each is unsupported, and writes none of them back, so the "
        "rules are gone from a file that looks untouched. Detected by reading "
        "the worksheet, not by a part name.",
        _never,
    ),
    HazardSpec(
        "pivot", "pivot tables / caches", SEV_DEGRADES, True,
        "openpyxl has read-support for pivots (preserved, light edits only); "
        "attached slicers still die and pivot charts follow chart fidelity.",
        _prefix("xl/pivotcache/", "xl/pivottables/"),
    ),
    HazardSpec(
        "charts", "charts", SEV_DEGRADES, True,
        "charts are re-serialized through openpyxl's chart model; features "
        "the model does not know are degraded or lost.",
        _prefix("xl/charts/"),
    ),
    HazardSpec(
        "vba", "VBA project", SEV_CONDITIONAL, True,
        "vbaProject.bin survives ONLY with load_workbook(keep_vba=True) and an "
        ".xlsm save; a default load drops it. Form controls break regardless.",
        _basename("vbaproject.bin"),
    ),
)

_SPEC_BY_KEY = {s.key: s for s in HAZARD_SPECS}


@dataclass
class Hazard:
    key: str
    label: str
    severity: str
    survives_openpyxl: bool
    note: str
    parts: list[str] = field(default_factory=list)


@dataclass
class HazardReport:
    path: str
    parts: list[str]
    hazards: list[Hazard]
    error: str | None = None
    #: uncompressed part sizes from the zip central directory (scan_path only;
    #: empty for scan_names). verify-after-write uses these to catch a fragile
    #: part REPLACED with empty content rather than dropped outright.
    sizes: dict[str, int] = field(default_factory=dict)
    #: vmlDrawing*.vml parts the content classifier verified as comment-only
    #: anchors and demoted from the drawings hazard (geriatric round, M-1).
    #: openpyxl regenerates them under its own commentsDrawing naming on
    #: save, so verify-after-write excuses exactly these names vanishing.
    comment_anchor_vml: list[str] = field(default_factory=list)
    #: WHY the scan failed: "corrupt" | "missing" | "locked" | None. A
    #: byte-range-locked (but healthy) file makes zipfile raise BadZipFile
    #: when the lock covers the central directory, and telling the owner of
    #: a 1-of-1 heirloom the file is CORRUPT when an indexer/AV merely holds
    #: a region lock invites a destructive "repair" (destroyer round, M-2).
    error_kind: str | None = None

    @property
    def clean(self) -> bool:
        """No fragile part present at all: openpyxl round-trip is safe."""
        return not self.hazards and self.error is None

    @property
    def lossy_keys(self) -> list[str]:
        """Hazards openpyxl would DROP outright (not merely degrade or
        conditionally preserve). These are the never-silently-lose set."""
        return [h.key for h in self.hazards if h.severity == SEV_DROPS]

    @property
    def would_lose(self) -> bool:
        return bool(self.lossy_keys)

    def labels(self) -> list[str]:
        return [h.label for h in self.hazards]

    def as_dict(self) -> dict:
        return {
            "path": self.path,
            "clean": self.clean,
            "would_lose": self.would_lose,
            "hazards": [
                {
                    "key": h.key,
                    "label": h.label,
                    "severity": h.severity,
                    "survives_openpyxl": h.survives_openpyxl,
                    "parts": h.parts,
                    "note": h.note,
                }
                for h in self.hazards
            ],
            "error": self.error,
        }


# --------------------------------------------------- chart-vs-shape drawings
#
# The Phase 2 flagged over-refusal: every part under xl/drawings/ matched the
# SEV_DROPS "drawings" (shape-loss) spec, so a workbook whose ONLY drawing is a
# chart was treated as unrecoverable shape loss and refused. But a chart's
# drawing is just its anchor: openpyxl re-models charts and round-trips both the
# chart part and its anchor drawing byte-for-byte (Phase 1 fidelity table +
# empirically re-confirmed here: chart.xlsx loses NO parts on an openpyxl
# round-trip). Only genuine shapes (textboxes, rectangles, form/ActiveX
# controls) and pictures are dropped.
#
# THE HEURISTIC (cheap, and honest about its limits): a drawing part is a pure
# chart anchor iff BOTH hold:
#   (a) its relationship part (xl/drawings/_rels/drawingN.xml.rels) exists and
#       EVERY relationship in it targets a chart (Type ending "/chart"), and
#   (b) the drawing XML itself contains no inline shape content (sp / pic /
#       grpSp / cxnSp elements). This second check matters: a textbox drawn
#       NEXT TO a chart in the same drawing needs NO relationship of its own,
#       so a rels-only test would call that drawing chart-only and let the
#       textbox be silently dropped. The re-audit closed that bypass.
#   - No rels part at all  -> inline shapes (textbox/rectangle), a real drop
#     (this is exactly shape.xlsx: xl/drawings/drawing1.xml with no rels).
#   - rels all chart AND no inline shape elements -> chart anchor, NOT a
#     shape-loss drop; chart fidelity is already tracked by the separate
#     SEV_DEGRADES "charts" spec.
#   - rels present with any image / oleObject / control target -> a drop
#     (this is image.xlsx: a /image relationship; media survival is Pillow- and
#     authorship-dependent per Phase 1, so it stays conservatively flagged).
#
# LIMITS: the classification needs to read the tiny rels part and the drawing
# part, so it is not decidable from the central-directory namelist ALONE.
# scan_path supplies a reader (two extra small-part reads only when drawings
# are present, still milliseconds); scan_names WITHOUT a reader cannot see
# part content and so stays CONSERVATIVE, keeping every drawing flagged as a
# potential drop (a false positive costs a needless raw/COM route, never
# silent loss). An unparseable drawing part is likewise conservatively
# flagged. A chart drawn with no relationship part (not produced by Excel or
# openpyxl in practice) would be conservatively flagged.

_CHART_REL_SUFFIX = "/chart"
_TYPE_RE = re.compile(r'Type="([^"]+)"')


def _drawing_rels_for(drawing_part: str) -> str:
    head, base = drawing_part.rsplit("/", 1)
    return f"{head}/_rels/{base}.rels"


#: Local element names that mean a drawing carries inline shape content
#: openpyxl would drop: shapes, pictures, shape groups, connectors.
_SHAPE_LOCALS = frozenset({"sp", "pic", "grpSp", "cxnSp"})


def _drawing_has_shape_content(data: bytes) -> bool | None:
    """True when the drawing XML contains inline shape elements (dropped by
    openpyxl even when every relationship targets a chart). None when the part
    cannot be parsed (the caller stays conservative)."""
    try:
        from xml.etree.ElementTree import fromstring
        root = fromstring(data)
    except Exception:
        return None
    for el in root.iter():
        tag = el.tag
        local = tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""
        if local in _SHAPE_LOCALS:
            return True
    return False


def _drawing_is_chart_only(
    drawing_part: str, nameset: set[str],
    rels_reader: Callable[[str], bytes] | None,
) -> bool | None:
    """True: pure chart anchor (survives). False: shape/picture/control drawing
    (drops). None: cannot tell without reading part content (namelist-only
    path, or an unreadable/unparseable part)."""
    rels = _drawing_rels_for(drawing_part)
    if rels not in nameset:
        return False  # no rels: inline shapes, a genuine shape-loss drop
    if rels_reader is None:
        return None   # rels exists but its content is not readable here
    try:
        data = rels_reader(rels)
    except Exception:
        return None
    if not data:
        return None
    types = _TYPE_RE.findall(data.decode("utf-8", "replace"))
    if not types:
        return False
    if not all(t.rstrip("/").lower().endswith(_CHART_REL_SUFFIX)
               for t in types):
        return False
    # Every relationship targets a chart; now confirm the drawing XML itself
    # holds no inline shape (a textbox beside the chart needs no rel).
    try:
        drawing_data = rels_reader(drawing_part)
    except Exception:
        return None
    has_shape = _drawing_has_shape_content(drawing_data)
    if has_shape is None:
        return None
    return not has_shape


def _refine_drawings(
    found: dict[str, "Hazard"], nameset: set[str],
    rels_reader: Callable[[str], bytes] | None,
) -> None:
    """Drop pure chart-anchor drawings from the SEV_DROPS drawings hazard.
    Chart fidelity is covered by the separate SEV_DEGRADES charts spec, so a
    chart-only workbook must not be treated as shape loss."""
    hz = found.get("drawings")
    if hz is None:
        return
    drawing_xmls = [
        p for p in hz.parts
        if p.lower().startswith("xl/drawings/")
        and "/_rels/" not in p.lower()
        and p.lower().endswith(".xml")
    ]
    if not drawing_xmls:
        return
    chart_only = {
        d for d in drawing_xmls
        if _drawing_is_chart_only(d, nameset, rels_reader) is True
    }
    exclude = set(chart_only)
    for d in chart_only:
        rels = _drawing_rels_for(d)
        if rels in nameset:
            exclude.add(rels)
    # Subtract ONLY the demoted chart anchors (and their rels). The old
    # code rebuilt the part list from the surviving .xml drawings, which
    # silently dropped every NON-xml member -- an Excel-authored
    # vmlDrawing*.vml sharing the hazard with a chart anchor vanished from
    # the report, was therefore never warned, and its loss at verify time
    # was unexcusable even under allow_loss: the advertised remedy could
    # never succeed on such a file (destroyer round, M-3).
    kept = [p for p in hz.parts if p not in exclude]
    if not kept:
        del found["drawings"]  # every drawing is a pure chart anchor
        return
    hz.parts = kept


def _normalized(name: str) -> str:
    r"""A part name reduced to the spelling the matchers expect.

    Zip member names are attacker-controlled text, and a package can carry the
    same part under a spelling the prefix matchers miss: './xl/slicers/x.xml',
    '/xl/slicers/x.xml', 'xl//slicers/x.xml', or backslash separators.
    Adversarial round: those four evaded the hazard scan, so the mutation was
    NOT refused up front (the default-fail inventory check in
    verify-after-write still caught the loss and refused the save, which is
    why nothing was destroyed). Normalizing for the MATCH restores the loud,
    specific up-front refusal; the reported part keeps its real name, because
    the inventory diff compares against the real one."""
    low = name.replace("\\", "/")
    while low.startswith("./"):
        low = low[2:]
    low = low.lstrip("/")
    while "//" in low:
        low = low.replace("//", "/")
    return low


def scan_names(names: Iterable[str], path: str = "<names>", *,
               part_reader: Callable[[str], bytes] | None = None,
               rels_reader: Callable[[str], bytes] | None = None,
               ) -> HazardReport:
    """Classify a pre-listed set of archive member names. Split out so the
    detection logic is testable without a real file on disk.

    `part_reader`, when supplied (scan_path always does), reads a member's
    bytes. It serves two content-dependent checks that a namelist cannot
    answer: whether a drawing is a pure chart anchor rather than shape loss
    (chart-vs-shape fix), and whether a worksheet carries a top-level extLst
    whose contents openpyxl drops. Without it the scan reports only the
    name-detectable hazards; `rels_reader` is the former name of the same
    argument and is still accepted."""
    reader = part_reader or rels_reader
    names = list(names)
    nameset = set(names)
    found: dict[str, Hazard] = {}
    for name in names:
        probe = _normalized(name)
        for spec in HAZARD_SPECS:
            if spec.matcher(name) or (probe != name and spec.matcher(probe)):
                h = found.get(spec.key)
                if h is None:
                    h = Hazard(
                        spec.key, spec.label, spec.severity,
                        spec.survives_openpyxl, spec.note,
                    )
                    found[spec.key] = h
                h.parts.append(name)
    _refine_drawings(found, nameset, reader)
    demoted_vml = _refine_comment_vml(found, nameset, reader)
    _detect_in_part_extensions(found, names, reader)
    ordered = [found[s.key] for s in HAZARD_SPECS if s.key in found]
    return HazardReport(path=path, parts=names, hazards=ordered,
                        comment_anchor_vml=demoted_vml)


def scan_path(path: str) -> HazardReport:
    """Open the .xlsx/.xlsm as a zip, classify its members, and return the
    report.

    The member list comes from the central directory alone (no decompression),
    and that is still all most hazards need. Two checks read part content: the
    chart-vs-shape refinement reads a drawing and its tiny rels part, and the
    in-part extension detector reads each worksheet to see whether openpyxl
    would drop a top-level extLst. The worksheet read is what makes this scan
    proportional to sheet size rather than the old flat milliseconds; the
    alternative is a data bar that vanishes without a word, so the read
    stays."""
    kind = ole_container_kind(path)
    if kind is not None:
        # An OLE container (legacy .xls, or an encrypted package). Modern
        # Excel's .xls files embed a zip fragment, so without this sniff the
        # zip open below SUCCEEDS on a BIFF file and the scan mistakes the
        # theme fragment for the workbook (geriatric round, H-1).
        return HazardReport(path=path, parts=[], hazards=[],
                            error=ole_refusal_text(kind, path),
                            error_kind="ole")
    try:
        with zipfile.ZipFile(path) as zf:
            infos = zf.infolist()
            names = [i.filename for i in infos]
            sizes = {i.filename: i.file_size for i in infos}
            # Only drawings are cached: the refinement reads a drawing and
            # its rels part and may revisit them. Worksheets are read once
            # and released, so a big workbook is not held in memory whole
            # just to look at its extLst.
            cache: dict[str, bytes] = {}

            def reader(member: str, _zf=zf, _cache=cache) -> bytes:
                if not member.lower().startswith("xl/drawings/"):
                    return _zf.read(member)
                if member not in _cache:
                    _cache[member] = _zf.read(member)
                return _cache[member]

            rep = scan_names(names, path=path, part_reader=reader)
            rep.sizes = sizes
            return rep
    except zipfile.BadZipFile:
        # zipfile answers "not a zip" for a file it could not fully READ
        # too: a byte-range lock over the central directory surfaces as
        # BadZipFile. Probe before diagnosing corruption; the wrong word
        # here sends a panicked owner to a destructive "repair".
        if file_read_blocked(path):
            return HazardReport(
                path=path, parts=[], hazards=[], error=_LOCKED_ERROR,
                error_kind="locked")
        return HazardReport(path=path, parts=[], hazards=[],
                            error="not a valid zip / OOXML package",
                            error_kind="corrupt")
    except FileNotFoundError:
        return HazardReport(path=path, parts=[], hazards=[],
                            error="file not found", error_kind="missing")
    except PermissionError:
        return HazardReport(
            path=path, parts=[], hazards=[], error=_LOCKED_ERROR,
            error_kind="locked")
    except OSError as exc:
        return HazardReport(
            path=path, parts=[], hazards=[],
            error=f"cannot read the file ({type(exc).__name__}: {exc})",
            error_kind="os")


_LOCKED_ERROR = (
    "another process holds a lock on part of this file (an antivirus "
    "scanner, indexer, or app with a byte-range lock); the file itself may "
    "be perfectly intact. Wait for the other process to release it, or "
    "close whatever holds it open, then retry")


def file_read_blocked(path: str) -> bool:
    """True when reading the file end to end fails with an access error,
    i.e. some process holds a byte-range or exclusive lock. Only consulted
    on failure paths, so the full read is paid rarely."""
    try:
        with open(path, "rb") as fh:
            while fh.read(1 << 20):
                pass
        return False
    except PermissionError:
        return True
    except OSError as exc:
        return getattr(exc, "errno", None) == 13


def route(
    report: HazardReport,
    *,
    edit: str = "surgical",
    com_available: bool = False,
    allow_loss: bool = False,
) -> tuple[str, str]:
    """The per-tool routing decision (DESIGN Section 3.1). Returns
    (route_verb, reason).

    edit: 'surgical' (single cells/values/formulas, raw OOXML can preserve) or
    'structural' (row/col shifts, beyond safe raw surgery).
    """
    if report.error is not None:
        return ROUTE_REFUSE, f"cannot scan package: {report.error}"
    if report.clean:
        return ROUTE_OPENPYXL, "no fragile parts; openpyxl round-trip is safe"
    if not report.would_lose:
        # only degrade/conditional hazards (charts, pivots, VBA). openpyxl can
        # proceed but the caller must opt into fidelity loss on charts/pivots.
        if allow_loss:
            return ROUTE_OPENPYXL, (
                "only degradable parts present; proceeding with allow_loss"
            )
        if com_available:
            return ROUTE_COM, (
                "degradable parts present (" + ", ".join(report.labels())
                + "); COM preserves full fidelity"
            )
        return ROUTE_REFUSE, (
            "degradable parts present (" + ", ".join(report.labels())
            + "); pass allow_loss to accept the degradation, or leave the "
              "workbook to Excel. The com pack drives Excel for "
              "recalculation, pivots, export, rendering and conversion, "
              "and writes no cells, so it cannot carry this edit"
        )
    # would_lose: parts openpyxl drops outright.
    if edit == "surgical" and _raw_safe(report):
        return ROUTE_RAW_OOXML, (
            "lossy parts present (" + ", ".join(
                _SPEC_BY_KEY[k].label for k in report.lossy_keys)
            + "); surgical edit routes to byte-preserving raw OOXML"
        )
    if com_available:
        return ROUTE_COM, (
            "lossy parts present (" + ", ".join(
                _SPEC_BY_KEY[k].label for k in report.lossy_keys)
            + "); COM lets Excel save with everything intact"
        )
    if allow_loss:
        return ROUTE_OPENPYXL, (
            "lossy parts present but allow_loss set; the backup is the safety "
            "net and the loss is explicit"
        )
    return ROUTE_REFUSE, (
        "would lose " + ", ".join(
            _SPEC_BY_KEY[k].label for k in report.lossy_keys)
        + "; a surgical raw-OOXML edit preserves them, or pass allow_loss "
        "with a backup to accept the loss. The com pack writes no cells "
        "and is not a route for this"
    )


# ------------------------------------------------- refusal routes (honest)
# The gate used to offer "enable COM (com pack)" as the remedy for any
# refused file-tier mutation. The com pack holds eleven tools -- recalculate,
# com_manage_pivot, com_goal_seek, com_export_pdf, com_render_sheet,
# com_convert_format, com_save_with_password, com_set_sparkline, com_autofit,
# com_validate_opens_clean, com_status -- and NONE of them writes a cell, a
# format, or a row. A caller who enabled the pack as instructed found nothing
# there that could perform the operation that was refused. Naming a remedy
# that does not exist is worse than naming none, in a product whose whole
# claim is that its labels are honest.
#
# So the routes are now the two that are real: accept the loss explicitly,
# with the exact cost stated per class, or do not put this workbook through
# the file tier. Preservation is the fix and it is not in this release.

#: The route strings, kept as data so the refusal text, the .detail payload,
#: and the tests all read the same list.
ROUTE_ALLOW_LOSS = "allow_loss:true (backed up first; the losses below are permanent in the saved file)"
ROUTE_LEAVE_ALONE = "leave this workbook to Excel: copy_workbook branches it byte-for-byte, and reads never touch it"


def loss_costs(keys) -> list[str]:
    """"<label>: <what openpyxl does to it>" for each refused hazard class.

    The caller deciding whether to pass allow_loss is deciding to destroy
    specific content, so the refusal states which content and what happens
    to it, rather than a class name the caller has to go look up."""
    out = []
    for k in keys:
        spec = _SPEC_BY_KEY.get(k)
        if spec is None:
            continue
        out.append(f"{spec.label}: {spec.note.rstrip('.')}")
    return out


def refusal_routes() -> list[str]:
    """The routes a refused file-tier mutation actually has."""
    return [ROUTE_ALLOW_LOSS, ROUTE_LEAVE_ALONE]


def _raw_safe(report: HazardReport) -> bool:
    """Raw OOXML surgery preserves untouched parts byte-identical, so it is
    safe for any lossy hazard as long as the edit itself is surgical. The one
    class raw surgery must not touch structurally is the pivot cache index,
    but a surgical CELL edit never touches it, so raw is safe here."""
    return True


__all__ = [
    "HazardReport", "Hazard", "HazardSpec", "HAZARD_SPECS",
    "EXT_DROP_LABELS", "IN_PART_EXT_KEY",
    "loss_costs", "refusal_routes", "ROUTE_ALLOW_LOSS", "ROUTE_LEAVE_ALONE",
    "ole_container_kind", "ole_refusal_text", "refuse_ole_container",
    "OLE_KIND_BIFF", "OLE_KIND_ENCRYPTED", "OLE_KIND_UNKNOWN",
    "scan_path", "scan_names", "route",
    "ROUTE_OPENPYXL", "ROUTE_RAW_OOXML", "ROUTE_COM", "ROUTE_REFUSE",
    "SEV_DROPS", "SEV_DEGRADES", "SEV_CONDITIONAL",
]

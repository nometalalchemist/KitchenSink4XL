"""core/hazard.py: the round-trip hazard scan and per-tool routing decision.

The single most important new module (PLAN reuse ledger 1.2, DESIGN Section
3.1). A cheap zip central-directory read detects the fragile parts openpyxl
would drop on save (slicers, timelines, drawings/shapes, ActiveX, threaded
comments, Power Query DataMashup, connections, dynamic-array metadata, pivot
caches, VBA) WITHOUT an object-model load, then routes each mutation to
openpyxl / raw OOXML surgery / COM / HAZARD_REFUSED.

The knowledge table in HAZARD_SPECS encodes what openpyxl preserves vs drops.
The Phase 1 fidelity harness (tests/fixtures + scripts) VALIDATES every
`survives_openpyxl` claim in this table empirically against real fixtures.
Any file whose parts openpyxl would drop MUST be flagged here; a lossy file
this scan calls clean is the exact incumbent failure the brand exists to
prevent, so the table errs toward flagging (a false positive costs a needless
raw/COM route; a false negative silently destroys user content).
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from typing import Callable, Iterable

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
        "charts are re-modeled, textboxes and shapes vanish.",
        _prefix("xl/drawings/"),
    ),
    HazardSpec(
        "media", "embedded media / images", SEV_DROPS, False,
        "images embedded in an existing file are not carried through the "
        "openpyxl load+save of that file.",
        _prefix("xl/media/"),
    ),
    HazardSpec(
        "threaded_comments", "threaded comments", SEV_DROPS, False,
        "openpyxl models legacy notes only; threadedComments parts are lost.",
        _prefix("xl/threadedcomments/"),
    ),
    HazardSpec(
        "power_query", "Power Query (DataMashup)", SEV_DROPS, False,
        "Power Query M is stored in customXml DataMashup, which openpyxl does "
        "not model; queries are dropped.",
        _prefix("customxml/"),
    ),
    HazardSpec(
        "connections", "data connections", SEV_DROPS, False,
        "xl/connections.xml is not modeled by openpyxl and is dropped.",
        _exact("xl/connections.xml"),
    ),
    HazardSpec(
        "rich_metadata", "dynamic-array / rich-value metadata", SEV_DROPS,
        False,
        "xl/metadata.xml carries dynamic-array spill and rich-value metadata; "
        "openpyxl does not model it, so spill ranges can lose their metadata.",
        _basename("metadata.xml", "richdata.xml"),
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
# chart anchor iff its relationship part (xl/drawings/_rels/drawingN.xml.rels)
# exists and EVERY relationship in it targets a chart (Type ending "/chart").
#   - No rels part at all  -> inline shapes (textbox/rectangle), a real drop
#     (this is exactly shape.xlsx: xl/drawings/drawing1.xml with no rels).
#   - rels present, all chart -> chart anchor, NOT a shape-loss drop; chart
#     fidelity is already tracked by the separate SEV_DEGRADES "charts" spec.
#   - rels present with any image / oleObject / control target -> a drop
#     (this is image.xlsx: a /image relationship; media survival is Pillow- and
#     authorship-dependent per Phase 1, so it stays conservatively flagged).
#
# LIMITS: the classification needs to read the tiny rels part, so it is not
# decidable from the central-directory namelist ALONE. scan_path supplies a
# reader (one extra small-part read only when drawings are present, still
# milliseconds); scan_names WITHOUT a reader cannot see rels content and so
# stays CONSERVATIVE, keeping every drawing flagged as a potential drop (a false
# positive costs a needless raw/COM route, never silent loss). A drawing that
# mixes a chart with a shape is (correctly) treated as a drop. A chart drawn
# with no relationship part (not produced by Excel or openpyxl in practice)
# would be conservatively flagged.

_CHART_REL_SUFFIX = "/chart"
_TYPE_RE = re.compile(r'Type="([^"]+)"')


def _drawing_rels_for(drawing_part: str) -> str:
    head, base = drawing_part.rsplit("/", 1)
    return f"{head}/_rels/{base}.rels"


def _drawing_is_chart_only(
    drawing_part: str, nameset: set[str],
    rels_reader: Callable[[str], bytes] | None,
) -> bool | None:
    """True: pure chart anchor (survives). False: shape/picture/control drawing
    (drops). None: cannot tell without reading the rels (namelist-only path)."""
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
    return all(t.rstrip("/").lower().endswith(_CHART_REL_SUFFIX) for t in types)


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
    shape_drawings = [
        d for d in drawing_xmls
        if _drawing_is_chart_only(d, nameset, rels_reader) is not True
    ]
    if not shape_drawings:
        del found["drawings"]  # every drawing is a pure chart anchor
        return
    keep: set[str] = set()
    for d in shape_drawings:
        keep.add(d)
        rels = _drawing_rels_for(d)
        if rels in nameset:
            keep.add(rels)
    hz.parts = [p for p in hz.parts if p in keep]


def scan_names(names: Iterable[str], path: str = "<names>", *,
               rels_reader: Callable[[str], bytes] | None = None) -> HazardReport:
    """Classify a pre-listed set of archive member names. Split out so the
    detection logic is testable without a real file on disk. `rels_reader`,
    when supplied (scan_path does), reads a drawing's tiny rels part so a
    chart-only drawing is not misflagged as shape loss (chart-vs-shape fix)."""
    names = list(names)
    nameset = set(names)
    found: dict[str, Hazard] = {}
    for name in names:
        for spec in HAZARD_SPECS:
            if spec.matcher(name):
                h = found.get(spec.key)
                if h is None:
                    h = Hazard(
                        spec.key, spec.label, spec.severity,
                        spec.survives_openpyxl, spec.note,
                    )
                    found[spec.key] = h
                h.parts.append(name)
    _refine_drawings(found, nameset, rels_reader)
    ordered = [found[s.key] for s in HAZARD_SPECS if s.key in found]
    return HazardReport(path=path, parts=names, hazards=ordered)


def scan_path(path: str) -> HazardReport:
    """Open the .xlsx/.xlsm as a zip, read ONLY the central directory (the
    namelist, no decompression, no XML parse), and classify. Milliseconds."""
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            has_drawings = any(
                n.lower().startswith("xl/drawings/") for n in names)
            reader = None
            if has_drawings:
                cache: dict[str, bytes] = {}

                def reader(member: str, _zf=zf, _cache=cache) -> bytes:
                    if member not in _cache:
                        _cache[member] = _zf.read(member)
                    return _cache[member]

                rep = scan_names(names, path=path, rels_reader=reader)
                return rep
    except zipfile.BadZipFile:
        return HazardReport(path=path, parts=[], hazards=[],
                            error="not a valid zip / OOXML package")
    except FileNotFoundError:
        return HazardReport(path=path, parts=[], hazards=[],
                            error="file not found")
    return scan_names(names, path=path)


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
            + "); enable COM or pass allow_loss to accept degradation"
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
        + "; enable COM, use a surgical raw-OOXML edit, or pass allow_loss "
        "with a backup"
    )


def _raw_safe(report: HazardReport) -> bool:
    """Raw OOXML surgery preserves untouched parts byte-identical, so it is
    safe for any lossy hazard as long as the edit itself is surgical. The one
    class raw surgery must not touch structurally is the pivot cache index,
    but a surgical CELL edit never touches it, so raw is safe here."""
    return True


__all__ = [
    "HazardReport", "Hazard", "HazardSpec", "HAZARD_SPECS",
    "scan_path", "scan_names", "route",
    "ROUTE_OPENPYXL", "ROUTE_RAW_OOXML", "ROUTE_COM", "ROUTE_REFUSE",
    "SEV_DROPS", "SEV_DEGRADES", "SEV_CONDITIONAL",
]

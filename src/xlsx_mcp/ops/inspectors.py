"""ops/inspectors.py: read-side inspectors (DESIGN Section 11; io and data).

Four read-only tools over parts openpyxl either models lightly or not at
all. Nothing here mutates; every function opens the package (or the model)
for reading and reports.

- get_external_links: xl/externalLinks/ parts (which openpyxl preserves)
  parsed raw: target workbook path from the rels, referenced sheet names,
  and the cached cell values Excel stored at last refresh. Cached values
  can carry data from files the reader has no access to, so the output is
  marked sensitive.
- inspect_vba: presence and size always; module names and stream sizes
  when vbaProject.bin parses as a real CFB container (a compact MS-CFB
  reader plus the MS-OVBA dir-stream decompressor live below, validated
  against spec-constructed fixtures). Read-only by design: this server
  never authors VBA. A blob that does not parse degrades honestly to
  presence + size + a note.
- get_pivot: existing pivot tables described from openpyxl's read model:
  location, source range, row/column/data/page fields by name, cache
  record count and last-refresh stamp. Description only; creation and
  refresh are COM-tier.
- get_connections: xl/connections.xml parsed raw (name, type, connection
  strings, refresh flags) plus Power Query presence evidence (DataMashup
  customXml, mashup providers). Connection strings can embed credentials,
  so the output is marked sensitive.
"""

from __future__ import annotations

import codecs
import os
import struct
import zipfile
from typing import Any
from xml.etree.ElementTree import fromstring

from ..core import hazard as _hazard
from ..core.errors import WorkbookCorrupt, WorkbookNotFound, XlMcpError
from ..core.sandbox import check_path
from . import gridio

_PKG_REL_NS = ("{http://schemas.openxmlformats.org/package/2006/"
               "relationships}")

_MAX_CACHED_CELLS_PER_LINK = 100


def _open_zip(path: str) -> zipfile.ZipFile:
    p = check_path(path, "open workbook")
    if not os.path.exists(p):
        raise WorkbookNotFound(f"no such workbook: {p}")
    # Modern .xls files embed a zip fragment, so the zip open below can
    # SUCCEED on a BIFF file and misread the theme package as the workbook;
    # the OLE sniff refuses first with the format and the remedy
    # (geriatric round, H-1).
    _hazard.refuse_ole_container(p)
    try:
        return zipfile.ZipFile(p)
    except zipfile.BadZipFile:
        raise WorkbookCorrupt(f"{p}: not a valid zip / OOXML package")


# --------------------------------------------------------- external links


def _local(tag) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def get_external_links(path: str) -> dict:
    """List external workbook links and their cached values. Read-only."""
    with _open_zip(path) as zf:
        names = set(zf.namelist())
        link_parts = sorted(
            n for n in names
            if n.lower().startswith("xl/externallinks/")
            and n.lower().endswith(".xml")
            and "/_rels/" not in n.lower())
        links: list[dict] = []
        for part in link_parts:
            head, base = part.rsplit("/", 1)
            rels_part = f"{head}/_rels/{base}.rels"
            target = None
            if rels_part in names:
                try:
                    rroot = fromstring(zf.read(rels_part))
                    for rel in rroot.iter(f"{_PKG_REL_NS}Relationship"):
                        target = rel.get("Target")
                        break
                except Exception:  # noqa: BLE001
                    target = None
            entry: dict[str, Any] = {"part": part, "target": target}
            try:
                root = fromstring(zf.read(part))
            except Exception:  # noqa: BLE001
                entry["note"] = "link part XML could not be parsed"
                links.append(entry)
                continue
            book = next((el for el in root
                         if _local(el.tag) == "externalBook"), None)
            if book is None:
                kinds = sorted({_local(el.tag) for el in root})
                entry["kind"] = kinds[0] if kinds else "unknown"
                entry["note"] = (
                    "not an externalBook link (ddeLink/oleLink); cached "
                    "values are not parsed for this kind")
                links.append(entry)
                continue
            entry["kind"] = "externalBook"
            sheets = [sn.get("val") for sn in book.iter()
                      if _local(sn.tag) == "sheetName"]
            entry["sheets"] = sheets
            cached: list[dict] = []
            total = 0
            for sd in book.iter():
                if _local(sd.tag) != "sheetData":
                    continue
                sheet_id = sd.get("sheetId")
                sheet_name = None
                try:
                    sheet_name = sheets[int(sheet_id)]
                except (TypeError, ValueError, IndexError):
                    pass
                for cell in sd.iter():
                    if _local(cell.tag) != "cell":
                        continue
                    total += 1
                    if len(cached) >= _MAX_CACHED_CELLS_PER_LINK:
                        continue
                    v = next((c.text for c in cell
                              if _local(c.tag) == "v"), None)
                    cached.append({
                        "sheet": sheet_name, "cell": cell.get("r"),
                        "type": cell.get("t", "n"), "value": v})
            entry["cached_values"] = cached
            entry["cached_value_count"] = total
            if total > len(cached):
                entry["truncated"] = True
            links.append(entry)
    return {
        "links": links, "count": len(links),
        "sensitive": bool(links),
        "note": ("cached values are copies of data from the linked "
                 "workbooks as of the last refresh; treat as sensitive"
                 if links else "no external links"),
    }


# ------------------------------------------------------------ VBA (CFB)

# A compact MS-CFB (compound file) reader, enough to walk the directory and
# read streams, plus the MS-OVBA 2.4.1 decompressor for the dir stream.
# Both are exercised by spec-constructed fixtures in the test suite. Any
# parse failure anywhere degrades to the honest presence-only report.

_CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_ENDOFCHAIN = 0xFFFFFFFE
_FREESECT = 0xFFFFFFFF


class _Cfb:
    def __init__(self, data: bytes):
        if data[:8] != _CFB_MAGIC:
            raise ValueError("not a CFB container")
        self.data = data
        (self.sector_shift,) = struct.unpack_from("<H", data, 30)
        (self.mini_shift,) = struct.unpack_from("<H", data, 32)
        (self.num_fat,) = struct.unpack_from("<I", data, 44)
        (self.dir_start,) = struct.unpack_from("<I", data, 48)
        (self.mini_cutoff,) = struct.unpack_from("<I", data, 56)
        (self.minifat_start,) = struct.unpack_from("<I", data, 60)
        (self.num_minifat,) = struct.unpack_from("<I", data, 64)
        (self.difat_start,) = struct.unpack_from("<I", data, 68)
        (self.num_difat,) = struct.unpack_from("<I", data, 72)
        self.sector_size = 1 << self.sector_shift
        self.mini_size = 1 << self.mini_shift
        self._load_fat()
        self._load_dir()
        self._load_minifat()

    def _sector(self, n: int) -> bytes:
        off = 512 + n * self.sector_size
        return self.data[off:off + self.sector_size]

    def _load_fat(self) -> None:
        difat: list[int] = list(struct.unpack_from("<109I", self.data, 76))
        nxt = self.difat_start
        per = self.sector_size // 4 - 1
        for _ in range(self.num_difat):
            if nxt in (_ENDOFCHAIN, _FREESECT):
                break
            raw = struct.unpack(f"<{per + 1}I", self._sector(nxt))
            difat.extend(raw[:per])
            nxt = raw[per]
        self.fat: list[int] = []
        per_fat = self.sector_size // 4
        for s in difat:
            if s in (_ENDOFCHAIN, _FREESECT):
                continue
            self.fat.extend(struct.unpack(f"<{per_fat}I", self._sector(s)))

    def _chain(self, start: int, fat: list[int]) -> list[int]:
        out, seen, cur = [], set(), start
        while cur not in (_ENDOFCHAIN, _FREESECT) and cur < len(fat):
            if cur in seen:
                raise ValueError("cyclic sector chain")
            seen.add(cur)
            out.append(cur)
            cur = fat[cur]
        return out

    def _read_chain(self, start: int, size: int | None = None) -> bytes:
        buf = b"".join(self._sector(s) for s in self._chain(start, self.fat))
        return buf if size is None else buf[:size]

    def _load_dir(self) -> None:
        raw = self._read_chain(self.dir_start)
        self.entries: list[dict] = []
        for off in range(0, len(raw) - 127, 128):
            ent = raw[off:off + 128]
            (name_len,) = struct.unpack_from("<H", ent, 64)
            obj_type = ent[66]
            if obj_type == 0 or name_len < 2:
                continue
            name = ent[:name_len - 2].decode("utf-16-le", "replace")
            (start,) = struct.unpack_from("<I", ent, 116)
            (size,) = struct.unpack_from("<Q", ent, 120)
            self.entries.append({"name": name, "type": obj_type,
                                 "start": start, "size": size})

    def _load_minifat(self) -> None:
        self.minifat: list[int] = []
        per = self.sector_size // 4
        cur = self.minifat_start
        for _ in range(self.num_minifat):
            if cur in (_ENDOFCHAIN, _FREESECT):
                break
            self.minifat.extend(
                struct.unpack(f"<{per}I", self._sector(cur)))
            cur = self.fat[cur] if cur < len(self.fat) else _ENDOFCHAIN
        root = next((e for e in self.entries if e["type"] == 5), None)
        self.ministream = (
            self._read_chain(root["start"], root["size"])
            if root is not None else b"")

    def read_stream(self, name: str) -> bytes | None:
        ent = next((e for e in self.entries
                    if e["type"] == 2 and e["name"] == name), None)
        if ent is None:
            return None
        if ent["size"] >= self.mini_cutoff:
            return self._read_chain(ent["start"], ent["size"])
        out, cur = b"", ent["start"]
        seen: set[int] = set()
        while cur not in (_ENDOFCHAIN, _FREESECT) and cur < len(self.minifat):
            if cur in seen:
                raise ValueError("cyclic mini chain")
            seen.add(cur)
            off = cur * self.mini_size
            out += self.ministream[off:off + self.mini_size]
            cur = self.minifat[cur]
        return out[:ent["size"]]


def _ovba_decompress(data: bytes) -> bytes:
    """MS-OVBA 2.4.1 decompression of a compressed container."""
    if not data or data[0] != 0x01:
        raise ValueError("not an MS-OVBA compressed container")
    out = bytearray()
    pos = 1
    while pos + 2 <= len(data):
        (hdr,) = struct.unpack_from("<H", data, pos)
        pos += 2
        chunk_len = (hdr & 0x0FFF) + 3
        compressed = bool(hdr & 0x8000)
        chunk_end = pos + chunk_len - 2
        chunk_start_out = len(out)
        if not compressed:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        while pos < chunk_end and pos < len(data):
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end or pos >= len(data):
                    break
                if not (flags >> bit) & 1:
                    out.append(data[pos])
                    pos += 1
                    continue
                (token,) = struct.unpack_from("<H", data, pos)
                pos += 2
                diff = len(out) - chunk_start_out
                bits = max(4, (diff - 1).bit_length() if diff > 1 else 4)
                bits = min(bits, 12)
                length = (token & ((1 << (16 - bits)) - 1)) + 3
                offset = (token >> (16 - bits)) + 1
                src = len(out) - offset
                if src < chunk_start_out:
                    raise ValueError("copy token reaches before the chunk")
                for _ in range(length):
                    out.append(out[src])
                    src += 1
    return bytes(out)


#: Fallback for MODULENAME bytes when the project declares no code page, or
#: declares one this interpreter has no codec for. Windows-1252 is the
#: overwhelmingly common case and decodes every byte, so a name is never lost
#: to a decode error; a wrong-but-present name still beats no modules at all.
_OVBA_FALLBACK_CODEPAGE = "cp1252"


def _ovba_codec(codepage: int | None) -> str:
    """The codec for MODULENAME bytes, from the project's own declaration.

    MS-OVBA stores module names in the code page the PROJECTCODEPAGE record
    (0x0003) names, NOT in the code page of whatever machine happens to be
    reading the file. This used to decode with `mbcs`, which is an alias for
    the host's Windows ANSI code page: wrong whenever the two differ, and on
    Linux not a codec at all, so every real VBA project degraded to "could
    not be extracted" on any non-Windows host.
    """
    if codepage:
        candidate = f"cp{int(codepage)}"
        try:
            codecs.lookup(candidate)
        except LookupError:
            pass
        else:
            return candidate
    return _OVBA_FALLBACK_CODEPAGE


def _parse_dir_stream(data: bytes) -> list[dict]:
    """Extract module records from a decompressed dir stream (MS-OVBA
    2.3.4.2): a TLV walk collecting PROJECTCODEPAGE (0x0003), MODULENAME
    (0x0019), MODULESTREAMNAME (0x001A), and the module type ids, honoring
    the PROJECTVERSION (0x0009) fixed-size quirk."""
    modules: list[dict] = []
    current: dict | None = None
    codec = _OVBA_FALLBACK_CODEPAGE
    pos = 0
    while pos + 6 <= len(data):
        (rec_id, size) = struct.unpack_from("<HI", data, pos)
        pos += 6
        if rec_id == 0x0009:  # PROJECTVERSION: size field is Reserved
            size = 6
        payload = data[pos:pos + size]
        pos += size
        if rec_id == 0x0003 and len(payload) >= 2:  # PROJECTCODEPAGE
            # Declared before any module record, so one forward pass is
            # enough and every name below is decoded with the right codec.
            (declared,) = struct.unpack_from("<H", payload, 0)
            codec = _ovba_codec(declared)
        elif rec_id == 0x0019:  # MODULENAME
            current = {"name": payload.decode(codec, "replace")}
            modules.append(current)
        elif rec_id == 0x001A and current is not None:  # MODULESTREAMNAME
            current["stream"] = payload.decode(codec, "replace")
        elif rec_id == 0x0021 and current is not None:
            current["type"] = "procedural"
        elif rec_id == 0x0022 and current is not None:
            current["type"] = "document_class_or_designer"
        elif rec_id == 0x0010 and current is None:
            break  # PROJECT terminator before any module
    return modules


def inspect_vba(path: str) -> dict:
    """Presence, size, and (when parseable) module names of the VBA
    project. Strictly read-only."""
    with _open_zip(path) as zf:
        names = zf.namelist()
        vba_parts = [n for n in names
                     if n.rsplit("/", 1)[-1].lower() == "vbaproject.bin"]
        if not vba_parts:
            return {"has_vba": False, "note": "no vbaProject.bin part"}
        part = vba_parts[0]
        data = zf.read(part)
        declared = False
        try:
            ct = zf.read("[Content_Types].xml").decode("utf-8", "replace")
            declared = "vnd.ms-office.vbaProject" in ct
        except KeyError:
            pass
        out: dict[str, Any] = {
            "has_vba": True, "part": part, "size": len(data),
            "declared_content_type": declared,
            "extension_matches": path.lower().endswith(
                (".xlsm", ".xltm", ".xlam")),
        }
        if not out["extension_matches"]:
            out["warning"] = (
                "vbaProject.bin inside a non-macro extension; a file-based "
                "save cannot preserve it (keep_vba binds to .xlsm)")
        try:
            cfb = _Cfb(data)
            dir_stream = cfb.read_stream("dir")
            if dir_stream is None:
                raise ValueError("no dir stream in the VBA project")
            modules = _parse_dir_stream(_ovba_decompress(dir_stream))
            sizes = {e["name"]: e["size"] for e in cfb.entries
                     if e["type"] == 2}
            for m in modules:
                m["stream_size"] = sizes.get(m.get("stream", m["name"]))
            out["modules"] = modules
            out["module_count"] = len(modules)
        except Exception as exc:  # noqa: BLE001
            out["modules"] = None
            out["note"] = (
                "module names could not be extracted (the container did "
                f"not parse as a well-formed VBA project: {exc}); presence "
                "and size above are still authoritative")
        return out


# ----------------------------------------------------------------- pivots


def _excel_serial_to_iso(serial) -> str | None:
    try:
        from openpyxl.utils.datetime import from_excel
        return from_excel(float(serial)).isoformat()
    except Exception:  # noqa: BLE001
        return None


def get_pivot(path: str, sheet: str | None = None) -> dict:
    """Describe existing pivot tables (openpyxl read model). Read-only."""
    wb = gridio.open_wb(path)
    try:
        if sheet is not None and sheet not in wb.sheetnames:
            raise XlMcpError(
                f"no sheet named {sheet!r}; sheets: {wb.sheetnames}")
        pivots: list[dict] = []
        sheets = wb.worksheets if sheet is None else [wb[sheet]]
        for ws in sheets:
            for pv in getattr(ws, "_pivots", []):
                cache = pv.cache
                field_names = [f.name for f in cache.cacheFields] \
                    if cache is not None else []

                def _names(fields):
                    out = []
                    for f in fields:
                        idx = getattr(f, "x", getattr(f, "fld", None))
                        if idx is None:
                            continue
                        if idx == -2:
                            out.append("(values)")
                        elif 0 <= idx < len(field_names):
                            out.append(field_names[idx])
                        else:
                            out.append(f"field#{idx}")
                    return out

                src = None
                if cache is not None and cache.cacheSource is not None:
                    wsrc = cache.cacheSource.worksheetSource
                    if wsrc is not None:
                        src = {"sheet": wsrc.sheet, "range": wsrc.ref,
                               "named_range": wsrc.name}
                pivots.append({
                    "sheet": ws.title,
                    "name": pv.name,
                    "location": pv.location.ref if pv.location else None,
                    "source": src,
                    "fields": field_names,
                    "row_fields": _names(pv.rowFields),
                    "column_fields": _names(pv.colFields),
                    "data_fields": [d.name for d in pv.dataFields],
                    "page_fields": _names(pv.pageFields),
                    "record_count": getattr(cache, "recordCount", None),
                    "refreshed_by": getattr(cache, "refreshedBy", None),
                    "refreshed_at": _excel_serial_to_iso(
                        getattr(cache, "refreshedDate", None)),
                    "refresh_on_load": bool(
                        getattr(cache, "refreshOnLoad", False)),
                })
        return {
            "pivots": pivots, "count": len(pivots),
            "note": ("description only: pivot creation, modification, and "
                     "refresh need the Excel application "
                     "(com_manage_pivot, com pack)"
                     if pivots else "no pivot tables"),
        }
    finally:
        wb.close()


# ------------------------------------------------------------ connections

_CONNECTION_TYPES = {
    "1": "odbc", "2": "dao", "3": "file", "4": "web_query", "5": "oledb",
    "6": "text", "7": "ado", "8": "dsp",
}


def get_connections(path: str) -> dict:
    """List data connections and Power Query presence. Read-only."""
    with _open_zip(path) as zf:
        names = set(zf.namelist())
        conns: list[dict] = []
        if "xl/connections.xml" in names:
            try:
                root = fromstring(zf.read("xl/connections.xml"))
            except Exception:  # noqa: BLE001
                root = None
            if root is not None:
                for conn in root.iter():
                    if _local(conn.tag) != "connection":
                        continue
                    entry: dict[str, Any] = {
                        "name": conn.get("name"),
                        "description": conn.get("description"),
                        "type_code": conn.get("type"),
                        "type": _CONNECTION_TYPES.get(
                            conn.get("type") or "", "unknown"),
                        "refresh_on_load": conn.get("refreshOnLoad") == "1",
                    }
                    for child in conn:
                        local = _local(child.tag)
                        if local == "dbPr":
                            entry["connection_string"] = child.get(
                                "connection")
                            entry["command"] = child.get("command")
                        elif local == "webPr":
                            entry["url"] = child.get("url")
                        elif local == "textPr":
                            entry["source_file"] = child.get("sourceFile")
                    conns.append(entry)
        pq_evidence: list[str] = []
        for n in sorted(names):
            low = n.lower()
            if low.startswith("customxml/") and low.endswith(".xml"):
                try:
                    head = zf.read(n)[:4096]
                except Exception:  # noqa: BLE001
                    continue
                if b"DataMashup" in head:
                    pq_evidence.append(f"{n} holds a DataMashup item")
        for c in conns:
            cs = (c.get("connection_string") or "")
            if "Microsoft.Mashup" in cs:
                pq_evidence.append(
                    f"connection {c['name']!r} uses the Mashup provider")
        query_tables = sorted(
            n for n in names if n.lower().startswith("xl/querytables/"))
    return {
        "connections": conns, "count": len(conns),
        "power_query": {"present": bool(pq_evidence),
                        "evidence": pq_evidence},
        "legacy_query_tables": query_tables,
        "sensitive": bool(conns or pq_evidence),
        "note": ("connection strings can embed server names and "
                 "credentials; treat as sensitive. Refresh needs the Excel "
                 "application (com pack)" if conns or pq_evidence
                 else "no data connections"),
    }


__all__ = ["get_external_links", "inspect_vba", "get_pivot",
           "get_connections"]

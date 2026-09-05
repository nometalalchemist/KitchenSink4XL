"""Read-side inspector tests: get_external_links, inspect_vba, get_pivot,
get_connections. All four are read-only, so every test also gets implicit
no-mutation coverage (the fixtures are opened fresh, never saved).

The inspect_vba tests validate the compact MS-CFB reader and the MS-OVBA
decompressor against SPEC-CONSTRUCTED fixtures built byte-by-byte in this
file (a v3 compound file holding a compressed dir stream with two
modules), plus hand-built compression vectors for the literal-token,
copy-token, and uncompressed-chunk paths. The corpus macro.xlsm carries a
synthetic non-CFB blob, which pins the honest presence-only fallback.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import struct
import zipfile

import pytest

from xlsx_mcp.core.errors import WorkbookNotFound, XlMcpError
from xlsx_mcp.ops import inspectors as _inspectors
from xlsx_mcp.ops import lifecycle as _lifecycle
from xlsx_mcp.ops.inspectors import _Cfb, _ovba_decompress

CORPUS = os.path.join(os.path.dirname(__file__), "..", "fixtures", "corpus")

END = 0xFFFFFFFE
FREE = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
_ODR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _md5(path):
    with open(path, "rb") as fh:
        return hashlib.md5(fh.read()).hexdigest()


# ------------------------------------------------- CFB fixture construction


def _ovba_compress_literals(data: bytes) -> bytes:
    """A spec-valid MS-OVBA container using literal tokens only (all-zero
    flag bytes)."""
    assert len(data) <= 4096
    chunk = bytearray()
    for i in range(0, len(data), 8):
        chunk.append(0x00)
        chunk += data[i:i + 8]
    hdr = ((len(chunk) + 2 - 3) & 0x0FFF) | 0xB000
    return b"\x01" + struct.pack("<H", hdr) + bytes(chunk)


def _dir_entry(name, etype, start, size, child=FREE):
    raw = name.encode("utf-16-le") + b"\x00\x00"
    ent = bytearray(128)
    ent[:len(raw)] = raw
    struct.pack_into("<H", ent, 64, len(raw))
    ent[66] = etype
    ent[67] = 1
    struct.pack_into("<I", ent, 68, FREE)
    struct.pack_into("<I", ent, 72, FREE)
    struct.pack_into("<I", ent, 76, child)
    struct.pack_into("<I", ent, 116, start)
    struct.pack_into("<Q", ent, 120, size)
    return bytes(ent)


def _tlv(rec_id, payload):
    return struct.pack("<HI", rec_id, len(payload)) + payload


def _build_dir_stream() -> bytes:
    out = b""
    out += _tlv(0x0001, struct.pack("<I", 0x409))
    # PROJECTVERSION quirk: the size field is Reserved(4), data is 6 bytes
    out += struct.pack("<HI", 0x0009, 4) + struct.pack("<IH", 3, 1)
    out += _tlv(0x000F, struct.pack("<H", 2))
    for name in ("Module1", "ThisWorkbook"):
        out += _tlv(0x0019, name.encode("ascii"))
        out += _tlv(0x001A, name.encode("ascii"))
        out += _tlv(0x0031, struct.pack("<I", 0))
        out += _tlv(0x0021 if name == "Module1" else 0x0022, b"")
        out += _tlv(0x002B, b"")
    out += _tlv(0x0010, b"")
    return out


MOD1_SRC = b"' Module1 source placeholder\r\nSub Demo()\r\nEnd Sub\r\n"


def build_vba_bin() -> bytes:
    """A minimal, spec-conformant v3 compound file: FAT sector, directory
    sector, one ministream sector holding dir + two module streams, one
    miniFAT sector."""
    dir_stream = _ovba_compress_literals(_build_dir_stream())
    mod2 = b"' ThisWorkbook placeholder\r\n"

    def minis(data):
        n = (len(data) + 63) // 64
        return data + b"\x00" * (n * 64 - len(data)), n

    d_pad, d_n = minis(dir_stream)
    m1_pad, m1_n = minis(MOD1_SRC)
    m2_pad, m2_n = minis(mod2)
    ministream = d_pad + m1_pad + m2_pad
    assert len(ministream) <= 512
    ministream += b"\x00" * (512 - len(ministream))

    minifat, idx = [], 0
    for n in (d_n, m1_n, m2_n):
        minifat.extend(range(idx + 1, idx + n))
        minifat.append(END)
        idx += n
    minifat += [FREE] * (128 - len(minifat))

    entries = b"".join([
        _dir_entry("Root Entry", 5, 2, 512, child=1),
        _dir_entry("dir", 2, 0, len(dir_stream)),
        _dir_entry("Module1", 2, d_n, len(MOD1_SRC)),
        _dir_entry("ThisWorkbook", 2, d_n + m1_n, len(mod2)),
    ])
    fat = [FATSECT, END, END, END] + [FREE] * 124

    header = bytearray(512)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<H", header, 24, 0x003E)
    struct.pack_into("<H", header, 26, 0x0003)
    struct.pack_into("<H", header, 28, 0xFFFE)
    struct.pack_into("<H", header, 30, 9)
    struct.pack_into("<H", header, 32, 6)
    struct.pack_into("<I", header, 44, 1)
    struct.pack_into("<I", header, 48, 1)
    struct.pack_into("<I", header, 56, 4096)
    struct.pack_into("<I", header, 60, 3)
    struct.pack_into("<I", header, 64, 1)
    struct.pack_into("<I", header, 68, END)
    struct.pack_into("<I", header, 72, 0)
    struct.pack_into("<I", header, 76, 0)
    for i in range(1, 109):
        struct.pack_into("<I", header, 76 + 4 * i, FREE)

    return (bytes(header) + struct.pack("<128I", *fat) + entries
            + ministream + struct.pack("<128I", *minifat))


# ------------------------------------------------------- MS-OVBA unit paths


def test_ovba_literal_tokens_round_trip():
    data = _build_dir_stream()
    assert _ovba_decompress(_ovba_compress_literals(data)) == data


def test_ovba_uncompressed_chunk():
    raw = b"A" * 4096
    hdr = ((4098 - 3) & 0x0FFF) | 0x3000  # flag bit 15 clear: uncompressed
    blob = b"\x01" + struct.pack("<H", hdr) + raw
    assert _ovba_decompress(blob) == raw


def test_ovba_copy_token():
    """Hand-built per MS-OVBA 2.4.1: literal 'a', then a copy token with
    offset 1, length 10 (diff=1 -> 4 offset bits; token 0x0007)."""
    chunk = bytes([0b00000010, ord("a")]) + struct.pack("<H", 0x0007)
    hdr = ((len(chunk) + 2 - 3) & 0x0FFF) | 0xB000
    blob = b"\x01" + struct.pack("<H", hdr) + chunk
    assert _ovba_decompress(blob) == b"a" * 11


def test_ovba_rejects_non_container():
    with pytest.raises(ValueError):
        _ovba_decompress(b"\x02junk")


def test_cfb_reader_walks_spec_fixture():
    blob = build_vba_bin()
    cfb = _Cfb(blob)
    names = {e["name"] for e in cfb.entries if e["type"] == 2}
    assert names == {"dir", "Module1", "ThisWorkbook"}
    assert cfb.read_stream("Module1") == MOD1_SRC
    assert cfb.read_stream("missing") is None
    with pytest.raises(ValueError):
        _Cfb(b"not a cfb at all" * 40)


# -------------------------------------------------------------- inspect_vba


def _xlsm_with_vba(tmp_path, blob):
    src = os.path.join(CORPUS, "macro.xlsm")
    dst = str(tmp_path / "real.xlsm")
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/vbaProject.bin":
                data = blob
            zout.writestr(item, data)
    return dst


def test_inspect_vba_extracts_modules_from_cfb(tmp_path):
    p = _xlsm_with_vba(tmp_path, build_vba_bin())
    out = _inspectors.inspect_vba(p)
    assert out["has_vba"] is True
    assert out["extension_matches"] is True
    # When the CFB parse degrades, inspect_vba answers honestly with a note
    # and no module keys. Assert on the note first, so a failure here says
    # WHY the container would not parse instead of a bare KeyError.
    assert out.get("modules") is not None, (
        f"module extraction degraded: {out.get('note')!r} "
        f"(size={out.get('size')}, part={out.get('part')!r})")
    assert out["module_count"] == 2
    mods = {m["name"]: m for m in out["modules"]}
    assert mods["Module1"]["type"] == "procedural"
    assert mods["Module1"]["stream_size"] == len(MOD1_SRC)
    assert mods["ThisWorkbook"]["type"] == "document_class_or_designer"


def test_inspect_vba_honest_fallback_on_unparseable_blob():
    """macro.xlsm carries a synthetic non-CFB blob: presence and size stay
    authoritative, modules degrade to None with a note, no guessing."""
    out = _inspectors.inspect_vba(os.path.join(CORPUS, "macro.xlsm"))
    assert out["has_vba"] is True
    assert out["size"] > 0
    assert out["modules"] is None
    assert "could not be extracted" in out["note"]


def test_inspect_vba_absent(tmp_path):
    p = str(tmp_path / "n.xlsx")
    _lifecycle.create_workbook(p)
    out = _inspectors.inspect_vba(p)
    assert out["has_vba"] is False


def test_inspect_vba_warns_on_wrong_extension(tmp_path):
    src = _xlsm_with_vba(tmp_path, build_vba_bin())
    dst = str(tmp_path / "disguised.xlsx")
    shutil.copy(src, dst)
    out = _inspectors.inspect_vba(dst)
    assert out["extension_matches"] is False
    assert "cannot preserve" in out["warning"]


# ----------------------------------------------------------- external links


def _inject(path, members: dict):
    with zipfile.ZipFile(path, "a") as zf:
        for name, data in members.items():
            zf.writestr(name, data)


EXT_LINK = f"""<?xml version="1.0"?>
<externalLink xmlns="{_MAIN}" xmlns:r="{_ODR}"><externalBook r:id="rId1">
<sheetNames><sheetName val="Prices"/></sheetNames>
<sheetDataSet><sheetData sheetId="0"><row r="1">
<cell r="A1" t="n"><v>42.5</v></cell>
<cell r="B1" t="str"><v>widget</v></cell>
</row></sheetData></sheetDataSet></externalBook></externalLink>"""

EXT_RELS = f"""<?xml version="1.0"?>
<Relationships xmlns="{_RELS}">
<Relationship Id="rId1" Type="{_ODR}/externalLinkPath"
 Target="file:///C:/data/prices.xlsx" TargetMode="External"/>
</Relationships>"""


def test_external_links_parsed(tmp_path):
    p = str(tmp_path / "x.xlsx")
    shutil.copy(os.path.join(CORPUS, "clean.xlsx"), p)
    _inject(p, {
        "xl/externalLinks/externalLink1.xml": EXT_LINK,
        "xl/externalLinks/_rels/externalLink1.xml.rels": EXT_RELS,
    })
    before = _md5(p)
    out = _inspectors.get_external_links(p)
    assert _md5(p) == before  # read-only
    assert out["count"] == 1 and out["sensitive"] is True
    link = out["links"][0]
    assert link["target"] == "file:///C:/data/prices.xlsx"
    assert link["sheets"] == ["Prices"]
    assert link["cached_value_count"] == 2
    cells = {c["cell"]: c for c in link["cached_values"]}
    assert cells["A1"]["value"] == "42.5" and cells["A1"]["sheet"] == "Prices"
    assert cells["B1"]["type"] == "str"


def test_external_links_dde_kind_noted(tmp_path):
    p = str(tmp_path / "d.xlsx")
    shutil.copy(os.path.join(CORPUS, "clean.xlsx"), p)
    _inject(p, {"xl/externalLinks/externalLink1.xml":
                f'<externalLink xmlns="{_MAIN}"><ddeLink ddeService="x" '
                'ddeTopic="y"/></externalLink>'})
    out = _inspectors.get_external_links(p)
    assert out["links"][0]["kind"] == "ddeLink"
    assert "cached values are not parsed" in out["links"][0]["note"]


def test_external_links_malformed_part_degrades(tmp_path):
    p = str(tmp_path / "m.xlsx")
    shutil.copy(os.path.join(CORPUS, "clean.xlsx"), p)
    _inject(p, {"xl/externalLinks/externalLink1.xml": "<broken"})
    out = _inspectors.get_external_links(p)
    assert out["links"][0]["note"] == "link part XML could not be parsed"


def test_external_links_none(tmp_path):
    p = str(tmp_path / "c.xlsx")
    _lifecycle.create_workbook(p)
    out = _inspectors.get_external_links(p)
    assert out["count"] == 0 and out["sensitive"] is False


def test_external_links_missing_workbook(tmp_path):
    with pytest.raises(WorkbookNotFound):
        _inspectors.get_external_links(str(tmp_path / "no.xlsx"))


# ------------------------------------------------------------------- pivots


def test_get_pivot_describes_fixture():
    p = os.path.join(CORPUS, "pivot.xlsx")
    before = _md5(p)
    out = _inspectors.get_pivot(p)
    assert _md5(p) == before  # read-only
    assert out["count"] == 1
    pv = out["pivots"][0]
    assert pv["name"] == "PivotTable1"
    assert pv["source"] == {"sheet": "Src", "range": "A1:C7",
                            "named_range": None}
    assert pv["fields"] == ["Category", "Region", "Amount"]
    assert pv["row_fields"] == ["Category"]
    assert pv["data_fields"] == ["Sum of Amount"]
    assert pv["record_count"] == 6
    assert pv["refreshed_by"] == "Zeta"
    assert pv["refreshed_at"]  # serial converted to ISO
    assert "com pack" in out["note"]


def test_get_pivot_sheet_filter_and_refusal(tmp_path):
    p = os.path.join(CORPUS, "pivot.xlsx")
    out = _inspectors.get_pivot(p, sheet="Src")
    assert out["count"] == 0
    with pytest.raises(XlMcpError, match="no sheet named"):
        _inspectors.get_pivot(p, sheet="Nope")
    q = str(tmp_path / "c.xlsx")
    _lifecycle.create_workbook(q)
    assert _inspectors.get_pivot(q)["count"] == 0


# -------------------------------------------------------------- connections


CONN_XML = f"""<?xml version="1.0"?>
<connections xmlns="{_MAIN}">
<connection id="1" name="Query - Sales" description="pq" type="5"
 refreshOnLoad="1"><dbPr
 connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$"
 command="Sales"/></connection>
<connection id="2" name="WebQ" type="4">
<webPr url="https://example.com/data"/></connection>
</connections>"""


def test_get_connections_and_power_query(tmp_path):
    p = str(tmp_path / "c.xlsx")
    shutil.copy(os.path.join(CORPUS, "clean.xlsx"), p)
    _inject(p, {"xl/connections.xml": CONN_XML,
                "customXml/item1.xml": "<x>DataMashup payload</x>"})
    out = _inspectors.get_connections(p)
    assert out["count"] == 2 and out["sensitive"] is True
    by_name = {c["name"]: c for c in out["connections"]}
    assert by_name["Query - Sales"]["type"] == "oledb"
    assert by_name["Query - Sales"]["refresh_on_load"] is True
    assert "Mashup" in by_name["Query - Sales"]["connection_string"]
    assert by_name["WebQ"]["url"] == "https://example.com/data"
    assert out["power_query"]["present"] is True
    assert len(out["power_query"]["evidence"]) == 2


def test_get_connections_absent(tmp_path):
    p = str(tmp_path / "n.xlsx")
    _lifecycle.create_workbook(p)
    out = _inspectors.get_connections(p)
    assert out["count"] == 0
    assert out["power_query"]["present"] is False
    assert out["sensitive"] is False

"""Generate bundle/icon.png for the KitchenSink4XL .mcpb bundle.

Pure stdlib (zlib + struct): 512x512 RGB PNG (Claude Desktop's recommended
icon size), brass square with a white lightning bolt, rendered with 4x4
supersampling for smooth edges. Brass and the bolt are the Electrical aisle's
own colors, so the extension tile matches the product page.

Deterministic: same input, same bytes, so a rebuild produces no spurious diff.

Usage: python -X utf8 make_icon.py
"""

import struct
import zlib
from pathlib import Path

SIZE = 512          # output pixels
SS = 4              # supersample factor
BG = (125, 92, 20)      # brass  #7d5c14
FG = (247, 251, 253)    # near-white bolt

# Bolt outline in normalized (0..1) coordinates, clockwise from the top.
BOLT = [
    (0.585, 0.085),
    (0.255, 0.545),
    (0.450, 0.545),
    (0.400, 0.915),
    (0.745, 0.435),
    (0.545, 0.435),
    (0.630, 0.085),
]
POLY = [(x * SIZE, y * SIZE) for x, y in BOLT]


def inside_bolt(x: float, y: float) -> bool:
    """Even-odd point-in-polygon test."""
    inside = False
    n = len(POLY)
    j = n - 1
    for i in range(n):
        xi, yi = POLY[i]
        xj, yj = POLY[j]
        if (yi > y) != (yj > y):
            xc = xi + (y - yi) * (xj - xi) / (yj - yi)
            if x < xc:
                inside = not inside
        j = i
    return inside


def render() -> bytes:
    scale = 1.0 / SS
    rows = []
    for oy in range(SIZE):
        row = bytearray()
        row.append(0)  # PNG filter type 0 (None)
        for ox in range(SIZE):
            hit = 0
            for sy in range(SS):
                for sx in range(SS):
                    x = (ox * SS + sx + 0.5) * scale
                    y = (oy * SS + sy + 0.5) * scale
                    if inside_bolt(x, y):
                        hit += 1
            cov = hit / (SS * SS)
            row.extend(
                round(BG[i] + (FG[i] - BG[i]) * cov) for i in range(3)
            )
        rows.append(bytes(row))
    return b"".join(rows)


def write_png(path: Path, raw: bytes) -> None:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    ihdr = struct.pack(">IIBBBBB", SIZE, SIZE, 8, 2, 0, 0, 0)  # 8-bit RGB
    png = (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(raw, 9))
        + chunk(b"IEND", b"")
    )
    path.write_bytes(png)


if __name__ == "__main__":
    out = Path(__file__).parent / "icon.png"
    write_png(out, render())
    print(f"wrote {out} ({out.stat().st_size} bytes)")

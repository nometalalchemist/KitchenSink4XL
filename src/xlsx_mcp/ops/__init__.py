"""ops/ for KS4XL: the engine families.

Each family module multiplexes its domain XML logic over openpyxl / raw
OOXML / xlsxwriter (DESIGN Section 2.2, Section 11). New code, but mechanical
once package + locate + refs + hazard + verify exist. Landed in the Phase 3
waves (cells, formulas, formatting, cond-format, validation, tables, names,
sort/filter, view, objects, comments, hyperlinks, protection, io, pagelayout,
powerquery, vba) and the Phase 4 view/batch layer.

Empty in Phase 0; server.py registers only the tiered-loading toggles and a
placeholder reader until the families land.
"""

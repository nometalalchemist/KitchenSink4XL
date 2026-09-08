## KitchenSink4XL Quickstart

**Install.** Claude Desktop: the KitchenSink4XL extension; the "Verify every
save with Excel" checkbox is the strongest guarantee if Excel is installed.
Anywhere else: `uvx kitchensink4xl` (add `[com]` for the Excel-powered
extras).

**First sheet.** Ask Claude: "Read the Q3 tab of budget.xlsx and total the
travel column." Every value it reports is labeled calculated, cached, or
missing, so a stale number is never repeated as a fresh one. Writes land
after a backup and are verified after saving.

**The habits that matter.** Ask for ranges, not whole workbooks; the server
reads what you name. A workbook open in your Excel is refused rather than
risked; close it first. `get_server_info` reports health, version, and the
weekly update check (off with KS4XL_UPDATE_CHECK=off). The optional folder
limit on the install screen scopes the server to one directory if you want a
hard boundary.

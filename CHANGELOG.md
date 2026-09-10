# Changelog

### 1.2.1
- The server says what it is. A connected agent now reads "KitchenSink4XL (kitchensink4xl on PyPI), part of the KitchenSink4AI suite" at the head of the instructions it receives, and `get_server_info` reports the product name, the package to install, the documentation homepage, and the three sibling packages.
- Published addresses moved to kitchensink4.ai, and the registry record is republished at this version.
- The tool figure is stated as measured on every surface: 129 workbook operations across 69 tools, 67 workbook tools plus the two pack toggles.

### 1.2.0
- Read responses are dramatically smaller (seven to eleven times on the heavy calls) with nothing lost; the token estimator now matches the wire exactly, and every published figure is re-measured.
- A performance bug that reopened the workbook per cell is fixed; large cell reads are near-instant.
- Update notice: a weekly, disclosed check for newer releases, off with KS4XL_UPDATE_CHECK=off.
- The install screen and info card are rewritten in plain language.

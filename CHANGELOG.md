# Changelog

### 1.2.2
- The license story is stated correctly and in one place. KitchenSink4XL is dual-licensed: AGPL-3.0 for anyone whose use meets the AGPL's terms, including its source-sharing obligations, and a commercial license for organizations that need to ship it inside a closed product (licensing@kitchensink4.ai). NOTICE.md previously said "free for any use (personal, academic, commercial) under AGPL terms", which contradicted the commercial model.
- Copyright is attributed: Alvut Consulting, LLC, named at the top of LICENSE, in NOTICE.md, in the README, in the package author field, and as the grantee in the CLA.
- The landing page serves its fonts itself instead of loading them from Google's CDN.

### 1.2.1
- The server says what it is. A connected agent now reads "KitchenSink4XL (kitchensink4xl on PyPI), part of the KitchenSink4AI suite" at the head of the instructions it receives, and `get_server_info` reports the product name, the package to install, the documentation homepage, and the three sibling packages.
- Published addresses moved to kitchensink4.ai, and the registry record is republished at this version.
- The tool figure is stated as measured on every surface: 129 workbook operations across 69 tools, 67 workbook tools plus the two pack toggles.

### 1.2.0
- Read responses are dramatically smaller (seven to eleven times on the heavy calls) with nothing lost; the token estimator now matches the wire exactly, and every published figure is re-measured.
- A performance bug that reopened the workbook per cell is fixed; large cell reads are near-instant.
- Update notice: a weekly, disclosed check for newer releases, off with KS4XL_UPDATE_CHECK=off.
- The install screen and info card are rewritten in plain language.

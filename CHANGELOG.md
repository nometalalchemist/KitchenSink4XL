# Changelog

### 1.2.0
- Read responses are dramatically smaller (seven to eleven times on the heavy calls) with nothing lost; the token estimator now matches the wire exactly, and every published figure is re-measured.
- A performance bug that reopened the workbook per cell is fixed; large cell reads are near-instant.
- Update notice: a weekly, disclosed check for newer releases, off with KS4XL_UPDATE_CHECK=off.
- The install screen and info card are rewritten in plain language.

"""`python -m xlsx_mcp` starts the same server the console scripts start.

The distribution is `kitchensink4xl` and the package is `xlsx_mcp`, so the
module name a person guesses first is the wrong one either way. `xl-mcp` and
`kitchensink4xl` both resolve to `xlsx_mcp.server:main`, and so does this;
`python -m xlsx_mcp.server` keeps working as it always has.

Not documented for 1.0.0: that release has no `__main__` and a README line
promising this route would be false for everybody who installed from PyPI.
The line lands with the next release. See V1.1_QUEUE.md.
"""

from __future__ import annotations

from xlsx_mcp.server import main

if __name__ == "__main__":
    main()

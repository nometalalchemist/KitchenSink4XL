"""Suite-wide setup for the KitchenSink4XL tests.

One job so far: NO NETWORK IN TESTS. get_server_info runs the on-demand
update check, so the suite turns that check off for every test by default.
The update tests switch it back on for themselves and mock the fetch, and
setdefault means a developer who exports the variable still wins.
"""

from __future__ import annotations

import os


def pytest_configure(config):
    os.environ.setdefault("KS4XL_UPDATE_CHECK", "off")

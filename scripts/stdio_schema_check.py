"""Verify tools/list over a REAL stdio session, not the in-process registry.

The guard test reads the tool objects this process registered. That proves
the annotations are right; it does not prove what crosses the wire, and the
defect being closed here is entirely about what a client receives. This
script launches the server as a subprocess, speaks MCP over stdin/stdout the
way Claude Desktop and every other client does, and inspects the actual
tools/list payload: the serverInfo version, the tool count, and every
parameter schema.

Run: python scripts/stdio_schema_check.py [--mode full]
Exit 0 when the handshake reports this package's version and no parameter
ships an empty schema.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys

NON_TYPING = {"title", "description", "default"}


def _send(proc, msg):
    proc.stdin.write(json.dumps(msg) + "\n")
    proc.stdin.flush()


def _read(proc, want_id):
    while True:
        line = proc.stdout.readline()
        if not line:
            raise SystemExit("server closed stdout before answering")
        line = line.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError:
            continue
        if msg.get("id") == want_id:
            return msg


def main() -> int:
    mode = "full"
    if "--mode" in sys.argv:
        mode = sys.argv[sys.argv.index("--mode") + 1]
    env = dict(os.environ, KS4XL_MODE=mode, PYTHONIOENCODING="utf-8")
    proc = subprocess.Popen(
        [sys.executable, "-X", "utf8", "-m", "xlsx_mcp"],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL, text=True, encoding="utf-8", env=env)
    try:
        _send(proc, {"jsonrpc": "2.0", "id": 1, "method": "initialize",
                     "params": {"protocolVersion": "2025-06-18",
                                "capabilities": {},
                                "clientInfo": {"name": "schema-check",
                                               "version": "0"}}})
        init = _read(proc, 1)
        info = init["result"]["serverInfo"]
        _send(proc, {"jsonrpc": "2.0", "method": "notifications/initialized"})
        _send(proc, {"jsonrpc": "2.0", "id": 2, "method": "tools/list",
                     "params": {}})
        tools = _read(proc, 2)["result"]["tools"]
    finally:
        proc.stdin.close()
        proc.terminate()

    from xlsx_mcp import __version__

    print(f"serverInfo: {info.get('name')} {info.get('version')}")
    print(f"tools/list: {len(tools)} tools")

    problems = []
    if str(info.get("version")) != __version__:
        problems.append(
            f"serverInfo.version is {info.get('version')!r}, "
            f"not this package's {__version__!r}")

    empty = []
    params = 0
    for tool in tools:
        props = (tool.get("inputSchema") or {}).get("properties") or {}
        for name, schema in props.items():
            params += 1
            body = {k: v for k, v in schema.items() if k not in NON_TYPING}
            if not body:
                empty.append(f"{tool['name']}.{name}")
    print(f"parameters: {params}, empty schemas: {len(empty)}")
    if empty:
        problems.append("empty parameter schemas: " + ", ".join(empty))

    blob = json.dumps(tools)
    if "$ref" in blob:
        problems.append("a schema crossed the wire with a $ref in it")

    for p in problems:
        print("FAIL:", p)
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import json
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from quick_stock_lookup import (
    DEFAULT_ENV,
    CONTROL_REGISTERS,
    MAIN_REGISTER,
    connect,
    query_register,
    summarize,
)


HOST = "127.0.0.1"
PORT = 8765


class StockLookupState:
    def __init__(self) -> None:
        self.base = None
        self.connected_server = ""

    def ensure_connected(self) -> None:
        if self.base is None:
            self.base, self.connected_server = connect(DEFAULT_ENV)


STATE = StockLookupState()


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/health":
            self._send({"ok": True, "connected": STATE.base is not None, "server": STATE.connected_server})
            return
        if parsed.path != "/stock":
            self._send({"ok": False, "error": "NOT_FOUND"}, status=404)
            return

        params = parse_qs(parsed.query)
        items = params.get("item", [])
        warehouse = first(params, "warehouse")
        control = first(params, "control") in {"1", "true", "yes"}
        if not items or not warehouse:
            self._send({"ok": False, "error": "item and warehouse are required"}, status=400)
            return
        if len(items) > 30:
            self._send({"ok": False, "error": "maximum 30 items"}, status=400)
            return

        try:
            STATE.ensure_connected()
            registers = [MAIN_REGISTER, *CONTROL_REGISTERS] if control else [MAIN_REGISTER]
            matches = []
            errors = []
            for item in items:
                for register in registers:
                    try:
                        rows = query_register(STATE.base, register, item, warehouse)
                        if rows:
                            matches.append({"item": item, "register": register, "rows": rows})
                    except Exception as exc:
                        errors.append({"item": item, "register": register, "error": str(exc)})
            self._send(
                {
                    "ok": True,
                    "server": STATE.connected_server,
                    "main_register": MAIN_REGISTER,
                    "control": control,
                    "summary": summarize(matches),
                    "matches": matches,
                    "errors": errors,
                }
            )
        except Exception as exc:
            self._send({"ok": False, "error": str(exc)}, status=500)

    def log_message(self, format: str, *args) -> None:
        return

    def _send(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def first(params: dict[str, list[str]], name: str) -> str:
    values = params.get(name, [])
    return values[0] if values else ""


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    print(f"Starting 1C stock lookup server on http://{HOST}:{PORT}")
    print("Opening 1C COM connection...")
    STATE.ensure_connected()
    print(f"Connected via {STATE.connected_server}")
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

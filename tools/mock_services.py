"""
mock_services.py
================
Local stand-ins for the external services the bot talks to, so the code paths that
reach outside the machine can be exercised end to end without accounts, API keys or
a network:

    HTTP  127.0.0.1:8765
      GET  /v2/api?module=account&action=txlist&chainid=..&address=..&apikey=..
           Etherscan V2. Serves the bundled feed for chain 1 or 137; apikey=INVALID
           returns Etherscan's real error envelope; an unknown address returns its
           real "No transactions found" envelope.
      GET  /explorer/address/<address>
           The explorer page, for WEB mode over real HTTP.
      POST /webhook/<name>
           Accepts and records a webhook (Teams, Slack, generic) and replies 200.

    SMTP  127.0.0.1:8025
           Accepts and records any message.

Everything received is written under the capture directory, one JSON file per
request, so a test can assert on exactly what the bot sent.

These are test doubles, not simulations of the bot: they only replace what is on the
other end of the wire. The HTTP Request activity, the URL building, the JSON parsing
and the SMTP client all run for real.

Run:  python tools/mock_services.py [--capture DIR] [--seconds N]
"""

from __future__ import annotations

import argparse
import asyncore
import json
import os
import smtpd
import sys
import threading
import time
import warnings
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

warnings.filterwarnings("ignore", category=DeprecationWarning)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MOCK = os.path.join(ROOT, "BlockchainLogisticsValidator", "Data", "Input", "MockChain")
FEEDS = {"1": "etherscan_txlist_response.json", "137": "etherscan_txlist_polygon.json"}
HTTP_PORT, SMTP_PORT = 8765, 8025

CAPTURE = os.path.join(os.environ.get("TEMP", "/tmp"), "blv-mock-capture")
_seq = 0
_lock = threading.Lock()


def record(kind: str, payload: dict) -> None:
    global _seq
    with _lock:
        _seq += 1
        n = _seq
    os.makedirs(CAPTURE, exist_ok=True)
    payload = {"kind": kind, "receivedUtc": datetime.now(timezone.utc).isoformat(), **payload}
    with open(os.path.join(CAPTURE, f"{n:03d}_{kind}.json"), "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)


class Handler(BaseHTTPRequestHandler):
    server_version = "BLV-MockServices/1.0"

    def log_message(self, fmt, *args):          # keep the console quiet
        pass

    def _send(self, code: int, body: bytes, ctype: str) -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code: int, obj) -> None:
        self._send(code, json.dumps(obj).encode("utf-8"), "application/json")

    def do_GET(self):
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}

        if u.path == "/v2/api":
            record("etherscan_request", {"path": u.path, "query": q})
            if q.get("module") != "account" or q.get("action") != "txlist":
                return self._json(200, {"status": "0", "message": "NOTOK", "result": "Unsupported module/action"})
            if not q.get("apikey"):
                return self._json(200, {"status": "0", "message": "NOTOK", "result": "Missing/Invalid API Key"})
            if q["apikey"] == "INVALID":
                return self._json(200, {"status": "0", "message": "NOTOK", "result": "Invalid API Key"})
            feed = FEEDS.get(q.get("chainid", ""))
            if feed is None:
                return self._json(200, {"status": "0", "message": "NOTOK", "result": "Missing or unsupported chainid parameter"})
            with open(os.path.join(MOCK, feed), encoding="utf-8") as f:
                data = json.load(f)
            contract = data["result"][0]["to"].lower() if data["result"] else ""
            if q.get("address", "").lower() != contract:
                return self._json(200, {"status": "0", "message": "No transactions found", "result": []})
            return self._json(200, data)

        if u.path.startswith("/explorer/address/"):
            record("explorer_request", {"path": u.path})
            with open(os.path.join(MOCK, "explorer_ethereum.html"), "rb") as f:
                return self._send(200, f.read(), "text/html; charset=utf-8")

        if u.path == "/health":
            return self._json(200, {"ok": True})

        self._json(404, {"error": "not found", "path": u.path})

    def do_POST(self):
        u = urlparse(self.path)
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length).decode("utf-8", "replace")
        if u.path.startswith("/webhook/"):
            try:
                body = json.loads(raw)
            except json.JSONDecodeError:
                body = None
            record("webhook", {"path": u.path, "contentType": self.headers.get("Content-Type"),
                               "json": body, "raw": None if body is not None else raw})
            return self._send(200, b"1", "text/plain")
        self._json(404, {"error": "not found", "path": u.path})


class CapturingSMTP(smtpd.SMTPServer):
    def process_message(self, peer, mailfrom, rcpttos, data, **kwargs):
        text = data.decode("utf-8", "replace") if isinstance(data, bytes) else data
        record("smtp", {"from": mailfrom, "to": rcpttos, "bytes": len(text), "message": text})


def main() -> int:
    global CAPTURE
    ap = argparse.ArgumentParser()
    ap.add_argument("--capture", default=CAPTURE)
    ap.add_argument("--seconds", type=int, default=0, help="stop after N seconds (0 = run until killed)")
    args = ap.parse_args()
    CAPTURE = args.capture
    os.makedirs(CAPTURE, exist_ok=True)

    httpd = ThreadingHTTPServer(("127.0.0.1", HTTP_PORT), Handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    CapturingSMTP(("127.0.0.1", SMTP_PORT), None)
    threading.Thread(target=lambda: asyncore.loop(timeout=0.5), daemon=True).start()

    print(f"mock services up: http://127.0.0.1:{HTTP_PORT}  smtp://127.0.0.1:{SMTP_PORT}  capture={CAPTURE}",
          flush=True)
    try:
        deadline = time.time() + args.seconds if args.seconds else None
        while deadline is None or time.time() < deadline:
            time.sleep(0.5)
    except KeyboardInterrupt:
        pass
    httpd.shutdown()
    return 0


if __name__ == "__main__":
    sys.exit(main())

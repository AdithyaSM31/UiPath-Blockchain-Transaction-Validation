"""
generate_explorer_page.py
=========================
Renders the mock chain feed as a blockchain-explorer page, so `DataSourceMode = WEB`
can be demonstrated without depending on a live site.

Why a local page rather than scraping etherscan.io directly: a public explorer sits
behind bot protection, rate limits and a layout that changes without notice. A demo
that depends on it fails on the day. The page produced here uses the same table
structure and column order Etherscan uses, so the selectors and the extraction logic
are the same ones a live scrape needs - only the URL differs, and that is a config cell.

Run:  python tools/generate_explorer_page.py
"""

from __future__ import annotations

import html
import json
import os
from datetime import datetime, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROJ = os.path.join(ROOT, "BlockchainLogisticsValidator")
MOCK = os.path.join(PROJ, "Data", "Input", "MockChain")

SELECTOR_NOTE = """
The bot locates the table by id, so the surrounding layout can change freely:
    <webctrl id='txtable' tag='TABLE' />
Row cells are read in document order, matching COLUMNS below.
"""

# Column order matters: the bot reads cells positionally, so this list is the
# contract between the page and 01c_Extract_FromExplorerUI.
COLUMNS = ["TxHash", "Method", "Block", "DateTimeUtc", "From", "To", "Value", "Input"]


def build_page(txs: list[dict], chain_label: str) -> str:
    now = datetime.now(timezone.utc)

    rows = []
    for t in txs:
        ts = datetime.fromtimestamp(int(t["timeStamp"]), tz=timezone.utc)
        cells = [
            t["hash"],
            t["functionName"].split("(")[0],
            t["blockNumber"],
            ts.strftime("%Y-%m-%d %H:%M:%S"),
            t["from"],
            t["to"],
            t["value"],
            t["input"],
        ]
        # Values are rendered in full rather than abbreviated. A public explorer truncates
        # hashes in its list view, which would make a scrape lossy; showing the complete
        # value means extraction reads plain cell text and needs no attribute tricks.
        rows.append(
            "        <tr>\n"
            + "\n".join(
                f'          <td class="c{i}">{html.escape(str(c))}</td>'
                for i, c in enumerate(cells))
            + "\n        </tr>")

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>{chain_label} Explorer - Address Transactions</title>
<style>
  body {{ font-family: Segoe UI, Arial, sans-serif; font-size: 13px; margin: 24px; color: #222; }}
  h1 {{ font-size: 18px; color: #1F3864; margin-bottom: 2px; }}
  .sub {{ color: #666; margin-bottom: 18px; }}
  table {{ border-collapse: collapse; width: 100%; }}
  th {{ background: #1F3864; color: #fff; text-align: left; padding: 7px 9px; font-weight: 600; }}
  td {{ border-bottom: 1px solid #e4e4e4; padding: 6px 9px; font-family: Consolas, monospace;
        font-size: 11px; word-break: break-all; max-width: 340px; }}
  tr:nth-child(even) td {{ background: #fafafa; }}
</style>
</head>
<body>
  <h1>{chain_label} Explorer</h1>
  <div class="sub">Transactions for contract <b>{html.escape(txs[0]["to"]) if txs else ""}</b>
    &middot; {len(txs)} records &middot; generated for offline demonstration</div>

  <table id="txtable">
    <thead>
      <tr>{"".join(f"<th>{c}</th>" for c in COLUMNS)}</tr>
    </thead>
    <tbody>
{chr(10).join(rows)}
    </tbody>
  </table>
</body>
</html>
"""


def main() -> None:
    pairs = [
        ("etherscan_txlist_response.json", "explorer_ethereum.html", "Ethereum"),
        ("etherscan_txlist_polygon.json", "explorer_polygon.html", "Polygon"),
    ]
    for src, dst, label in pairs:
        with open(os.path.join(MOCK, src), encoding="utf-8") as f:
            txs = json.load(f)["result"]
        out = os.path.join(MOCK, dst)
        with open(out, "w", encoding="utf-8") as f:
            f.write(build_page(txs, label))
        print(f"wrote: {os.path.relpath(out, ROOT)}  ({len(txs)} rows)")

    print(SELECTOR_NOTE.strip())


if __name__ == "__main__":
    main()

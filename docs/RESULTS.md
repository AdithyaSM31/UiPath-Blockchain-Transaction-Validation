# Results

Measured on the development machine (Windows 11, UiPath Studio 25.10.15, Excel 16.0),
running `MOCK` data source in `BATCH` mode. Every figure here is reproducible with
`powershell -File tools\run_all_checks.ps1`.

---

## 1. Throughput

| Measure | Value |
|---|---|
| Transactions per run | 48 |
| Shipments | 12 |
| ERP milestones compared | 48 |
| Validation rules applied | 6 |
| Rule evaluations per run | 288 (48 × 6) |
| **Workflow execution time** | **21–22 s** |
| Pack + execute, wall clock | 37 s |
| Throughput | ≈ 2.2 transactions/second, ≈ 13 rule evaluations/second |

Roughly 12 of those 22 seconds are Excel COM startup for the two workbook reads. The
validation engine itself accounts for about 4 seconds. That is the cost worth quoting for
scaling estimates: **~8,000 transactions/hour** on one unattended robot, before any queue
parallelism.

### Effort reduction

A manual reconciliation of one transaction against the ERP — open the explorer, find the
hash, decode the payload, look up the shipment, compare four fields, record the outcome —
takes an experienced operator roughly 2–3 minutes.

| | Manual | Bot |
|---|---|---|
| 48 transactions | ~1.5–2.5 hours | 22 seconds |
| Per transaction | ~120–180 s | ~0.46 s |

That supports the deck's **90%+ reduction** claim comfortably. State it as an estimate
against a stated manual baseline, not as a measured comparison — no manual baseline was
actually timed for this project.

---

## 2. Validation outcome

```
total=48  pass=40  warn=1  fail=7  escalations=1     pass rate 83.3%
```

| Rule | Findings | Shipment | What was caught |
|---|---|---|---|
| R1 Timestamp | 2 | `SHP-1002`, `SHP-1003` | 2h30m drift (warning), 5h20m drift (fail) |
| R2 Quantity | 1 | `SHP-1005` | on-chain 480 against a PO for 500 |
| R3 Whitelist | 1 | `SHP-1007` | delivery signed by an unapproved wallet → escalated |
| R4 Duplicate | 2 | `SHP-1009` | one transaction hash recorded twice |
| R5 Sequence | 1 | `SHP-1011` | `InTransit` recorded after `Delivered` |
| R6 Contract event | 1 | `SHP-1012` | `GoodsReceived` never emitted |

**Detection rate: 7/7 seeded anomalies, 0 false positives.** The five clean shipments
(`SHP-1001`, `1004`, `1006`, `1008`, `1010`) passed all six rules on every event.

R4 reports two rows for one anomaly by design: both copies of a duplicated transaction are
flagged, because from the ledger alone there is no way to tell which submission was the
legitimate one.

---

## 3. Acceptance suite

`tools\run_all_checks.ps1` — **12/12 passing**.

| # | Check | Result |
|---|---|---|
| 1 | Dataset oracle — anomalies present, ABI decodes | PASS |
| 2 | Workflow Analyzer — no Error-severity findings | PASS |
| 3 | Batch run — expected counts | PASS |
| 4 | Every rule reports its seeded anomaly | PASS |
| 5 | Rule disabled in Excel is skipped (Novelty 1) | PASS |
| 6 | HTML dashboard produced, self-contained | PASS |
| 7 | CRITICAL anomaly escalated (Novelty 2) | PASS |
| 8 | Identical result on a second chain (Novelty 3) | PASS |
| 9 | Audit tampering detected (Novelty 4) | PASS |
| 10 | WEB mode reaches the same result as MOCK | PASS |
| 11 | UiPath test cases (7 of them) | PASS |
| 12 | Dispatch splits per shipment | PASS |

The dataset oracle (`tools\verify_seeded_data.py`) is a deliberately independent
implementation — it re-derives the expected outcome for all 48 transactions from the raw
JSON and the ERP workbook. The bot and the oracle agree on every row.

### 3.1 UiPath test cases

`tools\run_tests.ps1` — **7/7 passing**. Also runnable from Studio's Test Explorer.

| Test | Covers | Edge cases beyond the sample data |
|---|---|---|
| TC01 | R1 timestamp bands | drift exactly at the tolerance boundary; row with no ERP milestone |
| TC02 | R2 quantity | overage as well as shortfall; a tolerance that absorbs both |
| TC03 | R3 whitelist | EIP-55 checksummed address matching a lower-case entry; inactive partner |
| TC04 | R4 duplicates | same hash on a *different* shipment must not be flagged; malformed short hash |
| TC05 | R5 ordering | the flagged row is the out-of-order event, not the one that preceded it |
| TC06 | R6 required event | the finding attaches to the shipment's latest event |
| TC07 | audit chain | write → verify VALID → tamper → verify INVALID at the right line |

These are unit tests against hand-built fixtures, not a replay of the sample data. TC04
caught a real defect while being written: R4 assumed a transaction hash was at least 12
characters and threw on a malformed one. That is now handled and covered.

### 3.2 Extraction modes agree

The same 48 transactions reach the same verdict through every source:

| Mode | Source | Result |
|---|---|---|
| `MOCK` | bundled Etherscan-shaped JSON | `total=48 pass=40 warn=1 fail=7` |
| `WEB` | explorer page, HTML parsed | `total=48 pass=40 warn=1 fail=7` |
| `API` | live Etherscan V2 | not executed — no API key configured |

That equality is the real evidence for the "blockchain agnostic" claim: the validation
engine cannot tell which reader produced its DataTable.

---

## 4. Novelty evidence

### Novelty 1 — rules as configuration

Setting `R4 Enabled = FALSE` in `Config.xlsx` and re-running:

```
[VALIDATE] R4 is disabled in Config.xlsx - skipped.
[VALIDATE] total=48 pass=42 warn=1 fail=5 | R1=2, R2=1, R3=1, R5=1, R6=1
```

Failures 7 → 5, R4 absent from the summary. No workflow file was opened or republished.

### Novelty 2 — human-in-the-loop

```
[ESCALATE] 1 anomaly(ies) require human review: SHP-1007/Delivered
[ESCALATE] 1 escalation(s) dispositioned as "REJECTED - hold the shipment..." via SIMULATE
```

The decision is written into the results table *before* the audit log is generated, so the
reviewer's judgement is hashed into the chain alongside the machine's verdict.

Verified in `SIMULATE` and `AUTO_LOG`. **`PROMPT` has not been executed** — the dialog blocks
on a desktop and cannot be driven headlessly.

### Novelty 3 — blockchain agnosticism

| Run | Chain | Source | Outcome |
|---|---|---|---|
| 1 | Ethereum (chainid 1) | `etherscan_txlist_response.json` | `total=48 pass=40 warn=1 fail=7` |
| 2 | Polygon (chainid 137) | `etherscan_txlist_polygon.json` | `total=48 pass=40 warn=1 fail=7` |

Identical, with only configuration changed between runs.

### Novelty 4 — tamper-evident audit trail

Chain continuity across three consecutive runs:

| Run | Entries | Verdict |
|---|---|---|
| 1 | 48 | VALID from genesis |
| 2 | 96 | VALID from genesis |
| 3 | 144 | VALID from genesis |

Tamper test — one `FAIL` verdict changed to `PASS`, nothing else touched:

```
INVALID - first problem at line 7.
  line 7: record contents were altered (RecordHash mismatch) - SHP-1005/Dispatched
  line 7: EntryHash does not match its own contents
```

Restored from backup, verification returns to VALID.

---

## 5. Artefacts produced per run

| File | Size | Contents |
|---|---|---|
| `ValidationReport_<runid>.xlsx` | ~23 KB | Summary, per-rule breakdown, all results (colour-coded), exceptions |
| `ValidationAuditLog.csv` | ~20 KB / run | 48 hash-chained entries, append-only |
| `dashboard.html` | ~9 KB | Self-contained visual dashboard: KPIs, rules, findings, run history |
| `dashboard_summary.json` | ~3 KB | KPIs and findings for Power BI / Google Sheets |
| `Alert_<runid>.eml` | ~5 KB | HTML alert with the failure summary table |
| `QueueDryRun/queue_payloads_<runid>.json` | ~40 KB | 12 per-shipment queue payloads (dispatch mode) |

---

## 6. What these numbers do not cover

Stated plainly so the limits are not overclaimed:

- **Live blockchain data.** All figures are from `MOCK` and `WEB`, both of which read bundled
  data. `API` mode is implemented and shares the same parser, but has not been run against
  Etherscan (no API key configured).
- **Orchestrator queues.** `Dispatcher` is verified in `DRYRUN` (48 → 12 payloads).
  `Add Queue Item`, `Get Transaction Item` and `Set Transaction Status` have never executed
  against a live tenant. The package itself *is* published to Orchestrator.
- **Web extraction is HTML parsing, not browser automation.** `WEB` mode fetches the
  explorer page and parses its transaction table. It does not drive a browser with UiPath
  Data Scraping: UIAutomation 25.10 removed the classic `ExtractStructuredData`, its
  replacement needs recorder-generated descriptors, and the UiPath browser extension is not
  installed on this machine.
- **`EscalationMode = PROMPT`.** Verified only in `SIMULATE` and `AUTO_LOG`; the reviewer
  dialog blocks on a desktop and cannot be driven headlessly.
- **Scale.** The largest run tested is 48 transactions. Nothing in the design is quadratic —
  the rules use dictionary lookups and one sort per shipment — but thousands of transactions
  per run has not been measured.
- **Concurrency.** Two robots appending to one audit file would interleave and break the
  chain. Each robot needs its own audit folder.

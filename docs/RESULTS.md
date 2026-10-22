# Results

## 0. Release after the Review 1 audit — 16 Sep 2026

The project was audited slide by slide against *RPA Project Review 1.pptx*
([DECK_TRACEABILITY.md](DECK_TRACEABILITY.md)). This release closes the gaps found:

| Gap against the deck | Now |
|---|---|
| No PDF audit log (slide 7, objective 5) | `08_Report_PDF` |
| No Teams / Slack alerts (objective 4), no Outlook mail (slide 7) | `09b_Alert_Webhooks`; `AlertMode = OUTLOOK`; SMTP now attaches the reports and honours `SmtpUseSsl` |
| No CSV or SQL ERP source; no Database Connect (slide 5) | `02_Extract_Logistics` EXCEL / CSV / SQL (Database Connect over ODBC) |
| `FieldMapping` sheet read but never applied; `MappingMode = LOOKUP` not implemented (slide 5, Novelty 3) | `01d_Apply_FieldMapping`; LOOKUP in `03`; a third, differently-shaped feed (Hyperledger Besu ethereum-etl export) |
| No processing times or trends on the dashboard (objective 6) | Phase timings, `RunHistory.csv`, trend chart; fail % and top anomaly types in the KPI JSON; Power BI / Sheets push |
| Review dialog rather than a form (Novelty 2) | Review form with every fact, reviewer name, note, timeout and evidence image |
| Audit entries did not carry an input-data hash (Novelty 4) | `InputDataHash` column, hashed into `RecordHash` |
| Performer could not run as its own process and produced no audit or reports | Standalone entry point, drains the queue, runs the full finalise pipeline |
| `Settings` rows that had no effect (`ProjectName`, `ExplorerAddressUrl`, `ShipmentsSheet`, `ShipmentEventsSheet`, `MappingMode`, `TxShipmentMapFile`, the three tolerance settings, `EscalationSeverity`, `SmtpUseSsl`, `LogLevel`, the three folder settings) | All 60 settings are read by the bot |
| `generate_sample_data.py` dropped six settings on regeneration, which would break the bot | Fixed; regenerated inputs are unchanged apart from the intended additions |
| A failed run left no evidence | Logged, screenshotted, re-thrown |

**Verified on 16 Sep 2026 (development laptop):**

| Check | Result |
|---|---|
| Dataset oracle, including the new feed, CSV and mapping-file cross-checks | all passed |
| `tools\compile_check.ps1` — every Invoke Code body in the workflows and tests | 59/59 compile |
| `UiRobot pack` — all workflows and all 11 test cases load and compile | passed |
| Workflow Analyzer | 0 Error-severity findings |
| End-to-end batch run with the new configuration loading, extraction, field mapping, ERP canonicalisation, mapping join and validation engine | `total=48 pass=40 warn=1 fail=7 escalations=1` — identical to the original release |

## 0.1 Re-verified on UiPath Studio 26.0.204 — 21 Oct 2026

The checks listed above as pending have now been executed. The project was run against
Studio and Robot **26.0.204** (the split `UiPathPlatform` install layout), not the 25.10.15
it was built on.

| Check | Result |
|---|---|
| `tools\compile_check.ps1` | 59/59 Invoke Code bodies compile |
| Pack / publish — every workflow and all 11 test cases | passed |
| Workflow Analyzer | 0 Error-severity findings |
| Acceptance suite, first full run on 26.0.204 | **18/19** — only `UiPath test cases` failed |
| `tools\run_tests.ps1` after the fixes below | **11/11 test cases passed** |

Four defects were found by running these for the first time, and fixed:

| Found by | Defect |
|---|---|
| suite #17 | A run that failed inside `00_Init_ReadConfig` wrote its failure screenshot to the robot's working directory — for a packed run, the extracted package under `%USERPROFILE%\.nuget\packages`. The only evidence a failed job left behind went into a disposable cache. `Main.xaml` now resolves the output root itself (`BLV_OUTPUT_ROOT` → `Config.local.json` → project default) without needing `00_Init` to have succeeded. |
| TC10 | `06_HumanInTheLoop_Escalation` never assigned its own `out_EscalationCount` output argument, so every caller that bound it received 0. Invisible until now because the `escalations=` figure in the run summary comes from the validation engine's own tally. |
| TC08 | Canonical `DateTime` columns were created with a `DataColumn`'s default `DateTimeMode` of `UnspecifiedLocal`, which strips `DateTimeKind.Utc` back to `Unspecified` on read. The transforms produce correct UTC values, but for a column named `EventTimestampUtc` anything comparing against `DateTime.UtcNow` or re-serialising would reinterpret it in the robot's local zone. Now `DataSetDateTime.Utc`. |
| TC08 | *Test* bug: `Invoke Code` runs its body by reflection, so the exception reaching the catch is a `TargetInvocationException` whose own `Message` is generic. The assertion that an incomplete mapping sheet is "rejected by name" was reading the wrapper, not the inner message. Now `GetBaseException().Message`. |

> **Not yet re-measured:** a full 19/19 suite run *after* the four fixes above. The run that
> measured 18/19 predates them; since then `11/11` test cases and a clean pack have been
> confirmed directly, but the complete suite has not been re-run end to end. Run
> `powershell -ExecutionPolicy Bypass -File tools\run_all_checks.ps1` and record the number here.

Two environment notes, both now handled in `tools\UiPathPaths.ps1` and documented in
[USER_GUIDE.md §2.1](USER_GUIDE.md#21-two-things-that-bite-on-a-fresh-machine): Studio 26
installs to `%LOCALAPPDATA%\Programs\UiPathPlatform\{Studio,Robot}\<version>\` rather than the
legacy path every script used to hardcode, and a default Windows client refuses to run `.ps1`
files at all without `-ExecutionPolicy Bypass`.

The sections below are the measurements from the original release (11 Sep 2026) and remain
the reference numbers.

---

Measured on the development machine (Windows 11, UiPath Studio 25.10.15, Excel 16.0),
running `MOCK` data source in `BATCH` mode. Every figure here is reproducible with
`powershell -ExecutionPolicy Bypass -File tools\run_all_checks.ps1`.

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
| `WorkQueue/queue_payloads_<runid>.json` | ~60 KB | Work queue: 12 per-shipment items (dispatch mode) |

---

## 6. What these numbers do not cover

Stated plainly so the limits are not overclaimed:

- **Live blockchain data.** All figures are from `MOCK` and `WEB`, both of which read bundled
  data. `API` mode is implemented and shares the same parser, but has not been run against
  Etherscan (no API key configured).
- **No Orchestrator.** The project is scoped to UiPath Studio. The dispatch / perform split
  is verified end to end through a local work-queue file (48 transactions → 12 items → the
  same verdict as a batch run); Orchestrator scheduling, triggers and queues are out of scope.
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

> The limits above were written for the original release. Since then the Performer path is
> exercised end to end through the local work-queue file (suite check #16) and alert delivery
> against a local capture server (suite check #15); a live Etherscan key and real mail/chat
> endpoints remain outside what the repository can test.

# User Guide and Handover Notes

For the logistics or compliance person who runs the bot, and for whoever maintains it next.
Everything operational is changed in `BlockchainLogisticsValidator\Data\Config.xlsx`; no
workflow needs to be opened to switch chains, data sources, rules or alert channels.

---

## 1. What a run does

1. Reads `Config.xlsx`, fails fast on any missing or invalid setting.
2. Extracts the audited contract's transactions (bundled feed, Etherscan API, or explorer page) and translates them through the selected **FieldMapping** sheet.
3. Reads the ERP/WMS shipment master and expected milestones (Excel, CSV or SQL database).
4. Links each transaction to its shipment milestone (call-data key or mapping file).
5. Applies the enabled rules R1–R6; every transaction gets PASS, WARNING or FAIL.
6. Pauses for a human on high-severity anomalies (or records a preset/deferred decision).
7. Appends every verdict to the hash-chained audit log and re-verifies the whole chain.
8. Writes the Excel report, the PDF audit report, the alerts, the KPI summary and the HTML dashboard.

A default run (bundled data) always ends with `total=48 pass=40 warn=1 fail=7 escalations=1`.

---

## 2. Before you start

| Need | Why |
|---|---|
| UiPath Studio 25.10 or 26.0 (or a Robot of the same version) | Runs the Windows (.NET 8) project. Verified on 25.10.15 and 26.0.204 |
| Microsoft Excel | `Config.xlsx`, the ERP workbook and the mapping file are read with Excel activities |
| Microsoft Access ODBC driver (installed with Office) | Only for `LogisticsSourceMode = SQL` against the bundled `LogisticsERP.accdb` |
| Python 3 + `openpyxl`, `pycryptodome` | Only for the maintenance tools (data generator, oracle, rule-toggle check) |
| Windows **Smart App Control off** on the machine that runs the bot | Smart App Control can block the compiled, unsigned workflow assemblies inside a UiPath package |

### 2.1 Two things that bite on a fresh machine

**Scripts must be allowed to run.** Windows clients default to an execution policy of
`Restricted`, which refuses every `.ps1` in this repository with *"running scripts is disabled
on this system"* — including `build.ps1` and the acceptance suite. Every command in these docs
therefore passes the flag explicitly:

```bash
powershell -ExecutionPolicy Bypass -File "tools\run_all_checks.ps1"
```

That flag applies to one process and changes no machine setting. If you would rather not type
it each time, `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` is the usual alternative —
but that is a security setting, so make that choice deliberately.

**UiPath moved where it installs itself.** Studio 25.10 and earlier put Studio and the Robot
together in `%LOCALAPPDATA%\Programs\UiPath\Studio`. Studio 26.0 (the Connected installer)
splits them into versioned folders:

```
%LOCALAPPDATA%\Programs\UiPathPlatform\Studio\<version>\
%LOCALAPPDATA%\Programs\UiPathPlatform\Robot\<version>\
```

`tools\UiPathPaths.ps1` finds whichever layout is present, and every script here dot-sources
it rather than hardcoding a path. It also works around the one behavioural difference: in the
split layout `UiRobot pack` cannot find the Studio command-line application it delegates to, so
`Invoke-UiPathPack` falls back to `UiPath.Studio.CommandLine.exe publish` into a local folder,
which produces the same package.

If a script reports *"UiRobot.exe not found"*, UiPath is installed somewhere neither layout
covers — set `$UiPathRobotExe` before dot-sourcing, or extend the candidate list in that file.

---

## 3. Running it

**From Studio:** open `BlockchainLogisticsValidator\project.json`, run `Main.xaml`.

**From a terminal** (packs and runs the same package Orchestrator would):

```bash
powershell -ExecutionPolicy Bypass -File "build.ps1"
```

Try a setting for one run without touching `Config.xlsx`:

```bash
powershell -ExecutionPolicy Bypass -File "build.ps1" -Set @{ LogisticsSourceMode = "CSV" }
```

`-Set` writes `Data\Config.local.json`, which overrides any `Settings` value. That file is
also where machine-specific paths and real secrets (API key, SMTP password, webhook URLs)
belong — it is git-ignored and never part of the hand-in.

**Outputs** land in `Data\Output` (or wherever `BLV_OUTPUT_ROOT` / `OutputRoot` points):

| File | What it is |
|---|---|
| `Reports\ValidationReport_<run>.xlsx` | Summary, per-rule breakdown, colour-coded results, exceptions |
| `Reports\AuditReport_<run>.pdf` | Formal audit document: run metadata and bot identity, outcome, rules, findings with human decisions, every transaction, audit-chain evidence, signature block |
| `Reports\Alert_<run>.eml` | The composed e-mail alert (always written, sent only in SMTP/OUTLOOK mode) |
| `Reports\Alert_<run>_teams.json`, `_slack.json` | The Teams and Slack payloads (always written, posted only when a URL is set) |
| `Reports\dashboard.html` | Self-contained dashboard: KPIs, processing time by phase, trends across runs, rules, findings, history |
| `Reports\dashboard_summary.json` | KPIs, fail %, top anomaly types and phase timings for Power BI / Sheets |
| `AuditLogs\ValidationAuditLog.csv` | Append-only, hash-chained evidence of every verdict |
| `AuditLogs\RunHistory.csv` | One row per run: duration, counts, per-rule findings, chain status |
| `Screenshots\Review_*.png` | What the reviewer saw on the review form |
| `Screenshots\Failure_*.png` | The screen at the moment a run failed |

---

## 4. `Config.xlsx` reference

### 4.1 `Settings`

Relative paths are relative to the project folder. Output folders are anchored under the
output root.

**Run identity and output**

| Setting | Default | Meaning |
|---|---|---|
| `ProjectName` | BlockchainLogisticsValidator | Shown on every report header |
| `BotIdentity` | BLV-BOT-01 | Recorded in every audit entry; part of the run id |
| `ReportsFolder`, `AuditLogFolder`, `ScreenshotFolder` | `Data\Output\...` | Where outputs go (under the output root) |
| `AuditLogFile` | ValidationAuditLog.csv | The hash-chained log |
| `RunHistoryFile` | RunHistory.csv | Per-run metrics for the trend charts |
| `LogLevel` | Info | `Trace` also logs every transaction's verdict |

**Blockchain source**

| Setting | Default | Meaning |
|---|---|---|
| `DataSourceMode` | MOCK | `MOCK` bundled feed · `API` Etherscan V2 · `WEB` explorer page |
| `MockChainFile` | Etherscan-shaped feed | File read in MOCK mode |
| `ChainRecordsPath` | result | JSON path of the transaction array in a MOCK/API response |
| `FieldMappingSheet` | FieldMapping | Which mapping sheet translates the feed (see 4.3) |
| `EtherscanBaseUrl`, `EtherscanApiKey`, `ChainId`, `MaxTransactions` | V2 endpoint, blank, 1, 500 | API mode |
| `ContractAddress` | registry contract | The contract whose transactions are audited |
| `ExplorerPageUrl`, `ExplorerAddressUrl` | local replica, etherscan.io | WEB mode; a blank page URL uses the address URL template |

**ERP / WMS source**

| Setting | Default | Meaning |
|---|---|---|
| `LogisticsSourceMode` | EXCEL | `EXCEL` · `CSV` · `SQL` |
| `LogisticsFile`, `ShipmentsSheet`, `ShipmentEventsSheet` | bundled workbook | EXCEL mode |
| `ShipmentsCsvFile`, `ShipmentEventsCsvFile` | `Data\Input\ERP_CSV\*.csv` | CSV mode |
| `DbProviderName`, `DbConnectionString` | System.Data.Odbc, Access file | SQL mode; `{ProjectRoot}` is substituted |
| `DbShipmentsQuery`, `DbShipmentEventsQuery` | `SELECT * FROM ...` | SQL mode; must return the columns listed in 4.4 |

**Mapping and rules**

| Setting | Default | Meaning |
|---|---|---|
| `MappingMode` | DECODE | `DECODE` shipment/PO from call data · `LOOKUP` from `TxShipmentMapFile` by TxHash |
| `TxShipmentMapFile` | `Data\Input\TxShipmentMap.xlsx` | LOOKUP mode |
| `TimestampToleranceHours`, `TimestampWarnToleranceHours`, `QuantityTolerance` | 2, 4, 0 | Used only when the matching `ValidationRules` Param cell is blank |

**Human review**

| Setting | Default | Meaning |
|---|---|---|
| `EscalationEnabled` | True | Off records every escalation as deferred |
| `EscalationSeverity` | CRITICAL | A FAIL at or above this severity escalates (as does `FailAction = ESCALATE`) |
| `EscalationMode` | SIMULATE | `PROMPT` show the review form · `SIMULATE` apply the preset decision · `AUTO_LOG` defer |
| `AttendedMode` | True | False downgrades PROMPT to AUTO_LOG so an unattended job never waits |
| `ReviewUi` | WINDOW | `WINDOW` full review form · `DIALOG` simple Input Dialog |
| `ReviewTimeoutMinutes` | 30 | An unanswered form closes and records DEFERRED |
| `SimulatedEscalationDecision` | REJECTED - ... | Decision applied in SIMULATE mode |

**Alerts and dashboard**

| Setting | Default | Meaning |
|---|---|---|
| `AlertMode` | EML | `EML` write to disk · `SMTP` send via `SmtpHost` · `OUTLOOK` send via desktop Outlook |
| `AlertFrom`, `AlertTo` | demo addresses | `AlertTo` is semicolon separated |
| `AlertAttachReports` | True | Attach the Excel and PDF reports to sent mail |
| `SmtpHost`, `SmtpPort`, `SmtpUseSsl`, `SmtpUser`, `SmtpPassword` | smtp.gmail.com, 587, True, blank | SMTP mode (keep credentials in `Config.local.json`) |
| `OutlookAccount`, `OutlookSaveAsDraft` | blank, False | OUTLOOK mode |
| `TeamsWebhookUrl`, `SlackWebhookUrl` | blank | Blank saves the payload instead of posting |
| `DashboardPushUrl`, `DashboardPushFormat` | blank, POWERBI | Push KPIs to a Power BI streaming dataset (`POWERBI`) or any JSON endpoint such as a Google Apps Script web app (`JSON`) |

**Scale**

| Setting | Default | Meaning |
|---|---|---|
| `RunMode` | BATCH | `DISPATCH` splits the run into one queue item per shipment |
| `QueueMode` | DRYRUN | `ORCHESTRATOR` real queue · `DRYRUN` payload files on disk |
| `QueueName` | BlockchainValidationQueue | Must match the Orchestrator queue |
| `QueueDryRunFile` | blank | Performer in DRYRUN: payload file (blank = newest) |

Any value outside its listed options stops the run at start-up with a message naming the setting.

### 4.2 `ValidationRules`

One row per rule: `Enabled` (TRUE/FALSE), `Severity` (LOW/MEDIUM/HIGH/CRITICAL), `Param1`,
`Param2`, `FailAction` (FLAG or ESCALATE). Disabling a rule marks its column SKIPPED in every
report; nothing else changes.

### 4.3 `FieldMapping` sheets

Each row turns one field of a raw chain record into one canonical column.

| Column | Meaning |
|---|---|
| `BlockchainField` | A field name in the feed, or `field:argN` for ABI argument N of that field's call data, or `field:selector` for its 4-byte function selector |
| `LogisticsField` | The canonical column: `TxHash`, `BlockNumber`, `EventTimestampUtc`, `SenderAddress`, `ContractAddress`, `EventSelector`, `ShipmentID`, `OnChainQty`, `PONumber`, `TxFailed` (all required), `RawInput` (optional) |
| `TransformType` | `NONE`, `LOWERCASE`, `UPPERCASE`, `INTEGER`, `HEX_TO_INTEGER`, `DECIMAL`, `EPOCH_TO_DATETIME`, `ISO_TO_DATETIME`, `BOOLEAN`, `BOOLEAN_INVERT`, `ABI_STRING`, `ABI_UINT256`, `ABI_ADDRESS` |

`FieldMapping` reads Etherscan `txlist` records; `FieldMapping_EvmEtl` reads an ethereum-etl
export (as used for Hyperledger Besu and the BigQuery crypto datasets). A sheet that misses a
required column, or names an unknown transform, stops the run with the exact problem.

### 4.4 ERP data columns

Whatever the source, the bot needs:

- **Shipments:** `ShipmentID`, `PONumber`, `Product`, `ExpectedQty`, `ReceiverAddress`
- **ShipmentEvents:** `ShipmentID`, `EventName`, `ExpectedEventTimestampUtc`, `HandlerPartner`, `HandlerAddress`

Extra columns are carried along. Timestamps may be real date values or `yyyy-MM-dd HH:mm:ss` text.

### 4.5 Other sheets

`ApprovedWallets` (R3 whitelist, `Active` TRUE/FALSE), `EventSequence` (R5 order and each
milestone's function selector), `EventSignatures` (reference: signatures and topic hashes).

---

## 5. Common tasks

**Validate a different chain on Etherscan.** Set `DataSourceMode = API`, `ChainId` (137 Polygon,
56 BSC …), `ContractAddress`, and put `EtherscanApiKey` in `Config.local.json`.

**Validate a feed with different field names.** Add a `FieldMapping_<name>` sheet, set
`FieldMappingSheet` to it and `ChainRecordsPath` to the array's location. Try the bundled
example: `-Set @{ MockChainFile = "Data\Input\MockChain\evm_etl_besu_export.json"; ChainRecordsPath = "transactions"; FieldMappingSheet = "FieldMapping_EvmEtl" }`.

**Read the ERP from SQL Server.** `LogisticsSourceMode = SQL`, `DbProviderName = System.Data.SqlClient`,
`DbConnectionString = Server=...;Database=...;Integrated Security=True;` and point the two
queries at your tables or views.

**Link transactions from a mapping file.** Fill `TxShipmentMap.xlsx` (`TxHash`, `ShipmentID`,
`PONumber`, `MethodId`) and set `MappingMode = LOOKUP`.

**Send real alerts.** Gmail: `AlertMode = SMTP`, port 587, `SmtpUseSsl = True`, an app
password in `Config.local.json`. Outlook: `AlertMode = OUTLOOK` (try `OutlookSaveAsDraft = True`
first). Teams: create a *Workflows → "Post to a channel when a webhook request is received"*
flow and paste its URL into `TeamsWebhookUrl`. Slack: an incoming-webhook URL into `SlackWebhookUrl`.

**Push KPIs to Power BI.** Create a streaming dataset whose fields match the pushed row:
`runId`, `timestampUtc`, `dataSourceMode`, `chainId` (text); `total`, `pass`, `warning`, `fail`,
`escalations`, `r1Findings` … `r6Findings` (whole numbers); `passPercent`, `failPercent`,
`durationSeconds` (decimal); `topAnomalyType` (text); `auditChainValid` (true/false). Copy the
dataset's push URL into `DashboardPushUrl`.

**Review escalations in person.** `EscalationMode = PROMPT`, `AttendedMode = True`. The bot
pauses on each escalated transaction and shows the review form. Choose a decision, confirm your
name, add a note, and click *Record decision and resume*. The decision, reviewer and note go
into the audit log and every report. Leaving the form unanswered records DEFERRED after
`ReviewTimeoutMinutes`.

**Run unattended.** `AttendedMode = False` and `EscalationMode = AUTO_LOG`; set `BLV_OUTPUT_ROOT`
on the robot machine (runbook §3.1).

**Add a rule.** Add a row to `ValidationRules`, a `Workflows\Rules\R7_*.xaml` that fills
`R7_Status`/`R7_Detail`, a case in `05_Validate_Engine`'s Switch, and `R7` to `RULE_COLUMNS` in
`tools\build_workflows.py`.

**Add a milestone.** Add it to `EventSequence` (with its function selector) and to the ERP data.

---

## 6. Proving the audit trail

```bash
powershell -ExecutionPolicy Bypass -File "build.ps1" -Entry "Workflows\12_Verify_AuditChain.xaml"
```

It recomputes every entry's `RecordHash` and `EntryHash`, re-walks `PrevHash`, and reports
`VALID` or the first line that no longer matches. `tools\tamper_test.ps1` demonstrates this by
editing one verdict and restoring it.

Each entry fixes: bot identity, timestamp, run id, transaction hash, shipment, milestone,
on-chain time and quantity, **InputDataHash** (SHA-256 of the transaction's call data),
verdict, severity, findings, and the human decision and reviewer. A log written by the first
release (without `InputDataHash`) is still verifiable; the bot sets such a file aside as
`*.legacy-<time>.csv` and starts a new chain rather than mixing layouts.

---

## 7. Troubleshooting

| Symptom | Cause and fix |
|---|---|
| `X must be one of ...` at start-up | A `Settings` value is misspelt; the message names it |
| `The FieldMapping sheet does not map required field(s): ...` | Add rows for the named canonical columns |
| `ERP ... data is missing column(s)` | The CSV / query / sheet lacks a required column (4.4) |
| `An Application Control policy has blocked this file` | Windows Smart App Control is blocking the package's compiled workflow assembly. Run on a machine where it is off. |
| Job hangs with no output | `EscalationMode = PROMPT` on an unattended robot — set `AttendedMode = False` |
| `The review window could not be shown` | No interactive desktop; use `AUTO_LOG`, or `ReviewUi = DIALOG` |
| Webhook `HTTP 4xx` warning | The URL is wrong or expired; the payload is still saved in `Reports` |
| `QueueMode is DRYRUN but no dispatcher payload file was found` | Run Main with `RunMode = DISPATCH` first |
| `[VERIFY] INVALID` | The audit log was edited, truncated, or written by two robots at once — give each robot its own output root |

---

## 8. Handover notes

**Where things are.** Workflows are generated by `tools\build_workflows.py` and test cases by
`tools\build_tests.py`, so argument signatures stay consistent across files. Regenerate after
editing a generator:

```bash
python tools\build_workflows.py
python tools\build_tests.py
```

If a workflow is later maintained by hand in Studio, remove it from the generator's
`GENERATORS` map first, or the next regeneration overwrites the edit.

**Sample data** is deterministic: `python tools\generate_sample_data.py`, then
`powershell -ExecutionPolicy Bypass -File tools\build_erp_database.ps1` for the Access copy of the ERP data.
`python tools\verify_seeded_data.py` independently re-derives every expected finding.

**Before any hand-in or demo:** `powershell -ExecutionPolicy Bypass -File tools\run_all_checks.ps1` (about 15 minutes).
Every check and what it proves is listed at the top of that script; the deck-to-check mapping
is in [DECK_TRACEABILITY.md](DECK_TRACEABILITY.md).

**Checks that don't run the robot:** `tools\compile_check.ps1` compiles every Invoke Code body
exactly as the robot would. It is the fastest way to catch a typo in a rarely used branch.

**Honest gaps** (also in the README):

- Orchestrator queue activities, a live Etherscan key, a real mail server, Outlook sending and
  real Teams/Slack channels need accounts this repository does not hold. Their code paths are
  exercised locally (DRYRUN queue, capture server, payload files); the final hop is untested.
- `WEB` mode parses the explorer's HTML rather than recording a browser scrape.
- The review form is built in the workflow rather than designed as a Studio `.uiform`.
- The audit log is single-writer: parallel robots each need their own output root.

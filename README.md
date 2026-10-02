# UiPath-Based Validation of Blockchain Transactions for Transparent Logistics

A UiPath RPA bot that reconciles on-chain logistics events against enterprise ERP records,
flags anomalies against six configurable rules, escalates critical findings to a human, and
writes a tamper-evident audit trail suitable for regulatory submission.

**Course:** Robotic Process Automation
**Team:** Adithya Sankar Menon (23BRS1079) · Nitin Nandakumar (23BRS1378)

---

## Quick start

Open `BlockchainLogisticsValidator\project.json` in UiPath Studio and run `Main.xaml`.
That's it — the default configuration runs entirely offline against bundled sample data.

To run it headlessly instead (no Studio window):

```bash
powershell -File "build.ps1"
```

To verify every claim this project makes, end to end:

```bash
powershell -File "tools\run_all_checks.ps1"
```

That suite is the single source of truth for whether the project works. It currently
reports **16/16 passing**, including the ten UiPath test cases, the outward-path
integration tests and the reviewer-dialog test. Add `-SkipDialog` on a machine nobody is
watching — the last check puts a dialog on screen for a few seconds.

To run just the test cases:

```bash
powershell -File "tools\run_tests.ps1"
```

---

## What a run does

```
Config.xlsx ──┐
              ├──> 00 Init ──> 01 Extract chain ──> 02 Extract ERP ──> 03 Map & join
mock/API/web ─┘        │                                                     │
                       └─ 01m Normalise (ChainProfiles + FieldMapping)       v
       12 Verify chain <── 10 Audit log <── 06 Escalate <── 05 Validate (R1..R6)
                 │                                                   │
                 ├──> 07 Excel report · 11 Dashboard JSON · 09 Alert <┘
                 └──> 13 HTML dashboard ──> 08 PDF audit report
```

A default run processes **48 transactions across 12 shipments** and reports
`pass=40 warn=1 fail=7 escalations=1`. Those numbers are deterministic — the sample data
contains exactly one seeded anomaly per rule, documented in
[docs/SEEDED_ANOMALIES.md](docs/SEEDED_ANOMALIES.md).

### The six validation rules

| Rule | Checks | Seeded failure |
|---|---|---|
| R1 Timestamp | On-chain time vs the ERP's expected milestone time | `SHP-1002` warns, `SHP-1003` fails |
| R2 Quantity | Decoded on-chain quantity vs the purchase order | `SHP-1005` short by 20 units |
| R3 Whitelist | Sender is an approved partner wallet | `SHP-1007` signed by an unknown wallet |
| R4 Duplicate | No transaction hash recorded twice per shipment | `SHP-1009` delivery replayed |
| R5 Sequence | Milestones occur in the configured order | `SHP-1011` delivered before in-transit |
| R6 Contract event | Every shipment emits the required event | `SHP-1012` never confirms receipt |

---

## Layout

```
BlockchainLogisticsValidator/     the UiPath project - open this in Studio
├── Main.xaml                     phase orchestration
├── Workflows/
│   ├── 00_Init_ReadConfig        config, validation, path resolution
│   ├── 01_Extract_Blockchain     dispatches to 01a mock / 01b API / 01c web
│   ├── 01m_Normalise_ChainRecords  the field-mapping engine: any JSON feed -> one schema
│   ├── 02_Extract_Logistics      ERP master + expected milestones (+ tx map in LOOKUP)
│   ├── 03_Preprocess_Map         join chain to ERP, by decoded payload or by tx map
│   ├── 05_Validate_Engine        runs the enabled rules from Config.xlsx
│   │   └── Rules/R1..R6          one workflow per rule
│   ├── 06_HumanInTheLoop_Escalation
│   ├── 07_Report_Excel           styled workbook with conditional formatting
│   ├── 08_Report_PDF             formal audit report, printed by headless Edge
│   ├── 09_Alert_Email            composes the alert; EML / SMTP / Teams-Slack webhook
│   ├── 10_AuditLog_Chained       hash-chained audit trail
│   ├── 11_Dashboard_Summary      KPI JSON for Power BI / Sheets
│   ├── 12_Verify_AuditChain      standalone integrity check
│   ├── 13_Report_Dashboard       self-contained HTML dashboard
│   └── Queue/                    Dispatcher, Performer (own entry point), ProcessShipment
├── Tests/TC01..TC10              UiPath test cases (Test Explorer)
├── Data/
│   ├── Config.xlsx               the control panel - see below
│   ├── Input/                    ERP workbook, tx map, mock chain feeds
│   └── Output/                   reports, audit logs (generated)
build.ps1                         pack + analyze + run headlessly
tools/                            data generator, oracle, test doubles, demo scripts, acceptance suite
docs/                             plan, runbook, demo script, results
```

---

## Config.xlsx — the control panel

Everything operational lives here, so the bot adapts without touching a workflow.

| Sheet | What it controls |
|---|---|
| `Settings` | Data source mode, chain id, API endpoint and key, file paths, escalation, alerting, queue |
| `ValidationRules` | Which rules run, their severity, thresholds and fail action |
| `ApprovedWallets` | The partner whitelist R3 checks against |
| `ChainProfiles` | Per source: where the records sit in the JSON and how it reports success |
| `FieldMapping` | Per profile: source field → canonical field, with the transform applied |
| `EventSequence` | The milestone order R5 enforces, and each one's function selector |
| `EventSignatures` | Function signatures, selectors and event topic hashes |

Settings worth knowing:

- **`DataSourceMode`** — `MOCK` (default, offline), `API` (live Etherscan V2), `WEB`
  (scrapes an explorer page). All three produce byte-identical validation results.
- **`ChainProfile`** — which `ChainProfiles` row describes the feed. `ETHERSCAN` (default)
  or `LEDGER_EXPORT`, which has different field names, nesting, hex block numbers and ISO
  timestamps. A new chain or export format is a new profile, not a new workflow.
- **`MappingMode`** — `DECODE` takes shipment and milestone from the transaction's ABI call
  data; `LOOKUP` takes them from `TxShipmentMap.xlsx`, for ledgers whose payload the bot
  cannot decode.
- **`RunMode`** — `BATCH` validates everything in one job; `DISPATCH` splits it into one
  Orchestrator queue item per shipment, which `Workflows\Queue\Performer.xaml` consumes.
- **`EscalationMode`** — `PROMPT` asks a reviewer (attended only), `SIMULATE` applies a preset
  decision, `AUTO_LOG` defers. `PROMPT` automatically downgrades to `AUTO_LOG` when
  `AttendedMode` is false, so an unattended job can never hang on a dialog.
- **`AlertMode`** — comma-separated channels: `EML` (archive only, default), `SMTP`,
  `WEBHOOK` (Teams Adaptive Card, Slack Block Kit or plain JSON, per `AlertWebhookFormat`).
  The `.eml` copy is always written.
- **`PdfReportEnabled`** — prints the formal audit report to PDF with headless Edge or Chrome.

### Local overrides

`Data\Config.local.json` overrides any `Settings` value and is **not** meant to be committed.
Machine-specific output paths and real API keys belong there, so `Config.xlsx` stays portable
and safe to hand in. `build.ps1` writes it automatically.

---

## Notable design decisions

**Extraction only fetches; one engine normalises, driven by Excel.** The mock reader, the
live Etherscan reader and the web scraper each just produce JSON. `01m_Normalise_ChainRecords`
turns any of it into one canonical `DataTable`, reading *where* the records are from
`ChainProfiles` and *how* each field is derived from `FieldMapping` — ten transforms, including
hex and epoch conversion and ABI decoding of `string` / `uint256` arguments. No source field
name appears in any workflow. That is what makes the bot blockchain-agnostic rather than
merely configurable — proven by `tools\chain_swap_demo.ps1`, which reaches the same verdict on
Ethereum, Polygon and a differently-shaped ledger export with no workflow change.

**Real ledgers contain noise, and noise is not an anomaly.** Value transfers, calls to other
contracts and malformed records are counted and set aside before decoding, with a summary
line saying how many of each. Every ABI offset and length is bounds-checked, so one corrupt
record is reported by hash rather than taking the run down.

**Rules are set-based, not row-based.** Each rule is its own workflow but receives the whole
results table. Duplicates (R4), ordering (R5) and required events (R6) compare rows against
one another and cannot be evaluated a row at a time. This also means 6 workflow invocations
per run instead of 288.

**The queue work-unit is a shipment, not a transaction.** For the same reason: a performer
handed a single transaction could not evaluate half the rules. No rule reaches across
shipments, so a shipment is the smallest split that keeps every rule meaningful.

**Structure in XAML, algorithms in `Invoke Code`.** Sequences, branching, Excel and mail
activities stay visual so the project reads as real RPA in the designer; ABI decoding, hashing
and the sequence state machine live in VB code activities. This is ordinary production UiPath
practice.

**The dashboard is generated by the bot, not hand-made.** `13_Report_Dashboard` writes a
single self-contained HTML file — no server, no CDN, no external stylesheet — so it opens
with a double click on any machine and survives being emailed. Its run-history table is
read back out of the append-only audit log rather than kept separately, so the history
cannot drift from the evidence it reports on.

**Reports are written with ClosedXML, not Excel activities.** UiPath exposes no
conditional-formatting activity, and ClosedXML needs no Excel process — so the styled report
is produced identically on an unattended robot with no Office installed.

---

## Verification

| Tool | What it proves |
|---|---|
| `tools\run_all_checks.ps1` | Everything below, in one command |
| `tools\verify_seeded_data.py` | Independent oracle: the anomalies exist and the ABI decodes |
| `tools\tamper_test.ps1` | Editing the audit log is detected at the exact line |
| `tools\chain_swap_demo.ps1` | Identical results across two blockchains and a ledger export, config change only |
| `tools\run_tests.ps1` | The ten UiPath test cases |
| `tools\integration_tests.ps1` | Etherscan API (chains 1 and 137), SMTP, Teams and Slack webhooks, bad key, empty contract, WEB over HTTP — against local test doubles |
| `tools\dialog_test.ps1` | The `PROMPT` reviewer dialog, answered through Windows UI Automation |
| `build.ps1 -Analyze` | UiPath Workflow Analyzer, no Error-severity findings |

The dataset oracle is a deliberately separate implementation from the generator. If the bot
and the oracle ever disagree, one of them is wrong and the run is not signed off.

---

## Regenerating

```bash
python tools\generate_sample_data.py     # sample data (deterministic)
python tools\build_workflows.py          # regenerate the .xaml files
```

The `.xaml` files are the deliverable and are edited in Studio like any others.
`build_workflows.py` exists so the boilerplate stayed consistent while the project was being
built out — **once you start editing a workflow by hand in Studio, remove it from that
script's `GENERATORS` map**, or the next run will overwrite your changes.

---

## Testing

Ten UiPath test cases live in `Tests/` and appear in Studio's **Test Explorer**
(Test tab → Run All Tests).

| Test | Covers |
|---|---|
| TC01–TC06 | One validation rule each, against a hand-built fixture |
| TC07 | Audit chain: write, verify, tamper, detect at the right line |
| TC08 | The field-mapping engine, on two invented feeds whose field names appear nowhere in the shipped config — every transform asserted |
| TC09 | Normaliser hardening: transfers, foreign calls, truncated call data and non-objects are set aside; an unknown transform or a missing required field is rejected by name |
| TC10 | `MappingMode = LOOKUP`: the transaction map overrides the call data, case-insensitively; unmapped transactions are flagged |

They are unit tests, not a replay of the sample data — they cover edge cases the sample
data does not: a drift exactly at the tolerance boundary, a quantity tolerance that
absorbs a shortfall, a checksummed address matching a lower-case whitelist entry, and the
same transaction hash on two different shipments (which must *not* count as a duplicate).

Each asserts with `VerifyExpression` for the Test Explorer report, then throws if any
assertion failed, so a regression is also visible as a non-zero exit code in CI.

> One packaging quirk worth knowing: a *process* package deliberately excludes anything
> registered in `designOptions.fileInfoCollection`, because test cases normally ship in a
> separate test package. `tools\run_tests.ps1` therefore unregisters them just long enough
> to pack, runs them as entry points, and restores `project.json` afterwards.

## Known gaps

- **Orchestrator queue activities are untested against a live tenant.** The whole
  dispatch → perform path runs in `DRYRUN`: `Dispatcher` writes 12 shipment payloads and
  `Performer`, started as its own entry point, consumes them and reaches the same
  `total=48 pass=40 warn=1 fail=7` as a batch run. Only the three activities that talk to
  Orchestrator — `Add Queue Item`, `Get Transaction Item`, `Set Transaction Status` — have
  never run against a real queue. See [docs/ORCHESTRATOR_RUNBOOK.md](docs/ORCHESTRATOR_RUNBOOK.md).
- **API mode has not been run against the real Etherscan.** It is exercised over HTTP
  against a local Etherscan V2 test double (`tools\mock_services.py`), including error
  responses, but no API key is configured.
- **WEB mode parses the explorer's HTML rather than driving a browser.** It works and is
  covered by the acceptance suite, but it is not UiPath Data Scraping: UIAutomation 25.10
  removed the classic `ExtractStructuredData`, its replacement needs recorder-generated
  descriptors, and the UiPath browser extension is not installed on this machine. See the
  docstring in `01c_Extract_FromExplorerUI` for the upgrade path.
- **Excel Application Scope drives interactive Excel COM** for reading `Config.xlsx` and the
  ERP workbook. Fine attended; for a fully unattended robot, consider moving those reads to
  ClosedXML as the report writer already does.

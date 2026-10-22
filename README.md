# UiPath-Based Validation of Blockchain Transactions for Transparent Logistics

A UiPath RPA bot that reconciles on-chain logistics events against enterprise ERP records,
flags anomalies against six configurable rules, escalates critical findings to a human, and
writes a tamper-evident audit trail and a formal PDF audit report suitable for regulatory
submission.

**Course:** Robotic Process Automation
**Team:** Adithya Sankar Menon (23BRS1079) · Nitin Nandakumar (23BRS1378)

| Document | For |
|---|---|
| [docs/USER_GUIDE.md](docs/USER_GUIDE.md) | Running and configuring the bot; handover notes |
| [docs/DECK_TRACEABILITY.md](docs/DECK_TRACEABILITY.md) | Every Review 1 commitment → where it is implemented → how it is checked |
| [docs/RESULTS.md](docs/RESULTS.md) | Measured results |
| [docs/SEEDED_ANOMALIES.md](docs/SEEDED_ANOMALIES.md) | Which record breaks which rule |
| [docs/DEMO_SCRIPT.md](docs/DEMO_SCRIPT.md) | Review walkthrough |
| [docs/FINAL_REVIEW_SPEAKING_NOTES.md](docs/FINAL_REVIEW_SPEAKING_NOTES.md) | What to say, slide by slide, for the final review deck |
| [docs/REVIEW_DAY_CHECKLIST.md](docs/REVIEW_DAY_CHECKLIST.md) | Opening the project in Studio, and the mechanics of review day |

---

## Quick start

Open `BlockchainLogisticsValidator\project.json` in UiPath Studio and run `Main.xaml`.
That's it — the default configuration runs entirely offline against bundled sample data.

**The project needs UiPath Studio and nothing else.** There is no Orchestrator dependency: no
queues, assets or published processes. Queue-style processing uses a local work-queue file,
and every entry point is run from Studio.
Verified on Studio 25.10.15 and 26.0.204; on 26.x, decline the project-upgrade prompt unless
you mean to re-pin every dependency.

The `.ps1` commands below all pass `-ExecutionPolicy Bypass`, because a default Windows client
refuses to run scripts at all — see
[docs/USER_GUIDE.md §2.1](docs/USER_GUIDE.md#21-two-things-that-bite-on-a-fresh-machine).

To run it headlessly instead (no Studio window):

```bash
powershell -ExecutionPolicy Bypass -File "build.ps1"
```

To verify every claim this project makes, end to end:

```bash
powershell -ExecutionPolicy Bypass -File "tools\run_all_checks.ps1"
```

That suite is the single source of truth for whether the project works: 21 checks,
including the twelve UiPath test cases, API mode against a local Etherscan stand-in, and the
review window answered through UI Automation (it appears on screen for a few seconds; add
`-SkipReviewWindow` if nobody should touch the machine). On a machine where Windows Smart App Control blocks
newly built workflow assemblies, run the checks that don't need the robot:

```bash
powershell -ExecutionPolicy Bypass -File "tools\run_all_checks.ps1" -StaticOnly
```

---

## What a run does

```
Config.xlsx ──> 00 Init ──> 01 Extract chain ──> 01d Field mapping ──> 02 Extract ERP ──> 03 Map & join
                             mock / API / web      FieldMapping sheet     Excel / CSV / SQL    DECODE / LOOKUP
                                                                                                    │
                                                                                                    v
   13 Dashboard <── 11 KPIs + history <── 09 E-mail · 09b Teams/Slack <── 07 Excel · 08 PDF <── 05 Validate R1..R6
                                                                                ^                   │
                                                               12 Verify chain <── 10 Audit log <── 06 Human review
```

Steps 06–13 are `14_Finalise_Run`, shared by batch runs and by the queue Performer. Every
phase is timed; the dashboard shows the processing time per phase and trends across runs.

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
├── Main.xaml                     phase orchestration, timing, failure screenshot
├── Workflows/
│   ├── 00_Init_ReadConfig        config, defaults, fail-fast validation, path resolution
│   ├── 01_Extract_Blockchain     dispatches to 01a mock / 01b API / 01c web
│   ├── 01d_Apply_FieldMapping    raw records -> canonical transactions, per FieldMapping sheet
│   ├── 02_Extract_Logistics      ERP from Excel, CSV or SQL (Database Connect)
│   ├── 03_Preprocess_Map         shipment linking (DECODE / LOOKUP) and join to ERP
│   ├── 05_Validate_Engine        runs the enabled rules from Config.xlsx
│   │   └── Rules/R1..R6          one workflow per rule
│   ├── 06_HumanInTheLoop_Escalation   review form / dialog / simulated / deferred
│   ├── 07_Report_Excel           styled workbook with conditional formatting
│   ├── 08_Report_PDF             formal PDF audit report
│   ├── 09_Alert_Email            EML, SMTP or Outlook, with reports attached
│   ├── 09b_Alert_Webhooks        Teams Adaptive Card and Slack message
│   ├── 10_AuditLog_Chained       hash-chained audit trail
│   ├── 11_Dashboard_Summary      KPI JSON, run history, Power BI / Sheets push
│   ├── 12_Verify_AuditChain      standalone integrity check (also an entry point)
│   ├── 13_Report_Dashboard       self-contained HTML dashboard
│   ├── 14_Finalise_Run           06 -> 13, shared by batch and queue mode
│   └── Queue/Dispatcher,Performer  work-queue mode via a local file (Performer is an entry point)
├── Tests/TC01..TC12              UiPath test cases
├── Data/
│   ├── Config.xlsx               the control panel
│   ├── Input/                    ERP workbook, CSV exports, Access DB, tx map, chain feeds
│   └── Output/                   reports, audit logs, screenshots (generated)
build.ps1                         pack + analyze + run headlessly
tools/                            generators, oracle, compile check, demos, acceptance suite
docs/                             user guide, traceability, results, demo script
```

---

## Config.xlsx — the control panel

Everything operational lives here, so the bot adapts without touching a workflow. Every
setting is documented in [docs/USER_GUIDE.md §4](docs/USER_GUIDE.md#4-configxlsx-reference).

| Sheet | What it controls |
|---|---|
| `Settings` | Sources (chain, ERP), mapping, escalation, alert channels, dashboard push, work queue, output |
| `ValidationRules` | Which rules run, their severity, thresholds and fail action |
| `ApprovedWallets` | The partner whitelist R3 checks against |
| `FieldMapping` | Etherscan record field → canonical field, with the transform applied |
| `FieldMapping_EvmEtl` | The same, for an ethereum-etl export (e.g. Hyperledger Besu) |
| `EventSequence` | The milestone order R5 enforces, and each one's function selector |
| `EventSignatures` | Function signatures, selectors and event topic hashes |

The switches worth knowing:

- **`DataSourceMode`** — `MOCK` (default, offline), `API` (live Etherscan V2), `WEB`
  (scrapes an explorer page). All reach identical validation results.
- **`FieldMappingSheet`** + **`ChainRecordsPath`** — read a feed with different field names.
- **`LogisticsSourceMode`** — `EXCEL`, `CSV` or `SQL`.
- **`MappingMode`** — `DECODE` (shipment key from call data) or `LOOKUP` (mapping file).
- **`EscalationMode`** — `PROMPT` shows the review form (attended only), `SIMULATE` applies a
  preset decision, `AUTO_LOG` defers. `PROMPT` downgrades to `AUTO_LOG` when `AttendedMode`
  is false, so an unattended job can never hang.
- **`AlertMode`** — `EML` writes the message to disk (default), `SMTP` or `OUTLOOK` sends it.
  `TeamsWebhookUrl` / `SlackWebhookUrl` post to chat alongside.
- **`RunMode`** — `BATCH` validates everything in one run; `DISPATCH` writes a work-queue
  file with one item per shipment, which `Workflows\Queue\Performer.xaml` then processes.

`Data\Config.local.json` overrides any `Settings` value and is **not** committed — real API
keys, SMTP passwords and webhook URLs belong there. `build.ps1 -Set @{...}` writes it.

---

## Notable design decisions

**Readers fetch; the mapping sheet translates.** The mock, API and web readers return the
feed's records untouched. `01d_Apply_FieldMapping` builds the canonical transaction table
purely from the selected FieldMapping sheet — field names, ABI argument positions, timestamp
formats. A chain or exporter with a different schema needs a different sheet, not a
different workflow; `tools\chain_swap_demo.ps1` proves it on three feeds and two schemas.

**Every source lands in the same typed tables.** Excel, CSV and SQL ERP data are
canonicalised once in `02`, so rules never care where the data came from.

**Rules are set-based, not row-based.** Each rule is its own workflow but receives the whole
results table. Duplicates (R4), ordering (R5) and required events (R6) compare rows against
one another and cannot be evaluated a row at a time.

**The queue work-unit is a shipment, not a transaction.** For the same reason. The Performer
drains the queue in one job and then runs the same finalise pipeline as a batch run, so queue
mode produces the same audit trail and reports.

**The human decision is evidence.** Review happens before the audit log is written, so the
reviewer's decision, name and note are hashed into the chain like the machine's verdict.

**Structure in XAML, algorithms in `Invoke Code`.** Sequences, branching, Excel, database,
mail and HTTP activities stay visual; ABI decoding, hashing, PDF layout and the sequence
state machine live in VB code activities. `tools\compile_check.ps1` compiles all of that code
without running the robot.

**Reports need no Office.** The Excel report is written with ClosedXML and the PDF directly
as PDF 1.4 with the standard fonts, so an unattended robot produces identical files.

**The dashboard is generated by the bot.** One self-contained HTML file — no server, no CDN.
Run history comes from the append-only audit log and `RunHistory.csv`, not from a separate
store that could drift.

---

## Verification

| Tool | What it proves |
|---|---|
| `tools\run_all_checks.ps1` | Everything below, in one command (21 checks) |
| `tools\verify_seeded_data.py` | Independent oracle: the anomalies exist, the ABI decodes, every alternative feed and ERP source carries identical data |
| `tools\compile_check.ps1` | Every Invoke Code body compiles against the robot's references |
| `build.ps1 -Analyze` | Workflow Analyzer, no Error-severity findings |
| `tools\run_tests.ps1` | The twelve UiPath test cases |
| `tools\review_window_test.ps1` | The PROMPT review window, answered through Windows UI Automation; the decision reaches the audit chain |
| `tools\tamper_test.ps1` | Editing the audit log is detected at the exact line |
| `tools\chain_swap_demo.ps1` | Identical results on Ethereum, Polygon and a Besu ethereum-etl export |
| `tools\capture_server.ps1` | Local SMTP, webhook and Etherscan V2 stand-in used to test alert delivery and API mode offline |

The dataset oracle is a deliberately separate implementation from the generator. If the bot
and the oracle ever disagree, one of them is wrong and the run is not signed off.

---

## Regenerating

```bash
python tools\generate_sample_data.py          # sample data (deterministic)
powershell -ExecutionPolicy Bypass -File tools\build_erp_database.ps1  # SQL copy of the ERP data
python tools\build_workflows.py               # the workflow .xaml files + entry points
python tools\build_tests.py                   # the test cases
```

The final review deck is generated too. It needs Node.js:

```bash
cd tools\deck && npm install && node build_final_deck.js --date "15 November 2026"
```

The `.xaml` files are the deliverable and are edited in Studio like any others. The
generators keep argument signatures consistent across files — **once you start editing a
workflow by hand in Studio, remove it from that script's `GENERATORS` map**, or the next run
will overwrite your changes.

---

## Testing

Twelve UiPath test cases live in `Tests/` and appear in Studio's **Test Explorer**
(Test tab → Run All Tests):

| Test | Covers |
|---|---|
| TC01–TC06 | One validation rule each, against hand-built fixtures and edge cases the sample data does not have |
| TC07 | Audit chain: write, verify, tamper, detect at the right line; input-data hash recorded |
| TC08 | Field mapping: the Etherscan and ethereum-etl feeds produce identical canonical transactions; an incomplete sheet is rejected |
| TC09 | PDF audit report: header, xref offsets, objects and page tree are consistent |
| TC10 | Review form: pauses, times out to DEFERRED, attributes the timeout, keeps an evidence image |
| TC11 | Teams Adaptive Card and Slack Block Kit payloads |
| TC12 | Ledger noise: a value transfer is skipped, a call to another function is ignored, truncated call data and a non-object are rejected — with and without the function table, and in LOOKUP mode |

Each asserts with `VerifyExpression` for the Test Explorer report, then throws if any
assertion failed, so a regression is also visible as a non-zero exit code.

> A *process* package deliberately excludes anything registered in
> `designOptions.fileInfoCollection`, because test cases normally ship in a separate test
> package. `tools\run_tests.ps1` therefore unregisters them just long enough to pack, runs
> them as entry points, and restores `project.json` afterwards.

## Known gaps

- **No Orchestrator.** The project is scoped to UiPath Studio. Scheduling, triggers and
  Orchestrator queues are out of scope; the dispatcher/performer split is kept, using a local
  work-queue file instead of an Orchestrator queue.
- **Services this repository holds no credentials for.** A live Etherscan API key, a real
  SMTP server or Outlook profile, and real Teams/Slack channels. Each path is exercised
  locally — alerts through `tools\capture_server.ps1` — but the final hop is untested.
- **WEB mode parses the explorer's HTML rather than driving a browser.** UIAutomation 25.10
  removed the classic `ExtractStructuredData`, its replacement needs recorder-generated
  descriptors, and the UiPath browser extension is not installed on the build machine. See the
  docstring in `01c_Extract_FromExplorerUI` for the upgrade path.
- **The review form is built in the workflow**, not designed as a Studio `.uiform`: UiPath
  Forms 25.10 files can only be authored in Studio's visual designer. `ReviewUi = DIALOG`
  falls back to Input Dialog.
- **Excel Application Scope drives interactive Excel COM** for `Config.xlsx`, the ERP
  workbook and the mapping file, so the robot machine needs Excel. The ERP data itself can
  come from CSV or SQL instead.
- **The audit log is single-writer.** Robots running in parallel each need their own output root.
- **Verification status of this release** is recorded at the end of
  [docs/DECK_TRACEABILITY.md](docs/DECK_TRACEABILITY.md#verification-status).

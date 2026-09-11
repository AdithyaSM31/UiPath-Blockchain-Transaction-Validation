# Implementation Plan
## UiPath-Based Validation of Blockchain Transactions for Transparent Logistics

**Team:** Adithya Sankar Menon (23BRS1079), Nitin Nandakumar (23BRS1378)
**Course:** Robotic Process Automation
**Plan date:** 10 September 2026

---

## 0. Verified environment

Checked on this machine before planning — everything needed is already installed, no downloads required:

| Item | Status |
|---|---|
| UiPath Studio | **25.10.15** (`%LOCALAPPDATA%\Programs\UiPath\Studio`) |
| Project target | Windows (.NET 8) — modern, not Windows-Legacy |
| UiPath.System.Activities | 26.2.7 — Deserialize JSON, Invoke Code, For Each Row |
| UiPath.WebAPI.Activities | 2.4.0 — **HTTP Request** (Etherscan API) |
| UiPath.Excel.Activities | 3.4.1 — Read/Write Range, formatting |
| UiPath.Mail.Activities | 2.7.12 — SMTP / Outlook alerts |
| UiPath.Form.Activities | 25.10.2 — **human-in-the-loop escalation form** |
| UiPath.UIAutomation.Activities | 25.10.33 — browser scraping fallback |
| UiPath.Word.Activities | 2.4.1 — PDF audit report export |
| UiPath.Testing.Activities | 25.10.1 — automated test cases |
| `UiRobot.exe execute` / `pack` | available — headless run + `.nupkg` build |
| `UiPath.Studio.CommandLine.exe analyze` | available — Workflow Analyzer from CLI |

**Why this matters:** the last two rows give a real verification loop. Every workflow can be
executed and analysed from the command line, so nothing gets handed over "looks right in the
designer" — it gets handed over having actually run.

---

## 1. What the deck promises, and how each promise gets met

| Deck claim | Concrete deliverable |
|---|---|
| Phase 1 — Blockchain extraction (API **or** web scraping) | `01a` mock, `01b` Etherscan REST, `01c` browser scraping — all three, selected by config |
| Phase 1 — Logistics extraction | `02_Extract_Logistics.xaml` from Excel; SQL path stubbed with the same output schema |
| Phase 2 — Preprocessing & mapping | `03_Preprocess_Normalize.xaml`, `04_Map_TxToShipment.xaml` |
| Phase 3 — Six-rule validation engine | `05_Validate_Engine.xaml` + `Rules/R1..R6.xaml`, PASS/FAIL/WARNING per record |
| Phase 4 — Excel + PDF + email + dashboard | `07`–`09`, `11` |
| Phase 5 — Orchestrator, schedules, queues | `Queue/Dispatcher.xaml`, `Queue/Performer.xaml`, published `.nupkg`, asset/trigger runbook |
| Novelty 1 — no-code rule config | `Config.xlsx` → `ValidationRules` sheet drives which rules run, severity, thresholds |
| Novelty 2 — human-in-the-loop escalation | `06_HumanInTheLoop_Escalation.xaml` — UiPath Form pauses bot on CRITICAL anomaly |
| Novelty 3 — blockchain agnosticism | `Settings` + `FieldMapping` sheets; swap chain by editing Excel, zero workflow edits |
| Novelty 4 — tamper-evident audit log | `10_AuditLog_Chained.xaml` (SHA-256 chained entries) + `12_Verify_AuditChain.xaml` |

---

## 2. Solution architecture

```
BlockchainLogisticsValidator/          <- UiPath Windows project
├── project.json
├── Main.xaml                          <- phase orchestration, Try/Catch, run summary
├── Workflows/
│   ├── 00_Init_ReadConfig.xaml            out: Config dictionary + rule/wallet/mapping tables
│   ├── 01_Extract_Blockchain.xaml         switch: MOCK | API | WEB   -> dtChain
│   │   ├── 01a_Extract_FromMockJson.xaml
│   │   ├── 01b_Extract_FromEtherscanApi.xaml
│   │   └── 01c_Extract_FromExplorerUI.xaml
│   ├── 02_Extract_Logistics.xaml          -> dtERP
│   ├── 03_Preprocess_Normalize.xaml       epoch->DateTime, address lowercase, wei->qty
│   ├── 04_Map_TxToShipment.xaml           -> dtJoined
│   ├── 05_Validate_Engine.xaml            For Each Row -> invoke enabled rules
│   │   └── Rules/ R1_TimestampCheck … R6_SmartContractEvent
│   ├── 06_HumanInTheLoop_Escalation.xaml
│   ├── 07_Report_Excel.xaml               green/red/amber conditional formatting
│   ├── 08_Report_PDF.xaml                 formal audit document
│   ├── 09_Alert_Email.xaml                SMTP, summary table of failures
│   ├── 10_AuditLog_Chained.xaml           SHA-256 chain-of-custody
│   ├── 11_Dashboard_Summary.xaml          pass %, fail %, top anomaly types
│   ├── 12_Verify_AuditChain.xaml          proves tamper-evidence (demo showpiece)
│   └── Queue/ Dispatcher.xaml, Performer.xaml
├── Data/
│   ├── Config.xlsx                    <- the "no-code" control panel
│   ├── Input/ LogisticsRecords.xlsx, TxShipmentMap.xlsx, MockChain/*.json
│   └── Output/ Reports/, AuditLogs/, Screenshots/
├── Tests/                             <- UiPath Testing Activities cases
└── Documentation/                     <- user guide, demo script, architecture diagram
```

### `Config.xlsx` sheets

| Sheet | Purpose |
|---|---|
| `Settings` | `DataSourceMode` (MOCK/API/WEB), chain id, API base URL + key, wallet under audit, file paths, SMTP host/port/from/to, `EscalationEnabled`, `TimestampToleranceHours` |
| `ValidationRules` | `RuleID, RuleName, Enabled, Severity, Param1, Param2, FailAction` — one row per rule |
| `ApprovedWallets` | `Address, PartnerName, Role` — the whitelist |
| `FieldMapping` | `BlockchainField, LogisticsField, TransformType` — what makes the bot chain-agnostic |
| `EventSequence` | ordered `Dispatched → InTransit → Delivered → GoodsReceived` |
| `EventSignatures` | event name → topic0 hash, for the smart-contract-event rule |

### Design decisions worth defending in the viva

1. **Etherscan-shaped JSON is the single internal contract.** The mock reader, the live API
   reader and the UI scraper all emit the same `DataTable` schema, so switching source modes
   changes nothing downstream. This is what makes Novelty 3 real rather than a slide claim.
2. **Structure in XAML, algorithms in `Invoke Code` (C#).** Sequences, If, For Each Row,
   Try/Catch, Excel and Mail activities stay visual so it reads as genuine RPA in the designer;
   hashing, the sequence state machine and JSON field extraction go in `Invoke Code`. This is
   normal production UiPath practice, not a shortcut.
3. **Ship offline-first.** `DataSourceMode = MOCK` is the default, so the demo cannot be broken
   by campus Wi-Fi, a rate limit, or an expired API key on presentation day. `API` mode is
   built and tested, and is one cell edit away.
4. **Seeded anomalies.** The sample dataset deliberately contains one failure per rule, so all
   six rules visibly fire in a single run — a clean demo instead of an all-PASS anticlimax.

---

## 3. Build stages

Each stage ends with something runnable. Stage exit criteria are checks I will actually run,
not "looks done".

### Stage 0 — Scaffold and prove the toolchain
Create the Windows project, dependency set, folder tree, and a trivial `Main.xaml`.
**Exit:** `UiRobot.exe execute` runs it and `UiPath.Studio.CommandLine.exe analyze` returns clean.
*This is the gate — if hand-authored XAML doesn't load, we find out here, not in week three.*

### Stage 1 — Data foundation
`Config.xlsx` (all six sheets), `LogisticsRecords.xlsx` (~40 shipments), Etherscan-shaped mock
JSON, `TxShipmentMap.xlsx`. Seed exactly one anomaly per rule plus a clean majority.
**Exit:** every seeded anomaly documented in a traceability table (which row breaks which rule).

### Stage 2 — Phases 1 & 2 (extraction + preprocessing)
`00`, `01`, `01a`, `01b`, `02`, `03`, `04`.
**Exit:** run in MOCK mode produces a joined `DataTable` with normalized timestamps and
lowercase addresses; log shows row counts at each hop. `01c` (browser scraping) built and
demonstrated separately.

### Stage 3 — Phase 3 (validation engine)
`05` plus `R1`–`R6`, driven entirely by the `ValidationRules` sheet.
**Exit:** all six rules fire on their seeded rows; disabling a rule in Excel and re-running
skips it with no workflow change. Results table carries `ValidationStatus` + `FailureReason`.

### Stage 4 — Phase 4 (reporting, alerting, audit)
`07` styled Excel, `08` PDF, `09` email, `10` chained audit log, `11` dashboard summary,
`12` chain verifier.
**Exit:** reports generated on disk; `12` returns VALID on a clean log and **INVALID at the
exact row** after a cell is tampered with. That tamper demo is the strongest thing in the deck.

### Stage 5 — Novelty features
`06` HITL escalation form; chain-swap walkthrough (Ethereum → Polygon by editing `Settings`).
**Exit:** unapproved-wallet row pauses the bot, shows the form, and honours Approve/Reject.

### Stage 6 — Phase 5 (Orchestrator)
`Queue/Dispatcher.xaml` + `Queue/Performer.xaml`, `UiRobot pack` to `.nupkg`, and a written
runbook for assets, queue creation, time trigger and monitoring.
**Exit:** package builds; queue workflows run against a connected tenant, or against a local
simulation plus the runbook if no tenant is available.

### Stage 7 — Tests, documentation, Review 2 deliverables
UiPath test cases per rule, README, user guide, demo script with timings, architecture diagram,
results/screenshots, and an updated presentation covering implementation and results.
**Exit:** a clean-machine dry run following only the README.

---

## 4. Suggested split between the two of you

Not binding, but the seams are clean:

| | Adithya | Nitin |
|---|---|---|
| Primary | Stages 2 & 3 — extraction, preprocessing, validation engine + rules | Stages 4 & 5 — reporting, alerting, audit chain, HITL form |
| Secondary | Stage 6 Orchestrator + packaging | Stage 1 datasets + Stage 7 tests & docs |
| Shared | Stage 0 scaffold, architecture decisions, demo rehearsal | |

---

## 5. Open decisions

These change what gets built or how far it can be verified:

1. **Etherscan API key** — free key + internet lets me test `API` mode against the live chain.
   Without one, `API` mode is built and code-reviewed but only exercised against a mock HTTP
   response. Note: Etherscan moved to the V2 multichain endpoint
   (`api.etherscan.io/v2/api?chainid=1&...`); the base URL is a config cell, so either works.
2. **UiPath Orchestrator / Automation Cloud tenant** — a community tenant makes Stage 6 a live
   deployment. Without one it becomes a package + runbook + local queue simulation.
3. **Email alerting** — a Gmail app password enables real SMTP sends. Otherwise `09` writes the
   composed message to disk as `.eml`, which still demos fine and is safer to leave in a repo.

Defaults if unanswered: MOCK data source, package-and-runbook for Orchestrator, `.eml` on disk
for email. All three are single-cell config flips later.

---

## 6. Risks

| Risk | Mitigation |
|---|---|
| Hand-authored XAML rejected by Studio | Stage 0 gate; every workflow analysed and executed via CLI as it's written |
| Etherscan rate limits / key expiry mid-demo | MOCK is the default path; live API is opt-in |
| Orchestrator trial expiry before evaluation | `.nupkg` + runbook + screenshots captured while the tenant is live |
| Form activity behaviour in unattended runs | `EscalationEnabled` config flag; unattended runs log-and-continue instead of blocking |
| Scope creep from the "optional dashboard" | Dashboard is a summary sheet in the Excel report; Power BI stays explicitly optional |

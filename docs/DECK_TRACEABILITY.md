# Review 1 deck → implementation traceability

Every commitment made in *RPA Project Review 1.pptx*, where it lives in the project, and
how it is checked. Checks marked **suite #n** are in `tools\run_all_checks.ps1`; **TCnn** are
the UiPath test cases under `BlockchainLogisticsValidator\Tests`.

Status key:

| | Meaning |
|---|---|
| **Done** | Implemented and covered by at least one check below |
| **Done – external** | Implemented; the final step needs a service this repository cannot hold credentials for (an Etherscan API key, a real mail server or Teams channel). The code path is exercised locally by the named check. |
| **Partial** | Implemented differently from the slide's wording; the difference is stated |
| **Descoped** | Deliberately out of scope for the final version, with the reason stated |

**Scope change since Review 1:** the project runs in UiPath Studio only. No Orchestrator
tenant is available to it, so everything that needed Orchestrator — deployment, time
triggers, Orchestrator queues and monitoring — is descoped. The dispatcher / performer
pattern is kept, using a local work-queue file instead of an Orchestrator queue.

See [Verification status](#verification-status) at the end for which checks have been run
on which machine.

---

## Slide 3 — Objectives

| # | Objective | Implementation | Checked by | Status |
|---|---|---|---|---|
| 1 | Extract hash, block, timestamp, sender/receiver, payload from explorer portals or REST APIs | `01a` mock feed, `01b` Etherscan V2 API (HTTP Request + Deserialize JSON), `01c` explorer page; all → `01d_Apply_FieldMapping` | suite #4, #9, #11, TC08 | Done |
| 2 | Reconcile against shipment IDs, quantities, timestamps in Excel / CSV / ERP | `02_Extract_Logistics` with `LogisticsSourceMode` = EXCEL, CSV or SQL; `03_Preprocess_Map` joins chain ↔ ERP | suite #4, #12, #13 | Done |
| 3 | Configurable rules: timestamp, quantity, whitelist, smart-contract event | `05_Validate_Engine` + `Rules\R1..R6`, driven by the `ValidationRules` sheet | suite #4, #5, TC01–TC06 | Done |
| 4 | Flag mismatches, exception reports, email / Teams / Slack alerts | Exceptions table; `09_Alert_Email` (EML, SMTP, Outlook) and `09b_Alert_Webhooks` (Teams Adaptive Card, Slack Block Kit) | suite #15, TC11 | Done – external (real mailbox / channel URLs) |
| 5 | Timestamped validation logs (PDF / Excel) as compliance evidence | `10_AuditLog_Chained` (CSV), `07_Report_Excel`, `08_Report_PDF` | suite #7, #10, TC07, TC09 | Done |
| 6 | Dashboard: pass/fail rates, anomaly trends, processing times | `11_Dashboard_Summary` (KPI JSON, `RunHistory.csv`, phase timings), `13_Report_Dashboard` (HTML with processing-time table and trend chart) | suite #6 | Done |

## Slide 4 — Architecture

| Layer | Deck | Implementation | Status |
|---|---|---|---|
| 1 | Orchestrator: scheduling, triggers, monitoring, logs | Studio only: three entry points run from Studio (`Main`, `Queue\Performer`, `12_Verify_AuditChain`); every phase writes a tagged log line, and each run is recorded in `RunHistory.csv` and the dashboard | **Descoped** — no Orchestrator tenant |
| 2 | Data extraction, preprocessing & mapping, validation engine, reporting | `01*`/`02`, `01d`/`03`, `05`+`Rules`, `07`/`08`/`09`/`11`/`13` | Done |
| 3 | Explorer/API, ERP/Excel/SQL, Excel & PDF reports, Email/Teams alerts | as above | Done |

## Slide 5 — Phases 1 & 2

| Deck item | Implementation | Checked by | Status |
|---|---|---|---|
| Option A — browser automation of an explorer, Data Scraping Wizard | `01c_Extract_FromExplorerUI` fetches the explorer page and parses its transaction table by column header | suite #11 | **Partial** — HTML parsing, not a recorded browser scrape. UIAutomation 25.10 removed the classic Data Scraping activity and the UiPath browser extension is not installed on the build machine; the upgrade path is documented in the workflow. Option B (API) is fully implemented. |
| Option B — HTTP Request to `api.etherscan.io … action=txlist`, Deserialize JSON | `01b_Extract_FromEtherscanApi` (Etherscan **V2** endpoint; V1 is retired) | suite #18 (HTTP request to a local Etherscan V2 stand-in), TC08 | Done – external (live API key) |
| Extracted data stored in a DataTable | `01d` output `dtChain` | TC08 | Done |
| Logistics: Excel Application Scope / Read Range | `02` EXCEL branch, sheet names from Config | suite #4 | Done |
| Logistics: Database Connect for ERP/SQL | `02` SQL branch: **Database Connect → Execute Query ×2 → Disconnect** (UiPath.Database.Activities), ODBC to `LogisticsERP.accdb` by default | suite #13 | Done |
| Timestamp normalisation (`epoch.AddSeconds`, `DateTime.ParseExact`) | `01d` EPOCH_TO_DATETIME / ISO_TO_DATETIME; `01c` ParseExact; `02` typed ERP timestamps | TC08 | Done |
| Address standardisation (lower-case) | `01d` LOWERCASE transform; `02` ERP addresses; R3 compares lower-case | TC03, oracle | Done |
| Field mapping: TxHash → ShipmentID via input-data key **or** mapping file | `MappingMode` DECODE (ABI arguments via FieldMapping) or LOOKUP (`TxShipmentMap.xlsx` in `03`) | suite #14, TC08 | Done |

## Slide 6 — Validation engine

| Rule | Deck condition | Implementation | Checked by | Status |
|---|---|---|---|---|
| Engine | If/Else, Switch, For Each Row; PASS/FAIL/WARNING; failures → Exceptions table with reason | `05`: For Each Row over rules, If enabled, Switch on RuleID; roll-up to ValidationStatus + FailureReasons; Exceptions table | suite #4 | Done |
| Timestamp | \|chain − ERP\| ≤ 2 h | R1, Param1 = 2 h pass, Param2 = 4 h warning ceiling | TC01 | Done |
| Quantity | on-chain = PO quantity | R2, Param1 tolerance | TC02 | Done |
| Address whitelist | sender ∈ approved partners | R3 against `ApprovedWallets` | TC03 | Done |
| Duplicate | no duplicate TxHash for the same shipment | R4 | TC04 | Done |
| Sequence | Dispatched → InTransit → Delivered | R5 against `EventSequence` | TC05 | Done |
| Smart-contract event | expected event (GoodsReceived) emitted | R6, matched by function selector | TC06 | Done |

## Slide 7 — Phases 4 & 5

| Deck item | Implementation | Checked by | Status |
|---|---|---|---|
| Styled Excel report, green/red/yellow | `07_Report_Excel` (ClosedXML) | suite #4 | Done |
| PDF audit log with timestamps, bot ID, full results | `08_Report_PDF`: run metadata and bot identity, outcome, rules, findings with human decisions, every transaction, audit-chain evidence, attestation block | suite #7, TC09 | Done |
| Send SMTP Mail / Send Outlook Mail with a summary table of failures | `09_Alert_Email`: `AlertMode` SMTP (SSL/STARTTLS or plain, report attachments) or OUTLOOK (send or draft); EML to disk by default | suite #15 (SMTP via local capture server) | Done – external (real mail server / Outlook profile) |
| Optional dashboard: totals, pass %, fail %, top anomaly types → Power BI / Google Sheets via API | `11_Dashboard_Summary`: `dashboard_summary.json` + HTTP POST to `DashboardPushUrl` (Power BI push rows or full JSON for an Apps Script) | suite #15 | Done – external (dataset URL) |
| Deploy to Orchestrator | — | — | **Descoped** — Studio only |
| Time triggers (every 4 h / daily 06:00) | — | — | **Descoped** — needs Orchestrator |
| Queue-based processing | `RunMode = DISPATCH` → `Queue\Dispatcher` writes a local work-queue file (one item per shipment); `Queue\Performer`, its own entry point, processes it and runs the full finalise pipeline | suite #16 | **Partial** — a local work-queue file instead of an Orchestrator queue |
| Monitor & audit in Orchestrator | Studio's Output panel and execution log; every phase logs a tagged line; `RunHistory.csv` and the dashboard track every run | suite #6 | **Partial** — monitored in Studio, not Orchestrator |

## Slide 8 — Novelties

| Novelty | Deck claim | Implementation | Checked by | Status |
|---|---|---|---|---|
| 1 | Rules in an external Excel config, no workflow edits | `Config.xlsx` → `ValidationRules`, `ApprovedWallets`, `EventSequence`, `FieldMapping*`, `Settings`; every Settings row is read by the bot | suite #5 | Done |
| 2 | Attended escalation: bot pauses on a high-severity anomaly, shows an interactive form, resumes on the human decision | `06_HumanInTheLoop_Escalation`: escalates on `FailAction = ESCALATE` or severity ≥ `EscalationSeverity`; `EscalationMode = PROMPT` shows the review form (every fact, decision options, reviewer name, note, auto-timeout, evidence image); decision hashed into the audit log | suite #8, #20, TC10 | **Partial** — the form is a WPF window built in the workflow rather than a Studio-designed `.uiform` (UiPath Forms 25.10 files can only be authored in Studio's visual designer). `ReviewUi = DIALOG` falls back to Input Dialog. |
| 3 | Any chain by changing the API endpoint and field mappings | `EtherscanBaseUrl`, `ChainId`, `ChainRecordsPath`, `FieldMappingSheet` + `01d` generic mapper; demo feeds for Ethereum, Polygon and a Hyperledger Besu ethereum-etl export with different field names | suite #9, TC08 | Done |
| 4 | Log with bot identity, timestamp, input-data hash, result, and hash of previous entry | `10_AuditLog_Chained`: `BotIdentity`, `TimestampUtc`, `InputDataHash`, `ValidationStatus`, `RecordHash`, `PrevHash`, `EntryHash`; `12_Verify_AuditChain` re-walks it | suite #10, TC07 | Done |

## Slide 9 — Technology stack

| Component | Deck | Used |
|---|---|---|
| RPA platform | UiPath Studio / Community | Studio 25.10 and 26.0, Windows (.NET 8) project |
| Orchestration | Orchestrator (Cloud) | Descoped — run from Studio; local work-queue file |
| Blockchain source | Etherscan API / Explorer UI | Etherscan V2 API, explorer page, bundled mock feeds |
| Logistics data | Excel / CSV / SQL | All three |
| Rules config | Excel rule file | `Config.xlsx` |
| Reporting | Excel, PDF, Email (SMTP/Outlook) | All |
| Alerting | Email, Teams webhook | Email, Teams, Slack |
| Dashboard | Power BI / Google Sheets | HTTP push to either; self-contained HTML dashboard |
| Version control | Git | Git |

## Slide 11 — 14-week plan deliverables

| Weeks | Deliverable | Where |
|---|---|---|
| 1–2 | Plan, requirements, design | [IMPLEMENTATION_PLAN.md](IMPLEMENTATION_PLAN.md) |
| 3–4 | Ingestion workflow, normalised DataTables | `00`–`02`, `01d` |
| 5–6 | Preprocessing and mapping layer | `01d`, `03`, `FieldMapping*` sheets, `TxShipmentMap.xlsx` |
| 7–9 | Six-rule engine | `05`, `Rules\` |
| 10–11 | Excel + PDF, audit logs, alerts, dashboard | `07`–`13` |
| 12 | Orchestrator deployment, schedules, queues | Descoped; `Queue\` dispatcher / performer over a local work-queue file |
| 13 | Tests and refinement | `Tests\TC01–TC12`, `tools\run_all_checks.ps1` |
| 14 | User guide, technical docs, handover notes | [USER_GUIDE.md](USER_GUIDE.md), [README.md](../README.md), this file |

## Slide 12 — Expected outcomes

| Outcome | Evidence |
|---|---|
| 90%+ reduction in manual verification effort | [RESULTS.md](RESULTS.md) §1 — 48 transactions in ~22 s against an estimated 2–3 min each by hand |
| Anomaly detection within minutes | A run completes in well under a minute |
| Audit-ready reports generated automatically | Excel, PDF, hash-chained CSV every run |
| Configurable, reusable bot | Novelties 1 and 3 above |
| Scalable via Orchestrator queues | Dispatcher/Performer split per shipment over a local work-queue file, suite #16 — the unit of work is ready for a queue, Orchestrator itself is descoped |

Slides 10 and 13 (future scope) describe later work, not Review 1 commitments.

---

## Verification status

**22 Oct 2026, development laptop, UiPath Studio and Robot 26.0.204, Studio only (no
Orchestrator):** the full acceptance suite passed **21/21**, including all twelve test cases
(TC01–TC12), API mode against a local Etherscan stand-in (suite #18) and the review window
answered through UI Automation (suite #20). Details and the history of earlier runs are in
[RESULTS.md](RESULTS.md).

To re-run everything:

```bash
powershell -ExecutionPolicy Bypass -File "toolsun_all_checks.ps1"
```

On a machine where Windows Smart App Control blocks newly built workflow assemblies, add
`-StaticOnly` to run the three checks that don't need the robot.

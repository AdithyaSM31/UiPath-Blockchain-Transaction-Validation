# Orchestrator Runbook — Phase 5

Deploying `BlockchainLogisticsValidator` to UiPath Automation Cloud: publishing, scheduling,
queue processing and monitoring.

> **Status.** The Robot on the development machine is already connected to Automation Cloud —
> the execution log identifies it as `<your-account>@vitstudent.ac.in-attended`, under an
> organisation that already applies policy rules to the project. So there is nothing to
> install; what remains is publishing and configuring.
>
> Everything in §1–§3 is standard and low-risk. **§4 (queues) contains the one part of this
> project that has never been executed** — see the warning there.

---

## 0. Before you start

| | |
|---|---|
| Cloud URL | https://cloud.uipath.com |
| Sign in as | the account Studio is already licensed with |
| Tenant | your Community tenant (usually `DefaultTenant`) |
| Folder | `Shared` is fine for coursework |

Confirm the connection: in **UiPath Assistant**, the account name should appear at the
top and the status should read *Connected*. If it does not, in Assistant use
**Preferences → Orchestrator Settings → Sign In**.

---

## 1. Publish the package

> **Republish required.** Version **1.0.92** was published to the tenant feed on 11 Sep 2026.
> This release adds the PDF report, Teams/Slack alerts, the SQL/CSV ERP sources, the field
> mapping layer, the standalone Performer and the `UiPath.Database.Activities` dependency, so
> publish the current version before creating processes. Steps 2 onwards are done in the web
> UI — the CLI publishes the package but cannot create a Process, Queue or Trigger.
>
> Republish after any change with:
>
> ```
> UiPath.Studio.CommandLine.exe publish -p <project.json> -g OrchestratorTenant -o Process
> ```

**From Studio (recommended)**

1. Open `BlockchainLogisticsValidator\project.json`.
2. **Publish** on the Design ribbon.
3. Publish to **Orchestrator Tenant Processes Feed** (or your folder's package feed).
4. Give it a version note, e.g. `Review 2 build`. Publish.

**From the command line**

```bash
powershell -File "build.ps1" -SkipRun
```

That writes a `.nupkg` to `%TEMP%\blv-packages\`. Upload it in Orchestrator under
**Tenant → Packages → Upload**.

> **Delete `Data\Config.local.json` and empty `Data\Output` before publishing for grading.**
> The first carries machine-specific paths (and any API key, SMTP password or webhook URL you
> added); the second holds local run outputs. Neither is part of the deliverable.
>
> **Build and run on machines without Windows Smart App Control.** Smart App Control can block
> the compiled, unsigned workflow assemblies inside the package, so a robot machine where it
> is on cannot execute the process.

---

## 2. Create the process

**Automations → Processes → Add process**

| Field | Value |
|---|---|
| Package | `BlockchainLogisticsValidator` |
| Version | the one you just published |
| Display name | `Blockchain Logistics Validator` |
| Job priority | Normal |

---

## 3. Set the output location, then schedule

### 3.1 Output location (do this first)

A published package is extracted to a fresh, version-named folder on every run. Without an
explicit output root, reports land there and are effectively lost, and the audit chain can
never span runs because its file disappears with each new version.

Set the environment variable **`BLV_OUTPUT_ROOT`** on the robot machine to a stable path:

```powershell
[Environment]::SetEnvironmentVariable("BLV_OUTPUT_ROOT", "C:\RPA\BlockchainValidator\Output", "Machine")
```

Then **restart the UiPath Robot service** so it picks the variable up — the service captures
its environment at start, which is why setting the variable in a shell has no effect on jobs
it launches.

`00_Init_ReadConfig` honours it ahead of everything else. Its precedence is:

```
BLV_OUTPUT_ROOT  >  Config.local.json  >  Config.xlsx Settings  >  project-relative default
```

### 3.2 Trigger

**Automations → Triggers → Add trigger → Time**

| Field | Value |
|---|---|
| Process | Blockchain Logistics Validator |
| Timezone | your local zone |
| Recurrence | every 4 hours (or daily 06:00 — the deck proposes both) |
| Stop if job overruns | on, 3 hours |

### 3.3 Unattended safety

Before the first unattended run, set these in `Config.xlsx`:

| Setting | Value | Why |
|---|---|---|
| `AttendedMode` | `False` | Downgrades `PROMPT` to `AUTO_LOG` so no job blocks on a dialog nobody can answer |
| `EscalationMode` | `AUTO_LOG` | Escalations are recorded as `DEFERRED` and reported, not waited on |
| `AlertMode` | `SMTP` | Only if you have configured credentials; otherwise leave `EML` |

---

## 4. Queue-based processing (high volume)

> ### ⚠️ The live queue hop is untested
>
> Everything except the three Orchestrator queue activities is exercised locally:
> `RunMode = DISPATCH` with `QueueMode = DRYRUN` writes the per-shipment payloads, and running
> `Workflows\Queue\Performer.xaml` in DRYRUN reads them back, validates every shipment, and
> produces the full audit log, reports, alerts and dashboard (acceptance suite check #16,
> expected `[PERFORM] total=48 pass=40 warn=1 fail=7 escalations=1`). `Add Queue Item`,
> `Get Transaction Item` and `Set Transaction Status` compile and pass the Workflow Analyzer
> but have **not run against a live Orchestrator queue**, because that needs a tenant.
> Budget time for that first live run.

### 4.1 Why a shipment, not a transaction

The queue work-unit is one **shipment**, carrying all of its events.

Three of the six rules compare rows against one another — duplicate detection (R4), milestone
ordering (R5) and required events (R6). A performer handed a single transaction could not
evaluate any of them. A shipment is the smallest unit that keeps every rule meaningful, and
no rule reaches across shipments, so nothing is lost by splitting there.

### 4.2 Create the queue

**Automations → Queues → Add queue**

| Field | Value |
|---|---|
| Name | `BlockchainValidationQueue` (must match `QueueName` in `Config.xlsx`) |
| Unique reference | **Yes** — the reference is `ShipmentID\|RunId`, so this blocks accidental double-dispatch |
| Auto retry | Yes, 2 retries |

### 4.3 Two processes

| Process | Entry point | Configuration |
|---|---|---|
| `BLV Dispatcher` | `Main.xaml` | `RunMode = DISPATCH`, `QueueMode = ORCHESTRATOR` |
| `BLV Performer` | `Workflows\Queue\Performer.xaml` | `QueueMode = ORCHESTRATOR` (it reads the same `Config.xlsx`) |

Both entry points are registered in `project.json`, so they appear in the entry-point list
when the process is created. The Performer needs no arguments: started on its own it loads
the configuration itself.

Each queue item carries:

| Key | Contents |
|---|---|
| `ShipmentID` | e.g. `SHP-1007` |
| `EventCount` | number of on-chain events for that shipment |
| `EventsJson` | the shipment's joined transactions, serialised |
| `RunId` | the dispatching run, for traceability |

### 4.4 Trigger the Performer from the queue

**Add trigger → Queue**

| Field | Value |
|---|---|
| Queue | `BlockchainValidationQueue` |
| Process | `BLV Performer` |
| Min items | 1 |
| Max pending jobs | 1 per output root (see below) |

One Performer job **drains the queue**: it loops `Get Transaction Item` until the queue is
empty, validates each shipment, marks each item Successful or Failed (with the exception
message), and then runs `14_Finalise_Run` once over everything it processed — human review,
audit log, chain verification, Excel and PDF reports, alerts and the dashboard. A shipment
that faults is reported to Orchestrator and does not stop the rest.

**Parallelism and the audit trail.** The hash-chained audit log is single-writer. To process
in parallel, run Performers on different robot machines, each with its own `BLV_OUTPUT_ROOT`
(§3.1) — every machine then keeps its own verifiable chain. Two jobs appending to one file
would interleave entries and the next verification would report a broken link.

### 4.5 Rehearse without a tenant

```bash
powershell -File "build.ps1" -Set @{ RunMode = "DISPATCH" }
powershell -File "build.ps1" -Entry "Workflows\Queue\Performer.xaml"
```

The first writes `Reports\QueueDryRun\queue_payloads_<run>.json`; the second processes that
file exactly as it would process the live queue.

---

## 5. Monitoring

| Where | What to look at |
|---|---|
| **Automations → Processes → BLV Performer** | Jobs per queue drain; the last log line of each reports the shipments processed and failed |
| **Automations → Jobs** | Success rate, duration, failure reasons |
| **Queues → Transactions** | Per-shipment status; failed items carry the exception message |
| **Tenant → Logs** | Every `[INIT]`, `[VALIDATE]`, `[AUDIT]`, `[VERIFY]` line the bot writes |
| Audit log file | `%BLV_OUTPUT_ROOT%\AuditLogs\ValidationAuditLog.csv` |

The single most useful monitoring signal is the last line of every job:

```
[VERIFY] VALID - N entries verified, chain intact from genesis to head <hash>
```

If that ever reads `INVALID`, the audit trail has been altered since it was written, and the
message names the exact line. Alert on it.

---

## 6. Screenshots to capture for the report

While the tenant is live, capture:

1. Packages list showing the published version
2. Process configuration
3. Trigger definition
4. A completed job with its log output
5. Queue with processed transactions (if §4 is working)
6. Job detail showing the `[VERIFY] VALID` line

Community tenants expire. Capture these before the deadline, not on the day.

---

## 7. Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `Config.xlsx is missing required setting(s)` | Sheet edited or a row deleted | Restore the row; `Settings` needs `DataSourceMode`, `LogisticsFile`, `BotIdentity`, `ContractAddress` |
| `Mock chain file not found` | Relative path resolved against the package folder | Check `MockChainFile`; it must be relative to the project root |
| Job hangs with no log output | `EscalationMode = PROMPT` on an unattended robot | Set `AttendedMode = False` |
| Reports written under `.nuget\packages\...` | `BLV_OUTPUT_ROOT` not set, or the Robot service was not restarted after setting it | See §3.1 |
| `[VERIFY] INVALID ... broken link` | The audit file was edited, truncated, or two robots appended concurrently | Give each robot its own audit folder |
| `DataSourceMode is API but EtherscanApiKey is blank` | API mode without a key | Add the key to `Config.local.json`, not `Config.xlsx` |
| `An Application Control policy has blocked this file` | Windows Smart App Control on the build/run machine | Build and run on a machine where it is off |
| `QueueMode is DRYRUN but no dispatcher payload file was found` | Performer started in DRYRUN before any dispatch | Set `QueueMode = ORCHESTRATOR` for live runs, or dispatch first |
| Performer finishes with `Nothing was processed` | Queue empty, or `QueueName` differs from the Orchestrator queue | Check the queue name in `Config.xlsx` |

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

> **Already done.** Version **1.0.92** was published to the tenant feed on 11 Sep 2026
> (adding the HTML dashboard workflow). Confirm it under **Tenant → Packages** in Orchestrator. Steps 2 onwards are
> still outstanding — the CLI publishes the package but cannot create a Process, Queue or
> Trigger; those are done in the web UI.
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

> **Delete `Data\Config.local.json` before publishing for grading.** It carries
> machine-specific paths (and would carry an API key if you added one). It is a local
> override file, not part of the deliverable.

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

> ### ⚠️ This section is untested
>
> `Dispatcher.xaml` and `Performer.xaml` compile and are structurally correct, and the
> dispatcher's real logic — grouping, serialisation and payload construction — is fully
> exercised by `RunMode = DISPATCH` with `QueueMode = DRYRUN` (verified: 48 transactions →
> 12 payloads). But `Add Queue Item`, `Get Transaction Item` and `Set Transaction Status`
> have **never been run against a live Orchestrator queue**, because that needs a tenant.
>
> Budget time to debug this. Everything else in the project has been executed end to end;
> this has not.

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

| Process | Entry point | Config |
|---|---|---|
| `BLV Dispatcher` | `Main.xaml` | `RunMode = DISPATCH`, `QueueMode = ORCHESTRATOR` |
| `BLV Performer` | `Workflows\Queue\Performer.xaml` | — |

Set the Performer's entry point under **Process → Edit → Entry point**.

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
| Max pending jobs | 3 (raise to scale out) |

The Performer as written processes **one item per job**. To drain the queue in a single job,
wrap its body in a `Do While` that loops until `Get Transaction Item` returns `Nothing` — the
"queue empty" branch already handles that case.

---

## 5. Monitoring

| Where | What to look at |
|---|---|
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

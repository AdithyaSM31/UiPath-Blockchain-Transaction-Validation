# Demo Script

A 10-minute walkthrough for the review. Timings are generous; the whole thing fits in 8
minutes at a normal pace, leaving room for questions.

**Rehearse once with the projector connected.** The one genuine risk is a dialog or a
spreadsheet opening on the wrong screen.

---

## Who speaks when

Speak to what you built — an examiner can tell, and "my teammate did that part" is a bad
answer to a direct question. Both of you should be able to answer anything, but lead on
your own half.

| Section | Lead | Why |
|---|---|---|
| 1 · The problem | **Nitin** | Sets up the thesis; no screen, so it opens on eye contact |
| 2 · Config.xlsx | **Adithya** | Feeds straight into the rules he built |
| 3 · The run | **Adithya** | Extraction, mapping and the validation engine |
| 4 · Dashboard | **Nitin** | He owns reporting and the audit trail |
| 4.5 · Excel report | **Nitin** | Same |
| 5.1 Rule toggle | **Adithya** | It is the rules engine responding to config |
| 5.2 Human-in-the-loop | **Nitin** | His workflow |
| 5.3 Chain swap | **Adithya** | It is an extraction-layer property |
| 5.4 Tamper test | **Nitin** | His workflow, and the strongest moment |
| 5.5 Tests | **Nitin** | He owns the test cases |
| 6 · Scale / queues | **Adithya** | Orchestrator and packaging |
| 7 · Close | **Either** | Whoever is steadier under time pressure |

Hand over explicitly — "Nitin will take the audit trail from here" — rather than drifting.
Two people talking over one screen reads as unprepared.

---

## Before you present

```bash
powershell -File "tools\run_all_checks.ps1"
```

Expect `16/16 checks passed` (that includes `10/10 test cases passed`). If anything is red,
fix it before the room fills. Budget about fifteen minutes for the run, and don't touch
the mouse near the end — the last check answers a reviewer dialog on screen.

Then reset to a clean state so the audit-log line counts are easy to narrate:

```bash
del "BlockchainLogisticsValidator\Data\Output\AuditLogs\ValidationAuditLog.csv"
```

Have open and ready:

- UiPath Studio with `Main.xaml` on screen
- `Data\Config.xlsx`
- A terminal in the project root
- File Explorer at `Data\Output\Reports`

---

## 1 · The problem (60s) — no computer

> "A pharmaceutical shipment moves through six companies between Frankfurt and Chennai.
> Each handoff is written to a blockchain, so the record is immutable. But immutable is not
> the same as *correct* — a tampered quantity, forged timestamp or unauthorised wallet is
> recorded just as permanently as a legitimate one. Today somebody cross-checks those records
> against the ERP by hand. At thousands of transactions a day, nobody can."

Land the distinction between **immutable** and **verified**. It is the point of the project
and the thing most likely to be probed in questions.

---

## 2 · The control panel (90s) — `Config.xlsx`

Show `Settings`, then `ValidationRules`.

> "Everything operational lives in this workbook. Which rules run, their thresholds, the
> approved-partner whitelist, which blockchain to read. A logistics manager changes behaviour
> here without opening Studio."

Point at `ApprovedWallets` and `EventSequence` briefly — they make R3 and R5 concrete.

---

## 3 · A run (2 min) — Studio

Run `Main.xaml`. Narrate the Output panel as it scrolls:

| Line | Say |
|---|---|
| `[EXTRACT/MOCK] 48 transactions decoded` | "Real ABI-encoded call data, decoded on the way in." |
| `[MAP] 48 rows joined, 0 unmatched` | "Every on-chain event matched to an ERP milestone." |
| `[VALIDATE] Applying R1..R6` | "Six rules, each its own workflow, driven by the spreadsheet." |
| `total=48 pass=40 warn=1 fail=7` | "Seven failures, one warning — and every one is a different rule firing." |
| `[ESCALATE] SHP-1007/Delivered` | "An unapproved wallet. Critical, so it goes to a human." |
| `[VERIFY] VALID` | "The bot just re-verified its own audit trail." |

---

## 4 · The dashboard (75s) — lead with this

Open `Data\Output\Reports\dashboard.html`. It was written by the bot on the run you just
watched.

Walk it top to bottom: the run's identity, the six KPI tiles, the pass/warning/fail bar,
then the green **audit trail verified** banner.

> "The bot generates this every run. It is one HTML file with nothing linked from outside —
> no server, no internet. The run history at the bottom is read back out of the audit log
> itself, so it cannot disagree with the evidence."

Scroll to **Findings**: eight rows, each naming the shipment, the milestone, the rule that
fired and the reason in plain English. Point out `SHP-1007`, which carries a human decision.

---

## 4.5 · The Excel report (45s)

Open the newest `ValidationReport_*.xlsx` — the same run, in the format a logistics team
would actually file.

- **Summary** — the run's identity, counters, and per-rule findings.
- **ValidationResults** — green/amber/red status column; scroll to a red row and read
  `FailureReasons` aloud. It is written in plain English on purpose.
- **Exceptions** — the eight rows a manager actually needs.

---

## 5 · The four novelties (4 min)

### 5.1 Rules are configuration, not code (45s)

In `Config.xlsx → ValidationRules`, set **R4 `Enabled` = FALSE**. Save. Re-run.

> "`R4 is disabled in Config.xlsx - skipped`. Failures drop from seven to five. No workflow
> was opened, nothing was republished."

**Set it back to TRUE.**

### 5.2 Human-in-the-loop (60s)

In `Config.xlsx`, set **`EscalationMode` = `PROMPT`** and **`AttendedMode` = `True`**. Re-run.

The bot pauses on `SHP-1007` and shows the reviewer dialog with the shipment, the wallet and
the finding. Choose **REJECTED**. The run continues.

> "The decision is captured before the audit log is written, so a human's judgement is hashed
> into the chain exactly like the machine's verdict."

**Set `EscalationMode` back to `SIMULATE`** afterwards.

This path is covered by `tools\dialog_test.ps1`, which answers the dialog through Windows UI
Automation and checks the decision lands in the audit chain. Still rehearse it once by hand —
the examiner will watch you click, not the script.

### 5.3 Any blockchain (45s)

```bash
powershell -File "tools\chain_swap_demo.ps1"
```

It runs three times: Ethereum, Polygon, and a ledger export with a completely different
shape. All three end `total=48 pass=40 warn=1 fail=7`.

> "Same bot, three sources, identical results. The third one is the interesting one — it has
> different field names, the block number is nested and in hex, timestamps are ISO text, and
> it contains transactions that aren't ours at all. We onboarded it by adding rows to two
> sheets in `Config.xlsx`. No workflow knows any source's field names."

Then open the `FieldMapping` sheet and point at the `LEDGER_EXPORT` rows. Point at the run's
`[NORMALISE]` line: *"52 records, 48 decoded, 2 value transfers ignored, 1 call to another
function ignored, 1 malformed and skipped"* — noise is counted, not mistaken for fraud.

### 5.4 Tamper-evident audit trail (90s) — *the strongest moment*

Open `Data\Output\AuditLogs\ValidationAuditLog.csv` in a text editor. Show the columns:
`RecordHash`, `PrevHash`, `EntryHash`.

> "Each entry carries the hash of the one before it. Same idea as the blockchain we're
> auditing, applied to the audit itself."

Then:

```bash
powershell -File "tools\tamper_test.ps1"
```

It verifies clean, flips one `FAIL` verdict to `PASS` — the exact edit somebody covering up an
anomaly would make — re-verifies, and reports:

```
INVALID - first problem at line 7: record contents were altered (RecordHash mismatch) - SHP-1005/Dispatched
```

then restores the file and verifies clean again.

> "One character changed, and the log names the exact row. You cannot quietly edit history
> without rewriting everything after it."

---

## 5.5 · Tests (30s) — optional, strong if there's time

Open the **Test** tab in Studio and run all tests. Ten go green.

> "Six of these drive one validation rule each, so a failure points at exactly one rule. The
> seventh tampers with an audit chain and asserts it's caught. The last three test the
> mapping engine — TC08 feeds it a made-up format whose field names appear nowhere in our
> config, to prove the engine has no built-in knowledge of any source. One of them found a
> real bug while I was writing it — the duplicate rule assumed a transaction hash was at
> least twelve characters."

---

## 6 · Scale (45s)

```bash
powershell -File "build.ps1" -Set @{ RunMode = "DISPATCH" }
```

> "For enterprise volume the same run splits into one Orchestrator queue item per shipment —
> 48 transactions become 12 work items processed in parallel. Shipment, not transaction,
> because three of the six rules compare events against each other."

Show `Data\Output\Reports\QueueDryRun\queue_payloads_*.json`. Then run the other half:

```bash
powershell -File "build.ps1" -Entry "Workflows\Queue\Performer.xaml"
```

> "That's the Performer — its own entry point, the way Orchestrator would start it from a
> queue trigger. It consumed the twelve payloads and reached exactly the same verdict as the
> single batch run."

Be straight about status: the whole dispatch-and-perform path is tested in dry-run; only the
three activities that talk to a live Orchestrator queue have not been run against a tenant.

---

## 7 · Close (30s)

> "Automated reconciliation of blockchain records against enterprise systems, six configurable
> rules, human escalation where it matters, and an audit trail that proves it wasn't altered
> afterwards. Onboarding a new chain is a spreadsheet change. The whole thing is verified by a
> sixteen-check acceptance suite that runs in one command, and the bot writes its own dashboard
> and a PDF audit report."

---

## If something breaks live

It probably won't — the suite is green — but know the recoveries so a hiccup costs ten
seconds, not your composure.

| Symptom | Do this |
|---|---|
| Studio run hangs with no output | Excel COM is stuck from an earlier kill. Close Excel, `build.ps1` clears stray executors. |
| A script errors | Fall back to `docs/sample-output/dashboard.html` — committed output from a real run. Say "here's the output from the run I did this morning" and carry on. |
| `[VERIFY] INVALID` unexpectedly | You are mid-tamper-test. Re-run `tamper_test.ps1`; it restores the file and ends VALID. |
| Reviewer dialog does not appear | `AttendedMode` is `False` in `Config.xlsx`. Set it `True`, or just describe it. |
| Projector cuts the right of the screen | The dashboard reflows — narrow the browser window rather than scrolling sideways. |

**The universal recovery:** everything you are demonstrating is already captured in
`docs/RESULTS.md` with real numbers. If the machine misbehaves, switch to that document and
keep talking. Never debug live in front of an examiner.

---

## Questions you should expect

**"Isn't blockchain data already trustworthy?"**
Immutable, not correct. The chain guarantees the record hasn't changed since it was written —
not that it was true when written, and not that it agrees with the ERP. Every seeded anomaly
is a perfectly valid on-chain transaction.

**"Why RPA instead of a Python script?"**
The rules, thresholds and whitelist are owned by logistics operations, not developers. A
spreadsheet plus a scheduled robot puts change control where the domain knowledge is. It also
inherits Orchestrator's scheduling, queueing, retry and audit logging for free.

**"Why hash the audit log when it's already on a blockchain?"**
The *transactions* are on-chain. The *validation* isn't — it happens off-chain, in the bot. The
chained log is what makes the verification step itself evidential rather than just a report
somebody could edit.

**"Is the data real?"**
The transactions are synthetic, but not fake in shape: genuine EIP-55 checksummed addresses,
real keccak-256 function selectors, and correctly ABI-encoded call data in an authentic
Etherscan V2 response envelope. `API` mode against the live chain uses the same parser.

**"Where are the tests?"**
Ten UiPath test cases in the Test tab — six drive one rule each against hand-built fixtures,
one tampers with an audit chain and asserts detection, and three test the mapping engine:
every transform, malformed input, and the lookup mode. They cover edge cases the sample data
doesn't: a drift exactly at the tolerance boundary, a checksummed address against a
lower-case whitelist, the same hash on two shipments, truncated call data. One of them found
a real bug while being written — R4 assumed a hash was at least 12 characters and crashed on
a malformed one. Beyond those, `integration_tests.ps1` checks what actually goes out over the
network — the API request, the email, the Teams and Slack messages — against local test
servers, and `dialog_test.ps1` answers the reviewer dialog through UI Automation.

**"How would you add a new blockchain?"**
If it has an Etherscan-style explorer: change `ChainId`, nothing else. If its export looks
different: add a row to `ChainProfiles` saying where the records are, and rows to
`FieldMapping` saying which field becomes which, with a transform. That's how the ledger
export in the chain-swap demo was onboarded. If its payload can't be decoded at all, set
`MappingMode = LOOKUP` and the logistics system's own transaction map decides which shipment
each transaction belongs to.

**"What doesn't work?"**
Three things, all about live external systems. The Orchestrator queue activities have never
run against a live queue — the package is published and the whole dispatch-and-perform path
is tested in dry-run, but `Add Queue Item`, `Get Transaction Item` and `Set Transaction
Status` are unexercised. `API` mode has only been run against a local test server that
speaks Etherscan's protocol, because we haven't configured a real API key. And `WEB` mode
parses the explorer's HTML rather than driving a browser, because UIAutomation 25.10 dropped
the classic scraping activity and the browser extension isn't installed. All listed under
"Known gaps" in the README.

*Give that answer plainly if asked. Knowing precisely what isn't finished reads as competence.*

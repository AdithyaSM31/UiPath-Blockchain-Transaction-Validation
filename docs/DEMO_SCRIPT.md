# Demo Script

A 10-minute walkthrough for the review. Timings are generous; the whole thing fits in 8
minutes at a normal pace, leaving room for questions.

**Rehearse once with the projector connected.** The one genuine risk is a dialog or a
spreadsheet opening on the wrong screen.

---

## Before you present

```bash
powershell -File "tools\run_all_checks.ps1"
```

Expect `12/12 checks passed` (that includes `7/7 test cases passed`). If anything is red,
fix it before the room fills. Budget about six minutes for the run.

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

> ⚠️ This is the one path not covered by the automated suite — a dialog can't be verified
> headlessly. Rehearse it at least once.

### 5.3 Any blockchain (45s)

```bash
powershell -File "tools\chain_swap_demo.ps1"
```

> "Same bot, run against Ethereum and then Polygon. Identical results. The only thing that
> changed was a chain ID and an endpoint — because extraction normalises every explorer
> response to one internal schema before validation ever sees it."

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

Open the **Test** tab in Studio and run all tests. Seven go green.

> "Six of these drive one validation rule each against fixtures built in the test, so a
> failure points at exactly one rule. The seventh writes an audit chain, tampers with it,
> and asserts the tampering is caught. One of them found a real bug while I was writing
> it — the duplicate rule assumed a transaction hash was at least twelve characters."

---

## 6 · Scale (45s)

```bash
powershell -File "build.ps1" -Set @{ RunMode = "DISPATCH" }
```

> "For enterprise volume the same run splits into one Orchestrator queue item per shipment —
> 48 transactions become 12 work items processed in parallel. Shipment, not transaction,
> because three of the six rules compare events against each other."

Show `Data\Output\Reports\QueueDryRun\queue_payloads_*.json`.

Be straight about status: the dispatcher's logic is tested; the Orchestrator queue activities
themselves have not been run against a live tenant.

---

## 7 · Close (30s)

> "Automated reconciliation of blockchain records against enterprise systems, six configurable
> rules, human escalation where it matters, and an audit trail that proves it wasn't altered
> afterwards. The whole thing is verified by a twelve-check acceptance suite that runs in one
> command, and the bot writes its own dashboard."

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
Seven UiPath test cases in the Test tab — six drive one rule each against hand-built
fixtures, the seventh writes an audit chain, tampers with it and asserts detection. They
cover edge cases the sample data doesn't: a drift exactly at the tolerance boundary, a
checksummed address against a lower-case whitelist, the same hash on two shipments. One of
them found a real bug while being written — R4 assumed a hash was at least 12 characters
and crashed on a malformed one.

**"What doesn't work?"**
Three things. The Orchestrator queue activities have never run against a live queue — the
package is published and the dispatcher's logic is tested in dry-run, but `Add Queue Item`
and `Get Transaction Item` are unexercised. `WEB` mode parses the explorer's HTML rather
than driving a browser, because UIAutomation 25.10 dropped the classic scraping activity
and the browser extension isn't installed. And the reviewer dialog has only been run in
simulated mode, since a dialog can't be tested headlessly. All listed under "Known gaps" in
the README.

*Give that answer plainly if asked. Knowing precisely what isn't finished reads as competence.*

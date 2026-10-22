# Final Review — what to say

Plain speaking notes for **RPA Project Final Review.pptx** (16 slides). Say it in your own
words; these are the points that must land, not a script to recite.

The deck also carries these as PowerPoint speaker notes, so they show in Presenter View.
This file is the fuller version, with the new-since-Review-2 material called out.

**Total: about 12 minutes of talking, plus 4 minutes of live demo.** If you are cut to ten,
drop slides 8 and 13 and shorten the future scope.

---

## Who leads which slide

| Slide | Lead | Why |
|---|---|---|
| 1 Cover · 3 The problem | **Nitin** | Sets up the thesis; no screen, so it opens on eye contact |
| 2 Since Review 2 | **Either** | Whoever is steadier opening on a list |
| 4 Architecture · 5 Mapping | **Adithya** | Extraction and the mapping engine are his |
| 6 Rules · 7 Results · 8 Findings | **Adithya** | The validation engine he built |
| 9 Config / Novelty 1 and 3 | **Adithya** | It is the rules engine responding to config |
| 10 Human-in-the-loop | **Nitin** | His workflow |
| 11 Audit trail | **Nitin** | His workflow, and the strongest moment |
| 12 Evidence pack | **Nitin** | He owns reporting |
| 13 Queue and Performer | **Adithya** | Orchestrator and packaging |
| 14 Verification · 15 Status | **Nitin** | He owns the tests and the honesty slide |
| 16 Close | **Either** | Whoever is steadier under time pressure |

Hand over out loud — "Nitin will take the audit trail from here". Do not drift between
speakers mid-slide; two people over one screen reads as unprepared.

---

## 1 · Cover (20s) — Nitin

> "This is the final review of the blockchain validation bot. One sentence before anything
> else: blockchain makes a record immutable, but it does not make it correct. That gap is
> the whole project."

Do not read the names and numbers off the slide. The examiner can see them.

---

## 2 · Since Review 2 (75s)

This slide exists because the examiner saw Review 2 and will ask what changed. Six things.

> "At Review 2 the pipeline worked end to end, but six commitments from Review 1 were still
> open. All six are now built."

Then go down the list, one clause each — do not explain them here, they each get their own
slide or demo later:

- **Field mapping engine** — "the readers got dumber and a spreadsheet got smarter"
- **Ledger export support** — "a third feed, with genuinely different field names"
- **LOOKUP mode** — "a second way to link a transaction to a shipment"
- **PDF audit report** — "the document you actually hand an auditor"
- **Teams and Slack alerts** — "alerts where the operations team already is"
- **Standalone Performer** — "queue mode now produces the same evidence as a batch run"

> "And the list of things we have not executed got shorter — that is slide fifteen."

---

## 3 · The problem (60s) — Nitin, no computer

> "A pharmaceutical shipment moves through six companies between Frankfurt and Chennai.
> Each handoff is written to a blockchain, so the record is immutable. But immutable is not
> the same as correct — a tampered quantity, a forged timestamp or an unauthorised wallet is
> recorded just as permanently as a legitimate one. Today somebody cross-checks those records
> against the ERP by hand. At thousands of transactions a day, nobody can."

Land **immutable vs verified**. It is the point of the project and the thing most likely to
be probed. Every anomaly the bot catches is a perfectly valid on-chain transaction.

---

## 4 · Architecture (75s) — Adithya

Four bands: ingest, validate, evidence, report. Twenty-eight workflows.

> "Main.xaml orchestrates phases and times each one. Each box is its own workflow file, so
> any one of them can be opened, read and tested on its own."

The one point worth dwelling on is the band at the bottom:

> "Steps 06 through 13 are a single workflow called 14_Finalise_Run. A normal batch run calls
> it, and so does the queue Performer. That is deliberate — it means queue mode is not a
> second code path that could drift, and it does not need its own separate evidence."

**New since Review 2:** `01d` field mapping in the ingest band, `08` PDF and `09b`
Teams/Slack in the report band, and `14_Finalise_Run` itself.

If asked why algorithms sit in `Invoke Code` rather than activities: three of the six rules
compare rows against one another, and a state machine drawn as activities would be
unreadable. Structure stays visual; algorithms are code.

---

## 5 · The mapping engine (90s) — Adithya — **new slide**

This is the most important new slide in the deck. It turns "blockchain agnostic" from a
claim into a property of the design.

> "At Review 2 we read a FieldMapping sheet from the config workbook, but nothing applied it —
> each reader built its own canonical table. That has changed. The readers now fetch and
> nothing else: they return the feed's records untouched. A single workflow, 01d, builds the
> canonical transaction table purely from whichever FieldMapping sheet the config selects —
> field names, ABI argument positions, timestamp formats, all of it."

Then the payoff:

> "So a chain or an exporter with a different schema needs a different spreadsheet sheet, not
> a different workflow. Three feeds, two schemas, identical results."

Point at the three-row table: Ethereum, Polygon and a **Hyperledger Besu ethereum-etl
export**. Stress that the Besu one is the convincing row:

> "The first two are the same shape from the same provider. The third is an ethereum-etl
> export — different field names, different timestamp format. Same verdict on all 48
> transactions, and the only thing we changed was which sheet translates it."

Then the right-hand card, **LOOKUP mode**:

> "There are two ways to link a transaction to a shipment. DECODE reads the shipment key
> straight out of the ABI-encoded call data. But plenty of real contracts do not put a
> business key on-chain at all — so LOOKUP resolves the transaction hash through a mapping
> file instead. Both reach the same 48 joined rows with nothing unmatched."

**Demo option:** `tools\chain_swap_demo.ps1` proves all three feeds in one command. If time
is short, just open `FieldMapping` and `FieldMapping_EvmEtl` side by side in `Config.xlsx`.

---

## 6 · The six rules (60s) — Adithya

Walk the table briskly — one clause per rule, do not read the columns aloud.

> "Six rules. Timestamp drift, quantity against the purchase order, sender against the
> approved-partner whitelist, duplicate transaction hashes, milestone ordering, and a
> required smart-contract event per shipment."

Then the two cards:

> "They are set-based, not row-by-row. Duplicates, ordering and required events have to
> compare rows against one another, so each rule receives the whole results table — that is
> six workflow invocations per run instead of two hundred and eighty-eight."

> "And whether a rule runs at all, its severity and its thresholds are cells in a spreadsheet,
> not constants in code."

Severity is not decoration — it decides whether a finding goes to a human.

---

## 7 · Results (60s) — Adithya

> "Forty-eight transactions across twelve shipments, in twenty-one seconds. Forty pass, one
> warning, seven failures, one escalation."

Then the honest framing of the headline claim:

> "Roughly twelve of those twenty-one seconds are Excel starting up. The validation engine
> itself is about four seconds, which is around eight thousand transactions an hour on one
> robot. Against a manual baseline of two to three minutes per transaction, that supports the
> ninety-percent reduction we claimed at Review 1 — and I'll say plainly that it is an
> estimate against a stated baseline, not a timed comparison. We never actually stopwatched a
> human."

Volunteering that is better than being caught on it.

> "And it is deterministic. The same 48 transactions give the same verdict every run, which is
> what makes an acceptance suite that asserts exact counts possible at all."

---

## 8 · Findings per rule (40s) — Adithya — drop if short on time

> "Every rule earns its place — each one fires on the sample data, and none fires on a clean
> shipment. Seven of seven seeded anomalies, zero false positives."

Explain the two bars that are not 1:

> "R1 shows two because one is a warning and one is a failure — that is the three-band logic
> working. R4 shows two for a single replayed delivery, because from the ledger alone there is
> no way to tell which copy was the legitimate one, so we flag both."

The point to make here is the oracle:

> "We did not just assert our own output was right. `verify_seeded_data.py` re-derives the
> expected outcome for all 48 transactions from the raw JSON and the ERP workbook — a
> deliberately separate implementation. If the bot and the oracle disagree, one of them is
> wrong and we do not sign the run off."

---

## 9 · Novelty 1 and 3 — configuration is the product (60s + demo) — Adithya

**Live demo.** In `Config.xlsx → ValidationRules`, set **R4 `Enabled` = FALSE**, save, re-run.

> "R4 is disabled in Config.xlsx — skipped. Failures drop from seven to five. No workflow was
> opened, nothing was recompiled, nothing was republished."

**Set it back to TRUE.**

> "The reason that matters is change control. The rules, the thresholds and the partner
> whitelist are owned by logistics operations, not by developers. A spreadsheet plus a
> scheduled robot puts the change where the domain knowledge already is."

On the config panel: **seven sheets and sixty settings** — up from six and forty at Review 2.

> "The new sheet is FieldMapping_EvmEtl, and it is the one that makes a differently-shaped
> ledger export work."

---

## 10 · Novelty 2 — human-in-the-loop (60s + demo) — Nitin

**Rehearse this by hand at least once before the review.** Set `EscalationMode = PROMPT` and
`AttendedMode = True` beforehand, and **set it back to `SIMULATE` afterwards.**

Four steps: detect, classify, pause, record.

> "R3 finds a delivery signed by a wallet that is not on the approved-partner list. The rule
> is marked CRITICAL in the spreadsheet with fail action ESCALATE, so the bot stops and shows
> a reviewer the shipment, the wallet, the hash and the finding."

Choose **REJECTED**, type your name and a short note, click **Record decision and resume**.

Then the point that actually matters:

> "The decision is captured before the audit log is written. So the reviewer's judgement — the
> name, the note, the verdict — is hashed into the chain exactly like the machine's. A human
> decision becomes evidence, not a comment in a side system."

And the safety property, which examiners like:

> "An unattended robot has nobody to answer a dialog. The bot detects that and downgrades
> PROMPT to AUTO_LOG, records the anomaly as DEFERRED and carries on — so a scheduled job can
> never hang waiting on a prompt nobody will ever see."

---

## 11 · Novelty 4 — the audit trail (90s + demo) — Nitin — **the strongest moment**

Open `Data\Output\AuditLogs\ValidationAuditLog.csv` in a text editor. Show the columns:
`BotIdentity`, `TimestampUtc`, `InputDataHash`, `RecordHash`, `PrevHash`, `EntryHash`.

> "Each entry carries the hash of the entry before it. Same idea as the blockchain we are
> auditing, applied to the audit itself."

Then run it:

```
powershell -ExecutionPolicy Bypass -File tools\tamper_test.ps1
```

It verifies clean, flips one `FAIL` verdict to `PASS` — the exact edit somebody covering up an
anomaly would make — re-verifies, and reports:

```
INVALID - first problem at line 7: record contents were altered (RecordHash mismatch) - SHP-1005/Dispatched
```

then restores the file and verifies clean again.

> "One character changed, and the log names the exact row. You cannot quietly edit history
> without rewriting everything after it."

If asked **why hash a log that is already about a blockchain** — this is the question to be
ready for:

> "The transactions are on-chain. The validation is not — it happens off-chain, in the bot.
> The chained log is what makes the verification step itself evidential, rather than a report
> somebody could edit afterwards."

Add that `12_Verify_AuditChain` is its own entry point, so an auditor can re-verify the log
without running the bot or trusting it.

---

## 12 · The evidence pack (60s) — Nitin

Open the newest `ValidationReport_*.xlsx` and the newest `AuditReport_*.pdf` from the run you
just did.

> "Every run leaves the same evidence pack behind."

- **Excel** — "Summary, per-rule breakdown, every result colour-coded, and an exceptions
  sheet." Scroll to a red row and read the `FailureReasons` cell aloud. It is plain English on
  purpose, because the person reading it is a logistics manager.
- **PDF audit report** — **new** — "This is the document you hand an auditor: the bot's
  identity, the configuration it ran with, every finding with the reviewer's decision, every
  transaction, the audit-chain head hash with its verification result, and a signature block."
- **Dashboard** — "One self-contained HTML file. No server, no CDN, no internet. Processing
  time per phase and trends across runs." Then the point worth making: "the run history at the
  bottom is read back out of the audit log itself, so it cannot disagree with the evidence."
- **Teams and Slack payloads** — **new** — open `Alert_*_teams.json` if alerts come up. "An
  Adaptive Card and a Block Kit message, written to disk so delivery is testable without a
  tenant."

---

## 13 · Queue mode and the Performer (45s) — Adithya — drop if short on time

> "For enterprise volume the same run splits into one Orchestrator queue item per shipment —
> forty-eight transactions become twelve work items. `RunMode` is one cell in Config.xlsx."

Why a shipment and not a transaction, which is the obvious question:

> "Three of the six rules compare events against each other. A worker handed a single
> transaction could not evaluate duplicates, ordering or required events at all. A shipment is
> the smallest unit that keeps every rule meaningful — it is a consequence of the rules, not
> an arbitrary choice."

**New since Review 2 — the Performer:**

> "The Performer is now its own entry point. It drains the queue in one job and then calls the
> same finalise pipeline a batch run uses, so queue mode produces the same audit trail, the
> same Excel and PDF reports and the same dashboard."

Then be straight, before you are asked:

> "Dispatch and perform are tested end to end in dry run — forty-eight transactions become
> twelve payloads and the Performer processes them into a full evidence pack. The three
> Orchestrator queue activities themselves have never run against a live queue. That is the
> one untested hop in the project."

---

## 14 · Verification (60s) — Nitin

> "Every claim in this deck is reproducible with one command."

Give the three numbers off the slide, then:

> "Eleven UiPath test cases in Studio's Test Explorer. Six drive one validation rule each
> against fixtures built inside the test, so a failing test points at exactly one rule. The
> other five cover audit-chain tampering, field-mapping equivalence across two schemas, the
> PDF's structure, the review form's timeout, and the Teams and Slack payloads."

**Volunteer the bug story even if nobody asks.** It is the most credible sentence available:

> "One of those tests found a real bug while I was writing it. R4 assumed a transaction hash
> was at least twelve characters long, and it crashed on a malformed one. Found, fixed, and
> now covered — which is what tests are for."

---

## 15 · What works and what does not (45s) — Nitin — **do not skip**

> "Stated plainly so nothing is overclaimed."

Left column briskly — do not read all twelve lines, say "everything in the pipeline, all six
rules, three feeds, two schemas, both mapping modes, three ERP sources, the full evidence
pack, and dispatch and perform in dry run."

Then the right column slowly, because this is the slide that earns trust:

> "Four things have not been executed. A live Orchestrator queue. A live Etherscan API key.
> Real browser automation — WEB mode parses the explorer's HTML, because UIAutomation 25.10
> removed the classic scraping activity and its replacement needs a recorded descriptor. And
> real mail and chat endpoints; delivery is tested against a local capture server."

Close the slide on the distinction that matters:

> "Nothing is stubbed. Every one of those is a hop into a system we hold no credentials for,
> and each path is exercised locally right up to that hop. And this list is shorter than it
> was at Review 2 — the PDF, the webhooks, the ledger export and the Performer all moved from
> this column to the other one."

*Knowing precisely what is unfinished reads as competence. Waffling reads as not knowing.*

---

## 16 · Close (30s)

> "Automated reconciliation of blockchain records against enterprise systems. Six rules a
> business team can configure. Escalation where human judgement belongs. And an audit trail
> that proves it was not altered afterwards."

Then return to the opening line and stop:

> "Immutable is not the same as verified. This is the bot that does the verifying. Thank you —
> happy to take questions."

---

## Questions you should expect

**"Isn't blockchain data already trustworthy?"**
Immutable, not correct. The chain guarantees the record has not changed since it was written —
not that it was true when written, and not that it agrees with the ERP. Every seeded anomaly is
a perfectly valid on-chain transaction.

**"Why RPA instead of a Python script?"**
The rules, thresholds and whitelist are owned by logistics operations, not developers. A
spreadsheet plus a scheduled robot puts change control where the domain knowledge is. It also
inherits Orchestrator's scheduling, queueing, retry and audit logging for free.

**"Why hash the audit log when it's already on a blockchain?"**
The transactions are on-chain. The validation isn't — it happens off-chain, in the bot. The
chained log is what makes the verification step itself evidential rather than a report somebody
could edit.

**"What does 'blockchain agnostic' actually mean here?"** *(expect this after slide 5)*
The readers return records untouched; one workflow builds the canonical table from a spreadsheet
sheet that declares field names, ABI argument positions and timestamp formats. A new schema is a
new sheet. Proven on three feeds and two schemas by `chain_swap_demo.ps1`.

**"Is the data real?"**
The transactions are synthetic, but not fake in shape: genuine EIP-55 checksummed addresses, real
keccak-256 function selectors, and correctly ABI-encoded call data in an authentic Etherscan V2
response envelope. API mode against the live chain uses the same parser.

**"Where are the tests?"**
Eleven UiPath test cases in the Test tab. Six drive one rule each against hand-built fixtures;
the others cover audit-chain tampering, field-mapping equivalence, PDF structure, the review
form's timeout, and the Teams and Slack payloads. They cover edge cases the sample data does not:
drift exactly at the tolerance boundary, a checksummed address against a lower-case whitelist, the
same hash on two shipments.

**"What doesn't work?"**
Nothing is stubbed, but four final hops need accounts we don't hold: a live Orchestrator queue
(dispatch and perform are tested in dry run), a live Etherscan key, real mail/Teams/Slack
endpoints (delivery is tested against a local capture server), and real browser automation. Also
the review form is built in the workflow rather than designed as a Studio `.uiform`, because those
can only be drawn in Studio's designer. All listed under "Known gaps" in the README.

---

## If something breaks live

| Symptom | Do this |
|---|---|
| Studio run hangs with no output | Excel COM is stuck from an earlier kill. Close Excel; `build.ps1` clears stray executors. |
| A script errors | Fall back to `docs/sample-output/dashboard.html` — committed output from a real run. "Here's the output from the run I did this morning," and carry on. |
| `[VERIFY] INVALID` unexpectedly | You are mid-tamper-test. Re-run `tamper_test.ps1`; it restores the file and ends VALID. |
| Reviewer dialog does not appear | `AttendedMode` is `False` in `Config.xlsx`. Set it `True`, or just describe it. |
| A script won't start at all | Running scripts is disabled on this machine. Every command needs `-ExecutionPolicy Bypass`. |
| Projector cuts the right of the screen | The dashboard reflows — narrow the browser window rather than scrolling sideways. |

**The universal recovery:** everything you are demonstrating is already captured in
[RESULTS.md](RESULTS.md) with real numbers. If the machine misbehaves, switch to that document
and keep talking. Never debug live in front of an examiner.

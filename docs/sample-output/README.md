# Sample output

Artefacts from one real run, committed so the project's output can be seen without
running it. Everything here is regenerated on every execution into
`BlockchainLogisticsValidator/Data/Output/` (which is git-ignored).

| File | What it is |
|---|---|
| `dashboard.html` | The dashboard the bot writes each run. Self-contained — download and open it, no server needed. GitHub shows HTML as source, so view it locally. |
| `ValidationReport_sample.xlsx` | The styled Excel report: Summary, per-rule breakdown, colour-coded results, exceptions. |

The run these came from: **48 transactions, 12 shipments — 40 passed, 1 warning, 7 failed,
1 escalated to a human.** Those seven failures are the seeded anomalies documented in
[../SEEDED_ANOMALIES.md](../SEEDED_ANOMALIES.md), one per validation rule.

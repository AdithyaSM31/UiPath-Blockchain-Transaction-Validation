# Review day checklist

Everything to do before and during the final review, in order. Tick straight down.

The talking points are in [FINAL_REVIEW_SPEAKING_NOTES.md](FINAL_REVIEW_SPEAKING_NOTES.md);
the demo walkthrough is in [DEMO_SCRIPT.md](DEMO_SCRIPT.md). This file is just the mechanics.

---

## The night before

- [ ] **Check UiPath Studio is licensed** on the demo laptop: open Studio, signed in, and
      no "unlicensed" banner. Without a licence the robot will not run anything.

- [ ] **Run the acceptance suite.** About fifteen minutes.

      powershell -ExecutionPolicy Bypass -File "tools\run_all_checks.ps1"

  Expect `21/21 checks passed`, including the twelve UiPath test cases. If anything is red,
  fix it now — not in the room.

  > Near the end the suite answers on-screen windows itself (check 17, failure handling, and
  > check 20, the review window).
  > **Do not touch the mouse or keyboard while it finishes**, or your click lands in the
  > dialog instead of its own.

- [ ] **Rehearse the reviewer dialog once, by hand.** In `Data\Config.xlsx → Settings`:
      set `EscalationMode = PROMPT` and `AttendedMode = True`, run `Main.xaml`, and click
      through the review form for `SHP-1007`: choose **REJECTED**, type your name, add a note,
      and press **Record decision and resume**.

- [ ] **Set it back.** `EscalationMode = SIMULATE`. Leave `AttendedMode` as you found it.
      Forgetting this is the single most likely way to stall the live run.

- [ ] **Reset the audit log** so the line counts are easy to narrate:

      del "BlockchainLogisticsValidator\Data\Output\AuditLogs\ValidationAuditLog.csv"

- [ ] **Do one clean run** after the reset, so there is a fresh dashboard, Excel report and
      PDF to open if a live run misbehaves.

- [ ] **Check the deck opens** and that Presenter View shows the speaker notes:
      `RPA Project Final Review.pptx`.

---

## Opening the project in Studio

- [ ] Open **`BlockchainLogisticsValidator\project.json`** — the `project.json`, not a loose
      `.xaml`. Opening a workflow on its own gives you a file with no dependencies and red
      error icons everywhere.

- [ ] **If Studio offers to upgrade the project**, that is expected on Studio 26: the project
      records `studioVersion` 25.10.15 and the dependencies are pinned. Accepting the upgrade
      is fine and the project still runs, but **decline it before a review** — there is no
      upside on the day and an upgrade rewrites `project.json` and every dependency pin.

- [ ] **Let the dependencies restore** before clicking anything. Watch the status bar. Every
      pinned package is already in the local NuGet cache, so this works offline, but it takes
      a few seconds.

- [ ] **Confirm the Project panel is clean** — no yellow warning triangles on
      `Main.xaml`, `Workflows\`, or `Tests\`.

- [ ] **Open the Test tab** once so Test Explorer populates all twelve test cases. It is slow
      the first time and instant afterwards, which matters if you demo them.

- [ ] Leave **`Main.xaml`** open on screen. That is the panel to present from: the phase
      sequence reads top to bottom and each box is named for what it does.

---

## Have open and ready

- [ ] UiPath Studio, `Main.xaml` on screen, Output panel visible and cleared
- [ ] `Data\Config.xlsx` — on the `Settings` sheet
- [ ] A terminal at the project root
- [ ] File Explorer at `Data\Output\Reports`
- [ ] `RPA Project Final Review.pptx` in Presenter View
- [ ] A browser with `Data\Output\Reports\dashboard.html` already loaded as a fallback

**Close everything else.** Notifications, chat windows and a second Excel instance are the
usual sources of an embarrassing moment.

---

## Projector

- [ ] **Rehearse once with the projector connected.** The one genuine risk in this demo is a
      dialog or a spreadsheet opening on the wrong screen.
- [ ] Check the right-hand edge is not cut off. If it is, the dashboard reflows — narrow the
      browser window rather than scrolling sideways.
- [ ] Set display scaling so the Output panel is readable from the back of the room. 125% or
      150% is usually right; the Output text is small at 100%.

---

## During the demo

Four live moments, in the order the deck reaches them:

| Deck slide | Live moment | Reset afterwards |
|---|---|---|
| 5 Mapping engine | `tools\chain_swap_demo.ps1`, or open `FieldMapping` beside `FieldMapping_EvmEtl` | nothing |
| 9 Novelty 1 | `ValidationRules → R4 Enabled = FALSE`, save, re-run | **set R4 back to TRUE** |
| 10 Novelty 2 | `EscalationMode = PROMPT`, `AttendedMode = True`, re-run, click through | **back to SIMULATE** |
| 11 Novelty 4 | `tools\tamper_test.ps1` | restores the log itself |

Every command needs the policy flag:

```bash
powershell -ExecutionPolicy Bypass -File "tools\tamper_test.ps1"
```

**If anything breaks:** switch to [RESULTS.md](RESULTS.md) or the committed
`docs/sample-output/dashboard.html` and keep talking. "Here's the output from the run I did
this morning" costs you nothing. Never debug live in front of an examiner.

---

## Afterwards

- [ ] Confirm `Config.xlsx` is back to `EscalationMode = SIMULATE` and all six rules enabled,
      so the next person to clone the repository gets the documented default run.
- [ ] Confirm `RunMode` is back to `BATCH` if you showed the queue split.

"""
build_final_deck.py - generate "RPA Project Final Review.pptx".

The Review 1 and Review 2 decks are deliverables in their own right and are never
touched by this script. This builds a third, standalone deck for the final review,
reusing the Review 2 design system (palette, type scale, slide chrome) so the three
read as one project.

    python tools\\build_final_deck.py
    python tools\\build_final_deck.py --date "21 October 2026"

Every figure in the deck comes from docs/RESULTS.md or from a counted artefact in the
repository - nothing here is an estimate unless the slide says so.
"""

import argparse
import os
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

# --------------------------------------------------------------------------- palette
INK        = RGBColor(0x24, 0x1C, 0x3A)   # body text on light
PURPLE     = RGBColor(0x2F, 0x1D, 0x63)   # headings, primary fill
PURPLE_MID = RGBColor(0x48, 0x36, 0x7C)
NIGHT      = RGBColor(0x0C, 0x06, 0x25)   # dark slide background
NIGHT_SOFT = RGBColor(0x1A, 0x10, 0x38)
PINK       = RGBColor(0xFF, 0x8A, 0xAF)   # accent
LILAC      = RGBColor(0x9A, 0x81, 0xDF)   # eyebrow
LILAC_PALE = RGBColor(0xCF, 0xC7, 0xE4)
HAIRLINE   = RGBColor(0xE2, 0xDC, 0xEC)   # card border
FOOTER     = RGBColor(0xA7, 0x9F, 0xC0)
MUTED      = RGBColor(0x6B, 0x64, 0x80)
PAPER      = RGBColor(0xF6, 0xF4, 0xFA)   # light slide background
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
GREEN      = RGBColor(0x1A, 0x7F, 0x4B)
GREEN_DEEP = RGBColor(0x15, 0x66, 0x3C)
GREEN_PALE = RGBColor(0xE4, 0xF5, 0xEC)
AMBER      = RGBColor(0x8A, 0x61, 0x00)
AMBER_DEEP = RGBColor(0x6B, 0x4B, 0x00)
AMBER_PALE = RGBColor(0xFD, 0xF2, 0xD8)
RED        = RGBColor(0xB3, 0x26, 0x1E)

BODY_FONT = "Calibri"
MONO_FONT = "Courier New"

L = 0.62          # left margin
W = 12.09         # content width
DECK_LABEL = "Final Review"


# --------------------------------------------------------------------------- helpers
def new_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.3333333)
    prs.slide_height = Inches(7.5)
    return prs


def add_slide(prs, dark=False):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = NIGHT if dark else PAPER
    return slide


def text(slide, x, y, w, h, body, size=11.5, bold=False, color=INK,
         align=PP_ALIGN.LEFT, font=BODY_FONT, spacing=None, anchor=MSO_ANCHOR.TOP):
    """A plain textbox. `body` is a string or a list of (string, bold, color, size)."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor

    lines = body if isinstance(body, list) else [body]
    for i, line in enumerate(lines):
        if isinstance(line, tuple):
            t, b, c, s = (line + (None,) * 4)[:4]
            b = bold if b is None else b
            c = color if c is None else c
            s = size if s is None else s
        else:
            t, b, c, s = line, bold, color, size
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        if spacing:
            para.space_after = Pt(spacing)
        run = para.add_run()
        run.text = t
        run.font.name = font
        run.font.size = Pt(s)
        run.font.bold = b
        run.font.color.rgb = c
    return box


def card(slide, x, y, w, h, fill=WHITE, border=HAIRLINE, radius=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    # python-pptx expresses the corner adjustment as a fraction of the shorter side;
    # 0.10in of radius at every size keeps the family of cards visually consistent.
    shp.adjustments[0] = radius if radius is not None else min(0.5, 0.10 / min(w, h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if border is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = border
        shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    return shp


def pill(slide, x, y, w, h, label, fill=PINK, color=NIGHT, size=9.5):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    shp.adjustments[0] = 0.5
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    text(slide, x, y + (h - 0.26) / 2, w, 0.26, label, size=size, bold=True,
         color=color, align=PP_ALIGN.CENTER)
    return shp


def circle(slide, x, y, d, label, fill=PURPLE, color=WHITE, size=12.5):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    text(slide, x, y + (d - 0.26) / 2, d, 0.26, label, size=size, bold=True,
         color=color, align=PP_ALIGN.CENTER)
    return shp


def header(slide, eyebrow, title, dark=False):
    text(slide, L, 0.36, W, 0.26, eyebrow, size=11.5, bold=True,
         color=LILAC if not dark else PINK)
    text(slide, L, 0.62, W, 0.62, title, size=30, bold=True,
         color=PURPLE if not dark else WHITE)


def footer(slide, n, dark=False):
    text(slide, L, 7.06, W, 0.26, f"{DECK_LABEL}  ·  {n}", size=9.5,
         color=FOOTER if not dark else RGBColor(0x6B, 0x5F, 0x8F))


def notes(slide, body):
    slide.notes_slide.notes_text_frame.text = body


def code_block(slide, x, y, w, lines, size=11.5, fill=NIGHT_SOFT, color=RGBColor(0x9F, 0xE3, 0xBC)):
    h = 0.22 + 0.24 * len(lines)
    card(slide, x, y, w, h, fill=fill, border=None, radius=0.04)
    for i, ln in enumerate(lines):
        text(slide, x + 0.2, y + 0.14 + 0.24 * i, w - 0.4, 0.24, ln,
             size=size, font=MONO_FONT, color=color)
    return h


# --------------------------------------------------------------------------- slides
def slide_cover(prs, date_label):
    s = add_slide(prs, dark=True)
    # Two offset discs, echoing the Review 1 and Review 2 covers.
    for x, y, d, c in ((9.9, -1.5, 5.6, PURPLE), (11.6, 3.7, 3.2, PURPLE_MID)):
        shp = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
        shp.fill.solid()
        shp.fill.fore_color.rgb = c
        shp.line.fill.background()
        shp.shadow.inherit = False

    text(s, L, 1.42, 8.4, 0.3, "FINAL REVIEW", size=12.5, bold=True, color=PINK)
    text(s, L, 1.86, 8.7, 1.9,
         "UiPath-Based Validation of Blockchain Transactions for Transparent Logistics",
         size=37, bold=True, color=WHITE)
    text(s, L, 3.96, 8.3, 0.9,
         "An RPA bot that reconciles on-chain logistics events against enterprise ERP "
         "records, flags anomalies against six configurable rules, escalates to a human, "
         "and proves its own audit trail was never altered.",
         size=14, color=LILAC_PALE)
    text(s, L, 5.12, 4.2, 0.3, "Adithya Sankar Menon  23BRS1079", size=13, bold=True, color=WHITE)
    text(s, L, 5.46, 4.2, 0.3, "Nitin Nandakumar  23BRS1378", size=13, bold=True, color=WHITE)
    text(s, L, 6.28, 8.4, 0.3, f"Robotic Process Automation  ·  {date_label}",
         size=11.5, color=FOOTER)
    notes(s, "Nitin opens. One line: blockchain makes a record immutable, it does not make "
             "it correct - that is the gap this project fills. Say that this is the final "
             "review of a finished system, and that everything in the deck is reproducible "
             "with one command.")


def slide_whats_new(prs, n):
    s = add_slide(prs)
    header(s, "SINCE REVIEW 2", "Six capabilities added, and the gap list shortened")
    rows = [
        ("Field mapping engine", "01d_Apply_FieldMapping",
         "Readers fetch; the mapping sheet translates. A new feed schema needs a sheet, not a workflow."),
        ("Ledger export support", "FieldMapping_EvmEtl",
         "A Hyperledger Besu ethereum-etl export - different field names, different timestamp format."),
        ("LOOKUP shipment linking", "MappingMode = LOOKUP",
         "Shipment keys from a mapping file when the call data does not carry one."),
        ("Formal PDF audit report", "08_Report_PDF",
         "The document handed to an auditor: config, findings, decisions, chain head, signature block."),
        ("Teams and Slack alerts", "09b_Alert_Webhooks",
         "An Adaptive Card and a Block Kit message alongside the e-mail."),
        ("Standalone Performer", "Queue\\Performer.xaml",
         "Its own entry point: works through the work queue, then runs the same finalise pipeline as a batch run."),
    ]
    y = 1.56
    for i, (title, where, why) in enumerate(rows, 1):
        card(s, L, y, W, 0.84)
        circle(s, L + 0.26, y + 0.25, 0.34, str(i), fill=PURPLE if i % 2 else PURPLE_MID)
        text(s, 1.4, y + 0.14, 3.5, 0.3, title, size=14, bold=True, color=PURPLE)
        text(s, 1.4, y + 0.46, 3.5, 0.26, where, size=10, font=MONO_FONT, color=MUTED)
        text(s, 5.1, y + 0.26, 7.3, 0.6, why, size=11.5, color=INK)
        y += 0.92
    notes(s, "This is the 'what changed' slide - keep it to about a minute. The one to dwell "
             "on is the mapping engine, because it is what turns the chain-agnostic claim from "
             "an assertion into a property of the design. The others are each a Review 1 "
             "commitment that was still open at Review 2.")
    footer(s, n)


def slide_problem(prs, n):
    s = add_slide(prs, dark=True)
    shp = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-1.8), Inches(4.4), Inches(4.6), Inches(4.6))
    shp.fill.solid()
    shp.fill.fore_color.rgb = PURPLE
    shp.line.fill.background()
    shp.shadow.inherit = False

    text(s, L, 1.22, W, 0.3, "THE PROBLEM, IN ONE LINE", size=12, bold=True, color=PINK)
    text(s, L, 1.74, 11.2, 1.9,
         [("Blockchain makes a record immutable.", True, WHITE, 34),
          ("It does not make it correct.", True, PINK, 34)], size=34)
    text(s, L, 3.92, 9.6, 1.0,
         "A tampered quantity, a forged timestamp or an unauthorised wallet is recorded just "
         "as permanently as a legitimate one. Every anomaly this bot catches is a perfectly "
         "valid on-chain transaction.",
         size=14.5, color=LILAC_PALE)

    for i, (t, d) in enumerate([
        ("Manual", "Operators cross-check hashes against ERP records by hand"),
        ("Siloed", "Explorers emit JSON and hex; ERP systems speak Excel and SQL"),
        ("Unprovable", "A manual verification log is not itself evidence"),
    ]):
        x = L + i * 4.03
        card(s, x, 5.28, 3.78, 1.22, fill=NIGHT_SOFT, border=PURPLE_MID)
        text(s, x + 0.24, 5.44, 3.3, 0.3, t, size=14, bold=True, color=PINK)
        text(s, x + 0.24, 5.76, 3.34, 0.62, d, size=11, color=LILAC_PALE)
    notes(s, "The intellectual core, and the question most likely to be probed. Land the "
             "distinction between immutable and verified before showing any code. If asked "
             "'isn't blockchain data already trustworthy' - immutable, not correct: the chain "
             "guarantees the record has not changed since it was written, not that it was true "
             "when written, and not that it agrees with the ERP.")
    footer(s, n, dark=True)


def slide_architecture(prs, n):
    s = add_slide(prs)
    header(s, "28 WORKFLOWS, ORCHESTRATED BY MAIN.XAML", "System architecture")
    bands = [
        ("INGEST", PURPLE_MID, [("00", "Read config"), ("01", "Extract chain"),
                                ("01d", "Field mapping"), ("02", "Extract ERP"),
                                ("03", "Map and join")]),
        ("VALIDATE", PURPLE, [("05", "Rules R1-R6"), ("06", "Escalate to human")]),
        ("EVIDENCE", PURPLE_MID, [("10", "Hash-chained log"), ("12", "Verify chain")]),
        ("REPORT", PURPLE_MID, [("07", "Excel report"), ("08", "PDF audit report"),
                                ("13", "HTML dashboard"), ("09", "Alert e-mail"),
                                ("09b", "Teams / Slack")]),
    ]
    for i, (label, fill, items) in enumerate(bands):
        x = L + i * 3.08
        card(s, x, 1.66, 2.92, 2.92)
        pill(s, x + 0.2, 1.86, 1.34, 0.34, label, fill=fill, color=WHITE)
        for j, (num, what) in enumerate(items):
            y = 2.42 + j * 0.42
            text(s, x + 0.2, y, 0.52, 0.28, num, size=11.5, bold=True, color=PINK)
            text(s, x + 0.76, y, 2.0, 0.28, what, size=11.5, color=INK)
        if i < 3:
            text(s, x + 2.9, 2.95, 0.24, 0.4, "›", size=26, bold=True, color=LILAC)

    card(s, L, 4.78, W, 0.88, fill=WHITE)
    text(s, L + 0.24, 4.94, 11.6, 0.6,
         [("Steps 06-13 are 14_Finalise_Run.", True, PURPLE, 12.0),
          ("A batch run and the queue Performer both call it, so queue mode produces the "
           "same audit trail, the same reports and the same dashboard as a single-job run.",
           False, INK, 11.5)], size=12.0)

    for i, (t, d) in enumerate([
        ("Structure in XAML", "Sequences, Excel, database, mail and HTTP activities stay visual."),
        ("Algorithms in Invoke Code", "ABI decoding, hashing, PDF layout and the sequence state machine."),
        ("Reports need no Office", "ClosedXML and raw PDF 1.4, so an unattended robot writes identical files."),
    ]):
        x = L + i * 4.03
        card(s, x, 5.78, 3.78, 1.06)
        text(s, x + 0.24, 5.96, 3.4, 0.3, t, size=12.5, bold=True, color=PURPLE)
        text(s, x + 0.24, 6.3, 3.4, 0.5, d, size=10.5, color=MUTED)
    notes(s, "Adithya leads. The point to emphasise is 14_Finalise_Run: it is why queue mode "
             "is not a second, divergent code path. If asked why Invoke Code rather than all "
             "activities - three of the six rules compare rows against one another, and a "
             "state machine drawn as activities would be unreadable.")
    footer(s, n)


def slide_mapping(prs, n):
    s = add_slide(prs)
    header(s, "THE CHAIN-AGNOSTIC CLAIM, AS A PROPERTY OF THE DESIGN",
           "Readers fetch. The mapping sheet translates.")

    stages = [("MOCK / API / WEB", "Three readers. Each returns the feed's records untouched."),
              ("01d FIELD MAPPING", "The selected FieldMapping sheet builds the canonical table."),
              ("R1-R6 VALIDATION", "Rules cannot tell which reader produced their DataTable.")]
    for i, (t, d) in enumerate(stages):
        x = L + i * 4.18
        card(s, x, 1.62, 3.9, 1.1)
        pill(s, x + 0.22, 1.8, 2.3, 0.32, t, fill=PINK, color=NIGHT)
        text(s, x + 0.22, 2.2, 3.46, 0.44, d, size=10.5, color=INK)
        if i < 2:
            text(s, x + 3.86, 1.97, 0.3, 0.4, "›", size=26, bold=True, color=LILAC)

    card(s, L, 2.92, 6.0, 2.36)
    text(s, L + 0.24, 3.1, 5.5, 0.3, "Three feeds, two schemas, one result", size=15,
         bold=True, color=PURPLE)
    rows = [("Ethereum", "chainid 1", "Etherscan", "48 / 40 / 1 / 7"),
            ("Polygon", "chainid 137", "Etherscan", "48 / 40 / 1 / 7"),
            ("Besu export", "ethereum-etl", "EvmEtl", "48 / 40 / 1 / 7")]
    text(s, L + 0.24, 3.52, 1.7, 0.24, "FEED", size=9, bold=True, color=MUTED)
    text(s, L + 2.0, 3.52, 1.4, 0.24, "SOURCE", size=9, bold=True, color=MUTED)
    text(s, L + 3.5, 3.52, 1.1, 0.24, "SHEET", size=9, bold=True, color=MUTED)
    text(s, L + 4.6, 3.52, 1.4, 0.24, "RESULT", size=9, bold=True, color=MUTED)
    for j, (a, b, c, d) in enumerate(rows):
        y = 3.82 + j * 0.34
        text(s, L + 0.24, y, 1.8, 0.26, a, size=11, bold=True, color=INK)
        text(s, L + 2.0, y, 1.5, 0.26, b, size=10.5, font=MONO_FONT, color=MUTED)
        text(s, L + 3.5, y, 1.1, 0.26, c, size=10.5, font=MONO_FONT, color=MUTED)
        text(s, L + 4.6, y, 1.5, 0.26, d, size=11, bold=True, color=GREEN)
    text(s, L + 0.24, 4.92, 5.5, 0.3,
         "Only config cells changed between these runs.", size=10.5, color=MUTED)

    card(s, 6.77, 2.92, 5.98, 2.36)
    text(s, 7.01, 3.1, 5.5, 0.3, "Linking a transaction to a shipment", size=15,
         bold=True, color=PURPLE)
    # Keep each description to a single line at this width, or it runs into the
    # summary line below.
    for j, (mode, d) in enumerate([
        ("DECODE", "The shipment key comes straight from the ABI call data."),
        ("LOOKUP", "No business key on-chain: a mapping file resolves it."),
    ]):
        y = 3.60 + j * 0.58
        pill(s, 7.01, y, 1.1, 0.3, mode, fill=PURPLE if j == 0 else PURPLE_MID, color=WHITE)
        text(s, 8.25, y + 0.03, 4.3, 0.3, d, size=11, color=INK)
    text(s, 7.01, 4.86, 5.5, 0.3,
         "Both modes reach the same 48 joined rows, 0 unmatched.", size=10.5,
         bold=True, color=GREEN)

    card(s, L, 5.42, W, 1.42, fill=GREEN_PALE, border=None, radius=0.03)
    text(s, L + 0.26, 5.6, 11.5, 0.3,
         "Why this is the strongest version of the claim", size=13, bold=True, color=GREEN)
    text(s, L + 0.26, 5.96, 11.5, 0.72,
         "A competitor's \"blockchain agnostic\" usually means a second reader workflow per "
         "chain. Here the readers are dumb and interchangeable, and every schema difference - "
         "field names, ABI argument positions, timestamp formats - is declared in a spreadsheet "
         "sheet. tools\\chain_swap_demo.ps1 proves it on three feeds and two schemas in one command.",
         size=11.5, color=GREEN_DEEP)
    notes(s, "Adithya leads. This slide did not exist at Review 2 - the FieldMapping sheet was "
             "read but never applied. Run chain_swap_demo.ps1 live if there is time, or open "
             "FieldMapping next to FieldMapping_EvmEtl in Config.xlsx to show the two translations "
             "side by side. The Besu export is the convincing one because its field names and "
             "timestamp format are genuinely different, not a renamed copy.")
    footer(s, n)


def slide_rules(prs, n):
    s = add_slide(prs)
    header(s, "SIX RULES, EVERY ONE DRIVEN BY CONFIG.XLSX", "The validation engine")
    cols = ["RULE", "WHAT IT CHECKS", "SEEDED FAILURE", "SEVERITY"]
    # Severity is read straight off the ValidationRules sheet in Config.xlsx - it is a
    # config value, not a property of the rule, so keep these two in step.
    rows = [
        ("R1  Timestamp", "On-chain time against the ERP's expected milestone time",
         "SHP-1002 warns, SHP-1003 fails", "HIGH"),
        ("R2  Quantity", "Decoded on-chain quantity against the purchase order",
         "SHP-1005 short by 20 units", "HIGH"),
        ("R3  Whitelist", "Sender is an approved partner wallet",
         "SHP-1007 signed by an unknown wallet", "CRITICAL"),
        ("R4  Duplicate", "No transaction hash recorded twice per shipment",
         "SHP-1009 delivery replayed", "MEDIUM"),
        ("R5  Sequence", "Milestones occur in the configured order",
         "SHP-1011 delivered before in-transit", "HIGH"),
        ("R6  Contract event", "Every shipment emits the required event",
         "SHP-1012 never confirms receipt", "HIGH"),
    ]
    widths = [2.4, 4.5, 3.3, 1.89]
    y = 1.62
    xs = [L]
    for w_ in widths[:-1]:
        xs.append(xs[-1] + w_)
    for i, (c, w_) in enumerate(zip(cols, widths)):
        text(s, xs[i] + 0.16, y, w_ - 0.2, 0.24, c, size=9, bold=True, color=MUTED)
    y += 0.32
    for r_i, row in enumerate(rows):
        card(s, L, y, W, 0.52, fill=WHITE if r_i % 2 == 0 else RGBColor(0xFC, 0xFB, 0xFE))
        for i, (val, w_) in enumerate(zip(row, widths)):
            bold = i == 0
            col = PURPLE if i == 0 else INK
            if i == 3:
                col = {"CRITICAL": RED, "HIGH": AMBER, "MEDIUM": MUTED}.get(val, MUTED)
                bold = True
            text(s, xs[i] + 0.16, y + 0.14, w_ - 0.26, 0.3, val,
                 size=11 if i != 3 else 9.5, bold=bold, color=col)
        y += 0.58

    for i, (t, d) in enumerate([
        ("Set-based, not row-by-row",
         "R4, R5 and R6 compare rows against one another, so each rule receives the whole "
         "table - 6 workflow invocations per run, not 288."),
        ("Rules are configuration, not code",
         "Whether a rule runs, its severity, its thresholds and its fail action all live in a "
         "spreadsheet a logistics manager can edit."),
    ]):
        x = L + i * 6.15
        card(s, x, 5.46, 5.94, 1.3)
        text(s, x + 0.24, 5.64, 5.5, 0.3, t, size=13.5, bold=True, color=PURPLE)
        text(s, x + 0.24, 5.98, 5.5, 0.66, d, size=11.5, color=INK)
    notes(s, "Each seeded failure is a real, valid on-chain transaction - that is the whole "
             "point. The sample data contains exactly one anomaly per rule, so a single run "
             "exercises all six. Severity is not decoration: it decides whether a finding is "
             "escalated to a human, and it is a cell in Config.xlsx.")
    footer(s, n)


def slide_results(prs, n):
    s = add_slide(prs)
    header(s, "48 TRANSACTIONS · 12 SHIPMENTS · 21 SECONDS", "Results from a real run")
    tiles = [("48", "Transactions validated", "288 rule evaluations"),
             ("40", "Passed all six rules", "83.3% pass rate"),
             ("8", "Flagged", "7 failed · 1 warning"),
             ("1", "Escalated to a human", "Unapproved wallet")]
    for i, (big, lab, sub) in enumerate(tiles):
        x = L + i * 3.0
        card(s, x, 1.66, 2.86, 1.62)
        text(s, x + 0.18, 1.79, 2.5, 0.62, big, size=34, bold=True, color=PURPLE)
        text(s, x + 0.18, 2.42, 2.5, 0.3, lab, size=11.5, color=INK)
        text(s, x + 0.18, 2.72, 2.5, 0.44, sub, size=10, color=MUTED)

    cd = CategoryChartData()
    cd.categories = ["Pass", "Warning", "Fail"]
    cd.add_series("Transactions", (40, 1, 7))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(L), Inches(3.52),
                            Inches(5.3), Inches(3.2), cd)
    ch = gf.chart
    ch.has_legend = False
    ch.has_title = False
    plot = ch.plots[0]
    plot.gap_width = 80
    plot.vary_by_categories = True
    for pt, c in zip(plot.series[0].points, (GREEN, AMBER, RED)):
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = c
    ch.value_axis.has_major_gridlines = True
    ch.font.size = Pt(11)
    ch.font.name = BODY_FONT

    card(s, 6.24, 3.52, 6.51, 3.2)
    text(s, 6.5, 3.72, 6.0, 0.32, "What the numbers mean", size=14.5, bold=True, color=PURPLE)
    text(s, 6.5, 4.12, 6.0, 2.4,
         [("7 of 7 seeded anomalies detected, with zero false positives.", True, INK, 11.5),
          ("The five clean shipments passed all six rules on every event.", False, INK, 11.5),
          ("R4 reports two rows for one anomaly by design - from the ledger alone there is no "
           "way to tell which copy of a duplicated transaction was legitimate, so both are flagged.",
           False, INK, 11.5),
          ("Deterministic: the same 48 transactions give the same verdict on every run, which "
           "is what makes the acceptance suite possible.", False, INK, 11.5),
          ("Roughly 12 of the 21 seconds are Excel COM starting up; the validation engine "
           "itself is about 4 seconds, or ~8,000 transactions/hour on one robot.", False, MUTED, 11.0)],
         spacing=7)
    notes(s, "21 seconds of workflow execution. Against a manual baseline of two to three "
             "minutes per transaction this supports the 90% reduction claim from Review 1 - "
             "present it as an estimate against a stated baseline, not as a measured "
             "comparison, because no manual baseline was actually timed. The determinism point "
             "is worth making: it is what lets a suite assert exact counts.")
    footer(s, n)


def slide_findings(prs, n):
    s = add_slide(prs)
    header(s, "FINDINGS PER RULE, FROM ONE RUN", "Every rule earns its place")
    cd = CategoryChartData()
    cd.categories = ["R1", "R2", "R3", "R4", "R5", "R6"]
    cd.add_series("Findings", (2, 1, 1, 2, 1, 1))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(L), Inches(1.72),
                            Inches(7.1), Inches(4.0), cd)
    ch = gf.chart
    ch.has_legend = False
    ch.has_title = False
    ch.plots[0].gap_width = 70
    ser = ch.plots[0].series[0]
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = PURPLE
    ch.font.size = Pt(11)
    ch.font.name = BODY_FONT

    card(s, 8.04, 1.72, 4.71, 1.86, fill=GREEN_PALE, border=None, radius=0.03)
    text(s, 8.28, 1.9, 4.2, 0.62, "7/7", size=34, bold=True, color=GREEN)
    text(s, 8.28, 2.54, 4.2, 0.28, "seeded anomalies detected", size=12, color=GREEN_DEEP)
    text(s, 8.28, 2.86, 4.2, 0.6,
         "No rule fires on a clean shipment, and no anomaly goes unreported.",
         size=11, color=GREEN_DEEP)

    card(s, 8.04, 3.74, 4.71, 1.98)
    text(s, 8.28, 3.92, 4.2, 0.3, "Why the counts differ", size=13.5, bold=True, color=PURPLE)
    text(s, 8.28, 4.28, 4.2, 1.3,
         [("R1 shows 2 - one warning and one failure, proving the three-band logic.", False, INK, 11.0),
          ("R4 shows 2 for a single replayed delivery, because both copies are flagged.", False, INK, 11.0),
          ("The other four fire exactly once each.", False, INK, 11.0)], spacing=6)

    card(s, L, 5.92, W, 0.88, fill=WHITE)
    text(s, L + 0.26, 6.08, 11.5, 0.6,
         [("Verified against an independent oracle.", True, PURPLE, 12.0),
          ("tools\\verify_seeded_data.py re-derives the expected outcome for all 48 "
           "transactions from the raw JSON and the ERP workbook. If the bot and the oracle "
           "ever disagree, one of them is wrong and the run is not signed off.", False, INK, 11.5)],
         size=12.0)
    notes(s, "The oracle is the point worth making here: we did not just assert our own output "
             "was right, we verified it against a deliberately separate implementation written "
             "from the data, not from the bot's code.")
    footer(s, n)


def slide_config(prs, n):
    s = add_slide(prs)
    header(s, "NO-CODE RULES · BLOCKCHAIN AGNOSTIC",
           "Novelty 1 and 3 - configuration is the product")

    card(s, L, 1.6, 5.94, 2.5)
    circle(s, L + 0.24, 1.8, 0.34, "1", fill=PURPLE)
    text(s, L + 0.74, 1.84, 4.8, 0.3, "Rules without code", size=15, bold=True, color=PURPLE)
    text(s, L + 0.24, 2.3, 5.5, 0.26, "Disable R4 in the spreadsheet, save, re-run:",
         size=11.5, color=INK)
    code_block(s, L + 0.24, 2.62, 5.5,
               ["[VALIDATE] R4 is disabled in Config.xlsx - skipped",
                "[VALIDATE] total=48 pass=42 warn=1 fail=5"], size=10.5)
    text(s, L + 0.24, 3.56, 5.5, 0.44,
         "Failures drop from 7 to 5. No workflow was opened and nothing was republished.",
         size=11, color=MUTED)

    card(s, 6.77, 1.6, 5.98, 2.5)
    circle(s, 7.01, 1.8, 0.34, "3", fill=PURPLE_MID)
    text(s, 7.51, 1.84, 4.8, 0.3, "Any blockchain", size=15, bold=True, color=PURPLE)
    hdr = ["NETWORK", "SOURCE", "RESULT"]
    xs2 = [7.01, 9.3, 11.0]
    for i, h in enumerate(hdr):
        text(s, xs2[i], 2.32, 2.2, 0.24, h, size=9, bold=True, color=MUTED)
    for j, row in enumerate([("Ethereum", "Etherscan V2", "48 / 40 / 1 / 7"),
                             ("Polygon", "Etherscan V2", "48 / 40 / 1 / 7"),
                             ("Besu", "ethereum-etl", "48 / 40 / 1 / 7")]):
        y = 2.62 + j * 0.36
        for i, v in enumerate(row):
            text(s, xs2[i], y, 2.3, 0.26, v, size=11,
                 bold=(i == 2), color=GREEN if i == 2 else INK)
    text(s, 7.01, 3.76, 5.5, 0.3,
         "Identical outcomes, with only config changed between runs.", size=11, color=MUTED)

    card(s, L, 4.26, W, 2.54)
    text(s, L + 0.26, 4.44, 11.5, 0.3, "Config.xlsx - the control panel", size=15,
         bold=True, color=PURPLE)
    text(s, L + 0.26, 4.78, 11.5, 0.3,
         "Seven sheets and sixty settings. Everything operational lives here, so the bot "
         "adapts without a developer.", size=11.5, color=INK)
    sheets = [("Settings", "Sources, mapping, escalation, alerts, dashboard, queue, output"),
              ("ValidationRules", "Which rules run, severity, thresholds, fail action"),
              ("ApprovedWallets", "The partner whitelist R3 checks against"),
              ("FieldMapping", "Etherscan record field to canonical field, with transform"),
              ("FieldMapping_EvmEtl", "The same, for an ethereum-etl ledger export"),
              ("EventSequence", "The milestone order R5 enforces, and each selector"),
              ("EventSignatures", "Function signatures, selectors and event topic hashes")]
    for j, (name, d) in enumerate(sheets):
        col, row = j % 2, j // 2
        x = L + 0.26 + col * 5.95
        y = 5.2 + row * 0.4
        text(s, x, y, 1.9, 0.26, name, size=11, bold=True, color=PINK, font=MONO_FONT)
        text(s, x + 2.0, y, 3.8, 0.26, d, size=10, color=MUTED)
    notes(s, "Demo moment: toggle R4 off live. Adithya leads. The point is that change control "
             "sits with the people who own the domain knowledge, not with developers. Note that "
             "the sheet count went from six to seven since Review 2 - FieldMapping_EvmEtl is "
             "the new one, and it is what makes a differently-shaped ledger export work.")
    footer(s, n)


def slide_human(prs, n):
    s = add_slide(prs)
    header(s, "CRITICAL ANOMALIES GO TO A PERSON, NOT A LOG FILE",
           "Novelty 2 - the bot knows when to stop")
    steps = [("Detect", "R3 finds a delivery signed by a wallet that is not on the approved-partner list."),
             ("Classify", "The rule is marked CRITICAL in Config.xlsx, with fail action ESCALATE."),
             ("Pause", "The review form shows the shipment, the wallet, the hash and the finding."),
             ("Record", "The decision, the reviewer's name and their note go into the results.")]
    for i, (t, d) in enumerate(steps):
        x = L + i * 3.0
        card(s, x, 1.6, 2.86, 1.72)
        circle(s, x + 0.2, 1.78, 0.34, str(i + 1), fill=PURPLE if i % 2 == 0 else PURPLE_MID)
        text(s, x + 0.2, 2.24, 2.5, 0.3, t, size=14.5, bold=True, color=PURPLE)
        text(s, x + 0.2, 2.58, 2.5, 0.64, d, size=11, color=INK)
        if i < 3:
            text(s, x + 2.84, 2.26, 0.24, 0.4, "›", size=26, bold=True, color=LILAC)

    card(s, L, 3.48, 6.0, 1.9, fill=GREEN_PALE, border=None, radius=0.03)
    text(s, L + 0.26, 3.66, 5.5, 0.32, "Why this matters", size=14, bold=True, color=GREEN)
    text(s, L + 0.26, 4.04, 5.5, 1.2,
         "The reviewer's decision is hashed into the audit chain alongside the machine's "
         "verdict. A human judgement becomes evidence, with the same tamper-evidence as "
         "everything the bot concluded on its own - not a note in a side system.",
         size=12, color=GREEN_DEEP)

    card(s, 6.77, 3.48, 5.98, 1.9, fill=AMBER_PALE, border=None, radius=0.03)
    text(s, 7.01, 3.66, 5.5, 0.32, "Safe when nobody is watching", size=14, bold=True, color=AMBER)
    text(s, 7.01, 4.04, 5.5, 1.2,
         "An unattended robot has no one to answer a dialog. The bot detects that and "
         "downgrades PROMPT to AUTO_LOG, recording the anomaly as DEFERRED and continuing, "
         "so a scheduled job can never hang on a prompt nobody will see.",
         size=12, color=AMBER_DEEP)

    card(s, L, 5.54, W, 1.26)
    text(s, L + 0.26, 5.72, 11.5, 0.3, "Four escalation modes, one config cell", size=13.5,
         bold=True, color=PURPLE)
    modes = [("PROMPT", "Shows the review form. Attended only."),
             ("SIMULATE", "Applies a preset decision, for tests and demos."),
             ("AUTO_LOG", "Defers and records, for unattended runs."),
             ("ReviewUi", "FORM or DIALOG, if no form runtime is present.")]
    for j, (m, d) in enumerate(modes):
        x = L + 0.26 + j * 2.95
        text(s, x, 6.08, 1.3, 0.26, m, size=10.5, bold=True, color=PINK, font=MONO_FONT)
        text(s, x, 6.36, 2.7, 0.4, d, size=10, color=MUTED)
    notes(s, "Nitin leads. Demo the form live if the projector allows - set AttendedMode True "
             "and EscalationMode PROMPT beforehand, and set it back to SIMULATE afterwards. "
             "Choose REJECTED so the decision shows up in the dashboard and the PDF. TC10 "
             "proves the form appears, times out to DEFERRED and keeps an evidence image, but "
             "a human clicking through it can only be rehearsed - do it at least once before "
             "the review.")
    footer(s, n)


def slide_audit(prs, n):
    s = add_slide(prs, dark=True)
    header(s, "THE STRONGEST CLAIM IN THE PROJECT",
           "Novelty 4 - an audit trail that proves itself", dark=True)
    text(s, L, 1.44, W, 0.6,
         "Every entry carries the hash of the entry before it - the same idea as the "
         "blockchain we are auditing, applied to the audit itself.",
         size=14, color=LILAC_PALE)

    parts = [("RecordHash", "fixes the verdict"), ("PrevHash", "links to the entry before"),
             ("EntryHash", "seals this entry")]
    for i, (t, d) in enumerate(parts):
        x = L + i * 4.18
        card(s, x, 2.14, 3.6, 0.92, fill=NIGHT_SOFT, border=PURPLE_MID)
        text(s, x + 0.24, 2.3, 3.1, 0.3, t, size=12.5, bold=True, color=PINK, font=MONO_FONT)
        text(s, x + 0.24, 2.62, 3.1, 0.3, d, size=11, color=LILAC_PALE)
        if i < 2:
            text(s, x + 3.72, 2.4, 0.3, 0.4, "→", size=24, bold=True, color=LILAC)

    text(s, L, 3.3, W, 0.26, "THE TEST: CHANGE ONE VERDICT FROM FAIL TO PASS",
         size=11.5, bold=True, color=PINK)
    code_block(s, L, 3.62, W,
               ["INVALID - first problem at line 7.",
                "  line 7: record contents were altered (RecordHash mismatch) - SHP-1005/Dispatched",
                "  line 7: EntryHash does not match its own contents"],
               size=11.5, fill=RGBColor(0x1A, 0x10, 0x38), color=RGBColor(0xFF, 0x9C, 0x93))

    stats = [("SHA-256", "over record, bot, machine, time"),
             ("1", "character changed is enough"),
             ("line 7", "named exactly, not just “invalid”")]
    for i, (big, lab) in enumerate(stats):
        x = L + i * 4.18
        card(s, x, 4.72, 3.9, 1.14, fill=NIGHT_SOFT, border=PURPLE_MID)
        text(s, x + 0.24, 4.88, 3.4, 0.5, big, size=27, bold=True, color=WHITE)
        text(s, x + 0.24, 5.42, 3.4, 0.3, lab, size=11.5, color=LILAC_PALE)

    text(s, L, 6.1, W, 0.6,
         "The log is append-only and spans every run. Altering any historical row breaks "
         "every link after it, so history cannot be quietly rewritten - and "
         "12_Verify_AuditChain is an entry point, so an auditor can re-verify without "
         "running the bot.",
         size=12, color=LILAC_PALE)
    notes(s, "This is the 90 seconds that separate the project from a data-comparison script. "
             "Open the CSV, show the three hash columns, then run tamper_test.ps1 live. It "
             "verifies clean, flips one FAIL to PASS - the exact edit somebody covering up an "
             "anomaly would make - detects it at the exact line, and restores the file. If "
             "asked why hash a log that is already about a blockchain: the transactions are "
             "on-chain, the validation is not. It happens off-chain, in the bot.")
    footer(s, n, dark=True)


def slide_evidence(prs, n):
    s = add_slide(prs)
    header(s, "WHAT THE BOT LEAVES BEHIND, EVERY RUN", "The evidence pack")
    items = [
        ("ValidationReport_*.xlsx", "Excel, ~23 KB",
         "Summary, per-rule breakdown, every result colour-coded, and an exceptions sheet."),
        ("AuditReport_*.pdf", "PDF 1.4, new",
         "The formal document for an auditor: bot identity, configuration, every finding with "
         "its reviewer decision, the chain head hash and a signature block."),
        ("dashboard.html", "Self-contained",
         "KPI tiles, pass/fail bar, findings, processing time per phase and trends across runs. "
         "No server, no CDN, no internet."),
        ("ValidationAuditLog.csv", "Append-only",
         "Hash-chained entries carrying BotIdentity, TimestampUtc, InputDataHash, RecordHash, "
         "PrevHash and EntryHash."),
        ("Alert_*.eml / _teams.json / _slack.json", "New",
         "The e-mail, the Teams Adaptive Card and the Slack Block Kit message, written to disk "
         "so delivery is testable without a real Teams or Slack channel."),
        ("dashboard_summary.json", "KPI export",
         "Findings and counters for a Power BI or Google Sheets push."),
    ]
    y = 1.56
    for i, (name, tag, d) in enumerate(items):
        card(s, L, y, W, 0.82)
        text(s, L + 0.26, y + 0.14, 4.2, 0.28, name, size=11.5, bold=True,
             color=PURPLE, font=MONO_FONT)
        text(s, L + 0.26, y + 0.46, 4.2, 0.26, tag, size=10, color=MUTED)
        text(s, 5.4, y + 0.2, 7.0, 0.56, d, size=11, color=INK)
        y += 0.88
    text(s, L, 6.86, W, 0.26,
         "The dashboard's run history is read back out of the audit log itself, so it cannot "
         "disagree with the evidence.", size=11, bold=True, color=GREEN)
    notes(s, "Nitin leads. Open the newest xlsx and the newest pdf from the run you just did. "
             "Read one FailureReasons cell aloud - it is written in plain English on purpose, "
             "because the person who reads it is a logistics manager, not a developer. The PDF "
             "and the webhook payloads are both new since Review 2.")
    footer(s, n)


def slide_queue(prs, n):
    s = add_slide(prs)
    header(s, "SAME PIPELINE, SPLIT PER SHIPMENT", "Queue mode and the Performer")

    card(s, L, 1.6, 5.94, 2.4)
    text(s, L + 0.24, 1.78, 5.5, 0.3, "How a run splits", size=15, bold=True, color=PURPLE)
    for j, (a, b) in enumerate([("BATCH", "One job validates all 48 transactions."),
                                ("DISPATCH", "One work item per shipment, in a local queue file: 48 → 12.")]):
        y = 2.2 + j * 0.74
        pill(s, L + 0.24, y, 1.3, 0.3, a, fill=PURPLE if j == 0 else PURPLE_MID, color=WHITE)
        text(s, L + 1.7, y + 0.02, 4.0, 0.6, b, size=11, color=INK)
    text(s, L + 0.24, 3.62, 5.5, 0.3, "RunMode is one cell in Config.xlsx.",
         size=10.5, color=MUTED)

    card(s, 6.77, 1.6, 5.98, 2.4)
    text(s, 7.01, 1.78, 5.5, 0.3, "Why a shipment, not a transaction", size=15,
         bold=True, color=PURPLE)
    text(s, 7.01, 2.18, 5.5, 1.6,
         "Three of the six rules compare events against each other, so a worker handed a "
         "single transaction could not evaluate duplicates, ordering or required events at "
         "all. A shipment is the smallest unit that keeps every rule meaningful - which is a "
         "design consequence of the rules, not an arbitrary choice.",
         size=11.5, color=INK)

    card(s, L, 4.16, W, 1.3, fill=WHITE)
    text(s, L + 0.26, 4.34, 11.5, 0.3, "The Performer is its own entry point", size=13.5,
         bold=True, color=PURPLE)
    text(s, L + 0.26, 4.68, 11.5, 0.66,
         "Run from Studio, it works through the queue and then calls 14_Finalise_Run - the same finalise "
         "pipeline a batch run uses. So queue mode produces the same audit trail, the same "
         "Excel and PDF reports and the same dashboard, rather than a second, divergent path "
         "that would need its own evidence.",
         size=11.5, color=INK)

    card(s, L, 5.62, W, 1.18, fill=AMBER_PALE, border=None, radius=0.03)
    text(s, L + 0.26, 5.8, 11.5, 0.3, "Runs in UiPath Studio - Orchestrator is out of scope",
         size=13, bold=True, color=AMBER)
    text(s, L + 0.26, 6.14, 11.5, 0.56,
         "The work queue is a local file. Dispatcher and Performer are verified end to end: "
         "48 transactions become 12 work items, and the Performer turns them into the same "
         "evidence pack as a batch run. Moving to an Orchestrator queue would swap that file "
         "for Add Queue Item and Get Transaction Item - the validation would not change.",
         size=11.5, color=AMBER_DEEP)
    notes(s, "Adithya leads. Say plainly that the queue is a local file and the project runs "
             "in Studio only - our Orchestrator access ended before the final review. The "
             "split per shipment is the design point. If there is time: set RunMode = DISPATCH "
             "in Config.xlsx, run Main, then open Workflows\\Queue\\Performer.xaml and Run "
             "File, and show that the counts and the audit trail match a batch run.")
    footer(s, n)


def slide_verification(prs, n, checks, tests, analyzer="0"):
    s = add_slide(prs)
    header(s, "ONE COMMAND VERIFIES EVERY CLAIM IN THIS DECK", "How we know it works")
    tiles = [(checks, "Acceptance checks passing"),
             (tests, "UiPath test cases passing"),
             (analyzer, "Workflow Analyzer errors"),
             ("58/58", "Invoke Code bodies compile")]
    for i, (big, lab) in enumerate(tiles):
        x = L + i * 3.0
        card(s, x, 1.6, 2.86, 1.3)
        text(s, x + 0.2, 1.74, 2.5, 0.56, big, size=30, bold=True, color=GREEN)
        text(s, x + 0.2, 2.36, 2.5, 0.4, lab, size=11, color=MUTED)

    card(s, L, 3.06, 6.5, 3.06)
    text(s, L + 0.24, 3.24, 6.0, 0.3, "Acceptance suite", size=15, bold=True, color=PURPLE)
    for j, line in enumerate([
        "Seeded anomalies present and the ABI decodes",
        "All six rules fire with the expected counts",
        "A rule disabled in Excel is skipped at runtime",
        "Identical result on a second chain and a third feed",
        "Excel, PDF, dashboard and KPI artefacts produced",
        "An edited audit log is detected at the right line",
        "All three extraction modes agree",
        "Dispatch splits 48 transactions into 12 work items",
        "The Performer works through the queue and finalises the run",
        "Alerts delivered to a local SMTP and webhook capture server",
    ]):
        y = 3.62 + j * 0.25
        text(s, L + 0.24, y, 0.2, 0.24, "✓", size=10.5, bold=True, color=GREEN)
        text(s, L + 0.5, y, 5.8, 0.24, line, size=10.5, color=INK)

    card(s, 7.33, 3.06, 5.42, 1.46)
    text(s, 7.57, 3.24, 5.0, 0.3, "Test cases in Studio's Test Explorer", size=14,
         bold=True, color=PURPLE)
    text(s, 7.57, 3.6, 5.0, 0.8,
         "Six drive one validation rule each against fixtures built inside the test, so a "
         "failure points at exactly one rule. The other six cover audit-chain tampering, "
         "field-mapping equivalence across two schemas, PDF structure, the review form's "
         "timeout, the Teams and Slack payloads, and ledger noise.",
         size=10.5, color=INK)

    card(s, 7.33, 4.66, 5.42, 1.46, fill=GREEN_PALE, border=None, radius=0.03)
    text(s, 7.57, 4.84, 5.0, 0.3, "A test found a real bug", size=13, bold=True, color=GREEN)
    text(s, 7.57, 5.2, 5.0, 0.8,
         "While writing the duplicate-detection test we discovered R4 assumed a transaction "
         "hash was at least twelve characters, and it crashed on a malformed one. Found, "
         "fixed and now covered - which is what tests are for.",
         size=10.5, color=GREEN_DEEP)

    card(s, L, 6.28, W, 0.52, fill=WHITE)
    text(s, L + 0.26, 6.42, 11.5, 0.3,
         [("powershell -ExecutionPolicy Bypass -File tools\\run_all_checks.ps1", True, PURPLE, 11.0)],
         font=MONO_FONT)
    notes(s, "Volunteer the bug story even if not asked - it is the most credible sentence "
             "available, because it is what actually happens when you test properly. Run the "
             "suite before the review, not during it: it takes about fifteen minutes and it "
             "answers an on-screen dialog itself near the end, so do not touch the mouse while "
             "it finishes.")
    footer(s, n)


def slide_status(prs, n):
    s = add_slide(prs)
    header(s, "STATED PLAINLY SO NOTHING IS OVERCLAIMED", "What works, and what does not")

    card(s, L, 1.6, 6.5, 4.5)
    pill(s, L + 0.24, 1.78, 2.2, 0.32, "EXECUTED AND VERIFIED", fill=GREEN, color=WHITE)
    done = [
        "Full pipeline end to end, 48 transactions in 21 seconds",
        "All six validation rules, against an independent oracle",
        "Three extraction modes producing identical results",
        "Three chain feeds across two different schemas",
        "DECODE and LOOKUP shipment linking",
        "Excel, CSV and SQL ERP sources",
        "Rule toggling from the spreadsheet with no code change",
        "Escalation: review window answered via UI Automation",
        "Hash-chained audit trail, tamper detected at the exact line",
        "Excel report, PDF audit report, HTML dashboard, KPI export",
        "Teams and Slack payloads, and alert delivery to a capture server",
        "Dispatcher and Performer end to end through a work queue",
    ]
    for j, line in enumerate(done):
        y = 2.26 + j * 0.3
        text(s, L + 0.24, y, 0.2, 0.26, "✓", size=11, bold=True, color=GREEN)
        text(s, L + 0.5, y, 5.8, 0.26, line, size=10.5, color=INK)

    card(s, 7.33, 1.6, 5.42, 4.5)
    pill(s, 7.57, 1.78, 1.9, 0.32, "NOT YET EXECUTED", fill=AMBER, color=WHITE)
    gaps = [
        ("Orchestrator - out of scope",
         "The project runs in UiPath Studio only. Queue mode uses a local work-queue file "
         "instead of an Orchestrator queue."),
        ("Live Etherscan key",
         "API mode is tested over HTTP against a local Etherscan stand-in, but no real key is "
         "configured, so it has not been run against the live chain."),
        ("Browser automation",
         "WEB mode parses the explorer's HTML. UIAutomation 25.10 removed the classic scraping "
         "activity and its replacement needs a recorded descriptor."),
        ("Real mail and chat endpoints",
         "Delivery is tested against a local capture server, not a real SMTP server, Outlook "
         "profile or Teams channel."),
    ]
    y = 2.24
    for t, d in gaps:
        text(s, 7.57, y, 5.0, 0.28, t, size=12, bold=True, color=AMBER)
        text(s, 7.57, y + 0.3, 5.0, 0.6, d, size=10.5, color=MUTED)
        y += 1.04

    card(s, L, 6.26, W, 0.56, fill=WHITE)
    text(s, L + 0.26, 6.42, 11.5, 0.3,
         "Nothing is stubbed. Each path above is exercised locally right up to the hop "
         "this project has no account or tool for.",
         size=11.5, bold=True, color=PURPLE)
    notes(s, "Do not skip this slide. Answer plainly if asked - waffling reads as not knowing. "
             "The honest summary: Orchestrator is out of scope because the project runs in "
             "Studio only, and the other gaps are accounts or tools, not capability - a live "
             "API key, the browser extension, and real mail and chat endpoints. "
             "Each path is exercised locally up to the final hop. This is also the slide that "
             "shortened most since Review 2, because the PDF, webhooks, ledger export and "
             "Performer all moved from the right column to the left.")
    footer(s, n)


def slide_close(prs, n):
    s = add_slide(prs, dark=True)
    # Sits low and right, clear of the closing card rather than cutting across it.
    shp = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(11.2), Inches(5.9), Inches(3.6), Inches(3.6))
    shp.fill.solid()
    shp.fill.fore_color.rgb = PURPLE
    shp.line.fill.background()
    shp.shadow.inherit = False

    header(s, "FUTURE SCOPE", "Where this goes next", dark=True)
    future = [
        ("Deploy to Orchestrator",
         "Publish the package, schedule it with time triggers, and swap the local work-queue "
         "file for an Orchestrator queue with parallel workers."),
        ("ML-based anomaly detection",
         "Use UiPath AI Center on historical transactions to catch fraud patterns no fixed rule "
         "anticipates."),
        ("IoT data fusion",
         "Validate on-chain records against temperature and GPS telemetry, so a cold-chain "
         "breach is caught as well as a clerical one."),
        ("Event-driven validation",
         "Replace scheduled polling with blockchain WebSocket subscriptions for sub-second "
         "latency."),
    ]
    for i, (t, d) in enumerate(future):
        x = L + (i % 2) * 6.15
        y = 1.5 + (i // 2) * 1.3
        card(s, x, y, 5.94, 1.14, fill=NIGHT_SOFT, border=PURPLE_MID)
        circle(s, x + 0.24, y + 0.2, 0.34, str(i + 1), fill=PINK, color=NIGHT)
        text(s, x + 0.74, y + 0.24, 4.9, 0.3, t, size=13.5, bold=True, color=WHITE)
        text(s, x + 0.74, y + 0.58, 4.9, 0.46, d, size=10.5, color=LILAC_PALE)

    card(s, L, 4.26, W, 1.36, fill=PURPLE, border=None, radius=0.03)
    text(s, L + 0.3, 4.46, 11.4, 1.0,
         [("Automated reconciliation of blockchain records against enterprise systems, six "
           "rules a business team can configure, escalation where human judgement belongs, and "
           "an audit trail that proves it was not altered afterwards.", True, WHITE, 14.5),
          ("Immutable is not the same as verified. This is the bot that does the verifying.",
           True, PINK, 13.0)], spacing=8)

    text(s, L, 6.3, W, 0.3,
         "github.com/AdithyaSM31/UiPath-Blockchain-Transaction-Validation",
         size=11.5, color=LILAC, font=MONO_FONT)
    text(s, L, 6.66, W, 0.3, "Thank you - questions welcome.", size=12, bold=True, color=WHITE)
    notes(s, "Close by returning to the opening line: immutable is not the same as verified, "
             "and this is the bot that does the verifying. Then invite questions. Expect: is "
             "blockchain data already trustworthy, why RPA rather than a Python script, why "
             "hash the audit log, is the data real, where are the tests, and what does not "
             "work - all six are answered in docs/DEMO_SCRIPT.md.")
    footer(s, n, dark=True)


# --------------------------------------------------------------------------- main
def build(path, date_label, checks, tests):
    prs = new_deck()
    slide_cover(prs, date_label)
    slide_whats_new(prs, 2)
    slide_problem(prs, 3)
    slide_architecture(prs, 4)
    slide_mapping(prs, 5)
    slide_rules(prs, 6)
    slide_results(prs, 7)
    slide_findings(prs, 8)
    slide_config(prs, 9)
    slide_human(prs, 10)
    slide_audit(prs, 11)
    slide_evidence(prs, 12)
    slide_queue(prs, 13)
    slide_verification(prs, 14, checks, tests)
    slide_status(prs, 15)
    slide_close(prs, 16)
    prs.save(path)
    return len(prs.slides.__iter__.__self__._sldIdLst)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default="21 October 2026",
                    help="date printed on the cover")
    ap.add_argument("--checks", default="21/21",
                    help="acceptance-suite result for the verification slide")
    ap.add_argument("--tests", default="12/12",
                    help="UiPath test-case result for the verification slide")
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out = args.out or os.path.join(root, "RPA Project Final Review.pptx")
    build(out, args.date, args.checks, args.tests)
    print(f"Wrote {out}")
    print(f"  cover date : {args.date}")
    print(f"  checks     : {args.checks}")
    print(f"  test cases : {args.tests}")

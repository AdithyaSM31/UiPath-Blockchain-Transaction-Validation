/*
    build_final_deck.js - generates "RPA Project Final Review.pptx" at the repository root.

        cd tools/deck
        npm install
        node build_final_deck.js [--date "15 November 2026"] [--checks 21/21] [--tests 12/12]

    The deck is built on a theme and two slide layouts (dark and content), so colours and
    fonts can be changed in one place. Speaker notes carry the talking points; the fuller
    version is docs/FINAL_REVIEW_SPEAKING_NOTES.md, which follows the same slide order.
    Images come from a real run: assets/review_window.png is the evidence image the review
    window saved, assets/dashboard.png the top of the dashboard the bot wrote.
*/
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");
const JSZip = require("jszip");

// ---------------------------------------------------------------- arguments
const arg = (name, dflt) => {
    const i = process.argv.indexOf("--" + name);
    return i > 0 && process.argv[i + 1] ? process.argv[i + 1] : dflt;
};
const DATE = arg("date", "October 2026");
const CHECKS = arg("checks", "21/21");
const TESTS = arg("tests", "12/12");
const OUT = arg("out", path.join(__dirname, "..", "..", "RPA Project Final Review.pptx"));
const ASSETS = path.join(__dirname, "assets");

// ---------------------------------------------------------------- theme
// Deep teal-ink and white. Teal means "verified", amber a warning, red a failure - the
// same three states the bot reports, so colour carries meaning on every slide.
const HEX = {
    ink: "16212B", deep: "0B2B33", white: "FFFFFF", mist: "EDF3F2",
    teal: "0E8577", amber: "C9821E", red: "C43D3D", green: "2B8F55",
    slate: "5D6C79", mint: "8FE0D2", deep2: "14404A", line: "D5E1DF",
};
const THEME = {
    name: "Ledger Teal",
    headFontFace: "Calibri",
    bodyFontFace: "Calibri",
    colors: {
        dk1: HEX.ink, lt1: HEX.white, dk2: HEX.deep, lt2: HEX.mist,
        accent1: HEX.teal, accent2: HEX.amber, accent3: HEX.red, accent4: HEX.green,
        accent5: HEX.slate, accent6: HEX.mint, hlink: HEX.teal, folHlink: HEX.slate,
    },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Final Project Review - Blockchain Logistics Validator";
pres.author = "Adithya Sankar Menon, Nitin Nandakumar";
pres.subject = "UiPath-Based Validation of Blockchain Transactions for Transparent Logistics";
const C = pres.SchemeColor;
const INK = C.text1, DEEP = C.text2, WHITE = C.background1, MIST = C.background2;
const TEAL = C.accent1, AMBER = C.accent2, RED = C.accent3, GREEN = C.accent4, SLATE = C.accent5, MINT = C.accent6;

const SW = 13.333, MX = 0.6, CW = SW - 2 * MX;
const MONO = "Consolas";

// ---------------------------------------------------------------- layouts
pres.defineSlideMaster({
    title: "DARK",
    background: { color: DEEP },
    objects: [],
});
pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: WHITE },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
        { placeholder: { options: { name: "kicker", type: "body", x: MX, y: 0.38, w: CW, h: 0.32,
            fontSize: 12, bold: true, color: TEAL, charSpacing: 2, margin: 0, valign: "middle" }, text: "" } },
        { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.72, w: CW, h: 0.78,
            fontSize: 32, bold: true, color: INK, margin: 0, valign: "top", align: "left", fontFace: THEME.headFontFace }, text: "" } },
        { text: { text: "Final Project Review  ·  Blockchain Logistics Validator",
            options: { x: MX, y: 7.02, w: 8, h: 0.3, fontSize: 10, color: SLATE, margin: 0, valign: "middle" } } },
    ],
    slideNumber: { x: SW - MX - 0.6, y: 7.08, w: 0.6, h: 0.25, fontSize: 10, color: SLATE, align: "right", margin: 0 },
});

// ---------------------------------------------------------------- helpers
const iconCache = {};
async function icon(name, hex) {
    const key = name + hex;
    if (iconCache[key]) return iconCache[key];
    const Comp = fa[name];
    if (!Comp) throw new Error("No react-icons/fa6 icon named " + name);
    const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { size: 256 }))
        .replace(/currentColor/g, "#" + hex);
    const buf = await sharp(Buffer.from(svg)).resize(256, 256, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
    return (iconCache[key] = "image/png;base64," + buf.toString("base64"));
}

/** An icon in a filled circle - the deck's visual motif. */
async function badge(slide, name, x, y, d, { fill = TEAL, fg = HEX.white, label } = {}) {
    slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" },
        objectName: label ? label + " badge" : undefined });
    const s = d * 0.5;
    slide.addImage({ data: await icon(name, fg), x: x + (d - s) / 2, y: y + (d - s) / 2, w: s, h: s,
        altText: label || name });
}

function card(slide, x, y, w, h, { fill = MIST, line, radius = 0.1, name } = {}) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: radius,
        fill: { color: fill }, line: line ? { color: line, width: 1 } : { type: "none" }, objectName: name });
}

function txt(slide, text, x, y, w, h, opts = {}) {
    slide.addText(text, { x, y, w, h, margin: 0, isTextBox: true, valign: "top", color: INK, fontSize: 14,
        fit: "none", ...opts });
}

function pill(slide, text, x, y, w, h, fill, color = WHITE, size = 12) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2, fill: { color: fill }, line: { type: "none" } });
    txt(slide, text, x, y, w, h, { fontSize: size, bold: true, color, align: "center", valign: "middle" });
}

function arrow(slide, x1, y1, x2, y2, color = HEX.slate, width = 2) {
    slide.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1) || 0.001,
        h: Math.abs(y2 - y1) || 0.001, flipV: y2 < y1, line: { color, width, endArrowType: "triangle" } });
}

function content(section, kicker, title) {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
    s.addText(kicker, { placeholder: "kicker" });
    s.addText(title, { placeholder: "title", align: "left" });
    return s;
}

const sections = new Set();
function section(title) {
    if (!sections.has(title)) { pres.addSection({ title }); sections.add(title); }
    return title;
}

// ================================================================ slides
async function build() {

    // ---- 1. Cover ------------------------------------------------------------
    {
        const sec = section("Opening");
        const s = pres.addSlide({ masterName: "DARK", sectionTitle: sec });
        txt(s, "FINAL PROJECT REVIEW  ·  ROBOTIC PROCESS AUTOMATION", 0.7, 0.95, 7.8, 0.35,
            { fontSize: 13, bold: true, color: MINT, charSpacing: 3 });
        txt(s, "UiPath-Based Validation of Blockchain Transactions for Transparent Logistics",
            0.7, 1.45, 7.6, 2.3, { fontSize: 38, bold: true, color: WHITE, fontFace: THEME.headFontFace, valign: "top" });
        txt(s, "Immutable is not the same as verified.", 0.7, 3.95, 7.6, 0.5,
            { fontSize: 22, italic: true, color: MINT });
        const people = [["Adithya Sankar Menon", "23BRS1079"], ["Nitin Nandakumar", "23BRS1378"]];
        people.forEach(([n, r], i) => {
            txt(s, n, 0.7 + i * 3.6, 5.2, 3.4, 0.4, { fontSize: 18, bold: true, color: WHITE });
            txt(s, r, 0.7 + i * 3.6, 5.6, 3.4, 0.35, { fontSize: 14, color: "A9C3C8" });
        });
        txt(s, DATE, 0.7, 6.45, 4, 0.35, { fontSize: 14, color: "A9C3C8" });

        // A short chain of blocks; the last one is the verified one.
        const blocks = [[8.6, 1.5], [10.05, 2.95], [11.5, 4.4]];
        for (let i = 0; i < blocks.length - 1; i++) {
            const [ax, ay] = blocks[i], [bx, by] = blocks[i + 1];
            s.addShape(pres.shapes.LINE, { x: ax + 0.6, y: ay + 0.6, w: bx - ax, h: by - ay, line: { color: HEX.mint, width: 2.5, dashType: "dash" } });
        }
        for (let i = 0; i < blocks.length; i++) {
            const [x, y] = blocks[i], last = i === blocks.length - 1;
            s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.2, h: 1.2, rectRadius: 0.16,
                fill: { color: last ? TEAL : HEX.deep2 }, line: { color: last ? HEX.teal : HEX.mint, width: 1.5 },
                objectName: "Chain block " + (i + 1) });
            s.addImage({ data: await icon(last ? "FaCheck" : "FaCube", last ? HEX.white : HEX.mint),
                x: x + 0.35, y: y + 0.35, w: 0.5, h: 0.5, altText: last ? "Verified block" : "Block" });
        }
        txt(s, "on-chain records", 8.6, 1.1, 2.4, 0.3, { fontSize: 12, color: "A9C3C8" });
        txt(s, "verified by the bot", 10.3, 5.7, 2.4, 0.3, { fontSize: 12, bold: true, color: MINT, align: "right" });
        s.addNotes("Nitin opens. One sentence before anything else: blockchain makes a record immutable, but it does not make it correct - that gap is the whole project. This is the final review of a finished system, and everything in the deck is reproducible with one command. Do not read the names off the slide.");
    }

    // ---- 2. The problem --------------------------------------------------------
    {
        const s = content(section("Opening"), "THE PROBLEM", "Blockchain makes a record immutable, not correct");
        txt(s, "A tampered quantity, a forged timestamp or an unauthorised wallet is recorded just as permanently as a legitimate one.",
            MX, 1.85, 5.7, 2.0, { fontSize: 24, color: INK });
        card(s, MX, 4.25, 5.7, 1.9, { name: "Key point" });
        await badge(s, "FaTriangleExclamation", MX + 0.35, 4.6, 0.8, { fill: AMBER, label: "Warning" });
        txt(s, "Every anomaly this bot catches is a perfectly valid on-chain transaction.",
            MX + 1.4, 4.55, 4.05, 1.3, { fontSize: 18, bold: true, color: INK, valign: "middle" });

        const rows = [
            ["FaHand", "Manual", "Operators cross-check transaction hashes against ERP records by hand, one at a time."],
            ["FaLayerGroup", "Siloed", "Explorers speak JSON and hex; ERP systems speak Excel and SQL."],
            ["FaFileCircleQuestion", "Unprovable", "A hand-kept verification log is not evidence: anyone can edit it afterwards."],
        ];
        for (let i = 0; i < rows.length; i++) {
            const y = 1.9 + i * 1.45, [ic, h, d] = rows[i];
            await badge(s, ic, 6.95, y, 0.75, { label: h });
            txt(s, h, 7.95, y - 0.02, 4.8, 0.4, { fontSize: 20, bold: true });
            txt(s, d, 7.95, y + 0.42, 4.8, 0.85, { fontSize: 15, color: SLATE });
        }
        s.addNotes("Nitin, no computer needed. A pharmaceutical shipment moves through six companies; each handoff is written to a blockchain, so the record is immutable. But immutable is not correct - a tampered quantity, a forged timestamp or an unauthorised wallet is recorded just as permanently as a legitimate one. Today somebody cross-checks those records against the ERP by hand. At thousands a day, nobody can. Land immutable vs verified - it is the point most likely to be probed.");
    }

    // ---- 3. Since Review 2 -------------------------------------------------------
    {
        const s = content(section("Opening"), "SINCE REVIEW 2", "Six gaps closed since the last review");
        const items = [
            ["FaTableCells", "Field-mapping engine", "Readers fetch; a spreadsheet translates. A new feed schema needs a sheet, not a workflow."],
            ["FaBroom", "Ledger noise handled", "Value transfers and calls to other functions are set aside, not mistaken for fraud."],
            ["FaLink", "LOOKUP linking", "When a contract carries no shipment key, a mapping file supplies it."],
            ["FaFilePdf", "PDF audit report", "The document handed to an auditor: findings, decisions, chain head, signature block."],
            ["FaBell", "Teams and Slack alerts", "An Adaptive Card and a Block Kit message alongside the e-mail."],
            ["FaLayerGroup", "Work queue in Studio", "Dispatcher and Performer split a run per shipment, with no Orchestrator."],
        ];
        const w = (CW - 2 * 0.3) / 3, h = 2.25;
        for (let i = 0; i < items.length; i++) {
            const x = MX + (i % 3) * (w + 0.3), y = 1.85 + Math.floor(i / 3) * (h + 0.3), [ic, t, d] = items[i];
            card(s, x, y, w, h, { name: t });
            await badge(s, ic, x + 0.3, y + 0.3, 0.65, { label: t });
            txt(s, t, x + 1.1, y + 0.3, w - 1.35, 0.65, { fontSize: 17, bold: true, valign: "middle" });
            txt(s, d, x + 0.3, y + 1.15, w - 0.6, 0.95, { fontSize: 14, color: SLATE });
        }
        s.addNotes("Either presenter, about a minute. At Review 2 the pipeline worked end to end, but six commitments from Review 1 were still open. Go down the list one clause each - the mapping engine ('the readers got dumber and a spreadsheet got smarter'), ledger noise, LOOKUP, the PDF, Teams and Slack, and the work queue that now runs entirely in Studio. Each gets its own slide or demo later.");
    }

    // ---- 4. Architecture ---------------------------------------------------------------
    {
        const s = content(section("The system"), "ARCHITECTURE", "28 workflows, orchestrated by Main.xaml");
        const lanes = [
            ["1 · INGEST", [["00", "Read config"], ["01", "Extract chain"], ["01d", "Field mapping"], ["02", "Extract ERP"], ["03", "Map and join"]]],
            ["2 · VALIDATE", [["05", "Rules R1-R6"], ["06", "Human review"]]],
            ["3 · EVIDENCE", [["10", "Hash-chained log"], ["12", "Verify chain"]]],
            ["4 · REPORT", [["07", "Excel report"], ["08", "PDF audit report"], ["13", "Dashboard"], ["09", "E-mail alert"], ["09b", "Teams / Slack"]]],
        ];
        const lw = (CW - 3 * 0.3) / 4, ly = 1.8, lh = 3.55;
        lanes.forEach(([head, chips], i) => {
            const x = MX + i * (lw + 0.3);
            card(s, x, ly, lw, lh, { name: head });
            txt(s, head, x + 0.25, ly + 0.2, lw - 0.5, 0.35, { fontSize: 14, bold: true, color: TEAL, charSpacing: 1 });
            chips.forEach(([code, name], j) => {
                const cy = ly + 0.7 + j * 0.55;
                s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.25, y: cy, w: lw - 0.5, h: 0.45, rectRadius: 0.08,
                    fill: { color: WHITE }, line: { color: HEX.line, width: 0.75 } });
                txt(s, [{ text: code + "   ", options: { bold: true, color: TEAL } }, { text: name, options: { color: INK } }],
                    x + 0.4, cy, lw - 0.8, 0.45, { fontSize: 14, valign: "middle" });
            });
            const note = { 1: "Six rule workflows, switched on and tuned from the ValidationRules sheet. CRITICAL findings pause for a person.",
                           2: "Every verdict is appended to a hash-chained CSV, then the whole chain is re-verified." }[i];
            if (note) txt(s, note, x + 0.25, ly + 0.7 + chips.length * 0.55 + 0.15, lw - 0.5, 1.6, { fontSize: 13, color: SLATE });
        });
        const principles = [
            ["FaDiagramProject", "Structure in XAML", "Sequences, Excel, database, mail and HTTP stay visual in Studio."],
            ["FaCode", "Algorithms in Invoke Code", "ABI decoding, hashing and PDF layout in VB code."],
            ["FaArrowsRotate", "One finalise pipeline", "Batch runs and the Performer share 14_Finalise_Run."],
        ];
        const pw = (CW - 2 * 0.3) / 3;
        for (let i = 0; i < principles.length; i++) {
            const x = MX + i * (pw + 0.3), y = 5.6, [ic, t, d] = principles[i];
            await badge(s, ic, x, y + 0.08, 0.6, { label: t });
            txt(s, t, x + 0.8, y, pw - 0.8, 0.4, { fontSize: 16, bold: true });
            txt(s, d, x + 0.8, y + 0.42, pw - 0.8, 0.75, { fontSize: 13, color: SLATE });
        }
        s.addNotes("Adithya leads. Four phases: ingest, validate, evidence, report. The point to emphasise is 14_Finalise_Run - batch mode and the queue Performer both call it, so queue mode is not a second, divergent code path. If asked why Invoke Code rather than all activities: three of the six rules compare rows against one another, and a state machine drawn as activities would be unreadable.");
    }

    // ---- 5. Mapping engine ------------------------------------------------------------
    {
        const s = content(section("The system"), "THE MAPPING ENGINE", "Readers fetch. A spreadsheet translates.");
        const srcs = [["FaEthereum", "Ethereum", "Etherscan V2"], ["FaCubesStacked", "Polygon", "Etherscan V2"], ["FaServer", "Hyperledger Besu", "ethereum-etl export"]];
        for (let i = 0; i < srcs.length; i++) {
            const y = 1.9 + i * 1.0, [ic, n, d] = srcs[i];
            card(s, MX, y, 2.6, 0.8, { fill: WHITE, line: HEX.line, name: n });
            s.addImage({ data: await icon(ic, HEX.teal), x: MX + 0.2, y: y + 0.2, w: 0.4, h: 0.4, altText: n });
            txt(s, n, MX + 0.75, y + 0.1, 1.8, 0.32, { fontSize: 14, bold: true });
            txt(s, d, MX + 0.75, y + 0.42, 1.8, 0.3, { fontSize: 12, color: SLATE });
            arrow(s, MX + 2.65, y + 0.4, 3.75, 3.3, HEX.slate, 1.5);
        }
        card(s, 3.8, 1.9, 2.15, 2.8, { fill: TEAL, name: "FieldMapping sheet" });
        s.addImage({ data: await icon("FaTableCells", HEX.white), x: 4.6, y: 2.15, w: 0.55, h: 0.55, altText: "Mapping sheet" });
        txt(s, "FieldMapping sheet", 3.95, 2.85, 1.85, 0.65, { fontSize: 16, bold: true, color: WHITE, align: "center" });
        txt(s, "field → column, with a transform", 3.95, 3.55, 1.85, 0.9, { fontSize: 13, color: WHITE, align: "center" });
        arrow(s, 6.0, 3.3, 6.45, 3.3, HEX.slate, 2);
        card(s, 6.5, 2.3, 1.7, 2.0, { fill: WHITE, line: HEX.teal, name: "Canonical table" });
        txt(s, "One canonical table", 6.6, 2.5, 1.5, 0.75, { fontSize: 15, bold: true, align: "center" });
        txt(s, "the rules never see the source", 6.6, 3.3, 1.5, 0.9, { fontSize: 12, color: SLATE, align: "center" });

        card(s, MX, 5.15, 7.6, 1.55, { name: "Transforms" });
        txt(s, "One transform per field, chosen in Excel", MX + 0.3, 5.3, 7.0, 0.4, { fontSize: 16, bold: true });
        txt(s, "hex → integer  ·  Unix epoch → UTC time  ·  ISO text → UTC time  ·  ABI string / uint256 / address decoding  ·  decimal  ·  boolean  ·  case",
            MX + 0.3, 5.72, 7.0, 0.85, { fontSize: 14, color: SLATE });

        const rx = 8.6, rw = SW - MX - rx;
        txt(s, "Linking a transaction to a shipment", rx, 1.85, rw, 0.4, { fontSize: 17, bold: true });
        const modes = [["DECODE", "The shipment ID is read straight from the transaction's ABI call data."],
                       ["LOOKUP", "No shipment key on-chain: a mapping file supplies it."]];
        modes.forEach(([m, d], i) => {
            const y = 2.4 + i * 1.3;
            card(s, rx, y, rw, 1.15, { name: m });
            pill(s, m, rx + 0.25, y + 0.25, 1.15, 0.36, TEAL, WHITE, 12);
            txt(s, d, rx + 1.55, y + 0.17, rw - 1.75, 0.85, { fontSize: 13.5, color: INK, valign: "middle" });
        });
        card(s, rx, 5.15, rw, 1.55, { fill: WHITE, line: HEX.line, name: "Ledger noise" });
        await badge(s, "FaBroom", rx + 0.25, 5.35, 0.6, { fill: AMBER, label: "Ledger noise" });
        txt(s, "Ledger noise is set aside", rx + 1.05, 5.3, rw - 1.25, 0.4, { fontSize: 15, bold: true });
        txt(s, "Transfers and calls to other functions are counted and skipped, not reported as fraud.",
            rx + 1.05, 5.7, rw - 1.25, 0.9, { fontSize: 13, color: SLATE });
        s.addNotes("Adithya leads - new since Review 2. The three readers now only fetch; 01d builds the canonical table purely from the selected FieldMapping sheet, so a feed with different field names needs a new sheet, not a new workflow. The Besu export has different field names, hex numbers and a different time format, and reaches the same verdict. DECODE reads the shipment key from call data; LOOKUP takes it from a mapping file. And real ledgers carry noise - transfers and other contract calls - which is set aside, not flagged (TC12).");
    }

    // ---- 6. Six rules -----------------------------------------------------------------
    {
        const s = content(section("The system"), "VALIDATION ENGINE", "Six rules, each catching a real kind of fraud");
        const rules = [
            ["R1", "Timestamp", "On-chain time vs the ERP's expected time, with a warning band and a fail band.", "SHP-1002 warns, SHP-1003 fails"],
            ["R2", "Quantity", "The decoded on-chain quantity must match the purchase order.", "SHP-1005 short by 20 units"],
            ["R3", "Wallet whitelist", "The signer must be an approved partner wallet. CRITICAL: escalates.", "SHP-1007, unapproved wallet"],
            ["R4", "Duplicates", "No transaction hash may be recorded twice for one shipment.", "SHP-1009 delivery replayed"],
            ["R5", "Sequence", "Milestones must happen in the order set in Config.xlsx.", "SHP-1011 out of order"],
            ["R6", "Contract event", "Every shipment must emit its required smart-contract events.", "SHP-1012 never confirmed receipt"],
        ];
        const w = (CW - 2 * 0.3) / 3, h = 2.3;
        rules.forEach(([id, n, d, caught], i) => {
            const x = MX + (i % 3) * (w + 0.3), y = 1.8 + Math.floor(i / 3) * (h + 0.3);
            card(s, x, y, w, h, { name: id });
            pill(s, id, x + 0.3, y + 0.3, 0.75, 0.42, TEAL, WHITE, 15);
            txt(s, n, x + 1.25, y + 0.3, w - 1.5, 0.42, { fontSize: 18, bold: true, valign: "middle" });
            txt(s, d, x + 0.3, y + 0.92, w - 0.6, 0.85, { fontSize: 14, color: SLATE });
            txt(s, [{ text: "Caught: ", options: { bold: true, color: RED } }, { text: caught, options: { color: INK } }],
                x + 0.3, y + 1.8, w - 0.6, 0.35, { fontSize: 13 });
        });
        s.addNotes("Adithya leads. Each rule is its own workflow, and each catches one seeded anomaly. R4 and R5 are set-based - they compare rows against each other - which is why rules take the whole table rather than one row at a time. R3 is CRITICAL, so it escalates to a human (slide 10). Every threshold and severity lives in Config.xlsx.");
    }

    // ---- 7. Results ---------------------------------------------------------------------
    {
        const s = content(section("The system"), "RESULTS", "One run: 48 transactions checked in about 24 seconds");
        const stats = [["48", "transactions validated", TEAL], ["40", "passed every rule", GREEN],
                       ["8", "flagged: 7 failed, 1 warning", RED], ["1", "escalated to a human", AMBER]];
        const sw = 2.75, sh = 1.5;
        stats.forEach(([n, l, col], i) => {
            const x = MX + (i % 2) * (sw + 0.3), y = 1.85 + Math.floor(i / 2) * (sh + 0.3);
            card(s, x, y, sw, sh, { name: l });
            txt(s, n, x + 0.3, y + 0.15, sw - 0.6, 0.8, { fontSize: 44, bold: true, color: col });
            txt(s, l, x + 0.3, y + 0.95, sw - 0.6, 0.45, { fontSize: 14, color: SLATE });
        });
        card(s, MX, 5.5, 2 * sw + 0.3, 1.2, { fill: WHITE, line: HEX.line, name: "Determinism" });
        await badge(s, "FaClock", MX + 0.25, 5.75, 0.7, { label: "Time" });
        txt(s, "About 24 s end to end on one robot, against 2-3 minutes per transaction by hand. The same input gives the same verdict every run.",
            MX + 1.15, 5.62, 2 * sw + 0.3 - 1.35, 1.0, { fontSize: 14, color: INK, valign: "middle" });

        s.addChart(pres.charts.DOUGHNUT, [{ name: "Verdicts", labels: ["Pass", "Warning", "Fail"], values: [40, 1, 7] }], {
            x: 7.0, y: 1.75, w: 5.7, h: 4.95, holeSize: 62,
            chartColors: [HEX.green, HEX.amber, HEX.red],
            showTitle: true, title: "Verdicts across 48 transactions", titleFontSize: 15, titleColor: HEX.ink, titleFontFace: "+mn-lt",
            showLegend: true, legendPos: "b", legendFontSize: 13, legendColor: HEX.ink, legendFontFace: "+mn-lt",
            showValue: true, showPercent: false, dataLabelColor: HEX.white, dataLabelFontSize: 14, dataLabelFontBold: true,
            dataLabelFontFace: "+mn-lt",
        });
        txt(s, "83.3%", 8.9, 3.65, 1.9, 0.6, { fontSize: 30, bold: true, color: GREEN, align: "center" });
        txt(s, "pass rate", 8.9, 4.2, 1.9, 0.35, { fontSize: 13, color: SLATE, align: "center" });
        s.addNotes("Adithya leads. About 24 seconds of workflow execution on this laptop. Against a manual baseline of two to three minutes per transaction this supports the 90%+ effort reduction from Review 1 - present it as an estimate against a stated baseline, not a measured comparison. The determinism point matters: the same 48 transactions give the same verdict every run, which is what lets the test suite assert exact counts.");
    }

    // ---- 8. Findings per rule --------------------------------------------------------
    {
        const s = content(section("The system"), "FINDINGS PER RULE", "Every rule earns its place: 7 of 7 anomalies found");
        s.addChart(pres.charts.BAR, [{ name: "Findings", labels: ["R1 Timestamp", "R2 Quantity", "R3 Wallet", "R4 Duplicates", "R5 Sequence", "R6 Contract event"],
            values: [2, 1, 1, 2, 1, 1] }], {
            x: MX, y: 1.75, w: 7.0, h: 4.95, barDir: "bar", barGapWidthPct: 55,
            chartColors: [HEX.teal],
            showTitle: true, title: "Rows flagged by each rule", titleFontSize: 15, titleColor: HEX.ink, titleFontFace: "+mn-lt",
            catAxisLabelColor: HEX.ink, catAxisLabelFontSize: 14, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
            catAxisLineShow: false, valAxisHidden: true, valAxisMaxVal: 2.5, valAxisMinVal: 0,
            valGridLine: { style: "none" }, catGridLine: { style: "none" },
            showValue: true, dataLabelPosition: "outEnd", dataLabelColor: HEX.ink, dataLabelFontSize: 14, dataLabelFontBold: true,
            dataLabelFontFace: "+mn-lt", showLegend: false,
        });
        const rx = 8.1, rw = SW - MX - rx;
        txt(s, "0", rx, 1.85, 1.2, 0.95, { fontSize: 60, bold: true, color: GREEN });
        txt(s, "false positives: the five clean shipments passed every rule on every event", rx + 1.25, 1.95, rw - 1.25, 0.85, { fontSize: 15, color: INK, valign: "middle" });
        card(s, rx, 3.05, rw, 1.65, { name: "Why the counts differ" });
        txt(s, "Why the counts differ", rx + 0.3, 3.2, rw - 0.6, 0.38, { fontSize: 15, bold: true });
        txt(s, "R1 shows a warning and a failure. R4 flags both copies of one replayed delivery: the ledger alone cannot tell which copy was real.",
            rx + 0.3, 3.6, rw - 0.6, 1.0, { fontSize: 13, color: SLATE });
        card(s, rx, 4.95, rw, 1.75, { fill: WHITE, line: HEX.line, name: "Oracle" });
        await badge(s, "FaScaleBalanced", rx + 0.25, 5.2, 0.6, { label: "Oracle" });
        txt(s, "Checked by an independent oracle", rx + 1.05, 5.12, rw - 1.25, 0.4, { fontSize: 15, bold: true });
        txt(s, "A separate script re-derives every expected verdict from the raw data. If it and the bot disagree, the run is not signed off.",
            rx + 1.05, 5.52, rw - 1.25, 1.1, { fontSize: 13, color: SLATE });
        s.addNotes("Adithya - drop if short on time. The oracle is the point: we did not just assert our own output was right, we checked it against a deliberately separate implementation (tools/verify_seeded_data.py) written from the data, not from the bot's code.");
    }

    // ---- 9. Novelty 1 and 3 -------------------------------------------------------------
    {
        const s = content(section("Novelties"), "NOVELTY 1 AND 3  ·  CONFIGURATION IS THE PRODUCT", "Change the spreadsheet, not the workflow");
        const head = ["Rule", "Name", "Enabled", "Severity"].map((t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: TEAL } } }));
        const data = [["R1", "TimestampCheck", "TRUE", "HIGH"], ["R2", "QuantityMatch", "TRUE", "HIGH"], ["R3", "AddressWhitelist", "TRUE", "CRITICAL"],
                      ["R4", "DuplicateDetection", "FALSE", "MEDIUM"], ["R5", "SequenceValidation", "TRUE", "HIGH"], ["R6", "SmartContractEvent", "TRUE", "HIGH"]];
        const rows = [head].concat(data.map((r) => r.map((c, j) => {
            const off = r[0] === "R4";
            return { text: c, options: { color: off && j === 2 ? RED : INK, bold: off && j === 2, fill: { color: off ? "FBE9E7" : WHITE } } };
        })));
        txt(s, "Config.xlsx → ValidationRules", MX, 1.8, 6.0, 0.35, { fontSize: 14, bold: true, color: SLATE });
        s.addTable(rows, { x: MX, y: 2.2, w: 6.0, colW: [0.8, 2.6, 1.2, 1.4], fontSize: 14, fontFace: THEME.bodyFontFace,
            rowH: 0.42, border: { type: "solid", pt: 0.75, color: HEX.line }, valign: "middle", margin: [0, 0.1, 0, 0.12] });
        card(s, MX, 5.45, 6.0, 1.15, { name: "Result" });
        txt(s, [{ text: "Failures 7 → 5", options: { bold: true, color: TEAL, fontSize: 22, breakLine: true } },
                { text: "R4 skipped at run time. No workflow opened, nothing republished.", options: { color: SLATE, fontSize: 14 } }],
            MX + 0.3, 5.45, 5.4, 1.15, { valign: "middle" });

        const rx = 7.1, rw = SW - MX - rx;
        txt(s, "Any blockchain: only config cells change", rx, 1.8, rw, 0.35, { fontSize: 14, bold: true, color: SLATE });
        const feeds = [["Feed", "Source", "Result"], ["Ethereum", "Etherscan V2", "48 / 40 / 1 / 7"], ["Polygon", "Etherscan V2", "48 / 40 / 1 / 7"], ["Hyperledger Besu", "ethereum-etl", "48 / 40 / 1 / 7"]];
        s.addTable(feeds.map((r, i) => r.map((c) => ({ text: c, options: i === 0 ? { bold: true, color: WHITE, fill: { color: TEAL } } : { color: INK, fill: { color: WHITE }, bold: c.startsWith("48") } }))),
            { x: rx, y: 2.2, w: rw, colW: [2.0, 1.8, rw - 3.8], fontSize: 14, fontFace: THEME.bodyFontFace, rowH: 0.42,
              border: { type: "solid", pt: 0.75, color: HEX.line }, valign: "middle", margin: [0, 0.1, 0, 0.12] });
        txt(s, "total / pass / warn / fail - identical on all three", rx, 3.95, rw, 0.3, { fontSize: 12, color: SLATE });
        card(s, rx, 4.4, rw, 2.3, { name: "Config sheets" });
        txt(s, "Config.xlsx is the control panel", rx + 0.3, 4.55, rw - 0.6, 0.4, { fontSize: 15, bold: true });
        txt(s, "7 sheets and 58 settings: rules, approved wallets, milestone order, function signatures, field mappings, data sources, alert channels. Change control sits with the people who own the domain knowledge.",
            rx + 0.3, 4.98, rw - 0.6, 1.6, { fontSize: 13.5, color: SLATE });
        s.addNotes("Adithya - demo moment: toggle R4 to FALSE in Config.xlsx, save, re-run from Studio. The log says 'R4 is disabled in Config.xlsx - skipped' and failures drop from 7 to 5. Then Novelty 3: Ethereum, Polygon and a Besu export reach identical results with only configuration changed (tools/chain_swap_demo.ps1). Set R4 back to TRUE afterwards.");
    }

    // ---- 10. Novelty 2 -------------------------------------------------------------------
    {
        const s = content(section("Novelties"), "NOVELTY 2  ·  HUMAN IN THE LOOP", "Critical anomalies go to a person, not a log file");
        const steps = [["Detect", "R3 finds a delivery signed by a wallet that is not an approved partner."],
                       ["Classify", "The rule is CRITICAL in Config.xlsx, with fail action ESCALATE."],
                       ["Pause", "The review window shows every fact; submit stays locked until a decision is chosen."],
                       ["Record", "Decision, reviewer and note are hashed into the audit trail."]];
        steps.forEach(([h, d], i) => {
            const y = 1.85 + i * 1.05;
            s.addShape(pres.shapes.OVAL, { x: MX, y, w: 0.6, h: 0.6, fill: { color: TEAL }, line: { type: "none" } });
            txt(s, String(i + 1), MX, y, 0.6, 0.6, { fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle" });
            txt(s, h, MX + 0.85, y - 0.04, 5.0, 0.38, { fontSize: 18, bold: true });
            txt(s, d, MX + 0.85, y + 0.36, 5.0, 0.65, { fontSize: 14, color: SLATE });
        });
        card(s, MX, 6.05, 5.85, 0.72, { fill: WHITE, line: HEX.line, name: "Unattended" });
        txt(s, "No one at the screen? PROMPT safely becomes AUTO_LOG and records the item as DEFERRED.", MX + 0.2, 6.05, 5.45, 0.72,
            { fontSize: 13, color: INK, valign: "middle" });
        const ih = 4.95, iw = ih * 766 / 670;
        const ix = 6.8 + (SW - MX - 6.8 - iw) / 2;
        s.addShape(pres.shapes.RECTANGLE, { x: ix, y: 1.85, w: iw, h: ih, fill: { color: WHITE }, line: { color: HEX.line, width: 1 },
            shadow: { type: "outer", color: "000000", opacity: 0.18, blur: 8, offset: 3, angle: 90 } });
        s.addImage({ path: path.join(ASSETS, "review_window.png"), x: ix, y: 1.85, w: iw, h: ih,
            altText: "The review window the bot shows for SHP-1007: shipment facts, the R3 finding, decision options, reviewer name and note" });
        s.addNotes("Nitin leads. Demo live if the projector allows: set EscalationMode = PROMPT and AttendedMode = True, run Main, and the window appears for SHP-1007. Choose REJECTED, type your name and a note, press Record decision and resume. The decision shows up in the dashboard and the PDF, and it is hashed into the audit chain like the machine's verdict. Set EscalationMode back to SIMULATE afterwards. tools/review_window_test.ps1 drives this same window automatically.");
    }

    // ---- 11. Novelty 4 ---------------------------------------------------------------------
    {
        const s = content(section("Novelties"), "NOVELTY 4  ·  TAMPER-EVIDENT AUDIT TRAIL", "An audit trail that proves it was never edited");
        const ents = [["Line 6", "SHP-1004 · PASS", "9f2c41", "b7a0e3", false], ["Line 7", "SHP-1005 · FAIL → PASS", "41ab9d", "9f2c41", true], ["Line 8", "SHP-1005 · PASS", "c07e5a", "41ab9d", false]];
        const ew = 3.55, gap = (CW - 3 * ew) / 2;
        ents.forEach(([ln, what, entry, prev, bad], i) => {
            const x = MX + i * (ew + gap), y = 1.85;
            card(s, x, y, ew, 1.95, { fill: bad ? "FBE9E7" : MIST, line: bad ? HEX.red : undefined, name: ln });
            txt(s, ln, x + 0.25, y + 0.18, 1.5, 0.35, { fontSize: 14, bold: true, color: bad ? RED : TEAL });
            txt(s, what, x + 0.25, y + 0.55, ew - 0.5, 0.35, { fontSize: 15, bold: true, color: INK });
            txt(s, [{ text: "PrevHash   " + prev + "…", options: { breakLine: true } }, { text: "EntryHash  " + entry + "…" }],
                x + 0.25, y + 1.0, ew - 0.5, 0.75, { fontSize: 13, fontFace: MONO, color: SLATE });
            if (i < ents.length - 1) arrow(s, x + ew + 0.05, y + 0.98, x + ew + gap - 0.05, y + 0.98, HEX.slate, 2);
        });
        s.addImage({ data: await icon("FaCircleXmark", HEX.red), x: MX + ew + gap + ew - 0.62, y: 2.0, w: 0.42, h: 0.42, altText: "Tampered" });

        card(s, MX, 4.15, 6.1, 2.55, { name: "Each entry carries" });
        txt(s, "Each entry carries", MX + 0.3, 4.3, 5.5, 0.4, { fontSize: 16, bold: true });
        const f = ["Bot identity and machine", "UTC timestamp", "Hash of the input data", "Verdict, severity, reasons", "Human decision and reviewer", "Three linked hashes"];
        f.forEach((t, i) => {
            const x = MX + 0.3 + (i % 2) * 2.85, y = 4.8 + Math.floor(i / 2) * 0.58;
            txt(s, [{ text: "✓  ", options: { color: GREEN, bold: true } }, { text: t }], x, y, 2.8, 0.5, { fontSize: 14, color: INK });
        });
        const dx = MX + 6.4, dw = SW - MX - dx;
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: dx, y: 4.15, w: dw, h: 2.55, rectRadius: 0.1, fill: { color: DEEP }, line: { type: "none" } });
        txt(s, "Change one verdict, re-verify:", dx + 0.3, 4.3, dw - 0.6, 0.4, { fontSize: 15, bold: true, color: WHITE });
        txt(s, [{ text: "INVALID - first problem at line 7", options: { color: "FF8A80", bold: true, breakLine: true } },
                { text: "record contents were altered", options: { color: MINT, breakLine: true } },
                { text: "(RecordHash mismatch) - SHP-1005", options: { color: MINT } }],
            dx + 0.3, 4.8, dw - 0.6, 1.1, { fontSize: 14, fontFace: MONO });
        txt(s, "You cannot quietly edit history without rewriting every line after it.", dx + 0.3, 5.95, dw - 0.6, 0.65,
            { fontSize: 13, color: "C9DCDD" });
        s.addNotes("Nitin - the strongest moment. Each entry carries the hash of the one before it: the same idea as the blockchain we are auditing, applied to the audit itself. Demo: run tools/tamper_test.ps1. It verifies clean, flips one FAIL verdict to PASS - the exact edit someone covering up an anomaly would make - re-verifies, and names the exact line. Then it restores the file and verifies clean again.");
    }

    // ---- 12. Evidence pack -------------------------------------------------------------------
    {
        const s = content(section("Evidence"), "EVIDENCE PACK", "Every run produces its own evidence");
        const iw = 7.3, ih = iw * 520 / 1140;
        s.addShape(pres.shapes.RECTANGLE, { x: MX, y: 1.85, w: iw, h: ih, fill: { color: WHITE }, line: { color: HEX.line, width: 1 },
            shadow: { type: "outer", color: "000000", opacity: 0.16, blur: 8, offset: 3, angle: 90 } });
        s.addImage({ path: path.join(ASSETS, "dashboard.png"), x: MX, y: 1.85, w: iw, h: ih,
            altText: "The HTML dashboard the bot writes: 48 validated, 40 passed, 1 warning, 7 failed, 1 escalated, 83.3% pass rate, audit trail verified" });
        txt(s, "dashboard.html: one self-contained file, no server, opens offline", MX, 1.95 + ih, iw, 0.35, { fontSize: 12, color: SLATE });
        card(s, MX, 5.55, iw, 1.15, { name: "Run history" });
        await badge(s, "FaClockRotateLeft", MX + 0.25, 5.8, 0.65, { label: "History" });
        txt(s, "Run history is read back out of the audit log, so the dashboard can never disagree with the evidence.",
            MX + 1.1, 5.62, iw - 1.3, 1.0, { fontSize: 14, color: INK, valign: "middle" });

        const rx = 8.35, rw = SW - MX - rx;
        const arts = [["FaFileExcel", "Excel report", "Colour-coded results and an exceptions sheet"],
                      ["FaFilePdf", "PDF audit report", "For the auditor, with a signature block"],
                      ["FaChartColumn", "HTML dashboard", "KPIs, phase timings and trends"],
                      ["FaBell", "Alerts", "E-mail, Teams card, Slack message"],
                      ["FaFingerprint", "Audit log", "Append-only and hash-chained"]];
        for (let i = 0; i < arts.length; i++) {
            const y = 1.85 + i * 0.97, [ic, t, d] = arts[i];
            await badge(s, ic, rx, y + 0.05, 0.62, { label: t });
            txt(s, t, rx + 0.85, y, rw - 0.85, 0.36, { fontSize: 16, bold: true });
            txt(s, d, rx + 0.85, y + 0.37, rw - 0.85, 0.5, { fontSize: 13, color: SLATE });
        }
        s.addNotes("Nitin leads. Open the newest Excel report and read one FailureReasons cell aloud - plain English on purpose, because the reader is a logistics manager. The PDF is what you hand an auditor: bot identity, configuration, every finding with the reviewer's decision, and the audit-chain head hash. The dashboard is one HTML file with no server and no internet.");
    }

    // ---- 13. Work queue -------------------------------------------------------------------------
    {
        const s = content(section("Evidence"), "SCALE  ·  WORK QUEUE", "Split per shipment, run entirely from Studio");
        const flow = [["FaPlay", "Main.xaml", "RunMode = DISPATCH"], ["FaListUl", "Work-queue file", "12 items, one per shipment"],
                      ["FaGears", "Performer.xaml", "its own entry point"], ["FaCircleCheck", "Same evidence", "48 / 40 / 1 / 7, chain VALID"]];
        const fw = 2.6, fg = (CW - 4 * fw) / 3;
        for (let i = 0; i < flow.length; i++) {
            const x = MX + i * (fw + fg), y = 1.9, [ic, t, d] = flow[i], last = i === flow.length - 1;
            card(s, x, y, fw, 1.85, { fill: last ? TEAL : MIST, name: t });
            s.addImage({ data: await icon(ic, last ? HEX.white : HEX.teal), x: x + 0.3, y: y + 0.28, w: 0.5, h: 0.5, altText: t });
            txt(s, t, x + 0.3, y + 0.92, fw - 0.6, 0.38, { fontSize: 17, bold: true, color: last ? WHITE : INK });
            txt(s, d, x + 0.3, y + 1.3, fw - 0.6, 0.45, { fontSize: 13, color: last ? WHITE : SLATE });
            if (i < flow.length - 1) arrow(s, x + fw + 0.08, y + 0.92, x + fw + fg - 0.08, y + 0.92, HEX.slate, 2.25);
        }
        const bw = (CW - 0.3) / 2;
        card(s, MX, 4.1, bw, 2.6, { fill: WHITE, line: HEX.line, name: "Why a shipment" });
        await badge(s, "FaTruckFast", MX + 0.3, 4.4, 0.7, { label: "Shipment" });
        txt(s, "Why a shipment, not a transaction", MX + 1.2, 4.42, bw - 1.45, 0.65, { fontSize: 17, bold: true, valign: "middle" });
        txt(s, "Three of the six rules compare events against each other: duplicates, ordering and required events. A shipment is the smallest unit that keeps every rule meaningful.",
            MX + 0.3, 5.25, bw - 0.6, 1.35, { fontSize: 14, color: SLATE });
        const x2 = MX + bw + 0.3;
        card(s, x2, 4.1, bw, 2.6, { fill: WHITE, line: HEX.line, name: "Studio only" });
        await badge(s, "FaLaptopCode", x2 + 0.3, 4.4, 0.7, { label: "Studio" });
        txt(s, "UiPath Studio only - no Orchestrator", x2 + 1.2, 4.42, bw - 1.45, 0.65, { fontSize: 17, bold: true, valign: "middle" });
        txt(s, "The queue is a local file. Moving to Orchestrator would swap it for Add Queue Item and Get Transaction Item; the validation would not change.",
            x2 + 0.3, 5.25, bw - 0.6, 1.35, { fontSize: 14, color: SLATE });
        s.addNotes("Adithya - drop if short on time. Set RunMode = DISPATCH, run Main: it writes one work item per shipment. Open Workflows/Queue/Performer.xaml and Run File: it processes them and produces the same audit trail and reports as a batch run. Say plainly that the queue is a local file and the project runs in Studio only - our Orchestrator access ended before the final review. Set RunMode back to BATCH.");
    }

    // ---- 14. Verification ----------------------------------------------------------------------
    {
        const s = content(section("Evidence"), "VERIFICATION", "One command verifies every claim in this deck");
        const tiles = [[CHECKS, "acceptance checks passing"], [TESTS, "UiPath test cases passing"], ["0", "Workflow Analyzer errors"], ["58/58", "Invoke Code bodies compile"]];
        const tw = (CW - 3 * 0.3) / 4;
        tiles.forEach(([n, l], i) => {
            const x = MX + i * (tw + 0.3);
            card(s, x, 1.8, tw, 1.45, { name: l });
            txt(s, n, x + 0.3, 1.9, tw - 0.6, 0.8, { fontSize: 40, bold: true, color: GREEN });
            txt(s, l, x + 0.3, 2.7, tw - 0.6, 0.4, { fontSize: 14, color: SLATE });
        });
        const lw = 7.5;
        txt(s, "The acceptance suite covers", MX, 3.5, lw, 0.4, { fontSize: 16, bold: true });
        const items = ["Seeded anomalies and the ABI decoding", "All six rules, exact counts", "A rule switched off in Excel",
                       "Three feeds, two schemas, one result", "Tampering found at the exact line", "WEB, API and CSV / SQL ERP sources",
                       "Excel, PDF, dashboard and alerts", "Dispatch and Performer", "Review window answered by UI Automation", "A failed run is logged and screenshotted"];
        items.forEach((t, i) => {
            const x = MX + (i % 2) * 3.75, y = 3.98 + Math.floor(i / 2) * 0.42;
            txt(s, [{ text: "✓  ", options: { color: GREEN, bold: true } }, { text: t }], x, y, 3.7, 0.4, { fontSize: 13.5, color: INK });
        });
        const rx = MX + lw + 0.4, rw = SW - MX - rx;
        card(s, rx, 3.5, rw, 2.6, { fill: "E6F4EC", name: "A test found a bug" });
        txt(s, "A test found a real bug", rx + 0.3, 3.68, rw - 0.6, 0.4, { fontSize: 16, bold: true, color: GREEN });
        txt(s, "Writing the duplicate-detection test showed R4 assumed a hash was at least 12 characters, and crashed on a malformed one. Found, fixed and now covered.",
            rx + 0.3, 4.12, rw - 0.6, 1.85, { fontSize: 14, color: INK });
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 6.25, w: CW, h: 0.5, rectRadius: 0.08, fill: { color: MIST }, line: { type: "none" } });
        txt(s, "powershell -ExecutionPolicy Bypass -File tools\\run_all_checks.ps1", MX + 0.3, 6.25, CW - 0.6, 0.5,
            { fontSize: 14, fontFace: MONO, bold: true, color: TEAL, valign: "middle" });
        s.addNotes("Nitin leads. Every claim in the deck is reproducible with one command - run it the night before, not during the review: it takes about fifteen minutes and answers on-screen windows itself near the end. Volunteer the bug story even if not asked; it is what actually happens when you test properly.");
    }

    // ---- 15. Status ----------------------------------------------------------------------------
    {
        const s = content(section("Close"), "STATED PLAINLY", "What works, and what is out of reach");
        const lw = 6.3;
        card(s, MX, 1.8, lw, 4.9, { name: "Verified" });
        pill(s, "EXECUTED AND VERIFIED", MX + 0.3, 2.05, 2.7, 0.4, GREEN, WHITE, 12);
        const done = ["Full pipeline, 48 transactions in about 24 s", "All six rules, against an independent oracle", "Three feeds across two schemas, one result",
                      "DECODE and LOOKUP shipment linking", "Excel, CSV and SQL ERP sources", "Rules switched on and off from Excel",
                      "Review window, answered via UI Automation", "Tamper detected at the exact line", "Excel, PDF, dashboard, Teams and Slack",
                      "API mode over HTTP to a local stand-in"];
        done.forEach((t, i) => {
            txt(s, [{ text: "✓  ", options: { color: GREEN, bold: true } }, { text: t }], MX + 0.3, 2.65 + i * 0.395, lw - 0.6, 0.38, { fontSize: 14, color: INK });
        });
        const rx = MX + lw + 0.3, rw = SW - MX - rx;
        card(s, rx, 1.8, rw, 4.9, { fill: WHITE, line: HEX.line, name: "Not executed" });
        pill(s, "NOT EXECUTED", rx + 0.3, 2.05, 1.9, 0.4, AMBER, WHITE, 12);
        const gaps = [["Orchestrator: out of scope", "The project runs in UiPath Studio; the work queue is a local file."],
                      ["Live Etherscan key", "API mode is tested over HTTP against a local stand-in, not the live chain."],
                      ["Browser automation", "WEB mode parses the explorer's HTML rather than recording a scrape."],
                      ["Real mail and chat", "Delivery is tested against a local capture server, not a real inbox."]];
        gaps.forEach(([t, d], i) => {
            const y = 2.65 + i * 1.0;
            txt(s, t, rx + 0.3, y, rw - 0.6, 0.36, { fontSize: 16, bold: true, color: AMBER });
            txt(s, d, rx + 0.3, y + 0.38, rw - 0.6, 0.55, { fontSize: 13, color: SLATE });
        });
        s.addNotes("Nitin - do not skip. Answer plainly; waffling reads as not knowing. Orchestrator is out of scope because the project runs in Studio only. The other gaps are accounts or tools, not capability: a live API key, the browser extension, and real mail and chat endpoints - each path is exercised locally up to that final hop. Nothing is stubbed.");
    }

    // ---- 16. Close --------------------------------------------------------------------------------
    {
        const s = pres.addSlide({ masterName: "DARK", sectionTitle: section("Close") });
        txt(s, "FUTURE SCOPE", MX, 0.42, 6, 0.32, { fontSize: 12, bold: true, color: MINT, charSpacing: 2 });
        txt(s, "Where this goes next", MX, 0.72, CW, 0.78, { fontSize: 32, bold: true, color: WHITE, fontFace: THEME.headFontFace });
        const fut = [["FaCloudArrowUp", "Deploy to Orchestrator", "Schedule with time triggers and swap the work-queue file for an Orchestrator queue."],
                     ["FaBrain", "ML anomaly detection", "AI Center on historical transactions, to catch patterns no fixed rule anticipates."],
                     ["FaTemperatureHalf", "IoT data fusion", "Check on-chain records against temperature and GPS, so a cold-chain breach is caught too."],
                     ["FaBolt", "Event-driven validation", "Blockchain WebSocket subscriptions instead of polling, for near-instant checks."]];
        const w = (CW - 0.3) / 2, h = 1.45;
        for (let i = 0; i < fut.length; i++) {
            const x = MX + (i % 2) * (w + 0.3), y = 1.8 + Math.floor(i / 2) * (h + 0.3), [ic, t, d] = fut[i];
            s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: { color: HEX.deep2 }, line: { type: "none" } });
            await badge(s, ic, x + 0.3, y + 0.35, 0.75, { fill: MINT, fg: HEX.deep, label: t });
            txt(s, t, x + 1.3, y + 0.22, w - 1.55, 0.4, { fontSize: 18, bold: true, color: WHITE });
            txt(s, d, x + 1.3, y + 0.65, w - 1.55, 0.7, { fontSize: 14, color: "C9DCDD" });
        }
        txt(s, "Immutable is not the same as verified.", MX, 5.35, CW, 0.5, { fontSize: 26, bold: true, color: WHITE });
        txt(s, "This is the bot that does the verifying.", MX, 5.85, CW, 0.45, { fontSize: 20, italic: true, color: MINT });
        txt(s, "Thank you - questions welcome", MX, 6.65, 6, 0.4, { fontSize: 16, bold: true, color: WHITE });
        txt(s, "github.com/AdithyaSM31/UiPath-Blockchain-Transaction-Validation", 6.3, 6.68, SW - MX - 6.3, 0.35,
            { fontSize: 12, fontFace: MONO, color: "A9C3C8", align: "right" });
        s.addNotes("Either presenter. Return to the opening line: immutable is not the same as verified, and this is the bot that does the verifying. Then invite questions. Expect: is blockchain data already trustworthy; why RPA rather than a Python script; why hash the audit log; is the data real; where are the tests; what does not work; why not Orchestrator. All are answered in docs/FINAL_REVIEW_SPEAKING_NOTES.md.");
    }
}

// ---------------------------------------------------------------- theme colours
async function applyTheme(file) {
    const zip = await JSZip.loadAsync(fs.readFileSync(file));
    const part = "ppt/theme/theme1.xml";
    const slots = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
    const scheme = `<a:clrScheme name="${THEME.name}">` + slots.map((k) => `<a:${k}><a:srgbClr val="${THEME.colors[k]}"/></a:${k}>`).join("") + "</a:clrScheme>";
    let xml = await zip.file(part).async("string");
    xml = xml.replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme)
             .replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, h) => h + THEME.name + '"');
    if (!xml.includes(scheme)) throw new Error("theme1.xml has no colour scheme to replace");
    zip.file(part, xml);
    for (const name of Object.keys(zip.files)) {
        if (!name.endsWith(".xml")) continue;
        const bad = (await zip.file(name).async("string")).match(/<a:srgbClr val="((?![0-9A-Fa-f]{6}")[^"]*)"/);
        if (bad) throw new Error(`${name}: a scheme colour "${bad[1]}" was given to a hex-only option`);
    }
    fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

(async () => {
    await build();
    await pres.writeFile({ fileName: OUT });
    await applyTheme(OUT);
    console.log(`Wrote ${OUT}\n  date: ${DATE}  checks: ${CHECKS}  tests: ${TESTS}  slides: ${pres.slides.length}`);
})().catch((e) => { console.error(e); process.exit(1); });

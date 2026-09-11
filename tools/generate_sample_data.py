"""
generate_sample_data.py
=======================
Generates every input artefact the BlockchainLogisticsValidator bot consumes:

    Data/Config.xlsx                              the no-code control panel (6 sheets)
    Data/Input/LogisticsRecords.xlsx              ERP/WMS master + event lines
    Data/Input/TxShipmentMap.xlsx                 TxHash -> ShipmentID lookup (LOOKUP mode)
    Data/Input/MockChain/etherscan_txlist_response.json
    Data/Input/MockChain/etherscan_txlist_polygon.json
    docs/SEEDED_ANOMALIES.md                      traceability: which row breaks which rule

Design notes
------------
* Addresses are genuine EIP-55 checksummed values and method IDs are real keccak-256
  function selectors, so the mock feed is indistinguishable in shape from a live
  Etherscan V2 `txlist` response. The bot's MOCK and API readers share one parser.
* Transaction `input` is properly ABI-encoded, so the preprocessing workflow performs
  a real decode rather than reading a convenience field.
* Exactly one anomaly is seeded per validation rule (plus one WARNING-band case), so a
  single run exercises R1-R6 and all three outcome states.

Run:  python tools/generate_sample_data.py
"""

from __future__ import annotations

import json
import os
from datetime import datetime, timedelta, timezone

from Crypto.Hash import keccak
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Paths
# --------------------------------------------------------------------------
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROJ = os.path.join(ROOT, "BlockchainLogisticsValidator")
DATA = os.path.join(PROJ, "Data")
INPUT = os.path.join(DATA, "Input")
MOCK = os.path.join(INPUT, "MockChain")
DOCS = os.path.join(ROOT, "docs")

for d in (DATA, INPUT, MOCK, DOCS):
    os.makedirs(d, exist_ok=True)


# --------------------------------------------------------------------------
# Ethereum primitives
# --------------------------------------------------------------------------
def keccak256(data: bytes) -> bytes:
    k = keccak.new(digest_bits=256)
    k.update(data)
    return k.digest()


def to_checksum_address(addr_hex: str) -> str:
    """EIP-55: capitalise a hex digit when the matching keccak nibble is >= 8."""
    a = addr_hex.lower().replace("0x", "")
    h = keccak256(a.encode("ascii")).hex()
    return "0x" + "".join(c.upper() if c.isalpha() and int(h[i], 16) >= 8 else c
                          for i, c in enumerate(a))


def derive_address(seed: str) -> str:
    """Deterministic, valid-looking address: last 20 bytes of keccak(seed), checksummed."""
    return to_checksum_address(keccak256(seed.encode("utf-8"))[-20:].hex())


def selector(signature: str) -> str:
    """First 4 bytes of keccak(signature) - the real Solidity function selector."""
    return "0x" + keccak256(signature.encode("ascii"))[:4].hex()


def topic0(event_signature: str) -> str:
    """Full 32-byte keccak of an event signature - what appears in log topics[0]."""
    return "0x" + keccak256(event_signature.encode("ascii")).hex()


# --------------------------------------------------------------------------
# Minimal ABI encoder (static uint256 + dynamic string)
# --------------------------------------------------------------------------
def _word(n: int) -> bytes:
    return n.to_bytes(32, "big")


def _enc_string(s: str) -> bytes:
    b = s.encode("utf-8")
    pad = (-len(b)) % 32
    return _word(len(b)) + b + b"\x00" * pad


def abi_encode(signature: str, arg_types: list[str], args: list) -> str:
    """Encode a call: 4-byte selector + head (offsets/statics) + tail (dynamic data)."""
    head, tail = b"", b""
    head_len = 32 * len(args)
    for t, v in zip(arg_types, args):
        if t == "string":
            head += _word(head_len + len(tail))
            tail += _enc_string(v)
        elif t == "uint256":
            head += _word(int(v))
        else:
            raise ValueError(f"unsupported ABI type {t}")
    return selector(signature) + (head + tail).hex()


# --------------------------------------------------------------------------
# Domain model: a pharma cold-chain contract on Ethereum
# --------------------------------------------------------------------------
CONTRACT = derive_address("PharmaChainLogisticsRegistry.v1")

PARTNERS = [
    ("NordPharma Manufacturing GmbH", "Manufacturer"),
    ("Maersk Freight Forwarding",     "FreightForwarder"),
    ("DHL Customs Brokerage",         "CustomsBroker"),
    ("ColdChain Warehousing Ltd",     "Warehouse"),
    ("LastMile Logistics India",      "Carrier"),
    ("Apollo Hospital Receiving",     "Consignee"),
]
PARTNER_ADDR = {name: derive_address(name) for name, _ in PARTNERS}

# Not on the whitelist - used to seed the R3 anomaly.
ROGUE_ADDR = derive_address("unregistered-reseller-wallet")

# Each logistics milestone maps to a distinct contract function and a distinct event.
EVENTS = [
    # event name,      function signature,                                arg types,                        emitted event signature
    ("Dispatched",     "recordDispatch(string,uint256,string)",           ["string", "uint256", "string"],  "Dispatched(string,uint256,address)"),
    ("InTransit",      "recordInTransit(string,uint256,string)",          ["string", "uint256", "string"],  "InTransit(string,uint256,address)"),
    ("Delivered",      "recordDelivery(string,uint256,string)",           ["string", "uint256", "string"],  "Delivered(string,uint256,address)"),
    ("GoodsReceived",  "confirmGoodsReceived(string,uint256,string)",     ["string", "uint256", "string"],  "GoodsReceived(string,uint256,address)"),
]
EVENT_META = {
    name: {
        "function": sig,
        "types": types,
        "methodId": selector(sig),
        "topic0": topic0(evsig),
        "eventSignature": evsig,
    }
    for name, sig, types, evsig in EVENTS
}
EVENT_HANDLER = {
    "Dispatched":    "NordPharma Manufacturing GmbH",
    "InTransit":     "Maersk Freight Forwarding",
    "Delivered":     "LastMile Logistics India",
    "GoodsReceived": "Apollo Hospital Receiving",
}

SHIPMENTS = [
    # id,        PO,         product,                       qty,  uom,     origin,             destination
    ("SHP-1001", "PO-88001", "Insulin Glargine 100IU",       500, "vials", "Frankfurt, DE",    "Chennai, IN"),
    ("SHP-1002", "PO-88002", "Amoxicillin 500mg",           1200, "boxes", "Mumbai, IN",       "Kochi, IN"),
    ("SHP-1003", "PO-88003", "mRNA Vaccine (-70C)",          800, "vials", "Frankfurt, DE",    "New Delhi, IN"),
    ("SHP-1004", "PO-88004", "Paracetamol IV 100ml",         950, "units", "Hyderabad, IN",    "Guwahati, IN"),
    ("SHP-1005", "PO-88005", "Rabies Immunoglobulin",        500, "vials", "Lyon, FR",         "Bengaluru, IN"),
    ("SHP-1006", "PO-88006", "Insulin Pen Cartridges",       300, "packs", "Copenhagen, DK",   "Pune, IN"),
    ("SHP-1007", "PO-88007", "Oncology Infusion Kit",        120, "kits",  "Basel, CH",        "Mumbai, IN"),
    ("SHP-1008", "PO-88008", "Saline IV Bags 500ml",        2400, "bags",  "Chennai, IN",      "Madurai, IN"),
    ("SHP-1009", "PO-88009", "Hepatitis B Vaccine",          640, "vials", "Seoul, KR",        "Kolkata, IN"),
    ("SHP-1010", "PO-88010", "Polyvalent Anti-Venom",         75, "vials", "Ahmedabad, IN",    "Bhubaneswar, IN"),
    ("SHP-1011", "PO-88011", "Drug-Eluting Cardiac Stents",  220, "units", "Galway, IE",       "Chennai, IN"),
    ("SHP-1012", "PO-88012", "Fresh Frozen Plasma",          460, "units", "Singapore, SG",    "Chennai, IN"),
]

BASE_TIME = datetime(2026, 9, 1, 6, 0, 0, tzinfo=timezone.utc)
EVENT_OFFSETS_H = {"Dispatched": 0, "InTransit": 18, "Delivered": 54, "GoodsReceived": 58}
GENESIS_BLOCK = 21_450_000

# --------------------------------------------------------------------------
# Seeded anomalies - exactly one per rule, plus one WARNING-band case
# --------------------------------------------------------------------------
ANOMALIES = [
    ("R1", "TimestampCheck",     "WARNING", "SHP-1002", "InTransit",
     "On-chain timestamp is 2h30m after the ERP expected time - inside the 4h warning band but past the 2h tolerance."),
    ("R1", "TimestampCheck",     "FAIL",    "SHP-1003", "InTransit",
     "On-chain timestamp is 5h20m after the ERP expected time, beyond the 4h fail threshold."),
    ("R2", "QuantityMatch",      "FAIL",    "SHP-1005", "Dispatched",
     "On-chain quantity 480 vials against a purchase order for 500 - a 20-unit shortfall."),
    ("R3", "AddressWhitelist",   "FAIL",    "SHP-1007", "Delivered",
     "Delivery recorded by a wallet that is not on the approved-partner whitelist. CRITICAL, so it escalates to a human."),
    ("R4", "DuplicateDetection", "FAIL",    "SHP-1009", "Delivered",
     "The same transaction hash is recorded twice for one shipment - a replayed or double-submitted event."),
    ("R5", "SequenceValidation", "FAIL",    "SHP-1011", "Delivered",
     "Delivered is timestamped before InTransit, so the milestone order is impossible."),
    ("R6", "SmartContractEvent", "FAIL",    "SHP-1012", "GoodsReceived",
     "No GoodsReceived confirmation was ever emitted - the chain of custody is left open."),
]


def build_chain_and_erp():
    """Produce the blockchain transaction list and the matching ERP event lines."""
    txs, erp_events = [], []
    nonce = 0

    for idx, (sid, po, product, qty, uom, origin, dest) in enumerate(SHIPMENTS):
        ship_start = BASE_TIME + timedelta(hours=6 * idx)

        for ev_name, _sig, types, _evsig in EVENTS:
            expected_at = ship_start + timedelta(hours=EVENT_OFFSETS_H[ev_name])
            actual_at = expected_at
            onchain_qty = qty
            handler_name = EVENT_HANDLER[ev_name]
            sender = PARTNER_ADDR[handler_name]

            # --- R1: timestamp drift ----------------------------------------
            if sid == "SHP-1002" and ev_name == "InTransit":
                actual_at = expected_at + timedelta(hours=2, minutes=30)
            if sid == "SHP-1003" and ev_name == "InTransit":
                actual_at = expected_at + timedelta(hours=5, minutes=20)

            # --- R2: quantity shortfall -------------------------------------
            if sid == "SHP-1005" and ev_name == "Dispatched":
                onchain_qty = 480

            # --- R3: unapproved wallet --------------------------------------
            if sid == "SHP-1007" and ev_name == "Delivered":
                sender = ROGUE_ADDR

            # --- R5: Delivered lands before InTransit -----------------------
            # The ERP expectation is moved to match, so the two systems agree on the
            # timestamp and R1 stays clean. Only the *ordering* is impossible. This is
            # what a consistently-falsified record looks like, and it is the reason
            # sequence logic is needed on top of pairwise field comparison.
            if sid == "SHP-1011" and ev_name == "Delivered":
                actual_at = ship_start + timedelta(hours=EVENT_OFFSETS_H["InTransit"] - 4)
                expected_at = actual_at

            # The ERP always expects every milestone. R6 works precisely because the
            # blockchain side is missing one that the ERP side still demands.
            emit_onchain = not (sid == "SHP-1012" and ev_name == "GoodsReceived")

            if emit_onchain:
                epoch = int(actual_at.timestamp())
                tx_seed = f"{sid}|{ev_name}|{epoch}|{onchain_qty}"
                tx_hash = "0x" + keccak256(tx_seed.encode()).hex()
                input_data = abi_encode(EVENT_META[ev_name]["function"], types, [sid, onchain_qty, po])

                txs.append({
                    "shipmentId": sid, "eventName": ev_name, "epoch": epoch,
                    "hash": tx_hash, "from": sender, "input": input_data,
                    "nonce": nonce, "qty": onchain_qty, "po": po,
                })
                nonce += 1

            erp_events.append({
                "ShipmentID": sid,
                "EventName": ev_name,
                "ExpectedEventTimestampUtc": expected_at.replace(tzinfo=None),
                "HandlerPartner": handler_name,
                "HandlerAddress": PARTNER_ADDR[handler_name].lower(),  # ERP stores lowercase
                "ExpectedQty": qty,
                "RecordedBy": "SAP-WMS",
            })

            # --- R4: replay the Delivered transaction for SHP-1009 ----------
            if emit_onchain and sid == "SHP-1009" and ev_name == "Delivered":
                dup = dict(txs[-1])
                dup["nonce"] = nonce
                nonce += 1
                txs.append(dup)   # identical hash - that is the anomaly

    txs.sort(key=lambda t: (t["epoch"], t["nonce"]))
    return txs, erp_events


def to_etherscan_json(txs, chain_label="ethereum"):
    """Wrap the transactions in a genuine Etherscan V2 `account/txlist` envelope."""
    latest = max(t["epoch"] for t in txs)
    result = []
    for t in txs:
        block = GENESIS_BLOCK + (t["epoch"] - int(BASE_TIME.timestamp())) // 12
        meta = EVENT_META[t["eventName"]]
        result.append({
            "blockNumber": str(block),
            "timeStamp": str(t["epoch"]),
            "hash": t["hash"],
            "nonce": str(t["nonce"]),
            "blockHash": "0x" + keccak256(f"block{block}{chain_label}".encode()).hex(),
            "transactionIndex": str(t["nonce"] % 120),
            "from": t["from"],
            "to": CONTRACT,
            "value": "0",
            "gas": "220000",
            "gasPrice": "24000000000",
            "isError": "0",
            "txreceipt_status": "1",
            "input": t["input"],
            "contractAddress": "",
            "cumulativeGasUsed": str(1_200_000 + t["nonce"] * 37),
            "gasUsed": str(96_400 + (t["nonce"] * 13) % 4000),
            "confirmations": str((latest - t["epoch"]) // 12 + 64),
            "methodId": meta["methodId"],
            "functionName": meta["function"].replace(
                "(string,uint256,string)", "(string shipmentId, uint256 quantity, string poNumber)"),
        })
    return {"status": "1", "message": "OK", "result": result}


# --------------------------------------------------------------------------
# Excel helpers
# --------------------------------------------------------------------------
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(color="FFFFFF", bold=True, size=11)


def write_sheet(ws, headers, rows, widths=None, freeze="A2"):
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", horizontal="left")
    for r in rows:
        ws.append(r)
    for i, h in enumerate(headers, start=1):
        w = (widths[i - 1] if widths and i - 1 < len(widths)
             else min(max(len(str(h)) + 4, 14), 46))
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            if isinstance(cell.value, datetime):
                cell.number_format = "yyyy-mm-dd hh:mm:ss"


# --------------------------------------------------------------------------
# Config.xlsx
# --------------------------------------------------------------------------
def write_config(path):
    wb = Workbook()

    # ---- Settings --------------------------------------------------------
    ws = wb.active
    ws.title = "Settings"
    settings = [
        ("ProjectName",                 "BlockchainLogisticsValidator", "Shown in reports and the audit log."),
        ("BotIdentity",                 "BLV-BOT-01",                   "Recorded in every audit entry."),
        ("DataSourceMode",              "MOCK",                         "MOCK | API | WEB - selects the blockchain extraction path."),
        ("MockChainFile",               r"Data\Input\MockChain\etherscan_txlist_response.json", "Used when DataSourceMode = MOCK."),
        ("EtherscanBaseUrl",            "https://api.etherscan.io/v2/api", "V2 multichain endpoint. V1 was retired."),
        ("EtherscanApiKey",             "",                             "Free key from etherscan.io/myapikey. Blank keeps API mode disabled."),
        ("ChainId",                     1,                              "1 Ethereum, 137 Polygon, 56 BSC. Change this alone to switch chain."),
        ("ContractAddress",             CONTRACT,                       "Logistics registry contract whose transactions are audited."),
        ("ExplorerAddressUrl",          "https://etherscan.io/address/{0}", "Used when DataSourceMode = WEB."),
        ("MaxTransactions",             500,                            "Upper bound per run."),
        ("LogisticsFile",               r"Data\Input\LogisticsRecords.xlsx", "ERP / WMS export."),
        ("ShipmentsSheet",              "Shipments",                    "Shipment master sheet name."),
        ("ShipmentEventsSheet",         "ShipmentEvents",               "Expected milestone lines."),
        ("TxShipmentMapFile",           r"Data\Input\TxShipmentMap.xlsx", "Used when MappingMode = LOOKUP."),
        ("MappingMode",                 "DECODE",                       "DECODE reads the ABI input data; LOOKUP uses the mapping file."),
        ("TimestampToleranceHours",     2,                              "R1: drift up to this is a PASS."),
        ("TimestampWarnToleranceHours", 4,                              "R1: drift past this is a FAIL; between the two is a WARNING."),
        ("QuantityTolerance",           0,                              "R2: permitted absolute difference in units."),
        ("EscalationEnabled",           "True",                         "Novelty 2: show the review form on a CRITICAL anomaly."),
        ("EscalationSeverity",          "CRITICAL",                     "Minimum severity that pauses the bot for a human."),
        ("AttendedMode",                "True",                         "False in unattended runs: log and continue instead of blocking."),
        ("AlertMode",                   "EML",                          "EML writes the message to disk; SMTP and OUTLOOK actually send."),
        ("AlertFrom",                   "rpa.bot@logistics-demo.local", ""),
        ("AlertTo",                     "logistics.manager@logistics-demo.local;compliance@logistics-demo.local", "Semicolon separated."),
        ("SmtpHost",                    "smtp.gmail.com",               "Only used when AlertMode = SMTP."),
        ("SmtpPort",                    587,                            ""),
        ("SmtpUseSsl",                  "True",                         ""),
        ("SmtpUser",                    "",                             "Leave blank unless AlertMode = SMTP."),
        ("SmtpPassword",                "",                             "App password. Never commit a real value."),
        ("ReportsFolder",               r"Data\Output\Reports",         ""),
        ("AuditLogFolder",              r"Data\Output\AuditLogs",       ""),
        ("AuditLogFile",                "ValidationAuditLog.csv",       "Append-only, hash-chained."),
        ("ScreenshotFolder",            r"Data\Output\Screenshots",     ""),
        ("LogLevel",                    "Info",                         "Info | Trace."),
    ]
    write_sheet(ws, ["Name", "Value", "Notes"],
                [[n, v, d] for n, v, d in settings], widths=[30, 62, 74])

    # ---- ValidationRules -------------------------------------------------
    ws = wb.create_sheet("ValidationRules")
    rules = [
        ["R1", "TimestampCheck",     "TRUE", "HIGH",     2, 4,  "FLAG",
         "Absolute gap between the on-chain timestamp and the ERP expected time. Param1 = pass tolerance (h), Param2 = warn ceiling (h)."],
        ["R2", "QuantityMatch",      "TRUE", "HIGH",     0, "", "FLAG",
         "ABI-decoded on-chain quantity against the purchase-order quantity. Param1 = permitted difference in units."],
        ["R3", "AddressWhitelist",   "TRUE", "CRITICAL", "", "", "ESCALATE",
         "Sender address must appear on the ApprovedWallets sheet. Escalates to a human reviewer."],
        ["R4", "DuplicateDetection", "TRUE", "MEDIUM",   "", "", "FLAG",
         "The same transaction hash must not be recorded twice for one shipment."],
        ["R5", "SequenceValidation", "TRUE", "HIGH",     "", "", "FLAG",
         "Milestones must occur in the order given on the EventSequence sheet."],
        ["R6", "SmartContractEvent", "TRUE", "HIGH",     "GoodsReceived", "", "FLAG",
         "Param1 lists the event(s) every completed shipment must emit, matched by function selector."],
    ]
    write_sheet(ws, ["RuleID", "RuleName", "Enabled", "Severity", "Param1", "Param2", "FailAction", "Description"],
                rules, widths=[10, 22, 10, 12, 16, 10, 13, 96])

    # ---- ApprovedWallets -------------------------------------------------
    ws = wb.create_sheet("ApprovedWallets")
    wallets = [[PARTNER_ADDR[n], n, r, "TRUE"] for n, r in PARTNERS]
    write_sheet(ws, ["Address", "PartnerName", "Role", "Active"],
                wallets, widths=[46, 34, 20, 10])

    # ---- FieldMapping ----------------------------------------------------
    ws = wb.create_sheet("FieldMapping")
    mapping = [
        ["hash",              "TxHash",            "LOWERCASE",         "Transaction identity."],
        ["blockNumber",       "BlockNumber",       "INTEGER",           ""],
        ["timeStamp",         "EventTimestampUtc", "EPOCH_TO_DATETIME", "Unix seconds to DateTime."],
        ["from",              "SenderAddress",     "LOWERCASE",         "Reconciles checksummed and non-checksummed forms."],
        ["to",                "ContractAddress",   "LOWERCASE",         ""],
        ["methodId",          "EventSelector",     "LOWERCASE",         "Real keccak-256 function selector."],
        ["input:arg0",        "ShipmentID",        "ABI_STRING",        "First ABI argument."],
        ["input:arg1",        "OnChainQty",        "ABI_UINT256",       "Second ABI argument."],
        ["input:arg2",        "PONumber",          "ABI_STRING",        "Third ABI argument."],
        ["isError",           "TxFailed",          "BOOLEAN_INVERT",    "Etherscan reports 0 for success."],
    ]
    write_sheet(ws, ["BlockchainField", "LogisticsField", "TransformType", "Notes"],
                mapping, widths=[22, 22, 22, 62])

    # ---- EventSequence ---------------------------------------------------
    ws = wb.create_sheet("EventSequence")
    seq = [[i, name, "TRUE", EVENT_META[name]["methodId"], EVENT_HANDLER[name]]
           for i, (name, *_rest) in enumerate(EVENTS, start=1)]
    write_sheet(ws, ["StepOrder", "EventName", "Required", "MethodId", "ExpectedHandlerRole"],
                seq, widths=[12, 18, 12, 16, 34])

    # ---- EventSignatures -------------------------------------------------
    ws = wb.create_sheet("EventSignatures")
    sigs = [[name, EVENT_META[name]["function"], EVENT_META[name]["methodId"],
             EVENT_META[name]["eventSignature"], EVENT_META[name]["topic0"]]
            for name, *_r in EVENTS]
    write_sheet(ws, ["EventName", "FunctionSignature", "MethodId", "EventSignature", "EventTopic0"],
                sigs, widths=[18, 46, 14, 40, 70])

    wb.save(path)
    return path


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    txs, erp_events = build_chain_and_erp()

    # ---- Mock chain feeds ------------------------------------------------
    eth = os.path.join(MOCK, "etherscan_txlist_response.json")
    with open(eth, "w", encoding="utf-8") as f:
        json.dump(to_etherscan_json(txs), f, indent=2)

    # A second chain proves Novelty 3: same schema, different network.
    poly = os.path.join(MOCK, "etherscan_txlist_polygon.json")
    with open(poly, "w", encoding="utf-8") as f:
        json.dump(to_etherscan_json(txs, "polygon"), f, indent=2)

    # ---- LogisticsRecords.xlsx ------------------------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "Shipments"
    ship_rows = []
    for idx, (sid, po, product, qty, uom, origin, dest) in enumerate(SHIPMENTS):
        start = BASE_TIME + timedelta(hours=6 * idx)
        ship_rows.append([
            sid, po, product, qty, uom, origin, dest,
            "NordPharma Manufacturing GmbH",
            PARTNER_ADDR["Apollo Hospital Receiving"].lower(),
            start.replace(tzinfo=None),
            (start + timedelta(hours=EVENT_OFFSETS_H["GoodsReceived"])).replace(tzinfo=None),
            "LastMile Logistics India",
            "2-8C" if "Vaccine" in product or "Insulin" in product else "Ambient",
            "Open",
        ])
    write_sheet(ws, ["ShipmentID", "PONumber", "Product", "ExpectedQty", "UoM", "Origin",
                     "Destination", "SupplierName", "ReceiverAddress", "DispatchTimestampUtc",
                     "ExpectedDeliveryUtc", "Carrier", "TempRange", "Status"],
                ship_rows,
                widths=[13, 12, 30, 13, 9, 18, 18, 32, 46, 22, 22, 26, 12, 10])

    ws2 = wb.create_sheet("ShipmentEvents")
    write_sheet(ws2, ["ShipmentID", "EventName", "ExpectedEventTimestampUtc", "HandlerPartner",
                      "HandlerAddress", "ExpectedQty", "RecordedBy"],
                [[e["ShipmentID"], e["EventName"], e["ExpectedEventTimestampUtc"],
                  e["HandlerPartner"], e["HandlerAddress"], e["ExpectedQty"], e["RecordedBy"]]
                 for e in erp_events],
                widths=[13, 16, 26, 32, 46, 13, 13])
    logistics = os.path.join(INPUT, "LogisticsRecords.xlsx")
    wb.save(logistics)

    # ---- TxShipmentMap.xlsx ---------------------------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "TxMap"
    seen = set()
    map_rows = []
    for t in txs:
        if t["hash"] in seen:
            continue
        seen.add(t["hash"])
        map_rows.append([t["hash"], t["shipmentId"], t["eventName"], t["po"],
                         EVENT_META[t["eventName"]]["methodId"]])
    write_sheet(ws, ["TxHash", "ShipmentID", "EventName", "PONumber", "MethodId"],
                map_rows, widths=[70, 13, 16, 12, 14])
    txmap = os.path.join(INPUT, "TxShipmentMap.xlsx")
    wb.save(txmap)

    # ---- Config.xlsx -----------------------------------------------------
    config = write_config(os.path.join(DATA, "Config.xlsx"))

    # ---- Anomaly traceability -------------------------------------------
    lines = [
        "# Seeded Anomalies",
        "",
        "Generated by `tools/generate_sample_data.py`. Every validation rule has at least one",
        "deliberately broken record, so a single run exercises R1-R6 and all three outcome states.",
        "Regenerating the data reproduces this table exactly - the generator is deterministic.",
        "",
        "| Rule | Name | Expected | Shipment | Event | What is wrong |",
        "|---|---|---|---|---|---|",
    ]
    for rid, rname, expected, sid, ev, desc in ANOMALIES:
        lines.append(f"| {rid} | {rname} | **{expected}** | `{sid}` | {ev} | {desc} |")

    total_ship = len(SHIPMENTS)
    flagged = {a[3] for a in ANOMALIES}
    clean = [s[0] for s in SHIPMENTS if s[0] not in flagged]
    lines += [
        "",
        "## Why each anomaly is isolated",
        "",
        "Every seeded fault trips exactly one rule, so a run reads cleanly and each rule's",
        "contribution is visible on its own. `SHP-1011` needed care: moving *Delivered* earlier",
        "than *InTransit* would also have opened a large gap against the ERP timestamp and tripped",
        "R1 as a side effect. The ERP expectation was therefore moved to match the on-chain time,",
        "so both systems agree on *when* the delivery happened and only the *order* is impossible.",
        "",
        "That is not a convenience - it is the realistic case. A falsified record is usually",
        "reconciled across systems, which is exactly why pairwise field comparison (R1, R2) is not",
        "sufficient and cross-record sequence logic (R5) has to exist alongside it.",
        "",
        "## Expected run totals",
        "",
        f"- Transactions on the mock feed: **{len(txs)}** (includes the one replayed transaction)",
        f"- Shipments: **{total_ship}** - {len(flagged)} carrying a seeded anomaly, {len(clean)} clean",
        f"- Clean shipments: {', '.join(f'`{c}`' for c in clean)}",
        "",
        "## Reference values",
        "",
        f"- Registry contract: `{CONTRACT}`",
        f"- Unapproved wallet used for the R3 anomaly: `{ROGUE_ADDR}`",
        "",
        "| Event | Function signature | Selector (methodId) |",
        "|---|---|---|",
    ]
    for name, *_r in EVENTS:
        m = EVENT_META[name]
        lines.append(f"| {name} | `{m['function']}` | `{m['methodId']}` |")
    lines += [
        "",
        "Selectors are the real first four bytes of `keccak256(signature)`, and every address is a",
        "valid EIP-55 checksummed value, so the mock feed matches a live Etherscan response in shape",
        "and in content conventions.",
        "",
    ]
    anomaly_doc = os.path.join(DOCS, "SEEDED_ANOMALIES.md")
    with open(anomaly_doc, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    # ---- Report ----------------------------------------------------------
    print(f"transactions generated : {len(txs)}")
    print(f"erp event lines        : {len(erp_events)}")
    print(f"shipments              : {len(SHIPMENTS)}")
    print(f"contract               : {CONTRACT}")
    print(f"rogue wallet (R3)      : {ROGUE_ADDR}")
    for p in (eth, poly, logistics, txmap, config, anomaly_doc):
        print("wrote:", os.path.relpath(p, ROOT))


if __name__ == "__main__":
    main()

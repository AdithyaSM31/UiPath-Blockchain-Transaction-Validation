"""
build_tests.py
==============
Generates UiPath test cases under `BlockchainLogisticsValidator/Tests/` and registers them
in `project.json` so Studio's Test Explorer picks them up.

Each test drives one validation rule against a hand-built fixture table rather than the
sample data, so it is a genuine unit test: it isolates the rule, covers the edge cases the
sample data does not (case-insensitive addresses, the same hash on two different shipments,
a tolerance boundary), and fails for exactly one reason.

Two assertion mechanisms on purpose:

* `VerifyExpression` - what Studio's Test Explorer reports on. Note this activity is NOT
  reachable through the `ui` xmlns: UiPath.Testing.Activities declares an XmlnsPrefix but no
  XmlnsDefinition, so it must be referenced through an explicit clr-namespace.
* A final `Throw` when any expectation failed - so `UiRobot execute` returns a non-zero exit
  code and the headless acceptance suite can tell pass from fail.

Run:  python tools/build_tests.py
"""

from __future__ import annotations

import json
import os
import uuid
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from xamlgen import (  # noqa: E402
    try_catch, assign, if_, invoke_code, invoke_workflow, lit, log, sequence, throw,
    variables, vb, workflow,
)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROJ = os.path.join(ROOT, "BlockchainLogisticsValidator")
TESTS = os.path.join(PROJ, "Tests")

DICT_SO = "scg:Dictionary(x:String, x:Object)"
TEST_NS = {"ut": "clr-namespace:UiPath.Testing.Activities;assembly=UiPath.Testing.Activities"}
TEST_REFS = ["UiPath.Testing.Activities", "UiPath.Testing"]
TEST_IMPORTS = ["UiPath.Testing.Activities"]


def verify(expression: str, title: str) -> str:
    """One assertion, reported in Studio's Test Explorer."""
    return (f'<ut:VerifyExpression DisplayName="{lit("Verify - " + title)}" '
            f'Expression="{vb(expression)}" Result="{vb("assertOk")}" />'
            + record_assertion(title))


def record_assertion(title: str) -> str:
    """Accumulate the assertion outcome so the test can fail the job at the end."""
    return (assign("failures",
                   f'If(assertOk, failures, failures & "{title}; ")',
                   name=f"Assign - track '{title}'"))


def test_case(stem: str, display: str, steps: str,
              extra_vars: tuple[tuple[str, str], ...] = ()) -> str:
    body = sequence(
        display,
        log(f'"[TEST] {display} - starting"')
        + steps
        # VerifyExpression alone does not fail the process, so an explicit throw is what
        # makes the headless acceptance suite able to detect a regression.
        + if_('failures <> ""',
              sequence("Fail the test",
                       throw(f'New Exception("{display} FAILED: " & failures)')),
              sequence("Passed", log(f'"[TEST] {display} - all assertions passed"')),
              name="If - any assertion failed"),
        variables(("x:Boolean", "assertOk"), ("x:String", "failures"), *extra_vars))
    return workflow(stem, body, extra_namespaces=TEST_NS,
                    extra_refs=TEST_REFS, extra_imports=TEST_IMPORTS)


# --------------------------------------------------------------------------
# Fixture builders - only the columns each rule actually reads
# --------------------------------------------------------------------------
def fixture(columns: list[tuple[str, str]], rows_vb: str, rule: str) -> str:
    """Invoke Code that builds a results-shaped table with just what `rule` needs."""
    cols = "\n".join(f'dt.Columns.Add("{n}", GetType({t}))' for n, t in columns)
    return invoke_code(f'''
Dim dt As New DataTable("Fixture")
{cols}
dt.Columns.Add("{rule}_Status", GetType(String))
dt.Columns.Add("{rule}_Detail", GetType(String))

{rows_vb}

out_dt = dt
'''.strip(), [("Out", "out_dt", "sd:DataTable", "dtFixture")],
        name="Invoke Code - build fixture")


def invoke_rule(rule_file: str, p1: str = "", p2: str = "",
                extra: list[tuple[str, str, str, str]] | None = None) -> str:
    args = [
        ("InOut", "io_dtResults", "sd:DataTable", "dtFixture"),
        ("In", "in_Param1", "x:String", f'"{p1}"'),
        ("In", "in_Param2", "x:String", f'"{p2}"'),
    ] + list(extra or [])
    return invoke_workflow(f"Workflows\\Rules\\{rule_file}.xaml", args,
                           name=f"Invoke rule under test - {rule_file}")


def status(row: int, rule: str) -> str:
    return f'Convert.ToString(dtFixture.Rows({row})("{rule}_Status"))'


# --------------------------------------------------------------------------
# TC01 - R1 timestamp bands
# --------------------------------------------------------------------------
def tc01() -> str:
    rows = '''
' Four rows spanning every band, including the boundary cases the sample data
' does not exercise: exactly at tolerance, and a row with no ERP counterpart.
Dim baseAt As New DateTime(2026, 9, 1, 6, 0, 0, DateTimeKind.Utc)
dt.Rows.Add("in tolerance",      baseAt, baseAt.AddHours(1.0),  1.0)
dt.Rows.Add("exactly at limit",  baseAt, baseAt.AddHours(2.0),  2.0)
dt.Rows.Add("warning band",      baseAt, baseAt.AddHours(3.0),  3.0)
dt.Rows.Add("beyond warn limit", baseAt, baseAt.AddHours(6.0),  6.0)

Dim rw As DataRow = dt.NewRow()
rw("Label") = "no ERP milestone"
rw("ChainTimestampUtc") = baseAt
rw("ErpExpectedTimestampUtc") = DBNull.Value
rw("DriftHours") = DBNull.Value
dt.Rows.Add(rw)
'''.strip()

    steps = (
        fixture([("Label", "String"), ("ChainTimestampUtc", "DateTime"),
                 ("ErpExpectedTimestampUtc", "Object"), ("DriftHours", "Object")],
                rows, "R1")
        + invoke_rule("R1_TimestampCheck", "2", "4")
        + verify(f'{status(0, "R1")} = "PASS"', "1.0h drift passes")
        + verify(f'{status(1, "R1")} = "PASS"', "exactly 2.0h is still a pass")
        + verify(f'{status(2, "R1")} = "WARNING"', "3.0h drift warns")
        + verify(f'{status(3, "R1")} = "FAIL"', "6.0h drift fails")
        + verify(f'{status(4, "R1")} = "SKIPPED"', "no ERP milestone is skipped, not failed")
        + verify('Convert.ToString(dtFixture.Rows(3)("R1_Detail")).Contains("drift")',
                 "failure detail explains the drift"))
    return test_case("TC01_R1_TimestampCheck", "TC01 R1 timestamp bands", steps,
                     (("sd:DataTable", "dtFixture"),))


# --------------------------------------------------------------------------
# TC02 - R2 quantity, including tolerance
# --------------------------------------------------------------------------
def tc02() -> str:
    rows = '''
dt.Rows.Add("exact match", 500L, 500L, "PO-1")
dt.Rows.Add("short by 20", 480L, 500L, "PO-2")
dt.Rows.Add("over by 5",   505L, 500L, "PO-3")

Dim rw As DataRow = dt.NewRow()
rw("Label") = "no PO quantity"
rw("OnChainQty") = 500L
rw("ErpExpectedQty") = DBNull.Value
rw("PONumber") = "PO-4"
dt.Rows.Add(rw)
'''.strip()

    steps = (
        fixture([("Label", "String"), ("OnChainQty", "Int64"),
                 ("ErpExpectedQty", "Object"), ("PONumber", "String")], rows, "R2")
        + invoke_rule("R2_QuantityMatch", "0")
        + verify(f'{status(0, "R2")} = "PASS"', "exact match passes")
        + verify(f'{status(1, "R2")} = "FAIL"', "shortfall fails at zero tolerance")
        + verify(f'{status(2, "R2")} = "FAIL"', "overage also fails - the check is absolute")
        + verify(f'{status(3, "R2")} = "SKIPPED"', "missing PO quantity is skipped")
        # Re-run the same fixture with a tolerance of 25 to prove Param1 is honoured.
        + invoke_rule("R2_QuantityMatch", "25")
        + verify(f'{status(1, "R2")} = "PASS"', "tolerance 25 absorbs the 20-unit shortfall")
        + verify(f'{status(2, "R2")} = "PASS"', "tolerance 25 absorbs the 5-unit overage"))
    return test_case("TC02_R2_QuantityMatch", "TC02 R2 quantity and tolerance", steps,
                     (("sd:DataTable", "dtFixture"),))


# --------------------------------------------------------------------------
# TC03 - R3 whitelist, including checksum-case handling
# --------------------------------------------------------------------------
def tc03() -> str:
    rows = '''
' The third row is the point of this test: the same wallet in EIP-55 checksummed
' form. Extraction lower-cases addresses, but the rule must not depend on that
' having happened - a mixed-case address has to match an approved lower-case one.
dt.Rows.Add("approved partner",  "0xaabbccddeeff00112233445566778899aabbccdd", "NordPharma")
dt.Rows.Add("unknown wallet",    "0x9999999999999999999999999999999999999999", "NordPharma")
dt.Rows.Add("approved, mixed case", "0xAABBCCDDEEFF00112233445566778899AABBCCDD", "NordPharma")
'''.strip()

    wallets = invoke_code('''
Dim w As New DataTable("ApprovedWallets")
w.Columns.Add("Address", GetType(String))
w.Columns.Add("PartnerName", GetType(String))
w.Columns.Add("Role", GetType(String))
w.Columns.Add("Active", GetType(String))
w.Rows.Add("0xaabbccddeeff00112233445566778899aabbccdd", "NordPharma", "Manufacturer", "TRUE")
w.Rows.Add("0x1111111111111111111111111111111111111111", "Retired Partner", "Carrier", "FALSE")
out_dt = w
'''.strip(), [("Out", "out_dt", "sd:DataTable", "dtWallets")],
        name="Invoke Code - build wallet whitelist")

    steps = (
        fixture([("Label", "String"), ("SenderAddress", "String"),
                 ("ErpHandlerPartner", "String")], rows, "R3")
        + wallets
        + invoke_rule("R3_AddressWhitelist",
                      extra=[("In", "in_dtWallets", "sd:DataTable", "dtWallets")])
        + verify(f'{status(0, "R3")} = "PASS"', "approved wallet passes")
        + verify(f'{status(1, "R3")} = "FAIL"', "unknown wallet fails")
        + verify(f'{status(2, "R3")} = "PASS"', "checksummed address matches its lower-case entry")
        + verify('Convert.ToString(dtFixture.Rows(0)("R3_Detail")).Contains("NordPharma")',
                 "passing detail names the partner"))
    return test_case("TC03_R3_AddressWhitelist", "TC03 R3 wallet whitelist", steps,
                     (("sd:DataTable", "dtFixture"), ("sd:DataTable", "dtWallets")))


# --------------------------------------------------------------------------
# TC04 - R4 duplicates, scoped per shipment
# --------------------------------------------------------------------------
def tc04() -> str:
    rows = '''
' Rows 0 and 1 are the duplicate pair. Row 3 is the important negative case: the
' SAME hash against a DIFFERENT shipment must not be flagged, because the rule is
' scoped per shipment, not globally.
Dim dupHash As String = "0x" & New String("a"c, 64)
Dim uniqueHash As String = "0x" & New String("b"c, 64)
dt.Rows.Add("SHP-A", dupHash)
dt.Rows.Add("SHP-A", dupHash)
dt.Rows.Add("SHP-A", uniqueHash)
dt.Rows.Add("SHP-B", dupHash)

' A short, malformed hash must be reported rather than crash the rule.
dt.Rows.Add("SHP-C", "0xshort")
'''.strip()

    steps = (
        fixture([("ShipmentID", "String"), ("TxHash", "String")], rows, "R4")
        + invoke_rule("R4_DuplicateDetection")
        + verify(f'{status(0, "R4")} = "FAIL"', "first copy of a duplicate is flagged")
        + verify(f'{status(1, "R4")} = "FAIL"', "second copy is flagged too")
        + verify(f'{status(2, "R4")} = "PASS"', "a unique hash passes")
        + verify(f'{status(3, "R4")} = "PASS"',
                 "same hash on another shipment is not a duplicate")
        + verify(f'{status(4, "R4")} = "PASS"',
                 "a malformed short hash is handled without crashing"))
    return test_case("TC04_R4_DuplicateDetection", "TC04 R4 duplicate detection", steps,
                     (("sd:DataTable", "dtFixture"),))


# --------------------------------------------------------------------------
# TC05 - R5 milestone ordering
# --------------------------------------------------------------------------
def tc05() -> str:
    rows = '''
Dim t As New DateTime(2026, 9, 1, 6, 0, 0, DateTimeKind.Utc)

' SHP-A is out of order: Delivered (step 3) is recorded before InTransit (step 2),
' so InTransit is the event that could not physically have happened when it did.
dt.Rows.Add(1, "SHP-A", "Dispatched", 1, t)
dt.Rows.Add(2, "SHP-A", "Delivered",  3, t.AddHours(10))
dt.Rows.Add(3, "SHP-A", "InTransit",  2, t.AddHours(20))

' SHP-B is a clean control.
dt.Rows.Add(4, "SHP-B", "Dispatched", 1, t)
dt.Rows.Add(5, "SHP-B", "InTransit",  2, t.AddHours(10))
dt.Rows.Add(6, "SHP-B", "Delivered",  3, t.AddHours(20))
'''.strip()

    steps = (
        fixture([("RowNo", "Int32"), ("ShipmentID", "String"), ("EventName", "String"),
                 ("StepOrder", "Int32"), ("ChainTimestampUtc", "DateTime")], rows, "R5")
        + invoke_rule("R5_SequenceValidation")
        + verify(f'{status(0, "R5")} = "PASS"', "first milestone passes")
        + verify(f'{status(1, "R5")} = "PASS"', "a later step arriving early is not itself the fault")
        + verify(f'{status(2, "R5")} = "FAIL"', "the out-of-order event is flagged")
        + verify('Convert.ToString(dtFixture.Rows(2)("R5_Detail")).Contains("Delivered")',
                 "detail names the event it came after")
        + verify(f'{status(3, "R5")} = "PASS" AndAlso {status(4, "R5")} = "PASS" '
                 f'AndAlso {status(5, "R5")} = "PASS"',
                 "an in-order shipment is untouched"))
    return test_case("TC05_R5_SequenceValidation", "TC05 R5 milestone ordering", steps,
                     (("sd:DataTable", "dtFixture"),))


# --------------------------------------------------------------------------
# TC06 - R6 required smart contract event
# --------------------------------------------------------------------------
def tc06() -> str:
    rows = '''
Dim t As New DateTime(2026, 9, 1, 6, 0, 0, DateTimeKind.Utc)

' SHP-A completes the chain of custody.
dt.Rows.Add("SHP-A", "Dispatched",    t)
dt.Rows.Add("SHP-A", "GoodsReceived", t.AddHours(10))

' SHP-B never confirms receipt. The finding attaches to its LATEST event, which is
' the point at which the chain of custody is left open.
dt.Rows.Add("SHP-B", "Dispatched",    t)
dt.Rows.Add("SHP-B", "Delivered",     t.AddHours(10))
'''.strip()

    steps = (
        fixture([("ShipmentID", "String"), ("EventName", "String"),
                 ("ChainTimestampUtc", "DateTime")], rows, "R6")
        + invoke_rule("R6_SmartContractEvent", "GoodsReceived")
        + verify(f'{status(0, "R6")} = "PASS" AndAlso {status(1, "R6")} = "PASS"',
                 "shipment with the required event passes")
        + verify(f'{status(2, "R6")} = "PASS"', "only the latest event carries the finding")
        + verify(f'{status(3, "R6")} = "FAIL"', "missing event flagged on the latest row")
        + verify('Convert.ToString(dtFixture.Rows(3)("R6_Detail")).Contains("GoodsReceived")',
                 "detail names the missing event"))
    return test_case("TC06_R6_SmartContractEvent", "TC06 R6 required contract event", steps,
                     (("sd:DataTable", "dtFixture"),))


# --------------------------------------------------------------------------
# TC07 - audit chain is tamper-evident (integration)
# --------------------------------------------------------------------------
def tc07() -> str:
    # Writes to its own log file so the project's real audit trail is untouched.
    redirect = assign('Config("AuditLogFile")', '"TestAuditLog.csv"', type_ref="x:Object",
                      name="Assign - isolate the test audit log")

    results_fixture = invoke_code('''
Dim dt As New DataTable("Results")
For Each c As String In New String() {"TxHash", "ShipmentID", "EventName", _
        "ValidationStatus", "Severity", "FailureReasons", "HumanDecision", "HumanDecidedBy"}
    dt.Columns.Add(c, GetType(String))
Next
dt.Columns.Add("ChainTimestampUtc", GetType(DateTime))
dt.Columns.Add("OnChainQty", GetType(Long))

Dim t As New DateTime(2026, 9, 1, 6, 0, 0, DateTimeKind.Utc)
dt.Rows.Add("0xaaa", "SHP-T1", "Dispatched", "PASS", "NONE", "", "", "")
dt.Rows.Add("0xbbb", "SHP-T2", "Delivered", "FAIL", "HIGH", "R2 FAIL: short by 20", "", "")
For Each r As DataRow In dt.Rows
    r("ChainTimestampUtc") = t
    r("OnChainQty") = 100L
Next
out_dt = dt
'''.strip(), [("Out", "out_dt", "sd:DataTable", "dtResults")],
        name="Invoke Code - two-entry results fixture")

    tamper = invoke_code('''
' Flip the FAIL verdict to PASS - the edit somebody covering up an anomaly makes.
Dim lines() As String = File.ReadAllLines(in_Path)
For i As Integer = 1 To lines.Length - 1
    If lines(i).Contains(",FAIL,") Then
        lines(i) = lines(i).Replace(",FAIL,HIGH,", ",PASS,NONE,")
        Exit For
    End If
Next
File.WriteAllLines(in_Path, lines, Encoding.UTF8)
out_Done = True
'''.strip(), [("In", "in_Path", "x:String", "auditPath"),
              ("Out", "out_Done", "x:Boolean", "tampered")],
        name="Invoke Code - tamper with one verdict")

    cleanup = invoke_code('''
If File.Exists(in_Path) Then File.Delete(in_Path)
out_Done = True
'''.strip(), [("In", "in_Path", "x:String", "auditPath"),
              ("Out", "out_Done", "x:Boolean", "tampered")],
        name="Invoke Code - remove the test audit log")

    steps = (
        assign("projectRoot", "Directory.GetCurrentDirectory()")
        + assign("configPath", 'Path.Combine(projectRoot, "Data", "Config.xlsx")')
        + invoke_workflow("Workflows\\00_Init_ReadConfig.xaml", [
            ("In", "in_ConfigPath", "x:String", "configPath"),
            ("In", "in_ProjectRoot", "x:String", "projectRoot"),
            ("Out", "out_Config", DICT_SO, "Config"),
        ], name="Invoke 00 - read configuration")
        + redirect
        + results_fixture
        + invoke_workflow("Workflows\\10_AuditLog_Chained.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("In", "in_dtResults", "sd:DataTable", "dtResults"),
            ("Out", "out_AuditPath", "x:String", "auditPath"),
            ("Out", "out_HeadHash", "x:String", "headHash"),
        ], name="Invoke 10 - write the audit chain")
        + invoke_workflow("Workflows\\12_Verify_AuditChain.xaml", [
            ("In", "in_AuditPath", "x:String", "auditPath"),
            ("Out", "out_IsValid", "x:Boolean", "chainValid"),
            ("Out", "out_Report", "x:String", "chainReport"),
        ], name="Invoke 12 - verify (expect VALID)")
        + verify("chainValid", "a freshly written chain verifies")
        + verify('headHash.Length = 64', "chain head is a SHA-256 hash")
        + tamper
        + invoke_workflow("Workflows\\12_Verify_AuditChain.xaml", [
            ("In", "in_AuditPath", "x:String", "auditPath"),
            ("Out", "out_IsValid", "x:Boolean", "chainValid"),
            ("Out", "out_FirstBadRow", "x:Int32", "firstBadRow"),
            ("Out", "out_Report", "x:String", "chainReport"),
        ], name="Invoke 12 - verify (expect INVALID)")
        + verify("Not chainValid", "an edited verdict is detected")
        + verify("firstBadRow > 1", "the failing line is identified")
        + verify('chainReport.Contains("RecordHash mismatch")',
                 "the report says the contents were altered")
        + cleanup)

    return test_case("TC07_AuditChain_TamperEvident", "TC07 audit chain tamper evidence",
                     steps,
                     (("x:String", "projectRoot"), ("x:String", "configPath"),
                      (DICT_SO, "Config"), ("sd:DataTable", "dtResults"),
                      ("x:String", "auditPath"), ("x:String", "headHash"),
                      ("x:Boolean", "chainValid"), ("x:Int32", "firstBadRow"),
                      ("x:String", "chainReport"), ("x:Boolean", "tampered")))



# --------------------------------------------------------------------------
# Shared fixtures for the normaliser tests
# --------------------------------------------------------------------------
SAMPLE_CALLDATA = "0x54201e35000000000000000000000000000000000000000000000000000000000000006000000000000000000000000000000000000000000000000000000000000001f400000000000000000000000000000000000000000000000000000000000000a000000000000000000000000000000000000000000000000000000000000000085348502d313030310000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000008504f2d3838303031000000000000000000000000000000000000000000000000"

# Two profiles with field names that appear NOWHERE in the shipped Config.xlsx, so a
# pass proves the engine knows nothing about any particular source.
NORMALISER_TABLES = r"""
Dim m As New DataTable("FieldMapping")
For Each c As String In New String() {"Profile", "BlockchainField", "LogisticsField", "TransformType"}
    m.Columns.Add(c, GetType(String))
Next
' NESTED: nested paths, hex numbers, ISO time, a JSON boolean, selector from call data.
m.Rows.Add("NESTED", "id", "TxHash", "LOWERCASE")
m.Rows.Add("NESTED", "blk.n", "BlockNumber", "HEX_TO_INTEGER")
m.Rows.Add("NESTED", "at", "EventTimestampUtc", "ISO_TO_DATETIME")
m.Rows.Add("NESTED", "by", "SenderAddress", "LOWERCASE")
m.Rows.Add("NESTED", "data", "EventSelector", "SELECTOR")
m.Rows.Add("NESTED", "data:arg0", "ShipmentID", "ABI_STRING")
m.Rows.Add("NESTED", "data:arg1", "OnChainQty", "ABI_UINT256")
m.Rows.Add("NESTED", "data:arg2", "PONumber", "ABI_STRING")
m.Rows.Add("NESTED", "failed", "TxFailed", "BOOLEAN")
' FLAT: decimal numbers, epoch seconds, a selector field, a 0/1 flag.
m.Rows.Add("FLAT", "h", "TxHash", "LOWERCASE")
m.Rows.Add("FLAT", "b", "BlockNumber", "INTEGER")
m.Rows.Add("FLAT", "t", "EventTimestampUtc", "EPOCH_TO_DATETIME")
m.Rows.Add("FLAT", "f", "SenderAddress", "LOWERCASE")
m.Rows.Add("FLAT", "sel", "EventSelector", "LOWERCASE")
m.Rows.Add("FLAT", "cd:arg0", "ShipmentID", "ABI_STRING")
m.Rows.Add("FLAT", "cd:arg1", "OnChainQty", "ABI_UINT256")
m.Rows.Add("FLAT", "err", "TxFailed", "BOOLEAN")

Dim p As New DataTable("ChainProfiles")
For Each c As String In New String() {"Profile", "RecordsPath", "StatusField", "StatusOkValue"}
    p.Columns.Add(c, GetType(String))
Next
p.Rows.Add("NESTED", "export.items", "", "")
p.Rows.Add("FLAT", "rows", "ok", "yes")

Dim g As New DataTable("EventSignatures")
g.Columns.Add("EventName", GetType(String))
g.Columns.Add("MethodId", GetType(String))
g.Rows.Add("Dispatched", "0x54201e35")
g.Rows.Add("Delivered", "0xae7b1f50")

out_dtMapping = m
out_dtProfiles = p
out_dtSignatures = g
""".strip()

TABLE_OUTS = [
    ("Out", "out_dtMapping", "sd:DataTable", "dtMapping"),
    ("Out", "out_dtProfiles", "sd:DataTable", "dtProfiles"),
    ("Out", "out_dtSignatures", "sd:DataTable", "dtSignatures"),
]
TABLE_VARS = (("sd:DataTable", "dtMapping"), ("sd:DataTable", "dtProfiles"),
              ("sd:DataTable", "dtSignatures"), ("sd:DataTable", "dtChain"),
              ("x:String", "json"), ("x:String", "summary"), ("x:String", "errorText"))


def normalise(profile: str) -> str:
    return invoke_workflow("Workflows\\01m_Normalise_ChainRecords.xaml", [
        ("In", "in_JsonText", "x:String", "json"),
        ("In", "in_Profile", "x:String", f'"{profile}"'),
        ("In", "in_SourceMode", "x:String", '"TEST"'),
        ("In", "in_LookupMode", "x:Boolean", "False"),
        ("In", "in_dtMapping", "sd:DataTable", "dtMapping"),
        ("In", "in_dtSignatures", "sd:DataTable", "dtSignatures"),
        ("In", "in_dtProfiles", "sd:DataTable", "dtProfiles"),
        ("Out", "out_dtChain", "sd:DataTable", "dtChain"),
        ("Out", "out_Summary", "x:String", "summary"),
    ], name=f"Invoke 01m - profile {profile}")


def cell(col: str, row: int = 0) -> str:
    return f'dtChain.Rows({row})("{col}")'


# --------------------------------------------------------------------------
# TC08 - every conversion the field-mapping engine offers
# --------------------------------------------------------------------------
def tc08() -> str:
    build_nested = invoke_code("""
' Mixed-case hash and sender, so LOWERCASE is observable.
Dim rec As New JObject()
rec("id") = "0x031CB4C9F4B810639F883327F2156740285E345C8B2D09A4088E19161D8E6CB1"
rec("blk") = New JObject(New JProperty("n", "0x1474d10"))
rec("at") = "2026-09-01T06:00:00Z"
rec("by") = "0x85EBC61D67700CCa8b0C35f9fF9F92d62A93eb68"
rec("data") = "{{SAMPLE_CALLDATA}}"
rec("failed") = False
out_Json = New JObject(New JProperty("export", New JObject(New JProperty("items", New JArray(rec))))).ToString()
""".strip().replace("{{SAMPLE_CALLDATA}}", SAMPLE_CALLDATA), [("Out", "out_Json", "x:String", "json")],
        name="Invoke Code - one nested record")

    build_flat = invoke_code("""
Dim rec As New JObject()
rec("h") = "0xABC"
rec("b") = "21450000"
rec("t") = "1788242400"
rec("f") = "0xFFFF"
rec("sel") = "0x54201E35"
rec("cd") = "{{SAMPLE_CALLDATA}}"
rec("err") = "1"
out_Json = New JObject(New JProperty("ok", "yes"), New JProperty("rows", New JArray(rec))).ToString()
""".strip().replace("{{SAMPLE_CALLDATA}}", SAMPLE_CALLDATA), [("Out", "out_Json", "x:String", "json")],
        name="Invoke Code - one flat record")

    expected_time = 'New DateTime(2026, 9, 1, 6, 0, 0, DateTimeKind.Utc)'
    steps = (
        invoke_code(NORMALISER_TABLES, TABLE_OUTS, name="Invoke Code - mapping, profile and signature tables")
        + build_nested + normalise("NESTED")
        + verify("dtChain.Rows.Count = 1", "nested record decodes")
        + verify(f'Convert.ToString({cell("TxHash")}) = "0x031cb4c9f4b810639f883327f2156740285e345c8b2d09a4088e19161d8e6cb1"',
                 "LOWERCASE on the hash")
        + verify(f'Convert.ToInt64({cell("BlockNumber")}) = 21450000', "HEX_TO_INTEGER from a nested path")
        + verify(f'Convert.ToDateTime({cell("EventTimestampUtc")}) = {expected_time}', "ISO_TO_DATETIME")
        + verify(f'Convert.ToString({cell("EventSelector")}) = "0x54201e35"', "SELECTOR derived from call data")
        + verify(f'Convert.ToString({cell("ShipmentID")}) = "SHP-1001"', "ABI_STRING argument 0")
        + verify(f'Convert.ToInt64({cell("OnChainQty")}) = 500', "ABI_UINT256 argument 1")
        + verify(f'Convert.ToString({cell("PONumber")}) = "PO-88001"', "ABI_STRING argument 2")
        + verify(f'Convert.ToBoolean({cell("TxFailed")}) = False', "BOOLEAN from a JSON false")
        + build_flat + normalise("FLAT")
        + verify(f'Convert.ToInt64({cell("BlockNumber")}) = 21450000', "INTEGER from a decimal string")
        + verify(f'Convert.ToDateTime({cell("EventTimestampUtc")}) = {expected_time}', "EPOCH_TO_DATETIME")
        + verify(f'Convert.ToString({cell("EventSelector")}) = "0x54201e35"', "LOWERCASE on a selector field")
        + verify(f'Convert.ToBoolean({cell("TxFailed")}) = True', "BOOLEAN from the string 1")
        + verify(f'IsDBNull({cell("PONumber")})', "an unmapped field is null, not invented"))
    return test_case("TC08_FieldMapping_Engine", "TC08 field-mapping engine", steps, TABLE_VARS)


# --------------------------------------------------------------------------
# TC09 - records the engine must refuse, and configuration it must reject
# --------------------------------------------------------------------------
def tc09() -> str:
    build = invoke_code("""
Dim good As New JObject()
good("h") = "0x01" : good("b") = "1" : good("t") = "1788242400" : good("f") = "0xaa"
good("sel") = "0x54201e35" : good("cd") = "{{SAMPLE_CALLDATA}}" : good("err") = "0"

Dim transfer As JObject = CType(good.DeepClone(), JObject)
transfer("h") = "0x02" : transfer("cd") = "0x"

' approve(address,uint256) - a real selector, but not one of this contract's functions.
Dim unrelated As JObject = CType(good.DeepClone(), JObject)
unrelated("h") = "0x03" : unrelated("cd") = "0x095ea7b3" & New String("0"c, 128)

' A known selector whose call data stops after one word.
Dim truncated As JObject = CType(good.DeepClone(), JObject)
truncated("h") = "0x04" : truncated("cd") = "0xae7b1f50" & New String("0"c, 62) & "60"

Dim arr As New JArray(good, transfer, unrelated, truncated, New JValue(42))
out_Json = New JObject(New JProperty("ok", "yes"), New JProperty("rows", arr)).ToString()
""".strip().replace("{{SAMPLE_CALLDATA}}", SAMPLE_CALLDATA), [("Out", "out_Json", "x:String", "json")],
        name="Invoke Code - one good record and four bad ones")

    def expect_config_error(label: str, mutate_vb: str, must_contain: str) -> str:
        mutate = invoke_code(mutate_vb, [("InOut", "io_dtMapping", "sd:DataTable", "dtMapping")],
                             name=f"Invoke Code - {label}")
        attempt = try_catch(
            sequence("Expect a configuration error",
                     normalise("FLAT") + assign("errorText", '"(no error raised)"')),
            assign("errorText", "exception.ToString()"),
            name=f"Try Catch - {label}")
        return (invoke_code(NORMALISER_TABLES, TABLE_OUTS, name="Invoke Code - fresh tables")
                + mutate + attempt
                + verify(f'errorText.Contains("{must_contain}")', f"rejected: {label}"))

    steps = (
        invoke_code(NORMALISER_TABLES, TABLE_OUTS, name="Invoke Code - mapping, profile and signature tables")
        + build + normalise("FLAT")
        + verify("dtChain.Rows.Count = 1", "only the well-formed logistics call is kept")
        + verify(f'Convert.ToString({cell("TxHash")}) = "0x01"', "and it is the right one")
        + verify('summary.Contains("5 record(s)")', "every record was looked at")
        + verify('summary.Contains("1 value transfer(s) ignored")', "a value transfer is ignored")
        + verify('summary.Contains("1 call(s) to other functions ignored")', "an unrelated call is ignored")
        + verify('summary.Contains("2 malformed and skipped")', "truncated call data and a non-object are skipped")
        + expect_config_error("an unknown transform",
                              'For Each r As DataRow In io_dtMapping.Rows\n'
                              '    If Convert.ToString(r("BlockchainField")) = "b" Then r("TransformType") = "BASE64"\n'
                              'Next',
                              "unknown TransformType")
        + expect_config_error("a missing required field",
                              'For Each r As DataRow In io_dtMapping.Select("BlockchainField = \'h\'")\n'
                              '    io_dtMapping.Rows.Remove(r)\n'
                              'Next',
                              "has no row for: TxHash"))
    return test_case("TC09_Normaliser_Hardening", "TC09 normaliser hardening", steps, TABLE_VARS)


# --------------------------------------------------------------------------
# TC10 - MappingMode LOOKUP overrides what the call data says
# --------------------------------------------------------------------------
def tc10() -> str:
    tables = invoke_code("""
Dim t0 As New DateTime(2026, 9, 1, 6, 0, 0, DateTimeKind.Utc)

' The chain says SHP-WRONG / Delivered. The logistics system's map disagrees.
Dim ch As New DataTable("Chain")
For Each c As String In New String() {"TxHash", "EventSelector", "ShipmentID", "PONumber", "SenderAddress"}
    ch.Columns.Add(c, GetType(String))
Next
ch.Columns.Add("BlockNumber", GetType(Long))
ch.Columns.Add("EventTimestampUtc", GetType(DateTime))
ch.Columns.Add("OnChainQty", GetType(Long))
ch.Columns.Add("TxFailed", GetType(Boolean))
ch.Rows.Add("0xaaa", "0xae7b1f50", "SHP-WRONG", "PO-1", "0x01", 1L, t0.AddHours(1), 10L, False)
ch.Rows.Add("0xbbb", "0xae7b1f50", "SHP-A", "PO-1", "0x01", 2L, t0.AddHours(2), 10L, False)

Dim map As New DataTable("TxMap")
For Each c As String In New String() {"TxHash", "ShipmentID", "EventName"}
    map.Columns.Add(c, GetType(String))
Next
map.Rows.Add("0xAAA", "SHP-A", "InTransit")   ' upper case: lookup must not be case-sensitive

Dim sh As New DataTable("Shipments")
sh.Columns.Add("ShipmentID", GetType(String)) : sh.Columns.Add("PONumber", GetType(String))
sh.Columns.Add("Product", GetType(String)) : sh.Columns.Add("ExpectedQty", GetType(Long))
sh.Rows.Add("SHP-A", "PO-1", "Test product", 10L)

Dim ev As New DataTable("ShipmentEvents")
For Each c As String In New String() {"ShipmentID", "EventName", "HandlerAddress", "HandlerPartner"}
    ev.Columns.Add(c, GetType(String))
Next
ev.Columns.Add("ExpectedEventTimestampUtc", GetType(DateTime))
ev.Rows.Add("SHP-A", "InTransit", "0x01", "Partner", t0)

Dim sq As New DataTable("EventSequence")
sq.Columns.Add("StepOrder", GetType(Integer)) : sq.Columns.Add("EventName", GetType(String))
sq.Rows.Add(1, "InTransit") : sq.Rows.Add(2, "Delivered")

Dim sg As New DataTable("EventSignatures")
sg.Columns.Add("EventName", GetType(String)) : sg.Columns.Add("MethodId", GetType(String))
sg.Rows.Add("Delivered", "0xae7b1f50")

out_Chain = ch : out_Map = map : out_Ship = sh : out_Events = ev : out_Seq = sq : out_Sig = sg
""".strip(), [
        ("Out", "out_Chain", "sd:DataTable", "dtChain"),
        ("Out", "out_Map", "sd:DataTable", "dtTxMap"),
        ("Out", "out_Ship", "sd:DataTable", "dtShipments"),
        ("Out", "out_Events", "sd:DataTable", "dtErpEvents"),
        ("Out", "out_Seq", "sd:DataTable", "dtSequence"),
        ("Out", "out_Sig", "sd:DataTable", "dtSignatures"),
    ], name="Invoke Code - chain rows, transaction map and ERP")

    def join(mode: str) -> str:
        return (assign("Config", 'New Dictionary(Of String, Object) From {{"MappingMode", "' + mode + '"}}',
                       type_ref="scg:Dictionary(x:String, x:Object)", name=f"Assign - MappingMode {mode}")
                + invoke_workflow("Workflows\\03_Preprocess_Map.xaml", [
                    ("In", "in_Config", "scg:Dictionary(x:String, x:Object)", "Config"),
                    ("In", "in_dtChain", "sd:DataTable", "dtChain"),
                    ("In", "in_dtShipments", "sd:DataTable", "dtShipments"),
                    ("In", "in_dtErpEvents", "sd:DataTable", "dtErpEvents"),
                    ("In", "in_dtSequence", "sd:DataTable", "dtSequence"),
                    ("In", "in_dtSignatures", "sd:DataTable", "dtSignatures"),
                    ("In", "in_dtTxMap", "sd:DataTable", "dtTxMap"),
                    ("Out", "out_dtJoined", "sd:DataTable", "dtJoined"),
                ], name=f"Invoke 03 - {mode}"))

    j = lambda col, row=0: f'Convert.ToString(dtJoined.Rows({row})("{col}"))'
    steps = (
        tables
        + join("LOOKUP")
        + verify(f'{j("ShipmentID")} = "SHP-A"', "LOOKUP: shipment comes from the map")
        + verify(f'{j("EventName")} = "InTransit"', "LOOKUP: milestone comes from the map")
        + verify(f'{j("MatchStatus")} = "MATCHED"', "LOOKUP: the mapped row joins to the ERP")
        + verify(f'Not IsDBNull(dtJoined.Rows(0)("DriftHours"))', "LOOKUP: and gets an ERP timestamp to compare")
        + verify(f'{j("MatchStatus", 1)} = "NOT_IN_TX_MAP"', "LOOKUP: an unmapped transaction is flagged")
        + join("DECODE")
        + verify(f'{j("ShipmentID")} = "SHP-WRONG" AndAlso {j("EventName")} = "Delivered"',
                 "DECODE: the same row follows the call data instead")
        + verify(f'{j("MatchStatus")} = "NO_ERP_SHIPMENT"', "DECODE: so it no longer joins"))
    return test_case("TC10_LookupMapping", "TC10 LOOKUP mapping mode", steps,
                     (("sd:DataTable", "dtChain"), ("sd:DataTable", "dtTxMap"),
                      ("sd:DataTable", "dtShipments"), ("sd:DataTable", "dtErpEvents"),
                      ("sd:DataTable", "dtSequence"), ("sd:DataTable", "dtSignatures"),
                      ("sd:DataTable", "dtJoined"), ("scg:Dictionary(x:String, x:Object)", "Config")))

# --------------------------------------------------------------------------
GENERATORS = {
    "TC01_R1_TimestampCheck.xaml": tc01,
    "TC02_R2_QuantityMatch.xaml": tc02,
    "TC03_R3_AddressWhitelist.xaml": tc03,
    "TC04_R4_DuplicateDetection.xaml": tc04,
    "TC05_R5_SequenceValidation.xaml": tc05,
    "TC06_R6_SmartContractEvent.xaml": tc06,
    "TC07_AuditChain_TamperEvident.xaml": tc07,
    "TC08_FieldMapping_Engine.xaml": tc08,
    "TC09_Normaliser_Hardening.xaml": tc09,
    "TC10_LookupMapping.xaml": tc10,
}


def register_in_project_json(file_names: list[str]) -> None:
    """
    Add the test cases to designOptions.fileInfoCollection so Studio's Test Explorer
    lists them. Idempotent - existing entries (and their assigned testCaseId, which
    Studio fills in when a test is linked to Orchestrator) are preserved.
    """
    path = os.path.join(PROJ, "project.json")
    with open(path, encoding="utf-8-sig") as f:
        data = json.load(f)

    design = data.setdefault("designOptions", {})
    existing = design.get("fileInfoCollection") or []
    by_name = {e.get("fileName"): e for e in existing}

    for name in file_names:
        rel = f"Tests\\{name}"
        if rel not in by_name:
            by_name[rel] = {"editingStatus": "InProgress", "testCaseId": "", "fileName": rel}

    design["fileInfoCollection"] = [by_name[k] for k in sorted(by_name)]

    # Also register them as entry points. Without this the process packager skips the
    # Tests folder entirely - test cases are normally shipped in a separate *test*
    # package - and `UiRobot execute --entry Tests\...` cannot find them.
    eps = data.setdefault("entryPoints", [])
    known = {e.get("filePath") for e in eps}
    for name in file_names:
        rel = f"Tests\{name}"
        if rel not in known:
            eps.append({
                "filePath": rel,
                "uniqueId": str(uuid.uuid5(uuid.NAMESPACE_URL, "blv-test/" + name)),
                "input": [],
                "output": [],
            })

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
        f.write("\n")
    print(f"registered {len(file_names)} test case(s) in project.json")


def main() -> None:
    os.makedirs(TESTS, exist_ok=True)
    for name, gen in GENERATORS.items():
        with open(os.path.join(TESTS, name), "w", encoding="utf-8") as f:
            f.write(gen())
        print("wrote:", os.path.relpath(os.path.join(TESTS, name), ROOT))
    register_in_project_json(list(GENERATORS))


if __name__ == "__main__":
    main()

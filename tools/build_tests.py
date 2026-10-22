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
    assign, excel_read, excel_scope, if_, invoke_code, invoke_workflow, lit, log, sequence,
    throw, try_catch, variables, vb, workflow,
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
        "ValidationStatus", "Severity", "FailureReasons", "HumanDecision", "HumanDecidedBy", "InputDataHash"}
    dt.Columns.Add(c, GetType(String))
Next
dt.Columns.Add("ChainTimestampUtc", GetType(DateTime))
dt.Columns.Add("OnChainQty", GetType(Long))

Dim t As New DateTime(2026, 9, 1, 6, 0, 0, DateTimeKind.Utc)
dt.Rows.Add("0xaaa", "SHP-T1", "Dispatched", "PASS", "NONE", "", "", "", New String("a"c, 64))
dt.Rows.Add("0xbbb", "SHP-T2", "Delivered", "FAIL", "HIGH", "R2 FAIL: short by 20", "", "", New String("b"c, 64))
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
        + verify('File.ReadAllLines(auditPath)(0).Contains(",InputDataHash,")',
                 "every entry records the input data hash")
        + verify('File.ReadAllLines(auditPath)(2).Contains("," & New String("b"c, 64) & ",")',
                 "the transaction's input data hash is written into its entry")
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
# Shared: load the real configuration inside a test
# --------------------------------------------------------------------------
def load_config() -> str:
    return (assign("projectRoot", "Directory.GetCurrentDirectory()")
            + assign("configPath", 'Path.Combine(projectRoot, "Data", "Config.xlsx")')
            + invoke_workflow("Workflows\\00_Init_ReadConfig.xaml", [
                ("In", "in_ConfigPath", "x:String", "configPath"),
                ("In", "in_ProjectRoot", "x:String", "projectRoot"),
                ("Out", "out_Config", DICT_SO, "Config"),
                ("Out", "out_dtRules", "sd:DataTable", "dtRules"),
                ("Out", "out_dtWallets", "sd:DataTable", "dtWallets"),
                ("Out", "out_dtMapping", "sd:DataTable", "dtMapping"),
                ("Out", "out_dtSequence", "sd:DataTable", "dtSequence"),
                ("Out", "out_dtSignatures", "sd:DataTable", "dtSignatures"),
            ], name="Invoke 00 - read configuration"))


CONFIG_VARS = (("x:String", "projectRoot"), ("x:String", "configPath"), (DICT_SO, "Config"),
               ("sd:DataTable", "dtRules"), ("sd:DataTable", "dtWallets"), ("sd:DataTable", "dtMapping"),
               ("sd:DataTable", "dtSequence"), ("sd:DataTable", "dtSignatures"))


# --------------------------------------------------------------------------
# TC08 - the field mapping makes the bot chain-agnostic (Novelty 3)
# --------------------------------------------------------------------------
def tc08() -> str:
    read_feeds = invoke_code('''
' The two real mock feeds: Etherscan txlist (Ethereum) and an ethereum-etl export
' (Hyperledger Besu). Every field name and two value formats differ between them.
Dim root As String = Directory.GetCurrentDirectory()
Dim eth As JObject = JObject.Parse(File.ReadAllText(Path.Combine(root, "Data", "Input", "MockChain", "etherscan_txlist_response.json")))
Dim etl As JObject = JObject.Parse(File.ReadAllText(Path.Combine(root, "Data", "Input", "MockChain", "evm_etl_besu_export.json")))
out_Eth = CType(eth("result"), JArray)
out_Etl = CType(etl("transactions"), JArray)
'''.strip(), [("Out", "out_Eth", "njl:JArray", "ethRecords"), ("Out", "out_Etl", "njl:JArray", "etlRecords")],
        name="Invoke Code - read both mock feeds")

    read_sheets = excel_scope("configPath", sequence(
        "Read both mapping sheets",
        excel_read("FieldMapping", "dtMapEth", name="Read Range - FieldMapping")
        + excel_read("FieldMapping_EvmEtl", "dtMapEtl", name="Read Range - FieldMapping_EvmEtl")),
        name="Excel Application Scope - Config.xlsx")

    def map_feed(records: str, mapping: str, out: str, label: str) -> str:
        return invoke_workflow("Workflows\\01d_Apply_FieldMapping.xaml", [
            ("In", "in_jRecords", "njl:JArray", records),
            ("In", "in_dtMapping", "sd:DataTable", mapping),
            ("In", "in_SourceMode", "x:String", '"MOCK"'),
            ("Out", "out_dtChain", "sd:DataTable", out),
        ], name=f"Invoke 01d - map the {label} feed")

    compare = invoke_code('''
' Compare every canonical column of every row. SourceMode is the only column allowed
' to be equal by construction, so it proves nothing and is left out.
Dim cols() As String = {"TxHash", "BlockNumber", "EventTimestampUtc", "SenderAddress", "ContractAddress", _
                        "EventSelector", "ShipmentID", "OnChainQty", "PONumber", "TxFailed", "RawInput"}
Dim diffs As New List(Of String)
If in_A.Rows.Count <> in_B.Rows.Count Then diffs.Add("row counts differ")
For i As Integer = 0 To Math.Min(in_A.Rows.Count, in_B.Rows.Count) - 1
    For Each c As String In cols
        If Not Object.Equals(in_A.Rows(i)(c), in_B.Rows(i)(c)) Then
            diffs.Add(String.Format("row {0} {1}: '{2}' vs '{3}'", i, c, in_A.Rows(i)(c), in_B.Rows(i)(c)))
        End If
    Next
Next
out_Diffs = String.Join("; ", diffs.Take(5))
'''.strip(), [("In", "in_A", "sd:DataTable", "dtEth"), ("In", "in_B", "sd:DataTable", "dtEtl"),
          ("Out", "out_Diffs", "x:String", "diffs")], name="Invoke Code - compare the two tables")

    broken_mapping = invoke_code('''
Dim m As DataTable = in_Mapping.Copy()
For i As Integer = m.Rows.Count - 1 To 0 Step -1
    If Convert.ToString(m.Rows(i)("LogisticsField")) = "ShipmentID" Then m.Rows.RemoveAt(i)
Next
out_Mapping = m
'''.strip(), [("In", "in_Mapping", "sd:DataTable", "dtMapEth"), ("Out", "out_Mapping", "sd:DataTable", "dtBroken")],
        name="Invoke Code - mapping without ShipmentID")

    steps = (
        assign("configPath", 'Path.Combine(Directory.GetCurrentDirectory(), "Data", "Config.xlsx")')
        + read_sheets
        + read_feeds
        + map_feed("ethRecords", "dtMapEth", "dtEth", "Etherscan")
        + map_feed("etlRecords", "dtMapEtl", "dtEtl", "ethereum-etl")
        + compare
        + verify("dtEth.Rows.Count = 48", "all 48 Etherscan records are mapped")
        + verify("dtEtl.Rows.Count = 48", "all 48 ethereum-etl records are mapped")
        + verify('diffs = ""', "both schemas produce identical canonical transactions")
        + verify('Convert.ToString(dtEtl.Rows(0)("ShipmentID")) = "SHP-1001"', "ABI string argument decoded")
        + verify('Convert.ToInt64(dtEtl.Rows(0)("OnChainQty")) = 500', "ABI uint256 argument decoded")
        + verify('Convert.ToDateTime(dtEtl.Rows(0)("EventTimestampUtc")).Kind = DateTimeKind.Utc',
                 "timestamps are normalised to UTC")
        + broken_mapping
        + try_catch(
            sequence("Map with an incomplete sheet",
                     map_feed("ethRecords", "dtBroken", "dtChain", "incomplete")
                     + assign("mappingError", '"no error raised"')),
            # Invoke Code runs the body by reflection, so whatever it throws arrives
            # here wrapped in a TargetInvocationException whose own Message is the
            # generic "Exception has been thrown by the target of an invocation."
            # The message naming the unmapped field is on the innermost exception.
            assign("mappingError", "exception.GetBaseException().Message"),
            name="Try Catch - incomplete mapping must be rejected")
        + verify('mappingError.Contains("ShipmentID")', "a mapping sheet missing a required field is rejected by name"))

    return test_case("TC08_FieldMapping_ChainAgnostic", "TC08 field mapping is chain agnostic", steps,
                     (("x:String", "configPath"), ("sd:DataTable", "dtMapEth"), ("sd:DataTable", "dtMapEtl"),
                      ("njl:JArray", "ethRecords"), ("njl:JArray", "etlRecords"),
                      ("sd:DataTable", "dtEth"), ("sd:DataTable", "dtEtl"), ("sd:DataTable", "dtBroken"),
                      ("sd:DataTable", "dtChain"), ("x:String", "diffs"), ("x:String", "mappingError")))


# --------------------------------------------------------------------------
# Shared fixture: a results table with every column the finalise steps read
# --------------------------------------------------------------------------
RESULTS_FIXTURE_CODE = '''
Dim dt As New DataTable("ValidationResults")
dt.Columns.Add("RowNo", GetType(Integer))
For Each c As String In New String() {"TxHash", "ShipmentID", "EventName", "SenderAddress", "ErpHandlerPartner", _
        "PONumber", "Product", "MatchStatus", "R1_Status", "R2_Status", "R3_Status", "R4_Status", "R5_Status", "R6_Status", _
        "ValidationStatus", "Severity", "FailureReasons", "HumanDecision", "HumanDecidedBy", "HumanDecidedAtUtc", "InputDataHash"}
    dt.Columns.Add(c, GetType(String))
Next
dt.Columns.Add("ChainTimestampUtc", GetType(DateTime))
dt.Columns.Add("ErpExpectedQty", GetType(Object))
dt.Columns.Add("OnChainQty", GetType(Long))
dt.Columns.Add("RequiresEscalation", GetType(Boolean))

Dim t As New DateTime(2026, 9, 1, 6, 0, 0, DateTimeKind.Utc)
Dim r1 As DataRow = dt.NewRow()
r1("RowNo") = 1 : r1("TxHash") = "0x" & New String("a"c, 64) : r1("ShipmentID") = "SHP-T1" : r1("EventName") = "Dispatched"
r1("SenderAddress") = "0x" & New String("1"c, 40) : r1("ErpHandlerPartner") = "NordPharma" : r1("PONumber") = "PO-1"
r1("Product") = "Test vials" : r1("MatchStatus") = "MATCHED"
For Each c As String In New String() {"R1_Status", "R2_Status", "R3_Status", "R4_Status", "R5_Status", "R6_Status"}
    r1(c) = "PASS"
Next
r1("ValidationStatus") = "PASS" : r1("Severity") = "NONE" : r1("FailureReasons") = ""
r1("HumanDecision") = "" : r1("HumanDecidedBy") = "" : r1("HumanDecidedAtUtc") = "" : r1("InputDataHash") = New String("c"c, 64)
r1("ChainTimestampUtc") = t : r1("ErpExpectedQty") = 500L : r1("OnChainQty") = 500L : r1("RequiresEscalation") = False
dt.Rows.Add(r1)

' The second row carries the awkward content: parentheses and a backslash that PDF
' strings must escape, and a non-ASCII dash that must not corrupt byte offsets.
Dim r2 As DataRow = dt.NewRow()
r2("RowNo") = 2 : r2("TxHash") = "0x" & New String("b"c, 64) : r2("ShipmentID") = "SHP-T2" : r2("EventName") = "Delivered"
r2("SenderAddress") = "0x" & New String("9"c, 40) : r2("ErpHandlerPartner") = "LastMile" : r2("PONumber") = "PO-2"
r2("Product") = "Test kits" : r2("MatchStatus") = "MATCHED"
For Each c As String In New String() {"R1_Status", "R2_Status", "R4_Status", "R5_Status", "R6_Status"}
    r2(c) = "PASS"
Next
r2("R3_Status") = "FAIL"
r2("ValidationStatus") = "FAIL" : r2("Severity") = "CRITICAL"
r2("FailureReasons") = "R3 FAIL: wallet (unregistered) \\ not approved " & ChrW(8212) & " expected LastMile " & New String("x"c, 300)
r2("HumanDecision") = "" : r2("HumanDecidedBy") = "" : r2("HumanDecidedAtUtc") = "" : r2("InputDataHash") = New String("d"c, 64)
r2("ChainTimestampUtc") = t.AddHours(50) : r2("ErpExpectedQty") = 120L : r2("OnChainQty") = 120L : r2("RequiresEscalation") = True
dt.Rows.Add(r2)

Dim stats As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
stats("Total") = 2 : stats("Pass") = 1 : stats("Warning") = 0 : stats("Fail") = 1 : stats("Escalations") = 1
stats("PassPercent") = 50.0 : stats("RuleHitSummary") = "R3(AddressWhitelist)=1"
For Each rid As String In New String() {"R1", "R2", "R3", "R4", "R5", "R6"}
    stats(rid & "_Hits") = If(rid = "R3", 1, 0)
Next

Dim ex As DataTable = dt.Clone()
ex.ImportRow(r2)

' A private output folder, so a test never touches the project's real reports.
Dim folder As String = Path.Combine(Path.GetTempPath(), "blv_tests_" & Guid.NewGuid().ToString("N"))
Directory.CreateDirectory(folder)

out_dt = dt
out_Exceptions = ex
out_Stats = stats
out_Folder = folder
'''.strip()


def results_fixture() -> str:
    return invoke_code(RESULTS_FIXTURE_CODE, [
        ("Out", "out_dt", "sd:DataTable", "dtResults"),
        ("Out", "out_Exceptions", "sd:DataTable", "dtExceptions"),
        ("Out", "out_Stats", DICT_SO, "Stats"),
        ("Out", "out_Folder", "x:String", "testFolder"),
    ], name="Invoke Code - results fixture")


FIXTURE_VARS = (("sd:DataTable", "dtResults"), ("sd:DataTable", "dtExceptions"),
                (DICT_SO, "Stats"), ("x:String", "testFolder"))


def isolate_outputs() -> str:
    return (assign('Config("ReportsFolder")', "testFolder", type_ref="x:Object", name="Assign - test reports folder")
            + assign('Config("ScreenshotFolder")', "testFolder", type_ref="x:Object", name="Assign - test screenshot folder"))


def remove_folder() -> str:
    return invoke_code('''
If Directory.Exists(in_Folder) Then Directory.Delete(in_Folder, True)
'''.strip(), [("In", "in_Folder", "x:String", "testFolder")], name="Invoke Code - remove the test folder")


# --------------------------------------------------------------------------
# TC09 - the PDF audit report is a structurally valid PDF
# --------------------------------------------------------------------------
def tc09() -> str:
    inspect = invoke_code('''
' Structural checks a PDF reader relies on: header, trailer, a cross-reference table
' whose every offset lands on the object it names, and a page count that matches.
Dim bytes() As Byte = File.ReadAllBytes(in_Path)
Dim text As String = Encoding.ASCII.GetString(bytes)
Dim problems As New List(Of String)
If Not text.StartsWith("%PDF-1.4") Then problems.Add("missing %PDF-1.4 header")
If Not text.EndsWith("%%EOF" & vbLf) Then problems.Add("missing %%EOF trailer")
If bytes.Any(Function(b) b > 127) Then problems.Add("non-ASCII byte present")

Dim sx As Integer = text.LastIndexOf("startxref")
Dim xrefAt As Integer = Integer.Parse(text.Substring(sx + 9).Trim().Split(ChrW(10))(0))
If Not text.Substring(xrefAt).StartsWith("xref") Then problems.Add("startxref does not point at the xref table")

Dim xrefLines() As String = text.Substring(xrefAt).Split(ChrW(10))
Dim count As Integer = Integer.Parse(xrefLines(1).Split(" "c)(1))
For i As Integer = 1 To count - 1
    Dim entry As String = xrefLines(1 + i + 1)
    Dim off As Integer = Integer.Parse(entry.Substring(0, 10))
    If Not text.Substring(off).StartsWith(i.ToString() & " 0 obj") Then problems.Add("xref entry " & i.ToString() & " is misplaced")
Next

Dim pages As Integer = System.Text.RegularExpressions.Regex.Matches(text, "/Type /Page /Parent").Count
Dim declared As Integer = Integer.Parse(System.Text.RegularExpressions.Regex.Match(text, "/Count (\\d+)").Groups(1).Value)
If pages = 0 OrElse pages <> declared Then problems.Add(String.Format("page count mismatch: {0} page objects, /Count {1}", pages, declared))

out_Problems = String.Join("; ", problems)
out_Pages = pages
'''.strip(), [("In", "in_Path", "x:String", "pdfPath"), ("Out", "out_Problems", "x:String", "problems"),
          ("Out", "out_Pages", "x:Int32", "pageCount")], name="Invoke Code - inspect the PDF structure")

    steps = (
        load_config()
        + results_fixture()
        + isolate_outputs()
        + invoke_workflow("Workflows\\08_Report_PDF.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("In", "in_dtResults", "sd:DataTable", "dtResults"),
            ("In", "in_dtRules", "sd:DataTable", "dtRules"),
            ("In", "in_Stats", DICT_SO, "Stats"),
            ("In", "in_AuditPath", "x:String", '"C:\\\\audit\\\\ValidationAuditLog.csv"'),
            ("In", "in_HeadHash", "x:String", 'New String("e"c, 64)'),
            ("In", "in_ChainValid", "x:Boolean", "True"),
            ("In", "in_ChainReport", "x:String", '"VALID - 2 entries verified (test)"'),
            ("In", "in_ReportPath", "x:String", '""'),
            ("Out", "out_PdfPath", "x:String", "pdfPath"),
        ], name="Invoke 08 - write the PDF")
        + verify("File.Exists(pdfPath)", "the PDF is written")
        + inspect
        + verify('problems = ""', "the PDF header, xref table, objects and page tree are consistent")
        + verify("pageCount >= 1", "the report has at least one page")
        + verify('Path.GetFileName(pdfPath).StartsWith("AuditReport_")', "the file is named after the run")
        + remove_folder())

    return test_case("TC09_PdfAuditReport", "TC09 PDF audit report structure", steps,
                     CONFIG_VARS + FIXTURE_VARS + (("x:String", "pdfPath"), ("x:String", "problems"),
                                                   ("x:Int32", "pageCount")))


# --------------------------------------------------------------------------
# TC10 - the review form pauses, and an unanswered form defers (Novelty 2)
# --------------------------------------------------------------------------
def tc10() -> str:
    configure = (
        assign('Config("EscalationEnabled")', '"True"', type_ref="x:Object", name="Assign - escalation on")
        + assign('Config("AttendedMode")', '"True"', type_ref="x:Object", name="Assign - attended")
        + assign('Config("EscalationMode")', '"PROMPT"', type_ref="x:Object", name="Assign - prompt the reviewer")
        + assign('Config("ReviewUi")', '"WINDOW"', type_ref="x:Object", name="Assign - review form")
        # Three seconds: long enough for the form to render, short enough for a test.
        + assign('Config("ReviewTimeoutMinutes")', '"0.05"', type_ref="x:Object", name="Assign - short timeout"))

    steps = (
        load_config()
        + results_fixture()
        + isolate_outputs()
        + configure
        + invoke_workflow("Workflows\\06_HumanInTheLoop_Escalation.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("InOut", "io_dtResults", "sd:DataTable", "dtResults"),
            ("Out", "out_EscalationCount", "x:Int32", "escalations"),
        ], name="Invoke 06 - show the review form")
        + verify("escalations = 1", "exactly the escalated transaction is reviewed")
        + verify('Convert.ToString(dtResults.Rows(1)("HumanDecision")).StartsWith("DEFERRED - no reviewer response")',
                 "an unanswered form records DEFERRED")
        + verify('Convert.ToString(dtResults.Rows(1)("HumanDecidedBy")) = "TIMEOUT"', "the timeout is attributed")
        + verify('Convert.ToString(dtResults.Rows(1)("HumanDecidedAtUtc")) <> ""', "the decision is time-stamped")
        + verify('Convert.ToString(dtResults.Rows(0)("HumanDecision")) = ""', "a passing transaction is not reviewed")
        + verify('File.Exists(Path.Combine(testFolder, "Review_" & Convert.ToString(Config("RunId")) & "_SHP-T2_Delivered.png"))',
                 "an image of the form the reviewer saw is kept as evidence")
        + remove_folder())

    return test_case("TC10_ReviewForm_Timeout", "TC10 review form timeout and evidence", steps,
                     CONFIG_VARS + FIXTURE_VARS + (("x:Int32", "escalations"),))


# --------------------------------------------------------------------------
# TC11 - Teams and Slack payloads are composed correctly
# --------------------------------------------------------------------------
def tc11() -> str:
    inspect = invoke_code('''
Dim problems As New List(Of String)
Dim runId As String = Convert.ToString(in_Config("RunId"))
Dim teams As JObject = JObject.Parse(File.ReadAllText(Path.Combine(in_Folder, "Alert_" & runId & "_teams.json")))
Dim slack As JObject = JObject.Parse(File.ReadAllText(Path.Combine(in_Folder, "Alert_" & runId & "_slack.json")))

If Convert.ToString(teams("type")) <> "message" Then problems.Add("teams: type is not message")
Dim att As JObject = CType(teams("attachments")(0), JObject)
If Convert.ToString(att("contentType")) <> "application/vnd.microsoft.card.adaptive" Then problems.Add("teams: not an Adaptive Card")
If Convert.ToString(att("content")("type")) <> "AdaptiveCard" Then problems.Add("teams: card type")
If Not att.ToString().Contains("SHP-T2") Then problems.Add("teams: finding missing")
If Not att.ToString().Contains("ACTION REQUIRED") Then problems.Add("teams: headline missing")

Dim blocks As JArray = CType(slack("blocks"), JArray)
If blocks.Count < 3 Then problems.Add("slack: expected header, fields and findings blocks")
If Convert.ToString(blocks(0)("type")) <> "header" Then problems.Add("slack: first block is not a header")
If Not slack.ToString().Contains("SHP-T2") Then problems.Add("slack: finding missing")

out_Problems = String.Join("; ", problems)
'''.strip(), [("In", "in_Config", DICT_SO, "Config"), ("In", "in_Folder", "x:String", "testFolder"),
          ("Out", "out_Problems", "x:String", "problems")], name="Invoke Code - inspect both payloads")

    steps = (
        load_config()
        + results_fixture()
        + isolate_outputs()
        + assign('Config("TeamsWebhookUrl")', '""', type_ref="x:Object", name="Assign - no Teams URL")
        + assign('Config("SlackWebhookUrl")', '""', type_ref="x:Object", name="Assign - no Slack URL")
        + invoke_workflow("Workflows\\09b_Alert_Webhooks.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("In", "in_dtExceptions", "sd:DataTable", "dtExceptions"),
            ("In", "in_Stats", DICT_SO, "Stats"),
            ("In", "in_ShouldSend", "x:Boolean", "True"),
        ], name="Invoke 09b - compose Teams and Slack alerts")
        + inspect
        + verify('problems = ""', "Teams gets an Adaptive Card and Slack gets Block Kit, both listing the finding")
        + remove_folder())

    return test_case("TC11_AlertWebhookPayloads", "TC11 Teams and Slack alert payloads", steps,
                     CONFIG_VARS + FIXTURE_VARS + (("x:String", "problems"),))



# --------------------------------------------------------------------------
# TC12 - ledger noise is set aside, not mapped and not fatal
# --------------------------------------------------------------------------
def tc12() -> str:
    build = invoke_code("""
' One genuine recordDispatch call from the real feed, and four records a real ledger
' also carries: a value transfer, a call to some other function, call data that stops
' short, and something that is not a record at all.
Dim root As String = Directory.GetCurrentDirectory()
Dim feed As JObject = JObject.Parse(File.ReadAllText(Path.Combine(root, "Data", "Input", "MockChain", "etherscan_txlist_response.json")))
Dim good As JObject = CType(CType(feed("result"), JArray)(0).DeepClone(), JObject)

Dim transfer As JObject = CType(good.DeepClone(), JObject)
transfer("hash") = "0x02" : transfer("input") = "0x" : transfer("methodId") = "0x"

' approve(address,uint256): a real selector, but not one of this contract's functions.
Dim other As JObject = CType(good.DeepClone(), JObject)
other("hash") = "0x03" : other("input") = "0x095ea7b3" & New String("0"c, 128) : other("methodId") = "0x095ea7b3"

' A known selector (Delivered) whose call data ends after one word.
Dim truncated As JObject = CType(good.DeepClone(), JObject)
truncated("hash") = "0x04" : truncated("input") = "0xae7b1f50" & New String("0"c, 62) & "60" : truncated("methodId") = "0xae7b1f50"

out_Records = New JArray(good, transfer, other, truncated, New JValue(42))
""".strip(), [("Out", "out_Records", "njl:JArray", "records")],
        name="Invoke Code - one real record and four kinds of noise")

    def map_records(label: str, signatures: str, lookup: str) -> str:
        args = [
            ("In", "in_jRecords", "njl:JArray", "records"),
            ("In", "in_dtMapping", "sd:DataTable", "dtMapping"),
            ("In", "in_SourceMode", "x:String", '"MOCK"'),
            ("In", "in_LookupMode", "x:Boolean", lookup),
            ("Out", "out_dtChain", "sd:DataTable", "dtChain"),
            ("Out", "out_Summary", "x:String", "summary"),
        ]
        if signatures:
            args.insert(3, ("In", "in_dtSignatures", "sd:DataTable", signatures))
        return invoke_workflow("Workflows\\01d_Apply_FieldMapping.xaml", args,
                               name=f"Invoke 01d - {label}")

    steps = (
        load_config()
        + build
        + map_records("with the contract's function table", "dtSignatures", "False")
        + verify("dtChain.Rows.Count = 1", "only the genuine logistics call is mapped")
        + verify('Convert.ToString(dtChain.Rows(0)("ShipmentID")) = "SHP-1001"', "and it decodes correctly")
        + verify('summary.Contains("1 skipped (no call data)")', "a value transfer is skipped")
        + verify('summary.Contains("1 call(s) to other functions ignored")', "a call to another function is ignored")
        + verify('summary.Contains("2 rejected")', "truncated call data and a non-object are rejected")
        + map_records("in LOOKUP mode", "dtSignatures", "True")
        + verify('summary.Contains("0 call(s) to other functions ignored")',
                 "LOOKUP mode leaves the decision to the mapping file")
        + map_records("without a function table", "", "False")
        + verify('summary.Contains("0 call(s) to other functions ignored") AndAlso summary.Contains("3 rejected")',
                 "without the table the unrelated call is rejected, not silently mapped"))

    return test_case("TC12_Mapping_IgnoresLedgerNoise", "TC12 ledger noise is set aside", steps,
                     CONFIG_VARS + (("njl:JArray", "records"), ("sd:DataTable", "dtChain"),
                                    ("x:String", "summary")))

# --------------------------------------------------------------------------
GENERATORS = {
    "TC01_R1_TimestampCheck.xaml": tc01,
    "TC02_R2_QuantityMatch.xaml": tc02,
    "TC03_R3_AddressWhitelist.xaml": tc03,
    "TC04_R4_DuplicateDetection.xaml": tc04,
    "TC05_R5_SequenceValidation.xaml": tc05,
    "TC06_R6_SmartContractEvent.xaml": tc06,
    "TC07_AuditChain_TamperEvident.xaml": tc07,
    "TC08_FieldMapping_ChainAgnostic.xaml": tc08,
    "TC09_PdfAuditReport.xaml": tc09,
    "TC10_ReviewForm_Timeout.xaml": tc10,
    "TC11_AlertWebhookPayloads.xaml": tc11,
    "TC12_Mapping_IgnoresLedgerNoise.xaml": tc12,
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
            # The same shape Studio writes when a workflow is marked as a test case.
            by_name[rel] = {"editingStatus": "InProgress",
                            "testCaseId": str(uuid.uuid5(uuid.NAMESPACE_URL, "blv-testcase/" + name)),
                            "testCaseType": "TestCase", "fileName": rel}

    design["fileInfoCollection"] = [by_name[k] for k in sorted(by_name)]

    # Also register them as entry points. Without this the process packager skips the
    # Tests folder entirely - test cases are normally shipped in a separate *test*
    # package - and `UiRobot execute --entry Tests\...` cannot find them.
    eps = data.setdefault("entryPoints", [])
    known = {e.get("filePath") for e in eps}
    for name in file_names:
        rel = f"Tests\\{name}"
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

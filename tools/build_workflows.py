"""
build_workflows.py
==================
Generates the project's .xaml workflows.

The .xaml files are the deliverable and are edited in Studio like any others - this
script exists so the boilerplate header, VB escaping and cross-workflow argument
signatures stay consistent while the project is being built out. Studio rewrites any
file it saves; regenerating afterwards would overwrite those edits, so once a workflow
is being maintained by hand, drop it from GENERATORS below.

Run:  python tools/build_workflows.py
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from xamlgen import (  # noqa: E402
    assign, deserialize_json, excel_read, excel_scope, for_each_row, if_,
    invoke_code, invoke_workflow, lit, log, read_text, sequence, switch, throw,
    try_catch, variables, vb, workflow, write_text,
)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROJ = os.path.join(ROOT, "BlockchainLogisticsValidator")
WF = os.path.join(PROJ, "Workflows")
RULES = os.path.join(WF, "Rules")
QUEUE = os.path.join(WF, "Queue")

DICT_SO = "scg:Dictionary(x:String, x:Object)"

# The reviewer's dialog. Kept as module constants because they are long VB expressions
# and embedding them inline makes the escalation workflow unreadable.
HITL_LABEL_EXPR = (
    '"A blockchain transaction failed a CRITICAL validation rule and needs your decision."'
    ' & vbCrLf & vbCrLf &'
    ' "Shipment:  " & Convert.ToString(CurrentRow("ShipmentID")) &'
    ' "  (" & Convert.ToString(CurrentRow("Product")) & ")" & vbCrLf &'
    ' "Milestone: " & Convert.ToString(CurrentRow("EventName")) & vbCrLf &'
    ' "Recorded:  " & Convert.ToDateTime(CurrentRow("ChainTimestampUtc")).ToString("yyyy-MM-dd HH:mm") & " UTC" & vbCrLf &'
    ' "Signed by: " & Convert.ToString(CurrentRow("SenderAddress")) & vbCrLf &'
    ' "Severity:  " & Convert.ToString(CurrentRow("Severity")) & vbCrLf & vbCrLf &'
    ' "Finding:" & vbCrLf & Convert.ToString(CurrentRow("FailureReasons"))'
)

EXPLORER_URL_EXPR = 'Convert.ToString(in_Config("ExplorerPageUrl"))'

HITL_OPTIONS_EXPR = (
    'New String() {'
    '"APPROVED - verified with the partner out of band", '
    '"REJECTED - hold the shipment and open an investigation", '
    '"ESCALATED - refer to the compliance team", '
    '"DEFERRED - decide at the next review"}'
)


# ==========================================================================
# 00_Init_ReadConfig
# ==========================================================================
def init_read_config() -> str:
    code = r'''
' ---- Flatten the Settings sheet into a case-insensitive dictionary -------
Dim d As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
For Each r As DataRow In in_dtSettings.Rows
    Dim k As String = Convert.ToString(r("Name")).Trim()
    If k <> "" Then
        d(k) = r("Value")
    End If
Next

' ---- Local, uncommitted overrides ---------------------------------------
' Data\Config.local.json (if present) overrides any Settings value. This is where a
' machine-specific output path or a real API key belongs, so Config.xlsx stays
' portable and safe to hand in.
Dim localCfg As String = Path.Combine(in_ProjectRoot, "Data", "Config.local.json")
If File.Exists(localCfg) Then
    Dim ov As JObject = JObject.Parse(File.ReadAllText(localCfg))
    For Each prop As JProperty In ov.Properties()
        d(prop.Name) = prop.Value.ToString()
    Next
End If

' ---- Defaults for settings added after the first release -----------------
' An older Config.xlsx without these rows keeps working exactly as before.
Dim defaults As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
defaults("ProjectName") = "BlockchainLogisticsValidator"
defaults("ChainRecordsPath") = "result"
defaults("FieldMappingSheet") = "FieldMapping"
defaults("LogisticsSourceMode") = "EXCEL"
defaults("ShipmentsSheet") = "Shipments"
defaults("ShipmentEventsSheet") = "ShipmentEvents"
defaults("MappingMode") = "DECODE"
defaults("EscalationMode") = "AUTO_LOG"
defaults("ReviewUi") = "WINDOW"
defaults("ReviewTimeoutMinutes") = "30"
defaults("AlertMode") = "EML"
defaults("AlertAttachReports") = "True"
defaults("OutlookSaveAsDraft") = "False"
defaults("DashboardPushFormat") = "POWERBI"
defaults("RunHistoryFile") = "RunHistory.csv"
defaults("AuditLogFile") = "ValidationAuditLog.csv"
defaults("RunMode") = "BATCH"
defaults("LogLevel") = "Info"
For Each kv As KeyValuePair(Of String, String) In defaults
    If Not d.ContainsKey(kv.Key) OrElse Convert.ToString(d(kv.Key)).Trim() = "" Then d(kv.Key) = kv.Value
Next

' ---- Resolve relative paths against the project root --------------------
' Paths are stored relative in Config.xlsx so the project stays portable; at
' runtime the working directory is the deployed package's content folder.
Dim pathKeys() As String = {"MockChainFile", "LogisticsFile", "TxShipmentMapFile", _
                            "ShipmentsCsvFile", "ShipmentEventsCsvFile", "QueueFile"}
For Each key As String In pathKeys
    If d.ContainsKey(key) Then
        Dim v As String = Convert.ToString(d(key)).Trim()
        If v <> "" AndAlso Not Path.IsPathRooted(v) Then
            d(key) = Path.GetFullPath(Path.Combine(in_ProjectRoot, v))
        End If
    End If
Next

' A database connection string may point at a file inside the project.
If d.ContainsKey("DbConnectionString") Then
    d("DbConnectionString") = Convert.ToString(d("DbConnectionString")).Replace("{ProjectRoot}", in_ProjectRoot)
End If

' ---- Where outputs go ----------------------------------------------------
' Studio runs with the project folder as the working directory, so the default puts
' reports in Data\Output as expected. A published package is extracted to a fresh,
' version-named folder under .nuget on every run, so relying on that default in a
' deployed run would scatter reports where nobody finds them - and the audit chain
' could never span runs, because its file would vanish with each new version.
'
' Precedence: BLV_OUTPUT_ROOT env var > Config.local.json > Settings sheet > default.
Dim outputRoot As String = Environment.GetEnvironmentVariable("BLV_OUTPUT_ROOT")
If String.IsNullOrWhiteSpace(outputRoot) AndAlso d.ContainsKey("OutputRoot") Then
    outputRoot = Convert.ToString(d("OutputRoot")).Trim()
End If
If String.IsNullOrWhiteSpace(outputRoot) Then
    outputRoot = Path.Combine(in_ProjectRoot, "Data", "Output")
End If
If Not Path.IsPathRooted(outputRoot) Then
    outputRoot = Path.GetFullPath(Path.Combine(in_ProjectRoot, outputRoot))
End If

d("OutputRoot") = outputRoot

' The three folder settings are honoured, but anchored under the output root: the
' sheet's default "Data\Output\Reports" becomes "<output root>\Reports", so a
' deployed robot writes wherever BLV_OUTPUT_ROOT points. An absolute path is used as is.
For Each spec As String() In New String()() { _
        New String() {"ReportsFolder", "Reports"}, _
        New String() {"AuditLogFolder", "AuditLogs"}, _
        New String() {"ScreenshotFolder", "Screenshots"}}
    Dim v As String = ""
    If d.ContainsKey(spec(0)) Then v = Convert.ToString(d(spec(0))).Trim().Replace("/", "\")
    If v = "" Then
        v = Path.Combine(outputRoot, spec(1))
    ElseIf Not Path.IsPathRooted(v) Then
        If v.StartsWith("Data\Output\", StringComparison.OrdinalIgnoreCase) Then v = v.Substring("Data\Output\".Length)
        v = Path.GetFullPath(Path.Combine(outputRoot, v))
    End If
    d(spec(0)) = v
Next

' ---- Make sure output folders exist before anything tries to write ------
For Each key As String In New String() {"ReportsFolder", "AuditLogFolder", "ScreenshotFolder"}
    If d.ContainsKey(key) Then
        Dim v As String = Convert.ToString(d(key)).Trim()
        If v <> "" Then Directory.CreateDirectory(v)
    End If
Next

' ---- Fail fast on missing or contradictory configuration ----------------
Dim missing As New List(Of String)
For Each key As String In New String() {"DataSourceMode", "BotIdentity", "ContractAddress"}
    If Not d.ContainsKey(key) OrElse Convert.ToString(d(key)).Trim() = "" Then
        missing.Add(key)
    End If
Next
If missing.Count > 0 Then
    Throw New InvalidOperationException("Config.xlsx is missing required setting(s): " & String.Join(", ", missing))
End If

' Every enumerated setting is normalised to upper case and checked, so a typo in the
' sheet stops the run with a clear message instead of silently taking a default branch.
Dim allowed As New Dictionary(Of String, String())(StringComparer.OrdinalIgnoreCase)
allowed("DataSourceMode") = New String() {"MOCK", "API", "WEB"}
allowed("LogisticsSourceMode") = New String() {"EXCEL", "CSV", "SQL"}
allowed("MappingMode") = New String() {"DECODE", "LOOKUP"}
allowed("EscalationMode") = New String() {"PROMPT", "SIMULATE", "AUTO_LOG"}
allowed("ReviewUi") = New String() {"WINDOW", "DIALOG"}
allowed("AlertMode") = New String() {"EML", "SMTP", "OUTLOOK"}
allowed("DashboardPushFormat") = New String() {"POWERBI", "JSON"}
allowed("RunMode") = New String() {"BATCH", "DISPATCH"}
allowed("LogLevel") = New String() {"INFO", "TRACE"}
For Each kv As KeyValuePair(Of String, String()) In allowed
    Dim v As String = Convert.ToString(d(kv.Key)).Trim().ToUpperInvariant()
    If Array.IndexOf(kv.Value, v) < 0 Then
        Throw New InvalidOperationException(String.Format("{0} must be one of {1}. Found: '{2}'", _
            kv.Key, String.Join(" | ", kv.Value), Convert.ToString(d(kv.Key))))
    End If
    d(kv.Key) = v
Next

Dim escSev As String = If(d.ContainsKey("EscalationSeverity"), Convert.ToString(d("EscalationSeverity")).Trim().ToUpperInvariant(), "")
If escSev <> "" AndAlso Array.IndexOf(New String() {"LOW", "MEDIUM", "HIGH", "CRITICAL"}, escSev) < 0 Then
    Throw New InvalidOperationException("EscalationSeverity must be LOW, MEDIUM, HIGH, CRITICAL or blank. Found: " & escSev)
End If
d("EscalationSeverity") = escSev

Dim reviewMinutes As Double
If Not Double.TryParse(Convert.ToString(d("ReviewTimeoutMinutes")), NumberStyles.Float, CultureInfo.InvariantCulture, reviewMinutes) OrElse reviewMinutes <= 0 Then
    Throw New InvalidOperationException("ReviewTimeoutMinutes must be a positive number. Found: " & Convert.ToString(d("ReviewTimeoutMinutes")))
End If

Dim mode As String = Convert.ToString(d("DataSourceMode"))
If mode = "API" AndAlso Convert.ToString(d("EtherscanApiKey")).Trim() = "" Then
    Throw New InvalidOperationException("DataSourceMode is API but EtherscanApiKey is blank in Config.xlsx.")
End If
If mode = "MOCK" AndAlso Not File.Exists(Convert.ToString(d("MockChainFile"))) Then
    Throw New IO.FileNotFoundException("Mock chain file not found: " & Convert.ToString(d("MockChainFile")))
End If

' The ERP source must exist for whichever mode is selected - and only that mode.
Dim erpMode As String = Convert.ToString(d("LogisticsSourceMode"))
Dim erpFiles As New List(Of String)
If erpMode = "EXCEL" Then erpFiles.Add("LogisticsFile")
If erpMode = "CSV" Then
    erpFiles.Add("ShipmentsCsvFile")
    erpFiles.Add("ShipmentEventsCsvFile")
End If
For Each key As String In erpFiles
    Dim f As String = If(d.ContainsKey(key), Convert.ToString(d(key)).Trim(), "")
    If f = "" OrElse Not File.Exists(f) Then
        Throw New IO.FileNotFoundException(String.Format("LogisticsSourceMode is {0} but {1} was not found: {2}", erpMode, key, f))
    End If
Next
If erpMode = "SQL" Then
    For Each key As String In New String() {"DbProviderName", "DbConnectionString", "DbShipmentsQuery", "DbShipmentEventsQuery"}
        If Not d.ContainsKey(key) OrElse Convert.ToString(d(key)).Trim() = "" Then
            Throw New InvalidOperationException("LogisticsSourceMode is SQL but " & key & " is blank in Config.xlsx.")
        End If
    Next
End If

If Convert.ToString(d("MappingMode")) = "LOOKUP" Then
    Dim mf As String = If(d.ContainsKey("TxShipmentMapFile"), Convert.ToString(d("TxShipmentMapFile")).Trim(), "")
    If mf = "" OrElse Not File.Exists(mf) Then
        Throw New IO.FileNotFoundException("MappingMode is LOOKUP but TxShipmentMapFile was not found: " & mf)
    End If
End If

' ---- Stamp this run -----------------------------------------------------
' The start time is captured before the workbook is opened, so the processing
' time reported on the dashboard includes reading the configuration.
Dim started As DateTime = DateTime.SpecifyKind(in_StartedUtc, DateTimeKind.Utc)
d("RunId") = started.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture) & "-" & Convert.ToString(d("BotIdentity"))
d("RunStartedUtc") = started
d("ConfigSource") = If(File.Exists(localCfg), "Config.xlsx + Data\Config.local.json", "Config.xlsx")
d("MachineName") = Environment.MachineName
d("WindowsUser") = Environment.UserName
d("ProjectRoot") = in_ProjectRoot

out_Config = d
out_Summary = String.Format("mode={0} erp={1} mapping={2}/{3} rules={4} wallets={5} milestones={6} runId={7}", _
    mode, erpMode, d("MappingMode"), d("FieldMappingSheet"), in_dtRules.Rows.Count, in_dtWallets.Rows.Count, _
    in_dtSequence.Rows.Count, d("RunId"))
'''.strip()

    excel_body = sequence("Read configuration sheets",
        excel_read("Settings", "dtSettings")
        + excel_read("ValidationRules", "out_dtRules")
        + excel_read("ApprovedWallets", "out_dtWallets")
        + excel_read("FieldMapping", "out_dtMapping")
        + excel_read("EventSequence", "out_dtSequence")
        + excel_read("EventSignatures", "out_dtSignatures"))

    # Novelty 3: a feed with a different shape is read through a different mapping sheet.
    # The default sheet is already loaded above; any other one is read on demand, so the
    # common case costs no second Excel session.
    alt_mapping = if_(
        'Not String.Equals(Convert.ToString(out_Config("FieldMappingSheet")), "FieldMapping", StringComparison.OrdinalIgnoreCase)',
        sequence("Load the selected mapping sheet",
                 log('"[INIT] Field mapping profile: " & Convert.ToString(out_Config("FieldMappingSheet"))')
                 + excel_scope("in_ConfigPath", sequence(
                     "Read the selected FieldMapping sheet",
                     f'<ui:ExcelReadRange DisplayName="{lit("Read Range - selected FieldMapping sheet")}" '
                     f'AddHeaders="True" DataTable="{vb("out_dtMapping")}" Range="" '
                     f'SheetName="{vb("Convert.ToString(out_Config(" + chr(34) + "FieldMappingSheet" + chr(34) + "))")}" />'),
                     name="Excel Application Scope - Config.xlsx (mapping profile)")),
        name="If - a non-default field mapping profile is selected")

    body = sequence(
        "00 - Initialise and read configuration",
        assign("startedAt", "DateTime.UtcNow", type_ref="s:DateTime", name="Assign - run start time")
        + log('"[INIT] Reading configuration from " & in_ConfigPath')
        + excel_scope("in_ConfigPath", excel_body, name="Excel Application Scope - Config.xlsx")
        + invoke_code(code, [
            ("In", "in_dtSettings", "sd:DataTable", "dtSettings"),
            ("In", "in_dtRules", "sd:DataTable", "out_dtRules"),
            ("In", "in_dtWallets", "sd:DataTable", "out_dtWallets"),
            ("In", "in_dtSequence", "sd:DataTable", "out_dtSequence"),
            ("In", "in_ProjectRoot", "x:String", "in_ProjectRoot"),
            ("In", "in_StartedUtc", "s:DateTime", "startedAt"),
            ("Out", "out_Config", DICT_SO, "out_Config"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - build config dictionary and validate")
        + alt_mapping
        + log('"[INIT] " & summary'),
        variables(("sd:DataTable", "dtSettings"), ("x:String", "summary"),
                  ("s:DateTime", "startedAt")))

    return workflow("00_Init_ReadConfig", body, members=[
        ("in_ConfigPath", "InArgument(x:String)"),
        ("in_ProjectRoot", "InArgument(x:String)"),
        ("out_Config", f"OutArgument({DICT_SO})"),
        ("out_dtRules", "OutArgument(sd:DataTable)"),
        ("out_dtWallets", "OutArgument(sd:DataTable)"),
        ("out_dtMapping", "OutArgument(sd:DataTable)"),
        ("out_dtSequence", "OutArgument(sd:DataTable)"),
        ("out_dtSignatures", "OutArgument(sd:DataTable)"),
    ])


# ==========================================================================
# 01a / 01b / 01c readers  +  01d_Apply_FieldMapping
# ==========================================================================
# The readers only FETCH. Each returns the feed's transaction records untouched, as a
# JArray of JSON objects in whatever shape the source uses. Turning those records into
# the canonical transaction table is 01d's job, driven by the FieldMapping sheet selected
# in Config.xlsx - so a chain or exporter with different field names is supported by a
# different mapping sheet, never by a different workflow. That, and not a shared
# Etherscan envelope, is what makes the bot blockchain-agnostic.
SELECT_RECORDS_CODE = r'''
' The transaction array's location is configuration (ChainRecordsPath), because
' explorers and exporters wrap it differently - Etherscan uses "result", an
' ethereum-etl export uses "transactions".
Dim p As String = If(in_RecordsPath, "").Trim()
If p = "" Then p = "result"
Dim tok As JToken = in_Json.SelectToken(p)
If tok Is Nothing OrElse tok.Type <> JTokenType.Array Then
    Throw New InvalidOperationException("No transaction array at ChainRecordsPath '" & p & _
        "' in the chain response. Check ChainRecordsPath in Config.xlsx.")
End If
out_Records = CType(tok, JArray)
out_Summary = String.Format("{0} record(s) read from '{1}'", out_Records.Count, p)
'''.strip()


def _select_records(json_var: str) -> str:
    return invoke_code(SELECT_RECORDS_CODE, [
        ("In", "in_Json", "njl:JObject", json_var),
        ("In", "in_RecordsPath", "x:String", 'Convert.ToString(in_Config("ChainRecordsPath"))'),
        ("Out", "out_Records", "njl:JArray", "out_jRecords"),
        ("Out", "out_Summary", "x:String", "summary"),
    ], name="Invoke Code - select the transaction records")


# Etherscan answers status "0" with a message for errors; a feed without a status
# field (an exporter's file) is not an error.
STATUS_ERROR_EXPR = ('jsonObj("status") IsNot Nothing AndAlso '
                     'Convert.ToString(jsonObj("status")) <> "1"')


def extract_from_mock() -> str:
    body = sequence(
        "01a - Extract from mock chain feed",
        log('"[EXTRACT/MOCK] Reading " & Convert.ToString(in_Config("MockChainFile"))')
        + read_text('Convert.ToString(in_Config("MockChainFile"))', "jsonText",
                    name="Read Text File - mock chain feed")
        + deserialize_json("jsonText", "jsonObj")
        + if_(STATUS_ERROR_EXPR,
              throw('New InvalidOperationException("Explorer returned status=" '
                    '& Convert.ToString(jsonObj("status")) & " message=" '
                    '& Convert.ToString(jsonObj("message")))'),
              name="If - explorer reported an error")
        + _select_records("jsonObj")
        + log('"[EXTRACT/MOCK] " & summary'),
        variables(("x:String", "jsonText"), ("njl:JObject", "jsonObj"), ("x:String", "summary")))

    return workflow("01a_Extract_FromMockJson", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("out_jRecords", "OutArgument(njl:JArray)"),
    ])


def extract_from_api() -> str:
    build_url = r'''
' Etherscan V2 is multichain: the network is selected by `chainid`, not by a
' different host. Switching Ethereum -> Polygon is therefore a single config cell,
' which is what makes the bot chain-agnostic in practice.
Dim baseUrl As String = Convert.ToString(in_Config("EtherscanBaseUrl")).TrimEnd("/"c)
Dim chainId As String = Convert.ToString(in_Config("ChainId"))
Dim address As String = Convert.ToString(in_Config("ContractAddress"))
Dim apiKey As String = Convert.ToString(in_Config("EtherscanApiKey"))
Dim maxTx As String = Convert.ToString(in_Config("MaxTransactions"))

out_Url = String.Format("{0}?chainid={1}&module=account&action=txlist&address={2}" & _
                        "&startblock=0&endblock=99999999&page=1&offset={3}&sort=asc&apikey={4}", _
                        baseUrl, chainId, address, maxTx, apiKey)

' Never log the key itself.
out_SafeUrl = out_Url.Replace(apiKey, "***REDACTED***")
'''.strip()

    body = sequence(
        "01b - Extract from Etherscan V2 API",
        invoke_code(build_url, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("Out", "out_Url", "x:String", "requestUrl"),
            ("Out", "out_SafeUrl", "x:String", "safeUrl"),
        ], name="Invoke Code - build request URL")
        + log('"[EXTRACT/API] GET " & safeUrl')
        + f'<ui:HttpClient DisplayName="{lit("HTTP Request - txlist")}" '
          f'EndPoint="{vb("requestUrl")}" Method="GET" AcceptFormat="JSON" '
          f'TimeoutMS="{vb("60000")}" Result="{vb("responseText")}" />'
        + deserialize_json("responseText", "jsonObj")
        + if_(STATUS_ERROR_EXPR,
              throw('New InvalidOperationException("Etherscan returned status=" '
                    '& Convert.ToString(jsonObj("status")) & " message=" '
                    '& Convert.ToString(jsonObj("message")) & ". A rate limit or an '
                    'invalid key is the usual cause.")'),
              name="If - Etherscan reported an error")
        + _select_records("jsonObj")
        + log('"[EXTRACT/API] " & summary'),
        variables(("x:String", "requestUrl"), ("x:String", "safeUrl"),
                  ("x:String", "responseText"), ("njl:JObject", "jsonObj"),
                  ("x:String", "summary")))

    return workflow("01b_Extract_FromEtherscanApi", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("out_jRecords", "OutArgument(njl:JArray)"),
    ])


# --------------------------------------------------------------------------
# 01d_Apply_FieldMapping
# --------------------------------------------------------------------------
FIELD_MAPPING_CODE = r'''
' Translate raw chain records into the canonical transaction table. Every column is
' produced by a row of the selected FieldMapping sheet:
'   BlockchainField  - a field of the record, or field:argN / field:selector to read
'                      ABI argument N or the 4-byte function selector from call data
'   LogisticsField   - the canonical column it fills
'   TransformType    - how the value is converted
' Nothing in this code knows any explorer's field names.
Dim canonical As New List(Of KeyValuePair(Of String, Type))
canonical.Add(New KeyValuePair(Of String, Type)("TxHash", GetType(String)))
canonical.Add(New KeyValuePair(Of String, Type)("BlockNumber", GetType(Long)))
canonical.Add(New KeyValuePair(Of String, Type)("EventTimestampUtc", GetType(DateTime)))
canonical.Add(New KeyValuePair(Of String, Type)("SenderAddress", GetType(String)))
canonical.Add(New KeyValuePair(Of String, Type)("ContractAddress", GetType(String)))
canonical.Add(New KeyValuePair(Of String, Type)("EventSelector", GetType(String)))
canonical.Add(New KeyValuePair(Of String, Type)("ShipmentID", GetType(String)))
canonical.Add(New KeyValuePair(Of String, Type)("OnChainQty", GetType(Long)))
canonical.Add(New KeyValuePair(Of String, Type)("PONumber", GetType(String)))
canonical.Add(New KeyValuePair(Of String, Type)("TxFailed", GetType(Boolean)))

Dim known() As String = {"NONE", "TEXT", "LOWERCASE", "UPPERCASE", "INTEGER", "HEX_TO_INTEGER", "DECIMAL", _
                         "EPOCH_TO_DATETIME", "ISO_TO_DATETIME", "BOOLEAN", "BOOLEAN_INVERT", _
                         "ABI_STRING", "ABI_UINT256", "ABI_ADDRESS"}

If in_dtMapping Is Nothing OrElse in_dtMapping.Rows.Count = 0 Then
    Throw New InvalidOperationException("The selected FieldMapping sheet is empty.")
End If

' ---- Read and check the mapping rules ------------------------------------
Dim rules As New List(Of String())
For Each m As DataRow In in_dtMapping.Rows
    Dim src As String = Convert.ToString(m("BlockchainField")).Trim()
    Dim dst As String = Convert.ToString(m("LogisticsField")).Trim()
    Dim tf As String = Convert.ToString(m("TransformType")).Trim().ToUpperInvariant()
    If src = "" OrElse dst = "" Then Continue For
    If tf = "" Then tf = "NONE"
    If Array.IndexOf(known, tf) < 0 Then
        Throw New InvalidOperationException(String.Format("FieldMapping: unknown TransformType '{0}' for {1} -> {2}.", tf, src, dst))
    End If
    rules.Add(New String() {src, dst, tf})
Next

Dim unmapped As New List(Of String)
For Each kv As KeyValuePair(Of String, Type) In canonical
    If Not rules.Any(Function(x) String.Equals(x(1), kv.Key, StringComparison.OrdinalIgnoreCase)) Then unmapped.Add(kv.Key)
Next
If unmapped.Count > 0 Then
    Throw New InvalidOperationException("The FieldMapping sheet does not map required field(s): " & String.Join(", ", unmapped))
End If

Dim dt As New DataTable("ChainTransactions")
For Each kv As KeyValuePair(Of String, Type) In canonical
    Dim col As DataColumn = dt.Columns.Add(kv.Key, kv.Value)
    ' The transforms hand us DateTimeKind.Utc values, but a DataColumn's default
    ' DateTimeMode is UnspecifiedLocal, which strips the Kind back to Unspecified on
    ' the way out. For a column called EventTimestampUtc that is worse than cosmetic:
    ' anything downstream that compares against DateTime.UtcNow, or re-serialises the
    ' value, would silently reinterpret it in the robot's local zone.
    If kv.Value Is GetType(DateTime) Then col.DateTimeMode = DataSetDateTime.Utc
Next
dt.Columns.Add("SourceMode", GetType(String))
dt.Columns.Add("RawInput", GetType(String))
For Each r As String() In rules
    If Not dt.Columns.Contains(r(1)) Then dt.Columns.Add(r(1), GetType(String))
Next

Dim epoch As New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)

' ---- ABI helpers -----------------------------------------------------------
' Call data is a 4-byte selector followed by 32-byte words. A static argument sits
' in its word; a dynamic string's word holds the byte offset of its length word.
Dim abiWords As Func(Of String, List(Of String)) = Function(hex As String) As List(Of String)
    Dim words As New List(Of String)
    Dim h As String = If(hex, "").Trim()
    If h.StartsWith("0x", StringComparison.OrdinalIgnoreCase) Then h = h.Substring(2)
    If h.Length < 8 Then Return words
    Dim payload As String = h.Substring(8)
    Dim i As Integer = 0
    While i + 64 <= payload.Length
        words.Add(payload.Substring(i, 64))
        i += 64
    End While
    Return words
End Function

Dim truthy As Func(Of String, Boolean) = Function(s As String) _
    New String() {"1", "true", "yes", "success", "ok"}.Contains(s.Trim().ToLowerInvariant())

Dim convertValue As Func(Of String, String, String, Object) = Function(raw As String, part As String, tf As String) As Object
    If raw Is Nothing Then Throw New FormatException("field is missing from the record")
    Dim value As String = raw

    If part = "selector" Then
        Dim h As String = raw.Trim()
        If Not h.StartsWith("0x", StringComparison.OrdinalIgnoreCase) Then h = "0x" & h
        If h.Length < 10 Then Throw New FormatException("no call data")
        value = h.Substring(0, 10)
    ElseIf part.StartsWith("arg") Then
        Dim n As Integer = Integer.Parse(part.Substring(3), CultureInfo.InvariantCulture)
        Dim words As List(Of String) = abiWords(raw)
        If words.Count <= n Then Throw New FormatException("no call data")
        If tf = "ABI_STRING" Then
            Dim offWord As Long = Convert.ToInt64(words(n), 16) \ 32
            If offWord < 0 OrElse offWord >= words.Count Then Throw New FormatException("ABI string offset out of range")
            Dim slen As Integer = Convert.ToInt32(words(CInt(offWord)), 16)
            Dim sb As New StringBuilder()
            Dim w As Integer = CInt(offWord) + 1
            While sb.Length < slen * 2 AndAlso w < words.Count
                sb.Append(words(w))
                w += 1
            End While
            If sb.Length < slen * 2 Then Throw New FormatException("ABI string is truncated")
            Dim hx As String = sb.ToString()
            Dim bytes(Math.Max(slen - 1, 0)) As Byte
            For k As Integer = 0 To slen - 1
                bytes(k) = Convert.ToByte(hx.Substring(k * 2, 2), 16)
            Next
            Return Encoding.UTF8.GetString(bytes, 0, slen)
        ElseIf tf = "ABI_UINT256" Then
            Dim wd As String = words(n)
            If wd.Substring(0, 49).Trim("0"c) <> "" Then Throw New OverflowException("uint256 value is too large for a quantity")
            Return Convert.ToInt64(wd.Substring(48), 16)
        ElseIf tf = "ABI_ADDRESS" Then
            Return "0x" & words(n).Substring(24).ToLowerInvariant()
        End If
        value = words(n)
    ElseIf tf.StartsWith("ABI_") Then
        Throw New FormatException(tf & " needs a source of the form field:argN")
    End If

    Select Case tf
        Case "NONE", "TEXT"
            Return value
        Case "LOWERCASE"
            Return value.Trim().ToLowerInvariant()
        Case "UPPERCASE"
            Return value.Trim().ToUpperInvariant()
        Case "INTEGER"
            Return Long.Parse(value.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture)
        Case "HEX_TO_INTEGER"
            Dim hx As String = value.Trim()
            If hx.StartsWith("0x", StringComparison.OrdinalIgnoreCase) Then hx = hx.Substring(2)
            Return Convert.ToInt64(hx, 16)
        Case "DECIMAL"
            Return Decimal.Parse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture)
        Case "EPOCH_TO_DATETIME"
            Return epoch.AddSeconds(Double.Parse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture))
        Case "ISO_TO_DATETIME"
            Dim s As String = value.Trim()
            If s.EndsWith(" UTC", StringComparison.OrdinalIgnoreCase) Then s = s.Substring(0, s.Length - 4)
            Return DateTime.Parse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal)
        Case "BOOLEAN"
            Return truthy(value)
        Case "BOOLEAN_INVERT"
            Return Not truthy(value)
    End Select
    Throw New FormatException("unsupported transform " & tf)
End Function

' ---- Calls this bot does not own -------------------------------------------
' A real ledger carries calls to functions that are not shipment milestones (an approve(),
' an admin call, another contract's traffic). They are out of scope, not bad data, so they
' are counted and set aside. LOOKUP mode skips this: there the mapping file, not the
' selector, decides what a transaction is.
Dim knownSel As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
If Not in_LookupMode AndAlso in_dtSignatures IsNot Nothing Then
    For Each sr As DataRow In in_dtSignatures.Rows
        Dim ms As String = Convert.ToString(sr("MethodId")).Trim()
        If ms <> "" Then knownSel.Add(ms)
    Next
End If
Dim callField As String = ""
For Each r As String() In rules
    Dim ci As Integer = r(0).IndexOf(":"c)
    If ci > 0 Then
        callField = r(0).Substring(0, ci).Trim()
        Exit For
    End If
Next

' ---- Map every record -------------------------------------------------------
Dim noCallData As Integer = 0
Dim foreignCalls As Integer = 0
Dim rejected As Integer = 0
Dim firstProblem As String = ""

For Each tok As JToken In in_Records
    Dim rec As JObject = TryCast(tok, JObject)
    If rec Is Nothing Then
        rejected += 1
        If firstProblem = "" Then firstProblem = "a record is not a JSON object"
        Continue For
    End If

    If knownSel.Count > 0 AndAlso callField <> "" Then
        Dim cdTok As JToken = rec.GetValue(callField, StringComparison.OrdinalIgnoreCase)
        Dim cd As String = If(cdTok Is Nothing OrElse cdTok.Type = JTokenType.Null, "", cdTok.ToString().Trim())
        If cd.Length >= 10 AndAlso Not knownSel.Contains(cd.Substring(0, 10)) Then
            foreignCalls += 1
            Continue For
        End If
    End If

    Dim rw As DataRow = dt.NewRow()
    Dim rowOk As Boolean = True
    Dim abiSource As String = Nothing

    For Each r As String() In rules
        Dim fieldName As String = r(0)
        Dim part As String = ""
        Dim ci As Integer = r(0).IndexOf(":"c)
        If ci > 0 Then
            fieldName = r(0).Substring(0, ci).Trim()
            part = r(0).Substring(ci + 1).Trim().ToLowerInvariant()
        End If

        Dim jt As JToken = rec.GetValue(fieldName, StringComparison.OrdinalIgnoreCase)
        Dim raw As String = Nothing
        If jt IsNot Nothing AndAlso jt.Type <> JTokenType.Null Then raw = jt.ToString()
        If part <> "" AndAlso abiSource Is Nothing Then abiSource = raw

        Try
            Dim v As Object = convertValue(raw, part, r(2))
            If dt.Columns(r(1)).DataType Is GetType(String) Then
                rw(r(1)) = Convert.ToString(v, CultureInfo.InvariantCulture)
            Else
                rw(r(1)) = v
            End If
        Catch ex As Exception
            rowOk = False
            ' A plain value transfer carries no call data - nothing to validate, so it
            ' is skipped rather than reported as bad data.
            If part <> "" AndAlso (raw Is Nothing OrElse raw.Trim().Length < 10) Then
                noCallData += 1
            Else
                rejected += 1
                If firstProblem = "" Then
                    firstProblem = String.Format("{0} -> {1} ({2}): {3}", r(0), r(1), r(2), ex.Message)
                End If
            End If
            Exit For
        End Try
    Next

    If Not rowOk Then Continue For
    rw("SourceMode") = in_SourceMode
    If IsDBNull(rw("RawInput")) OrElse Convert.ToString(rw("RawInput")) = "" Then
        rw("RawInput") = If(abiSource, "").Trim().ToLowerInvariant()
    End If
    dt.Rows.Add(rw)
Next

out_dtChain = dt
out_Rejected = rejected
out_Summary = String.Format("{0} transaction(s) mapped with {1} field rule(s), {2} skipped (no call data), {3} call(s) to other functions ignored, {4} rejected{5}", _
    dt.Rows.Count, rules.Count, noCallData, foreignCalls, rejected, If(firstProblem = "", "", " - first problem: " & firstProblem))
'''.strip()


def apply_field_mapping() -> str:
    body = sequence(
        "01d - Apply the configured field mapping",
        invoke_code(FIELD_MAPPING_CODE, [
            ("In", "in_Records", "njl:JArray", "in_jRecords"),
            ("In", "in_dtMapping", "sd:DataTable", "in_dtMapping"),
            ("In", "in_SourceMode", "x:String", "in_SourceMode"),
            ("In", "in_dtSignatures", "sd:DataTable", "in_dtSignatures"),
            ("In", "in_LookupMode", "x:Boolean", "in_LookupMode"),
            ("Out", "out_dtChain", "sd:DataTable", "out_dtChain"),
            ("Out", "out_Rejected", "x:Int32", "rejected"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - map records to the canonical schema")
        + if_("rejected > 0",
              log('"[MAP-FIELDS] " & summary', level="Warn"),
              log('"[MAP-FIELDS] " & summary'),
              name="If - any record rejected")
        + assign("out_Summary", "summary", name="Assign - summary for the caller"),
        variables(("x:String", "summary"), ("x:Int32", "rejected")))

    return workflow("01d_Apply_FieldMapping", body, members=[
        ("in_jRecords", "InArgument(njl:JArray)"),
        ("in_dtMapping", "InArgument(sd:DataTable)"),
        ("in_SourceMode", "InArgument(x:String)"),
        # Optional: the contract's function table. When given (and not in LOOKUP mode),
        # calls to any other function are ignored rather than mapped.
        ("in_dtSignatures", "InArgument(sd:DataTable)"),
        ("in_LookupMode", "InArgument(x:Boolean)"),
        ("out_dtChain", "OutArgument(sd:DataTable)"),
        ("out_Summary", "OutArgument(x:String)"),
    ])


# ==========================================================================
# 01_Extract_Blockchain  (dispatcher)
# ==========================================================================
def extract_blockchain() -> str:
    reader_args = [
        ("In", "in_Config", DICT_SO, "in_Config"),
        ("Out", "out_jRecords", "njl:JArray", "jRecords"),
    ]
    mock_branch = sequence("MOCK", invoke_workflow(
        "Workflows\\01a_Extract_FromMockJson.xaml", reader_args, name="Invoke 01a - mock chain feed"))
    api_branch = sequence("API", invoke_workflow(
        "Workflows\\01b_Extract_FromEtherscanApi.xaml", reader_args, name="Invoke 01b - Etherscan API"))
    web_branch = sequence("WEB", invoke_workflow(
        "Workflows\\01c_Extract_FromExplorerUI.xaml", reader_args, name="Invoke 01c - explorer UI scraping"))

    dispatch = switch(
        'Convert.ToString(in_Config("DataSourceMode")).ToUpperInvariant()',
        [("MOCK", mock_branch), ("API", api_branch), ("WEB", web_branch)],
        sequence("Default", throw(
            'New InvalidOperationException("Unsupported DataSourceMode: " '
            '& Convert.ToString(in_Config("DataSourceMode")))')),
        name="Switch - DataSourceMode")

    body = sequence(
        "01 - Extract blockchain transactions",
        log('"[EXTRACT] Source mode: " & Convert.ToString(in_Config("DataSourceMode")) '
            '& " | field mapping: " & Convert.ToString(in_Config("FieldMappingSheet"))')
        + dispatch
        # Phase 2 field mapping: raw records -> canonical table, whatever the source.
        + invoke_workflow("Workflows\\01d_Apply_FieldMapping.xaml", [
            ("In", "in_jRecords", "njl:JArray", "jRecords"),
            ("In", "in_dtMapping", "sd:DataTable", "in_dtMapping"),
            ("In", "in_SourceMode", "x:String", 'Convert.ToString(in_Config("DataSourceMode"))'),
            ("In", "in_dtSignatures", "sd:DataTable", "in_dtSignatures"),
            ("In", "in_LookupMode", "x:Boolean", 'Convert.ToString(in_Config("MappingMode")) = "LOOKUP"'),
            ("Out", "out_dtChain", "sd:DataTable", "out_dtChain"),
        ], name="Invoke 01d - apply field mapping")
        + if_("out_dtChain Is Nothing OrElse out_dtChain.Rows.Count = 0",
              log('"[EXTRACT] No transactions returned for contract "'
                  ' & Convert.ToString(in_Config("ContractAddress"))', level="Warn"),
              name="If - nothing returned"),
        variables(("njl:JArray", "jRecords")))

    return workflow("01_Extract_Blockchain", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtMapping", "InArgument(sd:DataTable)"),
        ("in_dtSignatures", "InArgument(sd:DataTable)"),
        ("out_dtChain", "OutArgument(sd:DataTable)"),
    ])


# ==========================================================================
# 01c_Extract_FromExplorerUI
# ==========================================================================
MAP_01C_CODE = r'''
' Parse the explorer's transaction table into records. Columns are located by their
' header text (TxHash, Method, Block, DateTimeUtc, From, To, Value, Input), falling
' back to position if the page has no header row.
'
' The explorer UI is Etherscan's, so each scraped row is emitted with Etherscan's
' txlist field names - the same records the API returns - and the same FieldMapping
' sheet then applies to both.
Dim tbodyStart As Integer = in_PageHtml.IndexOf("<tbody>", StringComparison.OrdinalIgnoreCase)
Dim tbodyEnd As Integer = in_PageHtml.IndexOf("</tbody>", StringComparison.OrdinalIgnoreCase)
If tbodyStart < 0 OrElse tbodyEnd < 0 Then
    Throw New InvalidOperationException("No transaction table found on the explorer page. " & _
        "If a live explorer is being targeted, its markup differs from the expected layout.")
End If
Dim tbody As String = in_PageHtml.Substring(tbodyStart, tbodyEnd - tbodyStart)

Dim rowRx As New Regex("<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
Dim cellRx As New Regex("<td[^>]*>(.*?)</td>", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
Dim headRx As New Regex("<th[^>]*>(.*?)</th>", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
Dim tagRx As New Regex("<[^>]+>")

Dim colIndex As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
Dim theadStart As Integer = in_PageHtml.IndexOf("<thead", StringComparison.OrdinalIgnoreCase)
If theadStart >= 0 AndAlso theadStart < tbodyStart Then
    Dim theadEnd As Integer = in_PageHtml.IndexOf("</thead>", theadStart, StringComparison.OrdinalIgnoreCase)
    If theadEnd > theadStart Then
        Dim pos As Integer = 0
        For Each hm As Match In headRx.Matches(in_PageHtml.Substring(theadStart, theadEnd - theadStart))
            colIndex(System.Net.WebUtility.HtmlDecode(tagRx.Replace(hm.Groups(1).Value, "")).Trim()) = pos
            pos += 1
        Next
    End If
End If
Dim col As Func(Of String, Integer, Integer) = Function(name As String, fallback As Integer) _
    If(colIndex.ContainsKey(name), colIndex(name), fallback)

Dim cHash As Integer = col("TxHash", 0)
Dim cBlock As Integer = col("Block", 2)
Dim cTime As Integer = col("DateTimeUtc", 3)
Dim cFrom As Integer = col("From", 4)
Dim cTo As Integer = col("To", 5)
Dim cValue As Integer = col("Value", 6)
Dim cInput As Integer = col("Input", 7)
Dim needed As Integer = New Integer() {cHash, cBlock, cTime, cFrom, cTo, cValue, cInput}.Max() + 1

Dim epoch As New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)
Dim records As New JArray()
Dim scrapedRows As Integer = 0

For Each rowMatch As Match In rowRx.Matches(tbody)
    Dim cells As New List(Of String)
    For Each cellMatch As Match In cellRx.Matches(rowMatch.Groups(1).Value)
        Dim rawCell As String = tagRx.Replace(cellMatch.Groups(1).Value, "")
        cells.Add(System.Net.WebUtility.HtmlDecode(rawCell).Trim())
    Next
    If cells.Count < needed Then
        Continue For
    End If
    scrapedRows += 1

    Dim at As DateTime = DateTime.SpecifyKind( _
        DateTime.ParseExact(cells(cTime), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), _
        DateTimeKind.Utc)
    Dim inputHex As String = cells(cInput)

    Dim rec As New JObject()
    rec("hash") = cells(cHash)
    rec("blockNumber") = cells(cBlock)
    rec("timeStamp") = CLng((at - epoch).TotalSeconds).ToString(CultureInfo.InvariantCulture)
    rec("from") = cells(cFrom)
    rec("to") = cells(cTo)
    rec("value") = cells(cValue)
    rec("input") = inputHex
    rec("methodId") = If(inputHex.Length >= 10, inputHex.Substring(0, 10), "0x")
    ' The explorer's table lists only mined, successful transactions.
    rec("isError") = "0"
    records.Add(rec)
Next

out_Records = records
out_Summary = String.Format("{0} table row(s) scraped", scrapedRows)
'''.strip()

RESOLVE_01C_CODE = r'''
' The setting may be a project-relative path or a live http(s) URL. When it is blank,
' the live explorer address page is built from ExplorerAddressUrl and ContractAddress.
Dim v As String = Convert.ToString(in_Config("ExplorerPageUrl")).Trim()
If v = "" Then
    Dim tmpl As String = If(in_Config.ContainsKey("ExplorerAddressUrl"), Convert.ToString(in_Config("ExplorerAddressUrl")).Trim(), "")
    If tmpl = "" Then
        Throw New InvalidOperationException("DataSourceMode is WEB but both ExplorerPageUrl and ExplorerAddressUrl are blank.")
    End If
    v = String.Format(tmpl, Convert.ToString(in_Config("ContractAddress")))
End If
If v.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse _
   v.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
    out_IsRemote = True
    out_Location = v
Else
    out_IsRemote = False
    If Path.IsPathRooted(v) Then
        out_Location = v
    Else
        out_Location = Path.GetFullPath(Path.Combine(Convert.ToString(in_Config("ProjectRoot")), v))
    End If
    If Not File.Exists(out_Location) Then
        Throw New IO.FileNotFoundException("Explorer page not found: " & out_Location & _
            ". Run tools/generate_explorer_page.py to create it.")
    End If
End If
'''.strip()


def extract_from_ui() -> str:
    """
    WEB extraction: fetch a blockchain explorer's address page and parse its transaction
    table into records, which 01d then maps exactly as it maps API records.

    What this is, precisely: it parses the explorer's HTML. It does NOT drive a browser
    with UiPath's Data Scraping wizard. Two concrete reasons:

      * UIAutomation 25.10 removed the classic `ExtractStructuredData` activity. Its
        replacement, NExtractData, is built around descriptors the Studio recorder
        generates against a live page - not something to hand-author and call done.
      * Browser automation needs the UiPath browser extension, which is not installed on
        the build machine, so a recorded version could be neither executed nor verified.

    Parsing the page is a genuine scrape and it is fully testable, which a hand-written
    descriptor would not be. To move to true browser automation: install the extension
    (Studio - Home - Tools - UiPath Extensions), record the table with the Data Scraping
    wizard, and convert its DataTable rows into the same records below. The field
    mapping is what carries the value and does not change.

    ExplorerPageUrl accepts a project-relative path or an http(s) URL; when blank,
    ExplorerAddressUrl is formatted with ContractAddress.
    """
    resolve = invoke_code(RESOLVE_01C_CODE, [
        ("In", "in_Config", DICT_SO, "in_Config"),
        ("Out", "out_IsRemote", "x:Boolean", "isRemote"),
        ("Out", "out_Location", "x:String", "location"),
    ], name="Invoke Code - resolve the explorer page location")

    fetch = if_(
        "isRemote",
        sequence("Fetch over HTTP",
                 f'<ui:HttpClient DisplayName="{lit("HTTP Request - explorer page")}" '
                 f'EndPoint="{vb("location")}" Method="GET" AcceptFormat="ANY" '
                 f'TimeoutMS="{vb("60000")}" Result="{vb("pageHtml")}" />'),
        sequence("Read the local page",
                 read_text("location", "pageHtml", name="Read Text File - explorer page")),
        name="If - remote or local page")

    body = sequence(
        "01c - Extract by scraping a blockchain explorer page",
        resolve
        + log('"[EXTRACT/WEB] Scraping " & location')
        + fetch
        + invoke_code(MAP_01C_CODE, [
            ("In", "in_PageHtml", "x:String", "pageHtml"),
            ("Out", "out_Records", "njl:JArray", "out_jRecords"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - parse the transaction table")
        + log('"[EXTRACT/WEB] " & summary'),
        variables(("x:Boolean", "isRemote"), ("x:String", "location"),
                  ("x:String", "pageHtml"), ("x:String", "summary")))

    return workflow("01c_Extract_FromExplorerUI", body,
                    members=[
                        ("in_Config", f"InArgument({DICT_SO})"),
                        ("out_jRecords", "OutArgument(njl:JArray)"),
                    ],
                    extra_imports=["System.Text.RegularExpressions", "System.Net"],
                    extra_refs=["System.Text.RegularExpressions"])


# ==========================================================================
# 02_Extract_Logistics
# ==========================================================================
# LogisticsSourceMode picks where the ERP/WMS data comes from:
#   EXCEL - Excel Application Scope + Read Range, sheet names from Config.xlsx
#   CSV   - Read CSV File, for flat-file ERP exports
#   SQL   - Database Connect + Execute Query + Disconnect, for ERP databases
# Every branch lands in the same two raw tables, and one canonicalisation step settles
# column names and types, so nothing downstream knows which source was used.
ERP_CANONICAL_CODE = r'''
' Whatever the source - Excel, CSV or SQL - the rest of the bot receives the same typed
' tables. CSV gives every value as text, a database gives real DATETIME and integer
' values, and Excel can give either, so types are settled here once rather than being
' guessed downstream.
Dim toDate As Func(Of Object, DateTime) = Function(v As Object) As DateTime
    If v Is Nothing OrElse IsDBNull(v) Then Throw New FormatException("empty timestamp")
    If TypeOf v Is DateTime Then Return DateTime.SpecifyKind(CType(v, DateTime), DateTimeKind.Utc)
    If TypeOf v Is Double Then Return DateTime.SpecifyKind(DateTime.FromOADate(CDbl(v)), DateTimeKind.Utc)
    Dim s As String = Convert.ToString(v, CultureInfo.InvariantCulture).Trim()
    Dim formats() As String = {"yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ssZ", "yyyy-MM-dd HH:mm", "yyyy-MM-dd"}
    Dim parsed As DateTime
    If DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture, _
            DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal, parsed) Then Return parsed
    Return DateTime.SpecifyKind(DateTime.Parse(s, CultureInfo.CurrentCulture), DateTimeKind.Utc)
End Function

Dim toLong As Func(Of Object, Long) = Function(v As Object) As Long
    If v Is Nothing OrElse IsDBNull(v) Then Throw New FormatException("empty quantity")
    Return Convert.ToInt64(Convert.ToDecimal(v, CultureInfo.InvariantCulture))
End Function

' Build a typed copy: the named columns get their canonical type, every other column
' in the source is carried across untouched.
Dim canonicalise As Func(Of DataTable, String, String(), DataTable) = _
    Function(src As DataTable, label As String, spec As String()) As DataTable
        If src Is Nothing Then Throw New InvalidOperationException("ERP source returned no " & label & " table.")
        Dim missingCols As New List(Of String)
        Dim dst As New DataTable(label)
        For Each s As String In spec
            Dim parts() As String = s.Split(":"c)
            If Not src.Columns.Contains(parts(0)) Then
                missingCols.Add(parts(0))
                Continue For
            End If
            Dim t As Type = GetType(String)
            If parts(1) = "date" Then t = GetType(DateTime)
            If parts(1) = "long" Then t = GetType(Long)
            dst.Columns.Add(parts(0), t)
        Next
        If missingCols.Count > 0 Then
            Dim have As New List(Of String)
            For Each c As DataColumn In src.Columns
                have.Add(c.ColumnName)
            Next
            Throw New InvalidOperationException(String.Format("ERP {0} data is missing column(s): {1}. Columns found: {2}", _
                label, String.Join(", ", missingCols), String.Join(", ", have)))
        End If
        For Each c As DataColumn In src.Columns
            If Not dst.Columns.Contains(c.ColumnName) Then dst.Columns.Add(c.ColumnName, GetType(Object))
        Next

        For Each r As DataRow In src.Rows
            ' Skip the blank rows a spreadsheet export often carries at the bottom.
            If IsDBNull(r(spec(0).Split(":"c)(0))) OrElse Convert.ToString(r(spec(0).Split(":"c)(0))).Trim() = "" Then Continue For
            Dim n As DataRow = dst.NewRow()
            For Each c As DataColumn In dst.Columns
                Dim v As Object = r(c.ColumnName)
                Try
                    If c.DataType Is GetType(DateTime) Then
                        n(c.ColumnName) = toDate(v)
                    ElseIf c.DataType Is GetType(Long) Then
                        n(c.ColumnName) = toLong(v)
                    ElseIf c.DataType Is GetType(String) Then
                        n(c.ColumnName) = If(IsDBNull(v), "", Convert.ToString(v, CultureInfo.InvariantCulture).Trim())
                    Else
                        n(c.ColumnName) = v
                    End If
                Catch ex As Exception
                    Throw New FormatException(String.Format("ERP {0}: cannot read {1} = '{2}' ({3})", _
                        label, c.ColumnName, Convert.ToString(v), ex.Message))
                End Try
            Next
            dst.Rows.Add(n)
        Next
        Return dst
    End Function

Dim ships As DataTable = canonicalise(in_dtShipRaw, "Shipments", _
    New String() {"ShipmentID:text", "PONumber:text", "Product:text", "ExpectedQty:long", "ReceiverAddress:text"})
Dim events As DataTable = canonicalise(in_dtEventsRaw, "ShipmentEvents", _
    New String() {"ShipmentID:text", "EventName:text", "ExpectedEventTimestampUtc:date", "HandlerPartner:text", "HandlerAddress:text"})

' ERP addresses are already lower-case, but normalise defensively - a hand-edited
' export is exactly the kind of thing that reintroduces mixed case.
For Each r As DataRow In events.Rows
    r("HandlerAddress") = Convert.ToString(r("HandlerAddress")).ToLowerInvariant()
Next
For Each r As DataRow In ships.Rows
    r("ReceiverAddress") = Convert.ToString(r("ReceiverAddress")).ToLowerInvariant()
Next

out_dtShipments = ships
out_dtErpEvents = events
out_Summary = String.Format("{0} shipments, {1} expected milestones (source {2})", _
    ships.Rows.Count, events.Rows.Count, in_Mode)
'''.strip()


def extract_logistics() -> str:
    def cfg(key: str) -> str:
        return f'Convert.ToString(in_Config("{key}"))'

    excel_branch = sequence(
        "EXCEL",
        log('"[ERP] Reading workbook " & ' + cfg("LogisticsFile"))
        + excel_scope(cfg("LogisticsFile"), sequence(
            "Read ERP sheets",
            f'<ui:ExcelReadRange DisplayName="{lit("Read Range - shipments sheet")}" AddHeaders="True" '
            f'DataTable="{vb("dtShipRaw")}" Range="" SheetName="{vb(cfg("ShipmentsSheet"))}" />'
            f'<ui:ExcelReadRange DisplayName="{lit("Read Range - shipment events sheet")}" AddHeaders="True" '
            f'DataTable="{vb("dtEventsRaw")}" Range="" SheetName="{vb(cfg("ShipmentEventsSheet"))}" />'),
            name="Excel Application Scope - ERP workbook"))

    csv_branch = sequence(
        "CSV",
        log('"[ERP] Reading CSV exports " & ' + cfg("ShipmentsCsvFile") + ' & " and " & ' + cfg("ShipmentEventsCsvFile"))
        + f'<ui:ReadCsvFile DisplayName="{lit("Read CSV - shipments")}" FilePath="{vb(cfg("ShipmentsCsvFile"))}" '
          f'DataTable="{vb("dtShipRaw")}" IncludeColumnNames="True" Delimitator="Comma" />'
        + f'<ui:ReadCsvFile DisplayName="{lit("Read CSV - shipment events")}" FilePath="{vb(cfg("ShipmentEventsCsvFile"))}" '
          f'DataTable="{vb("dtEventsRaw")}" IncludeColumnNames="True" Delimitator="Comma" />')

    # Disconnect in a Finally so a failed query never leaks the connection.
    disconnect = if_("dbConn IsNot Nothing",
                     f'<ui:DatabaseDisconnect DisplayName="{lit("Disconnect - ERP database")}" '
                     f'DatabaseConnection="{vb("dbConn")}" />',
                     name="If - connection is open")
    sql_branch = sequence(
        "SQL",
        log('"[ERP] Connecting to the ERP database via " & ' + cfg("DbProviderName"))
        + f'<ui:DatabaseConnect DisplayName="{lit("Connect - ERP database")}" '
          f'ProviderName="{vb(cfg("DbProviderName"))}" ConnectionString="{vb(cfg("DbConnectionString"))}" '
          f'DatabaseConnection="{vb("dbConn")}" />'
        + f'<TryCatch DisplayName="{lit("Try Finally - always disconnect")}"><TryCatch.Try>'
        + sequence("Query the ERP tables",
                   f'<ui:ExecuteQuery DisplayName="{lit("Execute Query - shipments")}" '
                   f'ExistingDbConnection="{vb("dbConn")}" Sql="{vb(cfg("DbShipmentsQuery"))}" '
                   f'DataTable="{vb("dtShipRaw")}" CommandType="Text" TimeoutMS="{vb("60000")}" />'
                   f'<ui:ExecuteQuery DisplayName="{lit("Execute Query - shipment events")}" '
                   f'ExistingDbConnection="{vb("dbConn")}" Sql="{vb(cfg("DbShipmentEventsQuery"))}" '
                   f'DataTable="{vb("dtEventsRaw")}" CommandType="Text" TimeoutMS="{vb("60000")}" />')
        + '</TryCatch.Try><TryCatch.Finally>' + disconnect + '</TryCatch.Finally></TryCatch>')

    source = switch(
        cfg("LogisticsSourceMode") + ".ToUpperInvariant()",
        [("EXCEL", excel_branch), ("CSV", csv_branch), ("SQL", sql_branch)],
        sequence("Default", throw('New InvalidOperationException("Unsupported LogisticsSourceMode: " & '
                                  + cfg("LogisticsSourceMode") + ')')),
        name="Switch - LogisticsSourceMode")

    body = sequence(
        "02 - Extract logistics (ERP/WMS) records",
        log('"[ERP] Source mode: " & ' + cfg("LogisticsSourceMode"))
        + source
        + invoke_code(ERP_CANONICAL_CODE, [
            ("In", "in_dtShipRaw", "sd:DataTable", "dtShipRaw"),
            ("In", "in_dtEventsRaw", "sd:DataTable", "dtEventsRaw"),
            ("In", "in_Mode", "x:String", cfg("LogisticsSourceMode")),
            ("Out", "out_dtShipments", "sd:DataTable", "out_dtShipments"),
            ("Out", "out_dtErpEvents", "sd:DataTable", "out_dtErpEvents"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - canonicalise ERP tables")
        + log('"[ERP] " & summary'),
        variables(("x:String", "summary"), ("sd:DataTable", "dtShipRaw"),
                  ("sd:DataTable", "dtEventsRaw"), ("ui:DatabaseConnection", "dbConn")))

    return workflow("02_Extract_Logistics", body,
                    members=[
                        ("in_Config", f"InArgument({DICT_SO})"),
                        ("out_dtShipments", "OutArgument(sd:DataTable)"),
                        ("out_dtErpEvents", "OutArgument(sd:DataTable)"),
                    ],
                    extra_imports=["UiPath.Database"],
                    extra_refs=["UiPath.Database", "UiPath.Database.Activities"])


# ==========================================================================
# 03_Preprocess_Map
# ==========================================================================
def preprocess_map() -> str:
    code = r'''
' Join the chain feed to the ERP expectation for the same (shipment, milestone).
' The milestone name is resolved from the function selector via the EventSequence
' sheet, so adding a new milestone is a config change, not a workflow change.

' ---- MappingMode = LOOKUP ---------------------------------------------------
' The deck allows two ways of linking a transaction to its shipment: a key embedded in
' the call data (DECODE - already produced by the field mapping) or a separate mapping
' file. LOOKUP takes ShipmentID, PO number and function selector from
' TxShipmentMapFile by transaction hash, overriding whatever the call data said.
Dim lookup As New Dictionary(Of String, DataRow)(StringComparer.OrdinalIgnoreCase)
Dim lookedUp As Integer = 0
Dim notInMap As Integer = 0
If in_Mode = "LOOKUP" Then
    If in_dtTxMap Is Nothing Then
        Throw New InvalidOperationException("MappingMode is LOOKUP but the mapping file was not read.")
    End If
    For Each req As String In New String() {"TxHash", "ShipmentID", "PONumber"}
        If Not in_dtTxMap.Columns.Contains(req) Then
            Throw New InvalidOperationException("TxShipmentMapFile is missing column " & req)
        End If
    Next
    For Each m As DataRow In in_dtTxMap.Rows
        Dim h As String = Convert.ToString(m("TxHash")).Trim().ToLowerInvariant()
        If h <> "" Then lookup(h) = m
    Next
End If

Dim selectorToEvent As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
Dim eventOrder As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
For Each r As DataRow In in_dtSequence.Rows
    Dim sel As String = Convert.ToString(r("MethodId")).Trim().ToLowerInvariant()
    Dim nm As String = Convert.ToString(r("EventName")).Trim()
    If sel <> "" Then selectorToEvent(sel) = nm
    eventOrder(nm) = Convert.ToInt32(r("StepOrder"))
Next

' ERP lookups, keyed the same way.
Dim erpEvent As New Dictionary(Of String, DataRow)(StringComparer.OrdinalIgnoreCase)
For Each r As DataRow In in_dtErpEvents.Rows
    erpEvent(Convert.ToString(r("ShipmentID")).Trim() & "|" & Convert.ToString(r("EventName")).Trim()) = r
Next
Dim shipment As New Dictionary(Of String, DataRow)(StringComparer.OrdinalIgnoreCase)
For Each r As DataRow In in_dtShipments.Rows
    shipment(Convert.ToString(r("ShipmentID")).Trim()) = r
Next

Dim dt As New DataTable("JoinedTransactions")
dt.Columns.Add("RowNo", GetType(Integer))
dt.Columns.Add("TxHash", GetType(String))
dt.Columns.Add("BlockNumber", GetType(Long))
dt.Columns.Add("ShipmentID", GetType(String))
dt.Columns.Add("EventName", GetType(String))
dt.Columns.Add("StepOrder", GetType(Integer))
dt.Columns.Add("EventSelector", GetType(String))
dt.Columns.Add("ChainTimestampUtc", GetType(DateTime))
dt.Columns.Add("ErpExpectedTimestampUtc", GetType(Object))
dt.Columns.Add("DriftHours", GetType(Object))
dt.Columns.Add("SenderAddress", GetType(String))
dt.Columns.Add("ErpHandlerAddress", GetType(String))
dt.Columns.Add("ErpHandlerPartner", GetType(String))
dt.Columns.Add("OnChainQty", GetType(Long))
dt.Columns.Add("ErpExpectedQty", GetType(Object))
dt.Columns.Add("PONumber", GetType(String))
dt.Columns.Add("ErpPONumber", GetType(String))
dt.Columns.Add("Product", GetType(String))
dt.Columns.Add("TxFailed", GetType(Boolean))
dt.Columns.Add("MatchStatus", GetType(String))
dt.Columns.Add("RawInput", GetType(String))

Dim hasRawInput As Boolean = in_dtChain.Columns.Contains("RawInput")
Dim unmatched As Integer = 0
Dim n As Integer = 0

For Each c As DataRow In in_dtChain.Rows
    n += 1
    Dim sel As String = Convert.ToString(c("EventSelector")).Trim().ToLowerInvariant()
    Dim sid As String = Convert.ToString(c("ShipmentID")).Trim()
    Dim po As String = Convert.ToString(c("PONumber")).Trim()

    If in_Mode = "LOOKUP" Then
        Dim h As String = Convert.ToString(c("TxHash")).Trim().ToLowerInvariant()
        If lookup.ContainsKey(h) Then
            Dim m As DataRow = lookup(h)
            sid = Convert.ToString(m("ShipmentID")).Trim()
            po = Convert.ToString(m("PONumber")).Trim()
            If m.Table.Columns.Contains("MethodId") AndAlso Convert.ToString(m("MethodId")).Trim() <> "" Then
                sel = Convert.ToString(m("MethodId")).Trim().ToLowerInvariant()
            End If
            lookedUp += 1
        Else
            ' Not in the map: the transaction cannot be tied to a shipment, which the
            ' join below reports as NO_ERP_SHIPMENT rather than guessing from call data.
            sid = ""
            notInMap += 1
        End If
    End If

    Dim evName As String = ""
    If selectorToEvent.ContainsKey(sel) Then evName = selectorToEvent(sel)

    Dim rw As DataRow = dt.NewRow()
    rw("RowNo") = n
    rw("TxHash") = c("TxHash")
    rw("BlockNumber") = c("BlockNumber")
    rw("ShipmentID") = sid
    rw("EventName") = evName
    rw("StepOrder") = If(evName <> "" AndAlso eventOrder.ContainsKey(evName), eventOrder(evName), 0)
    rw("EventSelector") = sel
    rw("ChainTimestampUtc") = c("EventTimestampUtc")
    rw("SenderAddress") = c("SenderAddress")
    rw("OnChainQty") = c("OnChainQty")
    rw("PONumber") = po
    rw("TxFailed") = c("TxFailed")
    rw("RawInput") = If(hasRawInput, Convert.ToString(c("RawInput")), "")

    Dim status As String = "MATCHED"
    If evName = "" Then status = "UNKNOWN_SELECTOR"

    Dim key As String = sid & "|" & evName
    If erpEvent.ContainsKey(key) Then
        Dim e As DataRow = erpEvent(key)
        Dim expectedAt As DateTime = Convert.ToDateTime(e("ExpectedEventTimestampUtc"))
        rw("ErpExpectedTimestampUtc") = expectedAt
        rw("DriftHours") = Math.Abs((Convert.ToDateTime(c("EventTimestampUtc")) - expectedAt).TotalHours)
        rw("ErpHandlerAddress") = Convert.ToString(e("HandlerAddress"))
        rw("ErpHandlerPartner") = Convert.ToString(e("HandlerPartner"))
    Else
        rw("ErpExpectedTimestampUtc") = DBNull.Value
        rw("DriftHours") = DBNull.Value
        rw("ErpHandlerAddress") = ""
        rw("ErpHandlerPartner") = ""
        If status = "MATCHED" Then status = "NO_ERP_MILESTONE"
    End If

    If shipment.ContainsKey(sid) Then
        Dim s As DataRow = shipment(sid)
        rw("ErpExpectedQty") = Convert.ToInt64(s("ExpectedQty"))
        rw("ErpPONumber") = Convert.ToString(s("PONumber"))
        rw("Product") = Convert.ToString(s("Product"))
    Else
        rw("ErpExpectedQty") = DBNull.Value
        rw("ErpPONumber") = ""
        rw("Product") = ""
        status = "NO_ERP_SHIPMENT"
    End If

    If status <> "MATCHED" Then unmatched += 1
    rw("MatchStatus") = status
    dt.Rows.Add(rw)
Next

out_dtJoined = dt
out_Summary = String.Format("{0} rows joined, {1} unmatched", dt.Rows.Count, unmatched)
If in_Mode = "LOOKUP" Then
    out_Summary &= String.Format(" | MappingMode LOOKUP: {0} resolved from the mapping file, {1} not in it", lookedUp, notInMap)
Else
    out_Summary &= " | MappingMode DECODE"
End If
'''.strip()

    read_map = if_(
        'Convert.ToString(in_Config("MappingMode")) = "LOOKUP"',
        sequence("Read the TxHash -> ShipmentID mapping file",
                 log('"[MAP] MappingMode LOOKUP - reading " & Convert.ToString(in_Config("TxShipmentMapFile"))')
                 + excel_scope('Convert.ToString(in_Config("TxShipmentMapFile"))',
                               sequence("Read mapping sheet",
                                        excel_read("TxMap", "dtTxMap", name="Read Range - TxMap")),
                               name="Excel Application Scope - TxShipmentMap.xlsx")),
        name="If - MappingMode is LOOKUP")

    body = sequence(
        "03 - Preprocess and map chain to ERP",
        log('"[MAP] Joining " & in_dtChain.Rows.Count.ToString() & " chain rows to ERP milestones"')
        + read_map
        + invoke_code(code, [
            ("In", "in_dtChain", "sd:DataTable", "in_dtChain"),
            ("In", "in_dtShipments", "sd:DataTable", "in_dtShipments"),
            ("In", "in_dtErpEvents", "sd:DataTable", "in_dtErpEvents"),
            ("In", "in_dtSequence", "sd:DataTable", "in_dtSequence"),
            ("In", "in_dtTxMap", "sd:DataTable", "dtTxMap"),
            ("In", "in_Mode", "x:String", 'Convert.ToString(in_Config("MappingMode"))'),
            ("Out", "out_dtJoined", "sd:DataTable", "out_dtJoined"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - map and join")
        + log('"[MAP] " & summary'),
        variables(("x:String", "summary"), ("sd:DataTable", "dtTxMap")))

    return workflow("03_Preprocess_Map", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtChain", "InArgument(sd:DataTable)"),
        ("in_dtShipments", "InArgument(sd:DataTable)"),
        ("in_dtErpEvents", "InArgument(sd:DataTable)"),
        ("in_dtSequence", "InArgument(sd:DataTable)"),
        ("out_dtJoined", "OutArgument(sd:DataTable)"),
    ])


# ==========================================================================
# 05_Validate_Engine + Rules/R1..R6
# ==========================================================================
# Each rule is its own workflow, but takes the WHOLE results table and annotates it,
# rather than being invoked once per row. Two reasons:
#   * R4 (duplicates) and R5 (ordering) are inherently set-based - they compare rows
#     against each other and cannot be evaluated one row at a time.
#   * 6 invocations instead of 6 x 48. Invoke Workflow File has real per-call overhead.
# Modularity is preserved: every rule is still a separate, separately-testable file.

RULE_COLUMNS = ["R1", "R2", "R3", "R4", "R5", "R6"]


def _skeleton_code() -> str:
    cols = "\n".join(
        f'dt.Columns.Add("{r}_Status", GetType(String))\ndt.Columns.Add("{r}_Detail", GetType(String))'
        for r in RULE_COLUMNS)
    init = "\n    ".join(f'rw("{r}_Status") = "SKIPPED"' for r in RULE_COLUMNS)
    return f'''
' Build the results table. Every rule column starts as SKIPPED; an enabled rule
' overwrites its own column, so a rule switched off in Config.xlsx is visibly
' skipped in the report rather than silently absent.
Dim dt As New DataTable("ValidationResults")
dt.Columns.Add("RowNo", GetType(Integer))
dt.Columns.Add("TxHash", GetType(String))
dt.Columns.Add("ShipmentID", GetType(String))
dt.Columns.Add("EventName", GetType(String))
dt.Columns.Add("StepOrder", GetType(Integer))
dt.Columns.Add("ChainTimestampUtc", GetType(DateTime))
dt.Columns.Add("ErpExpectedTimestampUtc", GetType(Object))
dt.Columns.Add("DriftHours", GetType(Object))
dt.Columns.Add("SenderAddress", GetType(String))
dt.Columns.Add("ErpHandlerPartner", GetType(String))
dt.Columns.Add("OnChainQty", GetType(Long))
dt.Columns.Add("ErpExpectedQty", GetType(Object))
dt.Columns.Add("PONumber", GetType(String))
dt.Columns.Add("Product", GetType(String))
dt.Columns.Add("MatchStatus", GetType(String))
{cols}
dt.Columns.Add("ValidationStatus", GetType(String))
dt.Columns.Add("Severity", GetType(String))
dt.Columns.Add("FailureReasons", GetType(String))
dt.Columns.Add("RequiresEscalation", GetType(Boolean))
dt.Columns.Add("HumanDecision", GetType(String))
dt.Columns.Add("HumanDecidedBy", GetType(String))
dt.Columns.Add("HumanDecidedAtUtc", GetType(String))
dt.Columns.Add("InputDataHash", GetType(String))

' Novelty 4 records a hash of each transaction's input data in the audit trail, so the
' exact call data a verdict was based on is fixed alongside the verdict itself.
Dim sha As SHA256 = SHA256.Create()
Dim hasRaw As Boolean = in_dtJoined.Columns.Contains("RawInput")

For Each j As DataRow In in_dtJoined.Rows
    Dim rw As DataRow = dt.NewRow()
    For Each c As DataColumn In in_dtJoined.Columns
        If dt.Columns.Contains(c.ColumnName) Then rw(c.ColumnName) = j(c.ColumnName)
    Next
    Dim rawInput As String = If(hasRaw, Convert.ToString(j("RawInput")).Trim().ToLowerInvariant(), "")
    rw("InputDataHash") = If(rawInput = "", "-", _
        BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(rawInput))).Replace("-", "").ToLowerInvariant())
    {init}
    rw("ValidationStatus") = "PASS"
    rw("Severity") = "NONE"
    rw("FailureReasons") = ""
    rw("RequiresEscalation") = False
    rw("HumanDecision") = ""
    rw("HumanDecidedBy") = ""
    rw("HumanDecidedAtUtc") = ""
    dt.Rows.Add(rw)
Next

out_dtResults = dt
out_Summary = dt.Rows.Count.ToString() & " rows staged for validation"
'''.strip()


def _finalise_code() -> str:
    # Severity ranking drives which failure decides the row's overall severity.
    checks = "\n".join(f'''
    If Convert.ToString(r("{c}_Status")) = "FAIL" OrElse Convert.ToString(r("{c}_Status")) = "WARNING" Then
        Dim st As String = Convert.ToString(r("{c}_Status"))
        Dim sev As String = "MEDIUM"
        If ruleSeverity.ContainsKey("{c}") Then sev = ruleSeverity("{c}")
        reasons.Add("{c} " & st & ": " & Convert.ToString(r("{c}_Detail")))
        If st = "FAIL" Then
            anyFail = True
            If rank(sev) > rank(worstSev) Then worstSev = sev
            If ruleAction.ContainsKey("{c}") AndAlso ruleAction("{c}") = "ESCALATE" Then escalate = True
            If escRank > 0 AndAlso rank(sev) >= escRank Then escalate = True
        Else
            anyWarn = True
        End If
    End If'''.rstrip() for c in RULE_COLUMNS)

    return f'''
' Roll the six per-rule outcomes up into one status per transaction, then split
' the failures out into an exceptions table for reporting and alerting.
Dim ruleSeverity As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
Dim ruleAction As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
Dim ruleName As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
For Each rr As DataRow In in_dtRules.Rows
    Dim id As String = Convert.ToString(rr("RuleID")).Trim()
    ruleSeverity(id) = Convert.ToString(rr("Severity")).Trim().ToUpperInvariant()
    ruleAction(id) = Convert.ToString(rr("FailAction")).Trim().ToUpperInvariant()
    ruleName(id) = Convert.ToString(rr("RuleName")).Trim()
Next

Dim order As New List(Of String)(New String() {{"NONE", "LOW", "MEDIUM", "HIGH", "CRITICAL"}})

' A FAIL escalates to a human when its rule says ESCALATE, or when its severity is at
' or above EscalationSeverity - "a high-severity anomaly pauses the bot".
Dim escSev As String = ""
If in_Config IsNot Nothing AndAlso in_Config.ContainsKey("EscalationSeverity") Then
    escSev = Convert.ToString(in_Config("EscalationSeverity")).Trim().ToUpperInvariant()
End If
Dim escRank As Integer = If(escSev = "", -1, order.IndexOf(escSev))

Dim dtEx As DataTable = in_dtResults.Clone()
dtEx.TableName = "Exceptions"

Dim nPass As Integer = 0, nWarn As Integer = 0, nFail As Integer = 0, nEsc As Integer = 0
Dim ruleHits As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

For Each r As DataRow In in_dtResults.Rows
    Dim reasons As New List(Of String)
    Dim anyFail As Boolean = False
    Dim anyWarn As Boolean = False
    Dim escalate As Boolean = False
    Dim worstSev As String = "NONE"
{checks}

    If anyFail Then
        r("ValidationStatus") = "FAIL"
        nFail += 1
    ElseIf anyWarn Then
        r("ValidationStatus") = "WARNING"
        nWarn += 1
    Else
        r("ValidationStatus") = "PASS"
        nPass += 1
    End If

    r("Severity") = worstSev
    r("FailureReasons") = String.Join(" | ", reasons)
    r("RequiresEscalation") = escalate
    If escalate Then nEsc += 1

    For Each rid As String In New String() {{"R1", "R2", "R3", "R4", "R5", "R6"}}
        Dim st As String = Convert.ToString(r(rid & "_Status"))
        If st = "FAIL" OrElse st = "WARNING" Then
            If Not ruleHits.ContainsKey(rid) Then ruleHits(rid) = 0
            ruleHits(rid) = ruleHits(rid) + 1
        End If
    Next

    If anyFail OrElse anyWarn Then dtEx.ImportRow(r)
Next

out_dtExceptions = dtEx

Dim stats As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
stats("Total") = in_dtResults.Rows.Count
stats("Pass") = nPass
stats("Warning") = nWarn
stats("Fail") = nFail
stats("Escalations") = nEsc
stats("PassPercent") = If(in_dtResults.Rows.Count = 0, 0.0, Math.Round(nPass * 100.0 / in_dtResults.Rows.Count, 1))
Dim hits As New List(Of String)
For Each rid As String In New String() {{"R1", "R2", "R3", "R4", "R5", "R6"}}
    Dim c As Integer = 0
    If ruleHits.ContainsKey(rid) Then c = ruleHits(rid)
    stats(rid & "_Hits") = c
    If c > 0 Then hits.Add(rid & "(" & ruleName(rid) & ")=" & c.ToString())
Next
stats("RuleHitSummary") = String.Join(", ", hits)
out_Stats = stats

out_Summary = String.Format("total={{0}} pass={{1}} warn={{2}} fail={{3}} escalations={{4}} | {{5}}", _
    in_dtResults.Rows.Count, nPass, nWarn, nFail, nEsc, String.Join(", ", hits))
'''.strip().replace("rank(", "order.IndexOf(")


def validate_engine() -> str:
    rule_args = [
        ("InOut", "io_dtResults", "sd:DataTable", "out_dtResults"),
        ("In", "in_Param1", "x:String", 'Convert.ToString(CurrentRule("Param1"))'),
        ("In", "in_Param2", "x:String", 'Convert.ToString(CurrentRule("Param2"))'),
    ]

    def param_with_fallback(n: int, setting: str) -> str:
        # A blank Param cell on the ValidationRules sheet falls back to the matching
        # Settings value, so both places in Config.xlsx are live and never contradict.
        cell = f'Convert.ToString(CurrentRule("Param{n}"))'
        return (f'If(String.IsNullOrWhiteSpace({cell}) AndAlso in_Config IsNot Nothing '
                f'AndAlso in_Config.ContainsKey("{setting}"), '
                f'Convert.ToString(in_Config("{setting}")), {cell})')

    def rule_case(rid: str, fname: str, extra=(), params=None) -> tuple[str, str]:
        args = list(rule_args) if params is None else [rule_args[0]] + list(params)
        args += list(extra)
        return rid, sequence(rid, invoke_workflow(
            f"Workflows\\Rules\\{fname}.xaml", args, name=f"Invoke {rid} - {fname}"))

    cases = [
        rule_case("R1", "R1_TimestampCheck", params=[
            ("In", "in_Param1", "x:String", param_with_fallback(1, "TimestampToleranceHours")),
            ("In", "in_Param2", "x:String", param_with_fallback(2, "TimestampWarnToleranceHours")),
        ]),
        rule_case("R2", "R2_QuantityMatch", params=[
            ("In", "in_Param1", "x:String", param_with_fallback(1, "QuantityTolerance")),
            ("In", "in_Param2", "x:String", 'Convert.ToString(CurrentRule("Param2"))'),
        ]),
        rule_case("R3", "R3_AddressWhitelist",
                  [("In", "in_dtWallets", "sd:DataTable", "in_dtWallets")]),
        rule_case("R4", "R4_DuplicateDetection"),
        rule_case("R5", "R5_SequenceValidation"),
        rule_case("R6", "R6_SmartContractEvent"),
    ]

    dispatch = switch(
        'Convert.ToString(CurrentRule("RuleID")).Trim().ToUpperInvariant()',
        cases,
        sequence("Unknown rule", log(
            '"[VALIDATE] Config.xlsx lists rule " & Convert.ToString(CurrentRule("RuleID")) '
            '& " but no workflow implements it - skipped."', level="Warn")),
        name="Switch - RuleID")

    per_rule = sequence(
        "Apply one configured rule",
        if_('Convert.ToString(CurrentRule("Enabled")).Trim().ToUpperInvariant() = "TRUE"',
            sequence("Enabled",
                     log('"[VALIDATE] Applying " & Convert.ToString(CurrentRule("RuleID")) '
                         '& " - " & Convert.ToString(CurrentRule("RuleName"))')
                     + dispatch),
            sequence("Disabled",
                     log('"[VALIDATE] " & Convert.ToString(CurrentRule("RuleID")) '
                         '& " is disabled in Config.xlsx - skipped."')),
            name="If - rule enabled in Config.xlsx"))

    body = sequence(
        "05 - Validation engine",
        invoke_code(_skeleton_code(), [
            ("In", "in_dtJoined", "sd:DataTable", "in_dtJoined"),
            ("Out", "out_dtResults", "sd:DataTable", "out_dtResults"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - stage results table")
        + log('"[VALIDATE] " & summary')
        + for_each_row("in_dtRules", per_rule, row_var="CurrentRule",
                       name="For Each Row - configured validation rules")
        + invoke_code(_finalise_code(), [
            ("In", "in_dtResults", "sd:DataTable", "out_dtResults"),
            ("In", "in_dtRules", "sd:DataTable", "in_dtRules"),
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("Out", "out_dtExceptions", "sd:DataTable", "out_dtExceptions"),
            ("Out", "out_Stats", DICT_SO, "out_Stats"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - roll up statuses and build exceptions")
        + log('"[VALIDATE] " & summary')
        + if_('in_Config IsNot Nothing AndAlso in_Config.ContainsKey("LogLevel") AndAlso '
              'Convert.ToString(in_Config("LogLevel")).ToUpperInvariant() = "TRACE"',
              for_each_row("out_dtResults",
                           log('"[TRACE] " & Convert.ToString(TraceRow("ShipmentID")) & "/" '
                               '& Convert.ToString(TraceRow("EventName")) & " " '
                               '& Convert.ToString(TraceRow("ValidationStatus")) & " severity=" '
                               '& Convert.ToString(TraceRow("Severity")) & " " '
                               '& Convert.ToString(TraceRow("FailureReasons"))', name="Log - [TRACE] verdict"),
                           row_var="TraceRow", name="For Each Row - trace every verdict"),
              name="If - LogLevel is Trace"),
        variables(("x:String", "summary")))

    return workflow("05_Validate_Engine", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtJoined", "InArgument(sd:DataTable)"),
        ("in_dtRules", "InArgument(sd:DataTable)"),
        ("in_dtWallets", "InArgument(sd:DataTable)"),
        ("out_dtResults", "OutArgument(sd:DataTable)"),
        ("out_dtExceptions", "OutArgument(sd:DataTable)"),
        ("out_Stats", f"OutArgument({DICT_SO})"),
    ])


def _rule(stem: str, display: str, code: str,
          extra_members: list[tuple[str, str]] | None = None,
          extra_args: list[tuple[str, str, str, str]] | None = None) -> str:
    args = [
        ("InOut", "io_dtResults", "sd:DataTable", "io_dtResults"),
        ("In", "in_Param1", "x:String", "in_Param1"),
        ("In", "in_Param2", "x:String", "in_Param2"),
    ] + list(extra_args or [])

    body = sequence(
        display,
        invoke_code(code, args + [("Out", "out_Summary", "x:String", "summary")],
                    name=f"Invoke Code - {display}")
        + log(f'"[{stem[:2]}] " & summary'),
        variables(("x:String", "summary")))

    return workflow(stem, body, members=[
        ("io_dtResults", "InOutArgument(sd:DataTable)"),
        ("in_Param1", "InArgument(x:String)"),
        ("in_Param2", "InArgument(x:String)"),
    ] + list(extra_members or []))


def r1_timestamp() -> str:
    code = r'''
' Drift between the on-chain timestamp and the ERP's expected milestone time.
' Param1 = hours still counted as a PASS, Param2 = ceiling for a WARNING.
' Anything beyond Param2 is a FAIL.
Dim tolPass As Double = 2.0
Dim tolWarn As Double = 4.0
Double.TryParse(in_Param1, tolPass)
Double.TryParse(in_Param2, tolWarn)
If tolWarn < tolPass Then tolWarn = tolPass

Dim nFail As Integer = 0, nWarn As Integer = 0, nSkip As Integer = 0

For Each r As DataRow In io_dtResults.Rows
    If IsDBNull(r("DriftHours")) Then
        r("R1_Status") = "SKIPPED"
        r("R1_Detail") = "No matching ERP milestone to compare against"
        nSkip += 1
        Continue For
    End If

    Dim drift As Double = Convert.ToDouble(r("DriftHours"))
    Dim detail As String = String.Format(CultureInfo.InvariantCulture, _
        "chain {0:yyyy-MM-dd HH:mm} vs ERP {1:yyyy-MM-dd HH:mm}, drift {2:F2}h", _
        Convert.ToDateTime(r("ChainTimestampUtc")), _
        Convert.ToDateTime(r("ErpExpectedTimestampUtc")), drift)

    If drift <= tolPass Then
        r("R1_Status") = "PASS"
        r("R1_Detail") = detail
    ElseIf drift <= tolWarn Then
        r("R1_Status") = "WARNING"
        r("R1_Detail") = detail & String.Format(CultureInfo.InvariantCulture, " (over {0:F1}h tolerance)", tolPass)
        nWarn += 1
    Else
        r("R1_Status") = "FAIL"
        r("R1_Detail") = detail & String.Format(CultureInfo.InvariantCulture, " (over {0:F1}h limit)", tolWarn)
        nFail += 1
    End If
Next

out_Summary = String.Format("timestamp check: {0} fail, {1} warning, {2} skipped (pass<={3}h, warn<={4}h)", _
    nFail, nWarn, nSkip, tolPass, tolWarn)
'''.strip()
    return _rule("R1_TimestampCheck", "R1 - Timestamp consistency", code)


def r2_quantity() -> str:
    code = r'''
' On-chain quantity against the purchase-order quantity. Param1 is the permitted
' absolute difference in units, so a tolerance can be granted without code changes.
Dim tol As Long = 0
Long.TryParse(in_Param1, tol)

Dim nFail As Integer = 0, nSkip As Integer = 0

For Each r As DataRow In io_dtResults.Rows
    If IsDBNull(r("ErpExpectedQty")) Then
        r("R2_Status") = "SKIPPED"
        r("R2_Detail") = "No ERP purchase-order quantity for this shipment"
        nSkip += 1
        Continue For
    End If

    Dim onChain As Long = Convert.ToInt64(r("OnChainQty"))
    Dim expected As Long = Convert.ToInt64(r("ErpExpectedQty"))
    Dim diff As Long = Math.Abs(onChain - expected)

    If diff <= tol Then
        r("R2_Status") = "PASS"
        r("R2_Detail") = String.Format("on-chain {0} = PO {1}", onChain, expected)
    Else
        r("R2_Status") = "FAIL"
        r("R2_Detail") = String.Format("on-chain {0} against PO {1} for {2} - shortfall of {3} unit(s)", _
            onChain, expected, Convert.ToString(r("PONumber")), diff)
        nFail += 1
    End If
Next

out_Summary = String.Format("quantity match: {0} fail, {1} skipped (tolerance {2})", nFail, nSkip, tol)
'''.strip()
    return _rule("R2_QuantityMatch", "R2 - Quantity match", code)


def r3_whitelist() -> str:
    code = r'''
' Sender address must be an approved partner wallet. Addresses were lower-cased
' during extraction, so a plain comparison is safe here - comparing the raw
' EIP-55 checksummed form against the ERP's lower-case form would fail every row.
Dim approved As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
For Each w As DataRow In in_dtWallets.Rows
    If Convert.ToString(w("Active")).Trim().ToUpperInvariant() = "TRUE" Then
        approved(Convert.ToString(w("Address")).Trim().ToLowerInvariant()) = Convert.ToString(w("PartnerName"))
    End If
Next

Dim nFail As Integer = 0

For Each r As DataRow In io_dtResults.Rows
    Dim addr As String = Convert.ToString(r("SenderAddress")).Trim().ToLowerInvariant()
    If approved.ContainsKey(addr) Then
        r("R3_Status") = "PASS"
        r("R3_Detail") = "signed by " & approved(addr)
    Else
        r("R3_Status") = "FAIL"
        r("R3_Detail") = String.Format("{0} is not an approved partner wallet (expected {1})", _
            addr, If(Convert.ToString(r("ErpHandlerPartner")) = "", "an approved partner", Convert.ToString(r("ErpHandlerPartner"))))
        nFail += 1
    End If
Next

out_Summary = String.Format("wallet whitelist: {0} fail against {1} approved partner(s)", nFail, approved.Count)
'''.strip()
    return _rule("R3_AddressWhitelist", "R3 - Approved wallet whitelist", code,
                 extra_members=[("in_dtWallets", "InArgument(sd:DataTable)")],
                 extra_args=[("In", "in_dtWallets", "sd:DataTable", "in_dtWallets")])


def r4_duplicate() -> str:
    code = r'''
' A transaction hash must appear at most once per shipment. Both copies are flagged,
' not just the second: from the ledger alone there is no way to tell which submission
' was the legitimate one, and quietly passing the first would hide that.
Dim counts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
For Each r As DataRow In io_dtResults.Rows
    Dim k As String = Convert.ToString(r("ShipmentID")) & "|" & Convert.ToString(r("TxHash"))
    If Not counts.ContainsKey(k) Then counts(k) = 0
    counts(k) = counts(k) + 1
Next

Dim nFail As Integer = 0

For Each r As DataRow In io_dtResults.Rows
    Dim k As String = Convert.ToString(r("ShipmentID")) & "|" & Convert.ToString(r("TxHash"))
    Dim c As Integer = counts(k)
    If c > 1 Then
        r("R4_Status") = "FAIL"
        ' Shorten for readability, but never assume a minimum length - a malformed or
        ' truncated hash from a bad feed must still be reported, not crash the rule.
        Dim h As String = Convert.ToString(r("TxHash"))
        Dim shortHash As String = If(h.Length > 12, h.Substring(0, 12) & "...", h)
        r("R4_Detail") = String.Format("transaction {0} recorded {1} times for {2} - replayed or double-submitted", _
            shortHash, c, Convert.ToString(r("ShipmentID")))
        nFail += 1
    Else
        r("R4_Status") = "PASS"
        r("R4_Detail") = "unique for this shipment"
    End If
Next

out_Summary = String.Format("duplicate detection: {0} row(s) in duplicate groups", nFail)
'''.strip()
    return _rule("R4_DuplicateDetection", "R4 - Duplicate detection", code)


def r5_sequence() -> str:
    code = r'''
' Milestones must occur in the order defined on the EventSequence sheet. Walk each
' shipment chronologically and flag any event whose stepNo number is lower than the
' highest already seen - that event physically could not follow what preceded it.
Dim byShipment As New Dictionary(Of String, List(Of DataRow))(StringComparer.OrdinalIgnoreCase)
For Each r As DataRow In io_dtResults.Rows
    Dim sid As String = Convert.ToString(r("ShipmentID"))
    If Not byShipment.ContainsKey(sid) Then byShipment(sid) = New List(Of DataRow)
    byShipment(sid).Add(r)
    r("R5_Status") = "PASS"
    r("R5_Detail") = "in order"
Next

Dim nFail As Integer = 0

For Each kv As KeyValuePair(Of String, List(Of DataRow)) In byShipment
    Dim rows As List(Of DataRow) = kv.Value
    rows.Sort(Function(a, b)
                  Dim c As Integer = Convert.ToDateTime(a("ChainTimestampUtc")).CompareTo(Convert.ToDateTime(b("ChainTimestampUtc")))
                  If c <> 0 Then Return c
                  Return Convert.ToInt32(a("RowNo")).CompareTo(Convert.ToInt32(b("RowNo")))
              End Function)

    Dim maxStep As Integer = 0
    Dim maxName As String = ""
    Dim maxAt As DateTime = DateTime.MinValue

    For Each r As DataRow In rows
        Dim stepNo As Integer = Convert.ToInt32(r("StepOrder"))
        If stepNo = 0 Then Continue For
        If stepNo < maxStep Then
            r("R5_Status") = "FAIL"
            r("R5_Detail") = String.Format(CultureInfo.InvariantCulture, _
                "{0} (step {1}) recorded {2:yyyy-MM-dd HH:mm}, after {3} (step {4}) at {5:yyyy-MM-dd HH:mm}", _
                Convert.ToString(r("EventName")), stepNo, Convert.ToDateTime(r("ChainTimestampUtc")), _
                maxName, maxStep, maxAt)
            nFail += 1
        Else
            maxStep = stepNo
            maxName = Convert.ToString(r("EventName"))
            maxAt = Convert.ToDateTime(r("ChainTimestampUtc"))
        End If
    Next
Next

out_Summary = String.Format("sequence validation: {0} out-of-order event(s) across {1} shipment(s)", _
    nFail, byShipment.Count)
'''.strip()
    return _rule("R5_SequenceValidation", "R5 - Milestone sequence", code)


def r6_contract_event() -> str:
    code = r'''
' Every shipment must emit the milestone(s) named in Param1, matched by function
' selector during extraction. A missing event has no row of its own, so the finding
' is attached to the shipment's latest transaction - the point at which the chain of
' custody is left open.
Dim required As New List(Of String)
For Each p As String In in_Param1.Split(","c)
    If p.Trim() <> "" Then required.Add(p.Trim())
Next

Dim byShipment As New Dictionary(Of String, List(Of DataRow))(StringComparer.OrdinalIgnoreCase)
For Each r As DataRow In io_dtResults.Rows
    Dim sid As String = Convert.ToString(r("ShipmentID"))
    If Not byShipment.ContainsKey(sid) Then byShipment(sid) = New List(Of DataRow)
    byShipment(sid).Add(r)
    r("R6_Status") = "PASS"
    r("R6_Detail") = "required event(s) present: " & String.Join(", ", required)
Next

Dim nFail As Integer = 0

For Each kv As KeyValuePair(Of String, List(Of DataRow)) In byShipment
    Dim present As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    For Each r As DataRow In kv.Value
        present.Add(Convert.ToString(r("EventName")))
    Next

    Dim missing As New List(Of String)
    For Each req As String In required
        If Not present.Contains(req) Then missing.Add(req)
    Next

    If missing.Count > 0 Then
        Dim last As DataRow = kv.Value(0)
        For Each r As DataRow In kv.Value
            If Convert.ToDateTime(r("ChainTimestampUtc")) > Convert.ToDateTime(last("ChainTimestampUtc")) Then last = r
        Next
        last("R6_Status") = "FAIL"
        last("R6_Detail") = String.Format("{0} never emitted {1} - chain of custody incomplete after {2}", _
            kv.Key, String.Join(", ", missing), Convert.ToString(last("EventName")))
        nFail += 1
    End If
Next

out_Summary = String.Format("smart contract event: {0} shipment(s) missing {1}", nFail, String.Join(", ", required))
'''.strip()
    return _rule("R6_SmartContractEvent", "R6 - Required smart contract event", code)


# ==========================================================================
# 06_HumanInTheLoop_Escalation  -  Novelty 2
# ==========================================================================
# A CRITICAL anomaly stops the bot and asks a person, rather than being logged and
# forgotten. The decision is captured before the audit log is written, so the human
# judgement becomes part of the tamper-evident record rather than a note beside it.
#
# EscalationMode decides how the question gets asked:
#   PROMPT   - show the reviewer form and wait. Attended runs only. ReviewUi = DIALOG
#              uses a plain Input Dialog instead of the full form.
#   SIMULATE - apply a pre-set decision. For rehearsals and automated tests.
#   AUTO_LOG - record DEFERRED and carry on. The correct behaviour unattended, where
#              there is nobody at the machine and blocking would hang the job.

REVIEW_WINDOW_CODE = r'''
' The reviewer form (Novelty 2). It pauses the bot on the anomaly, shows the reviewer
' every fact behind it, and returns their decision, their name and their reasoning,
' which the workflow records into the results before the tamper-evident audit log is
' written. It is built in code rather than in a designer so it can be compiled and
' checked like the rest of the project; ReviewUi = DIALOG falls back to Input Dialog.
'
' WPF needs a single-threaded apartment, so the window runs on its own STA thread and
' this activity waits for it. A timer closes an unanswered form after
' ReviewTimeoutMinutes, recording DEFERRED rather than leaving the job hanging.
' Local copies: the form's event handlers are lambdas, and a lambda may not capture
' an argument that could be passed by reference.
Dim projectName As String = in_ProjectName
Dim runId As String = in_RunId
Dim botIdentity As String = in_BotIdentity
Dim shipmentId As String = in_ShipmentId
Dim product As String = in_Product
Dim eventName As String = in_EventName
Dim chainTime As String = in_ChainTime
Dim sender As String = in_Sender
Dim expectedPartner As String = in_ExpectedPartner
Dim txHash As String = in_TxHash
Dim severity As String = in_Severity
Dim reasons As String = in_Reasons
Dim options() As String = in_Options
Dim evidencePath As String = in_EvidencePath

Dim decision As String = ""
Dim reviewer As String = ""
Dim note As String = ""
Dim outcome As String = "CLOSED"
Dim uiError As Exception = Nothing
Dim timeoutSeconds As Integer = CInt(Math.Max(2.0, Math.Round(in_TimeoutMinutes * 60.0)))

Dim showForm As System.Threading.ThreadStart = Sub()
    Try
        Dim ink As New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(22, 32, 44))
        Dim muted As New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(104, 119, 138))
        Dim navy As New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(31, 56, 100))
        Dim failBg As New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(253, 236, 234))
        Dim failInk As New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(179, 38, 30))

        Dim win As New System.Windows.Window()
        win.Title = "Human review required - " & projectName
        win.Width = 780
        win.SizeToContent = System.Windows.SizeToContent.Height
        win.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        win.ResizeMode = System.Windows.ResizeMode.NoResize
        win.Topmost = True
        win.Background = New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 246, 249))
        win.FontFamily = New System.Windows.Media.FontFamily("Segoe UI")
        win.FontSize = 13

        Dim root As New System.Windows.Controls.StackPanel()

        ' ---- Header band -------------------------------------------------------
        Dim header As New System.Windows.Controls.Border()
        header.Background = navy
        header.Padding = New System.Windows.Thickness(20, 16, 20, 16)
        Dim headStack As New System.Windows.Controls.StackPanel()
        Dim title As New System.Windows.Controls.TextBlock()
        title.Text = "The bot has paused for your decision"
        title.FontSize = 20
        title.FontWeight = System.Windows.FontWeights.SemiBold
        title.Foreground = System.Windows.Media.Brushes.White
        headStack.Children.Add(title)
        Dim subtitle As New System.Windows.Controls.TextBlock()
        subtitle.Text = String.Format("Run {0} - bot {1}. A blockchain transaction failed a {2}-severity validation rule.", runId, botIdentity, severity)
        subtitle.Foreground = New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(214, 224, 238))
        subtitle.Margin = New System.Windows.Thickness(0, 4, 0, 0)
        subtitle.TextWrapping = System.Windows.TextWrapping.Wrap
        headStack.Children.Add(subtitle)
        header.Child = headStack
        root.Children.Add(header)

        ' ---- Facts ---------------------------------------------------------------
        Dim grid As New System.Windows.Controls.Grid()
        grid.Margin = New System.Windows.Thickness(20, 14, 20, 4)
        Dim c0 As New System.Windows.Controls.ColumnDefinition()
        c0.Width = New System.Windows.GridLength(150)
        Dim c1 As New System.Windows.Controls.ColumnDefinition()
        c1.Width = New System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
        grid.ColumnDefinitions.Add(c0)
        grid.ColumnDefinitions.Add(c1)

        Dim facts As New List(Of String())
        facts.Add(New String() {"Shipment", shipmentId & "  (" & product & ")", ""})
        facts.Add(New String() {"Milestone", eventName, ""})
        facts.Add(New String() {"Recorded on chain", chainTime & " UTC", ""})
        facts.Add(New String() {"Signed by wallet", sender, "mono"})
        facts.Add(New String() {"Expected partner", If(expectedPartner = "", "-", expectedPartner), ""})
        facts.Add(New String() {"Transaction hash", txHash, "mono"})
        facts.Add(New String() {"Severity", severity, ""})

        For i As Integer = 0 To facts.Count - 1
            Dim rd As New System.Windows.Controls.RowDefinition()
            rd.Height = System.Windows.GridLength.Auto
            grid.RowDefinitions.Add(rd)
            Dim lbl As New System.Windows.Controls.TextBlock()
            lbl.Text = facts(i)(0)
            lbl.Foreground = muted
            lbl.Margin = New System.Windows.Thickness(0, 3, 10, 3)
            System.Windows.Controls.Grid.SetRow(lbl, i)
            System.Windows.Controls.Grid.SetColumn(lbl, 0)
            grid.Children.Add(lbl)
            Dim val As New System.Windows.Controls.TextBox()
            val.Text = facts(i)(1)
            val.IsReadOnly = True
            val.BorderThickness = New System.Windows.Thickness(0)
            val.Background = System.Windows.Media.Brushes.Transparent
            val.Foreground = ink
            val.TextWrapping = System.Windows.TextWrapping.Wrap
            val.Margin = New System.Windows.Thickness(0, 3, 0, 3)
            If facts(i)(2) = "mono" Then val.FontFamily = New System.Windows.Media.FontFamily("Consolas")
            System.Windows.Controls.Grid.SetRow(val, i)
            System.Windows.Controls.Grid.SetColumn(val, 1)
            grid.Children.Add(val)
        Next
        root.Children.Add(grid)

        ' ---- Finding -------------------------------------------------------------
        Dim findingBox As New System.Windows.Controls.Border()
        findingBox.Background = failBg
        findingBox.BorderBrush = New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(243, 194, 190))
        findingBox.BorderThickness = New System.Windows.Thickness(1)
        findingBox.CornerRadius = New System.Windows.CornerRadius(6)
        findingBox.Padding = New System.Windows.Thickness(12, 10, 12, 10)
        findingBox.Margin = New System.Windows.Thickness(20, 8, 20, 4)
        Dim finding As New System.Windows.Controls.TextBlock()
        finding.Text = "Finding: " & reasons.Replace(" | ", Environment.NewLine)
        finding.Foreground = failInk
        finding.TextWrapping = System.Windows.TextWrapping.Wrap
        findingBox.Child = finding
        root.Children.Add(findingBox)

        ' ---- Decision --------------------------------------------------------------
        Dim decisionLabel As New System.Windows.Controls.TextBlock()
        decisionLabel.Text = "Your decision"
        decisionLabel.FontWeight = System.Windows.FontWeights.SemiBold
        decisionLabel.Margin = New System.Windows.Thickness(20, 12, 20, 4)
        root.Children.Add(decisionLabel)

        Dim radios As New List(Of System.Windows.Controls.RadioButton)
        For Each opt As String In options
            Dim rb As New System.Windows.Controls.RadioButton()
            rb.Content = opt
            rb.GroupName = "decision"
            rb.Margin = New System.Windows.Thickness(24, 3, 20, 3)
            radios.Add(rb)
            root.Children.Add(rb)
        Next

        Dim reviewerLabel As New System.Windows.Controls.TextBlock()
        reviewerLabel.Text = "Reviewer name"
        reviewerLabel.FontWeight = System.Windows.FontWeights.SemiBold
        reviewerLabel.Margin = New System.Windows.Thickness(20, 12, 20, 4)
        root.Children.Add(reviewerLabel)
        Dim reviewerBox As New System.Windows.Controls.TextBox()
        reviewerBox.Text = Environment.UserName
        reviewerBox.Margin = New System.Windows.Thickness(20, 0, 20, 0)
        reviewerBox.Padding = New System.Windows.Thickness(4)
        root.Children.Add(reviewerBox)

        Dim noteLabel As New System.Windows.Controls.TextBlock()
        noteLabel.Text = "Reason / note (recorded in the audit trail)"
        noteLabel.FontWeight = System.Windows.FontWeights.SemiBold
        noteLabel.Margin = New System.Windows.Thickness(20, 12, 20, 4)
        root.Children.Add(noteLabel)
        Dim noteBox As New System.Windows.Controls.TextBox()
        noteBox.Height = 60
        noteBox.TextWrapping = System.Windows.TextWrapping.Wrap
        noteBox.AcceptsReturn = False
        noteBox.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        noteBox.Margin = New System.Windows.Thickness(20, 0, 20, 0)
        noteBox.Padding = New System.Windows.Thickness(4)
        root.Children.Add(noteBox)

        ' ---- Footer: countdown + submit ------------------------------------------
        Dim footer As New System.Windows.Controls.DockPanel()
        footer.Margin = New System.Windows.Thickness(20, 14, 20, 18)
        Dim submit As New System.Windows.Controls.Button()
        submit.Content = "Record decision and resume"
        submit.Padding = New System.Windows.Thickness(16, 6, 16, 6)
        submit.IsEnabled = False
        submit.IsDefault = True
        System.Windows.Controls.DockPanel.SetDock(submit, System.Windows.Controls.Dock.Right)
        footer.Children.Add(submit)
        Dim countdown As New System.Windows.Controls.TextBlock()
        countdown.Foreground = muted
        countdown.VerticalAlignment = System.Windows.VerticalAlignment.Center
        footer.Children.Add(countdown)
        root.Children.Add(footer)

        Dim remaining As Integer = timeoutSeconds
        Dim renderCountdown As Action = Sub()
            countdown.Text = String.Format("Closes automatically in {0}:{1:00} and records DEFERRED.", remaining \ 60, remaining Mod 60)
        End Sub
        renderCountdown()

        Dim refreshSubmit As Action = Sub()
            submit.IsEnabled = radios.Any(Function(x) x.IsChecked.GetValueOrDefault()) AndAlso reviewerBox.Text.Trim() <> ""
        End Sub
        For Each rb As System.Windows.Controls.RadioButton In radios
            AddHandler rb.Checked, Sub(s, e) refreshSubmit()
        Next
        AddHandler reviewerBox.TextChanged, Sub(s, e) refreshSubmit()

        AddHandler submit.Click, Sub(s, e)
            Dim chosen As System.Windows.Controls.RadioButton = radios.First(Function(x) x.IsChecked.GetValueOrDefault())
            decision = Convert.ToString(chosen.Content)
            reviewer = reviewerBox.Text.Trim()
            note = noteBox.Text.Trim()
            outcome = "SUBMITTED"
            win.Close()
        End Sub

        Dim timer As New System.Windows.Threading.DispatcherTimer()
        timer.Interval = TimeSpan.FromSeconds(1)
        AddHandler timer.Tick, Sub(s, e)
            remaining -= 1
            If remaining <= 0 Then
                timer.Stop()
                If outcome <> "SUBMITTED" Then outcome = "TIMEOUT"
                win.Close()
            Else
                renderCountdown()
            End If
        End Sub

        ' Keep an image of exactly what the reviewer was shown, next to the audit trail.
        AddHandler win.ContentRendered, Sub(s, e)
            Try
                If evidencePath <> "" Then
                    Dim el As System.Windows.FrameworkElement = CType(win.Content, System.Windows.FrameworkElement)
                    Dim bmp As New System.Windows.Media.Imaging.RenderTargetBitmap( _
                        CInt(Math.Ceiling(el.ActualWidth)), CInt(Math.Ceiling(el.ActualHeight)), 96, 96, _
                        System.Windows.Media.PixelFormats.Pbgra32)
                    bmp.Render(el)
                    Dim enc As New System.Windows.Media.Imaging.PngBitmapEncoder()
                    enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bmp))
                    Directory.CreateDirectory(Path.GetDirectoryName(evidencePath))
                    Using fs As New FileStream(evidencePath, FileMode.Create)
                        enc.Save(fs)
                    End Using
                End If
            Catch
                ' Evidence capture must never block the review itself.
            End Try
            timer.Start()
        End Sub

        win.Content = root
        win.ShowDialog()
        timer.Stop()
    Catch ex As Exception
        uiError = ex
    End Try
End Sub

Dim uiThread As New System.Threading.Thread(showForm)
uiThread.SetApartmentState(System.Threading.ApartmentState.STA)
uiThread.IsBackground = True
uiThread.Start()
uiThread.Join()

If uiError IsNot Nothing Then
    Throw New InvalidOperationException("The review window could not be shown (" & uiError.Message & _
        "). On a machine without an interactive desktop set EscalationMode = AUTO_LOG, or ReviewUi = DIALOG.", uiError)
End If

Dim minutesText As String = in_TimeoutMinutes.ToString("0.##", CultureInfo.InvariantCulture)
Select Case outcome
    Case "SUBMITTED"
        Dim cleanNote As String = note.Replace(vbCr, " ").Replace(vbLf, " ")
        out_Decision = If(cleanNote = "", decision, decision & " [note: " & cleanNote & "]")
        out_DecidedBy = reviewer
    Case "TIMEOUT"
        out_Decision = "DEFERRED - no reviewer response within " & minutesText & " minute(s)"
        out_DecidedBy = "TIMEOUT"
    Case Else
        out_Decision = "DEFERRED - review window closed without a decision"
        out_DecidedBy = Environment.UserName
End Select
out_Outcome = outcome
'''.strip()


def human_in_the_loop() -> str:
    count_code = r'''
Dim n As Integer = 0
Dim detail As New List(Of String)
For Each r As DataRow In io_dtResults.Rows
    If Convert.ToBoolean(r("RequiresEscalation")) Then
        n += 1
        detail.Add(Convert.ToString(r("ShipmentID")) & "/" & Convert.ToString(r("EventName")))
    End If
Next
out_Count = n
out_Summary = If(n = 0, "no anomalies met the escalation threshold", _
                 n.ToString() & " anomaly(ies) require human review: " & String.Join(", ", detail))
'''.strip()

    apply_code = r'''
' Non-interactive dispositions. SIMULATE applies the decision configured in
' Config.xlsx so a rehearsal produces the same artefacts as a real review;
' AUTO_LOG defers, which is what an unattended run must do.
Dim decision As String
Dim decidedBy As String
If in_Mode = "SIMULATE" Then
    decision = in_SimulatedDecision
    If String.IsNullOrWhiteSpace(decision) Then decision = "APPROVED - simulated"
    decidedBy = "SIMULATED (" & Environment.UserName & ")"
Else
    decision = "DEFERRED - unattended run, no reviewer present"
    decidedBy = "UNATTENDED"
End If

Dim n As Integer = 0
For Each r As DataRow In io_dtResults.Rows
    If Convert.ToBoolean(r("RequiresEscalation")) Then
        r("HumanDecision") = decision
        r("HumanDecidedBy") = decidedBy
        r("HumanDecidedAtUtc") = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
        n += 1
    End If
Next

out_Summary = String.Format("{0} escalation(s) dispositioned as ""{1}"" via {2}", n, decision, in_Mode)
'''.strip()

    # Interactive branch: iterate the results in place so the reviewer's answer lands
    # on the real row rather than on a filtered copy.
    def cfg(key: str) -> str:
        return f'Convert.ToString(in_Config("{key}"))'

    window_review = sequence(
        "Review form",
        assign("evidencePath",
               'Path.Combine(' + cfg("ScreenshotFolder") + ', "Review_" & ' + cfg("RunId") + ' & "_" '
               '& Convert.ToString(CurrentRow("ShipmentID")) & "_" & Convert.ToString(CurrentRow("EventName")) & ".png")',
               name="Assign - evidence image path")
        + invoke_code(REVIEW_WINDOW_CODE, [
            ("In", "in_ProjectName", "x:String", cfg("ProjectName")),
            ("In", "in_RunId", "x:String", cfg("RunId")),
            ("In", "in_BotIdentity", "x:String", cfg("BotIdentity")),
            ("In", "in_ShipmentId", "x:String", 'Convert.ToString(CurrentRow("ShipmentID"))'),
            ("In", "in_Product", "x:String", 'Convert.ToString(CurrentRow("Product"))'),
            ("In", "in_EventName", "x:String", 'Convert.ToString(CurrentRow("EventName"))'),
            ("In", "in_ChainTime", "x:String",
             'Convert.ToDateTime(CurrentRow("ChainTimestampUtc")).ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)'),
            ("In", "in_Sender", "x:String", 'Convert.ToString(CurrentRow("SenderAddress"))'),
            ("In", "in_ExpectedPartner", "x:String", 'Convert.ToString(CurrentRow("ErpHandlerPartner"))'),
            ("In", "in_TxHash", "x:String", 'Convert.ToString(CurrentRow("TxHash"))'),
            ("In", "in_Severity", "x:String", 'Convert.ToString(CurrentRow("Severity"))'),
            ("In", "in_Reasons", "x:String", 'Convert.ToString(CurrentRow("FailureReasons"))'),
            ("In", "in_Options", "s:String[]", HITL_OPTIONS_EXPR),
            ("In", "in_TimeoutMinutes", "x:Double",
             'Double.Parse(' + cfg("ReviewTimeoutMinutes") + ', CultureInfo.InvariantCulture)'),
            ("In", "in_EvidencePath", "x:String", "evidencePath"),
            ("Out", "out_Decision", "x:String", "decision"),
            ("Out", "out_DecidedBy", "x:String", "decidedBy"),
            ("Out", "out_Outcome", "x:String", "reviewOutcome"),
        ], name="Invoke Code - show the review form and wait")
        + log('"[ESCALATE] Review form " & reviewOutcome & " - evidence image " & evidencePath'))

    dialog_review = sequence(
        "Input Dialog (fallback)",
        f'<ui:InputDialog DisplayName="{lit("Input Dialog - anomaly review")}" '
        f'Title="{vb(chr(34) + "Blockchain anomaly requires review" + chr(34))}" '
        f'Label="{vb(HITL_LABEL_EXPR)}" '
        f'Options="{vb(HITL_OPTIONS_EXPR)}" '
        # No IsPassword: InputDialog rejects it alongside Options.
        # Result accepts ONLY property-element syntax with an explicit
        # x:TypeArguments - as a plain attribute it throws at XAML load,
        # whatever type the bound variable has.
        f'><ui:InputDialog.Result>'
        f'<OutArgument x:TypeArguments="x:String">{vb("decision")}</OutArgument>'
        f'</ui:InputDialog.Result></ui:InputDialog>'
        + assign("decidedBy", "Environment.UserName", name="Assign - reviewer"))

    prompt_body = sequence(
        "Ask the reviewer",
        if_('Convert.ToBoolean(CurrentRow("RequiresEscalation"))',
            sequence(
                "Escalate this transaction",
                if_(cfg("ReviewUi") + ' = "DIALOG"', dialog_review, window_review,
                    name="If - ReviewUi is DIALOG")
                + assign('CurrentRow("HumanDecision")', "decision",
                         type_ref="x:Object", name="Assign - record decision")
                + assign('CurrentRow("HumanDecidedBy")', "decidedBy",
                         type_ref="x:Object", name="Assign - record reviewer")
                + assign('CurrentRow("HumanDecidedAtUtc")',
                         'DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)',
                         type_ref="x:Object", name="Assign - record decision time")
                + log('"[ESCALATE] " & Convert.ToString(CurrentRow("ShipmentID")) '
                      '& " -> " & decision & " (by " & decidedBy & ")"')),
            name="If - this row needs review"))

    non_interactive = sequence(
        "Disposition without a reviewer",
        invoke_code(apply_code, [
            ("InOut", "io_dtResults", "sd:DataTable", "io_dtResults"),
            ("In", "in_Mode", "x:String", "mode"),
            ("In", "in_SimulatedDecision", "x:String",
             'Convert.ToString(in_Config("SimulatedEscalationDecision"))'),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - apply non-interactive decision")
        + log('"[ESCALATE] " & summary'))

    handle = if_(
        'mode = "PROMPT"',
        sequence("Interactive review",
                 log('"[ESCALATE] Pausing for human review - " & escalationCount.ToString() & " item(s)."', level="Warn")
                 + for_each_row("io_dtResults", prompt_body, row_var="CurrentRow",
                                name="For Each Row - transactions awaiting review")),
        non_interactive,
        name="If - interactive review available")

    body = sequence(
        "06 - Human-in-the-loop escalation",
        invoke_code(count_code, [
            ("InOut", "io_dtResults", "sd:DataTable", "io_dtResults"),
            ("Out", "out_Count", "x:Int32", "escalationCount"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - count escalations")
        + log('"[ESCALATE] " & summary')
        # PROMPT is downgraded to AUTO_LOG whenever escalation is switched off or the
        # run is unattended, so an unattended job can never block on a dialog nobody
        # is there to answer.
        + assign("mode",
                 'If(Convert.ToString(in_Config("EscalationEnabled")).Trim().ToUpperInvariant() <> "TRUE", "AUTO_LOG", '
                 'If(Convert.ToString(in_Config("AttendedMode")).Trim().ToUpperInvariant() <> "TRUE" '
                 'AndAlso Convert.ToString(in_Config("EscalationMode")).Trim().ToUpperInvariant() = "PROMPT", "AUTO_LOG", '
                 'Convert.ToString(in_Config("EscalationMode")).Trim().ToUpperInvariant()))',
                 name="Assign - effective escalation mode")
        + if_("escalationCount > 0",
              sequence("Handle escalations",
                       log('"[ESCALATE] Effective mode: " & mode')
                       + handle),
              sequence("Nothing to escalate",
                       log('"[ESCALATE] No human review required for this run."')),
              name="If - anything to escalate")
        # The count lived only in a local until now, so every caller that bound this
        # output argument silently received 0. The run summary quotes the validation
        # engine's own tally, which is why the published numbers were right and this
        # stayed invisible - but the argument still has to mean what it says.
        + assign("out_EscalationCount", "escalationCount", type_ref="x:Int32",
                 name="Assign - report how many items were escalated"),
        variables(("x:Int32", "escalationCount"), ("x:String", "summary"),
                  ("x:String", "mode"), ("x:String", "decision"),
                  ("x:String", "decidedBy"), ("x:String", "reviewOutcome"),
                  ("x:String", "evidencePath")))

    return workflow("06_HumanInTheLoop_Escalation", body,
                    members=[
                        ("in_Config", f"InArgument({DICT_SO})"),
                        ("io_dtResults", "InOutArgument(sd:DataTable)"),
                        ("out_EscalationCount", "OutArgument(x:Int32)"),
                    ],
                    extra_refs=["System.Xaml", "PresentationFramework", "PresentationCore", "WindowsBase"])


# ==========================================================================
# 10_AuditLog_Chained  -  Novelty 4
# ==========================================================================
# Each entry carries the hash of the entry before it, so the log forms a chain.
# Altering any historical row changes its RecordHash, which changes its EntryHash,
# which breaks every link after it - the tampering cannot be patched up without
# rewriting the entire remainder of the file.
#
# CSV column order is deliberate: every hash-relevant field is comma-free and comes
# first, and the one free-text field (FailureReasons) is quoted and last. A plain
# Split(","c) therefore recovers fields 0..17 correctly no matter what the free text
# contains, and the remainder rejoins to give field 18. No CSV state machine needed.
#
# Each entry carries what the deck's Novelty 4 lists: bot identity, timestamp, a hash of
# the transaction's input data (InputDataHash) and the result - all fixed by RecordHash -
# plus the hash of the previous entry (PrevHash).
AUDIT_HEADER = ("Seq,TimestampUtc,RunId,BotIdentity,MachineName,TxHash,ShipmentID,EventName,"
                "ChainTimestampUtc,OnChainQty,ValidationStatus,Severity,HumanDecision,"
                "HumanDecidedBy,RecordHash,PrevHash,EntryHash,InputDataHash,FailureReasons")
# The first release's layout, still verifiable so archived logs stay provable.
AUDIT_HEADER_V1 = ("Seq,TimestampUtc,RunId,BotIdentity,MachineName,TxHash,ShipmentID,EventName,"
                   "ChainTimestampUtc,OnChainQty,ValidationStatus,Severity,HumanDecision,"
                   "HumanDecidedBy,RecordHash,PrevHash,EntryHash,FailureReasons")


def audit_log_chained() -> str:
    code = r'''
Dim folder As String = Convert.ToString(in_Config("AuditLogFolder"))
Dim outPath As String = Path.Combine(folder, Convert.ToString(in_Config("AuditLogFile")))
Directory.CreateDirectory(folder)

Dim GENESIS As String = New String("0"c, 64)
Dim prevHash As String = GENESIS
Dim seq As Long = 0
Dim HEADER As String = "''' + AUDIT_HEADER + r'''"
Dim rotated As String = ""

' A log written in an older column layout cannot be extended without mixing formats in
' one chain. It is set aside untouched - still verifiable on its own - and a new chain
' starts from genesis. Nothing is ever deleted or rewritten.
If File.Exists(outPath) Then
    Dim firstLine As String = ""
    Using rd As New StreamReader(outPath)
        firstLine = rd.ReadLine()
    End Using
    If firstLine IsNot Nothing AndAlso firstLine.TrimStart(ChrW(&HFEFF)) <> HEADER Then
        rotated = Path.Combine(folder, Path.GetFileNameWithoutExtension(outPath) & ".legacy-" & _
            DateTime.UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture) & Path.GetExtension(outPath))
        File.Move(outPath, rotated)
    End If
End If

' Continue an existing chain rather than starting a fresh one - the audit trail is
' append-only and spans every run of the bot.
If File.Exists(outPath) Then
    Dim lines() As String = File.ReadAllLines(outPath)
    If lines.Length > 1 Then
        Dim last As String = lines(lines.Length - 1)
        If last.Trim() <> "" Then
            ' Fields 0..16 are comma-free by construction, so Split is safe for them;
            ' only FailureReasons (17, quoted) may contain commas. EntryHash is 16 -
            ' it MUST track the column order above, or every run resumes the chain from
            ' the wrong value and the links break at each run boundary.
            Dim p() As String = last.Split(","c)
            seq = Convert.ToInt64(p(0))
            prevHash = p(16).Trim()
        End If
    End If
End If

Dim sha As SHA256 = SHA256.Create()
Dim outLines As New List(Of String)
If Not File.Exists(outPath) Then outLines.Add(HEADER)

Dim runId As String = Convert.ToString(in_Config("RunId"))
Dim bot As String = Convert.ToString(in_Config("BotIdentity"))
Dim machine As String = Convert.ToString(in_Config("MachineName"))
Dim hasInputHash As Boolean = in_dtResults.Columns.Contains("InputDataHash")

For Each r As DataRow In in_dtResults.Rows
    seq += 1
    Dim stamp As String = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
    Dim chainTs As String = Convert.ToDateTime(r("ChainTimestampUtc")).ToString("s", CultureInfo.InvariantCulture)
    Dim reasons As String = Convert.ToString(r("FailureReasons"))
    ' Blank placeholders keep every hash-relevant field comma-free and non-empty.
    ' Line breaks would split an entry across lines; commas would shift the fields.
    Dim humanDecision As String = Convert.ToString(r("HumanDecision")).Replace(",", ";").Replace(vbCr, " ").Replace(vbLf, " ")
    Dim humanBy As String = Convert.ToString(r("HumanDecidedBy")).Replace(",", ";").Replace(vbCr, " ").Replace(vbLf, " ")
    If humanDecision = "" Then humanDecision = "-"
    If humanBy = "" Then humanBy = "-"
    reasons = reasons.Replace(vbCr, " ").Replace(vbLf, " ")
    ' SHA-256 of the transaction's raw call data, computed by the validation engine.
    Dim inputHash As String = "-"
    If hasInputHash Then inputHash = Convert.ToString(r("InputDataHash")).Trim()
    If inputHash = "" Then inputHash = "-"

    ' RecordHash fixes the verdict and the evidence behind it.
    Dim recordInput As String = String.Join("|", New String() { _
        runId, Convert.ToString(r("TxHash")), Convert.ToString(r("ShipmentID")), _
        Convert.ToString(r("EventName")), chainTs, Convert.ToString(r("OnChainQty")), inputHash, _
        Convert.ToString(r("ValidationStatus")), Convert.ToString(r("Severity")), reasons, _
        humanDecision, humanBy})
    ' Hash the SANITISED values, not the raw ones: the verifier can only ever see what
    ' was written to the file, so hashing anything else guarantees a false mismatch.
    Dim recordHash As String = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(recordInput))).Replace("-", "").ToLowerInvariant()

    ' EntryHash chains this entry to the one before it.
    Dim entryInput As String = String.Join("|", New String() { _
        seq.ToString(), stamp, bot, machine, recordHash, prevHash})
    Dim entryHash As String = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(entryInput))).Replace("-", "").ToLowerInvariant()

    Dim quoted As String = """" & reasons.Replace("""", """""") & """"
    outLines.Add(String.Join(",", New String() { _
        seq.ToString(), stamp, runId, bot, machine, _
        Convert.ToString(r("TxHash")), Convert.ToString(r("ShipmentID")), _
        Convert.ToString(r("EventName")), chainTs, Convert.ToString(r("OnChainQty")), _
        Convert.ToString(r("ValidationStatus")), Convert.ToString(r("Severity")), _
        humanDecision, humanBy, recordHash, prevHash, entryHash, inputHash, quoted}))

    prevHash = entryHash
Next

File.AppendAllLines(outPath, outLines, Encoding.UTF8)

out_AuditPath = outPath
out_HeadHash = prevHash
out_Summary = String.Format("{0} entries appended, chain head {1}...", in_dtResults.Rows.Count, prevHash.Substring(0, 16))
If rotated <> "" Then
    out_Summary &= " | an older-format log was set aside as " & Path.GetFileName(rotated) & " and a new chain started"
End If
'''.strip()

    body = sequence(
        "10 - Append tamper-evident audit log",
        invoke_code(code, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtResults", "sd:DataTable", "in_dtResults"),
            ("Out", "out_AuditPath", "x:String", "out_AuditPath"),
            ("Out", "out_HeadHash", "x:String", "out_HeadHash"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - hash-chain the validation results")
        + log('"[AUDIT] " & summary'),
        variables(("x:String", "summary")))

    return workflow("10_AuditLog_Chained", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtResults", "InArgument(sd:DataTable)"),
        ("out_AuditPath", "OutArgument(x:String)"),
        ("out_HeadHash", "OutArgument(x:String)"),
    ])


# ==========================================================================
# 12_Verify_AuditChain
# ==========================================================================
def verify_audit_chain() -> str:
    code = r'''
' Independently recompute every hash in the log and re-walk the chain. This is the
' demonstration that the audit trail is tamper-evident rather than merely tamper-
' resistant: edit any cell of the CSV and this reports the exact row that broke.
If Not File.Exists(in_AuditPath) Then
    out_IsValid = False
    out_FirstBadRow = 0
    out_Report = "Audit log not found: " & in_AuditPath
    Exit Sub
End If

Dim lines() As String = File.ReadAllLines(in_AuditPath)
Dim sha As SHA256 = SHA256.Create()
Dim GENESIS As String = New String("0"c, 64)

' Current layout carries InputDataHash (field 17); the first release's did not. Both
' remain verifiable, so an archived log can still be proven years later.
Dim headerLine As String = If(lines.Length > 0, lines(0).TrimStart(ChrW(&HFEFF)), "")
Dim v2 As Boolean
If headerLine = "''' + AUDIT_HEADER + r'''" Then
    v2 = True
ElseIf headerLine = "''' + AUDIT_HEADER_V1 + r'''" Then
    v2 = False
Else
    out_IsValid = False
    out_FirstBadRow = 1
    out_Report = "INVALID - first problem at line 1. Unrecognised audit log header."
    Exit Sub
End If
Dim reasonsField As Integer = If(v2, 18, 17)
Dim expectedPrev As String = GENESIS
Dim expectedSeq As Long = 0
Dim bad As Integer = 0
Dim problems As New List(Of String)
Dim checked As Integer = 0

For i As Integer = 1 To lines.Length - 1
    Dim line As String = lines(i)
    If line.Trim() = "" Then Continue For
    checked += 1

    Dim p() As String = line.Split(","c)
    If p.Length < reasonsField + 1 Then
        problems.Add("line " & (i + 1).ToString() & ": malformed (only " & p.Length.ToString() & " fields)")
        If bad = 0 Then bad = i + 1
        Continue For
    End If

    ' Fields before FailureReasons are comma-free by construction; FailureReasons is the
    ' quoted free-text tail, which is rejoined and unescaped here.
    Dim seq As String = p(0)
    Dim stamp As String = p(1)
    Dim runId As String = p(2)
    Dim bot As String = p(3)
    Dim machine As String = p(4)
    Dim txHash As String = p(5)
    Dim shipment As String = p(6)
    Dim evName As String = p(7)
    Dim chainTs As String = p(8)
    Dim qty As String = p(9)
    Dim status As String = p(10)
    Dim severity As String = p(11)
    Dim humanDecision As String = p(12)
    Dim humanBy As String = p(13)
    Dim recordHash As String = p(14)
    Dim prevHash As String = p(15)
    Dim entryHash As String = p(16)
    Dim inputHash As String = If(v2, p(17), "")

    Dim reasons As String = String.Join(",", p, reasonsField, p.Length - reasonsField)
    If reasons.StartsWith("""") AndAlso reasons.EndsWith("""") AndAlso reasons.Length >= 2 Then
        reasons = reasons.Substring(1, reasons.Length - 2).Replace("""""", """")
    End If

    expectedSeq += 1
    If Convert.ToInt64(seq) <> expectedSeq Then
        problems.Add("line " & (i + 1).ToString() & ": sequence jumped, expected " & expectedSeq.ToString() & " found " & seq)
        If bad = 0 Then bad = i + 1
    End If

    Dim recordFields() As String
    If v2 Then
        recordFields = New String() {runId, txHash, shipment, evName, chainTs, qty, inputHash, status, severity, reasons, humanDecision, humanBy}
    Else
        recordFields = New String() {runId, txHash, shipment, evName, chainTs, qty, status, severity, reasons, humanDecision, humanBy}
    End If
    Dim recomputedRecord As String = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes( _
        String.Join("|", recordFields)))).Replace("-", "").ToLowerInvariant()

    If recomputedRecord <> recordHash Then
        problems.Add("line " & (i + 1).ToString() & ": record contents were altered (RecordHash mismatch) - " & shipment & "/" & evName)
        If bad = 0 Then bad = i + 1
    End If

    If prevHash <> expectedPrev Then
        problems.Add("line " & (i + 1).ToString() & ": broken link, PrevHash does not match the previous entry")
        If bad = 0 Then bad = i + 1
    End If

    Dim recomputedEntry As String = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes( _
        String.Join("|", New String() {seq, stamp, bot, machine, recomputedRecord, prevHash}) _
        ))).Replace("-", "").ToLowerInvariant()

    If recomputedEntry <> entryHash Then
        problems.Add("line " & (i + 1).ToString() & ": EntryHash does not match its own contents")
        If bad = 0 Then bad = i + 1
    End If

    expectedPrev = entryHash
Next

out_IsValid = (bad = 0)
out_FirstBadRow = bad
If bad = 0 Then
    out_Report = String.Format("VALID - {0} entries verified, chain intact from genesis to head {1}...", _
        checked, expectedPrev.Substring(0, 16))
Else
    out_Report = String.Format("INVALID - first problem at line {0}. {1}", bad, String.Join(" ; ", problems.Take(5)))
End If
'''.strip()

    # Runnable with no arguments so it doubles as a standalone compliance check: an
    # auditor can verify the trail without running the validation bot at all.
    resolve_default = sequence(
        "Resolve the audit log from configuration",
        assign("projectRoot", "Directory.GetCurrentDirectory()")
        + assign("configPath", 'Path.Combine(projectRoot, "Data", "Config.xlsx")')
        + invoke_workflow("Workflows\\00_Init_ReadConfig.xaml", [
            ("In", "in_ConfigPath", "x:String", "configPath"),
            ("In", "in_ProjectRoot", "x:String", "projectRoot"),
            ("Out", "out_Config", DICT_SO, "Config"),
        ], name="Invoke 00 - read configuration")
        + assign("auditPath",
                 'Path.Combine(Convert.ToString(Config("AuditLogFolder")), '
                 'Convert.ToString(Config("AuditLogFile")))'))

    body = sequence(
        "12 - Verify audit chain integrity",
        if_("String.IsNullOrWhiteSpace(in_AuditPath)",
            resolve_default,
            sequence("Use the supplied path", assign("auditPath", "in_AuditPath")),
            name="If - no audit path supplied")
        + log('"[VERIFY] Checking " & auditPath')
        + invoke_code(code, [
            ("In", "in_AuditPath", "x:String", "auditPath"),
            ("Out", "out_IsValid", "x:Boolean", "out_IsValid"),
            ("Out", "out_FirstBadRow", "x:Int32", "out_FirstBadRow"),
            ("Out", "out_Report", "x:String", "out_Report"),
        ], name="Invoke Code - recompute and re-walk the hash chain")
        + if_("out_IsValid",
              log('"[VERIFY] " & out_Report'),
              log('"[VERIFY] " & out_Report', level="Error"),
              name="If - chain intact"),
        variables(("x:String", "projectRoot"), ("x:String", "configPath"),
                  ("x:String", "auditPath"), (DICT_SO, "Config")))

    return workflow("12_Verify_AuditChain", body, members=[
        ("in_AuditPath", "InArgument(x:String)"),
        ("out_IsValid", "OutArgument(x:Boolean)"),
        ("out_FirstBadRow", "OutArgument(x:Int32)"),
        ("out_Report", "OutArgument(x:String)"),
    ])


# ==========================================================================
# 07_Report_Excel
# ==========================================================================
# Written with ClosedXML rather than Excel activities: UiPath exposes no conditional-
# formatting activity, and ClosedXML needs no Excel process, so the report is produced
# identically on a machine with no Office installed.
def report_excel() -> str:
    code = r'''
Dim folder As String = Convert.ToString(in_Config("ReportsFolder"))
Directory.CreateDirectory(folder)
Dim runId As String = Convert.ToString(in_Config("RunId"))
Dim outPath As String = Path.Combine(folder, "ValidationReport_" & runId & ".xlsx")

Dim GREEN As XLColor = XLColor.FromHtml("#C6EFCE")
Dim RED As XLColor = XLColor.FromHtml("#FFC7CE")
Dim AMBER As XLColor = XLColor.FromHtml("#FFEB9C")
Dim HDR As XLColor = XLColor.FromHtml("#1F3864")

Dim wb As New XLWorkbook()

' ---------- Summary -------------------------------------------------------
Dim ws = wb.Worksheets.Add("Summary")
ws.Cell(1, 1).Value = "Blockchain Transaction Validation Report - " & Convert.ToString(in_Config("ProjectName"))
ws.Cell(1, 1).Style.Font.SetBold(True).Font.SetFontSize(16)

Dim rows As New List(Of String())
rows.Add(New String() {"Run ID", runId})
rows.Add(New String() {"Bot identity", Convert.ToString(in_Config("BotIdentity"))})
rows.Add(New String() {"Machine", Convert.ToString(in_Config("MachineName"))})
rows.Add(New String() {"Generated (UTC)", DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss")})
rows.Add(New String() {"Data source mode", Convert.ToString(in_Config("DataSourceMode"))})
rows.Add(New String() {"Chain ID", Convert.ToString(in_Config("ChainId"))})
rows.Add(New String() {"Contract audited", Convert.ToString(in_Config("ContractAddress"))})
rows.Add(New String() {"", ""})
rows.Add(New String() {"Transactions validated", Convert.ToString(in_Stats("Total"))})
rows.Add(New String() {"Passed", Convert.ToString(in_Stats("Pass"))})
rows.Add(New String() {"Warnings", Convert.ToString(in_Stats("Warning"))})
rows.Add(New String() {"Failed", Convert.ToString(in_Stats("Fail"))})
rows.Add(New String() {"Pass rate (%)", Convert.ToString(in_Stats("PassPercent"))})
rows.Add(New String() {"Escalated to human review", Convert.ToString(in_Stats("Escalations"))})
rows.Add(New String() {"", ""})
rows.Add(New String() {"Rule findings", Convert.ToString(in_Stats("RuleHitSummary"))})
' Appended after the counters so the coloured cells above keep their fixed rows.
rows.Add(New String() {"", ""})
rows.Add(New String() {"Field mapping profile", Convert.ToString(in_Config("FieldMappingSheet"))})
rows.Add(New String() {"Shipment mapping", Convert.ToString(in_Config("MappingMode"))})
rows.Add(New String() {"ERP / WMS source", Convert.ToString(in_Config("LogisticsSourceMode"))})
rows.Add(New String() {"Configuration", Convert.ToString(in_Config("ConfigSource"))})

Dim rIdx As Integer = 3
For Each kv As String() In rows
    ws.Cell(rIdx, 1).Value = kv(0)
    ws.Cell(rIdx, 1).Style.Font.SetBold(True)
    ws.Cell(rIdx, 2).Value = kv(1)
    rIdx += 1
Next

' Colour the headline counters so the outcome reads at a glance.
ws.Cell(11, 2).Style.Fill.SetBackgroundColor(GREEN)
ws.Cell(12, 2).Style.Fill.SetBackgroundColor(AMBER)
ws.Cell(13, 2).Style.Fill.SetBackgroundColor(RED)
ws.Columns(1, 2).AdjustToContents()

' ---------- Per-rule breakdown -------------------------------------------
Dim br As Integer = rIdx + 1
ws.Cell(br, 1).Value = "Rule"
ws.Cell(br, 2).Value = "Name"
ws.Cell(br, 3).Value = "Enabled"
ws.Cell(br, 4).Value = "Severity"
ws.Cell(br, 5).Value = "Findings"
For c As Integer = 1 To 5
    ws.Cell(br, c).Style.Font.SetBold(True).Font.SetFontColor(XLColor.White)
    ws.Cell(br, c).Style.Fill.SetBackgroundColor(HDR)
Next
Dim rr As Integer = br + 1
For Each rule As DataRow In in_dtRules.Rows
    Dim rid As String = Convert.ToString(rule("RuleID"))
    ws.Cell(rr, 1).Value = rid
    ws.Cell(rr, 2).Value = Convert.ToString(rule("RuleName"))
    ws.Cell(rr, 3).Value = Convert.ToString(rule("Enabled"))
    ws.Cell(rr, 4).Value = Convert.ToString(rule("Severity"))
    Dim hits As Integer = 0
    If in_Stats.ContainsKey(rid & "_Hits") Then hits = Convert.ToInt32(in_Stats(rid & "_Hits"))
    ws.Cell(rr, 5).Value = hits
    If hits > 0 Then ws.Cell(rr, 5).Style.Fill.SetBackgroundColor(AMBER)
    rr += 1
Next
ws.Columns(1, 5).AdjustToContents()

' ---------- Detail sheets -------------------------------------------------
Dim sheets As New List(Of Object())
sheets.Add(New Object() {"ValidationResults", in_dtResults})
sheets.Add(New Object() {"Exceptions", in_dtExceptions})

For Each spec As Object() In sheets
    Dim nm As String = Convert.ToString(spec(0))
    Dim src As DataTable = CType(spec(1), DataTable)
    Dim d = wb.Worksheets.Add(nm)

    If src Is Nothing OrElse src.Rows.Count = 0 Then
        d.Cell(1, 1).Value = "No rows"
        Continue For
    End If

    d.Cell(1, 1).InsertTable(src, nm & "Table", True)

    Dim statusCol As Integer = src.Columns.IndexOf("ValidationStatus") + 1
    If statusCol > 0 Then
        For i As Integer = 0 To src.Rows.Count - 1
            Dim st As String = Convert.ToString(src.Rows(i)("ValidationStatus"))
            Dim cell = d.Cell(i + 2, statusCol)
            If st = "PASS" Then
                cell.Style.Fill.SetBackgroundColor(GREEN)
            ElseIf st = "FAIL" Then
                cell.Style.Fill.SetBackgroundColor(RED)
            ElseIf st = "WARNING" Then
                cell.Style.Fill.SetBackgroundColor(AMBER)
            End If
        Next
    End If

    d.SheetView.FreezeRows(1)
    d.Columns().AdjustToContents()
Next

wb.SaveAs(outPath)

out_ReportPath = outPath
out_Summary = String.Format("report written: {0} ({1} results, {2} exceptions)", _
    Path.GetFileName(outPath), in_dtResults.Rows.Count, If(in_dtExceptions Is Nothing, 0, in_dtExceptions.Rows.Count))
'''.strip()

    body = sequence(
        "07 - Generate styled Excel report",
        invoke_code(code, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtResults", "sd:DataTable", "in_dtResults"),
            ("In", "in_dtExceptions", "sd:DataTable", "in_dtExceptions"),
            ("In", "in_dtRules", "sd:DataTable", "in_dtRules"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("Out", "out_ReportPath", "x:String", "out_ReportPath"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - build the workbook with ClosedXML")
        + log('"[REPORT] " & summary'),
        variables(("x:String", "summary")))

    return workflow("07_Report_Excel", body,
                    members=[
                        ("in_Config", f"InArgument({DICT_SO})"),
                        ("in_dtResults", "InArgument(sd:DataTable)"),
                        ("in_dtExceptions", "InArgument(sd:DataTable)"),
                        ("in_dtRules", "InArgument(sd:DataTable)"),
                        ("in_Stats", f"InArgument({DICT_SO})"),
                        ("out_ReportPath", "OutArgument(x:String)"),
                    ],
                    extra_imports=["ClosedXML.Excel"],
                    extra_refs=["ClosedXML"])


# ==========================================================================
# 11_Dashboard_Summary
# ==========================================================================
def dashboard_summary() -> str:
    code = r'''
' Emit the run's KPIs as JSON. This is the hand-off point for an external dashboard -
' Power BI, Google Sheets or anything else - without coupling the bot to one of them.
Dim o As New JObject()
o("runId") = Convert.ToString(in_Config("RunId"))
o("generatedUtc") = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
o("botIdentity") = Convert.ToString(in_Config("BotIdentity"))
o("machine") = Convert.ToString(in_Config("MachineName"))
o("dataSourceMode") = Convert.ToString(in_Config("DataSourceMode"))
o("chainId") = Convert.ToString(in_Config("ChainId"))
o("contract") = Convert.ToString(in_Config("ContractAddress"))
o("total") = Convert.ToInt32(in_Stats("Total"))
o("pass") = Convert.ToInt32(in_Stats("Pass"))
o("warning") = Convert.ToInt32(in_Stats("Warning"))
o("fail") = Convert.ToInt32(in_Stats("Fail"))
o("escalations") = Convert.ToInt32(in_Stats("Escalations"))
o("passPercent") = Convert.ToDouble(in_Stats("PassPercent"))
Dim totalCount As Integer = Convert.ToInt32(in_Stats("Total"))
o("failPercent") = If(totalCount = 0, 0.0, Math.Round(Convert.ToInt32(in_Stats("Fail")) * 100.0 / totalCount, 1))
o("warningPercent") = If(totalCount = 0, 0.0, Math.Round(Convert.ToInt32(in_Stats("Warning")) * 100.0 / totalCount, 1))
o("auditChainHead") = in_HeadHash
o("auditChainValid") = in_ChainValid

' ---- Processing time (Objective 6) -------------------------------------------------
' Main stamps the end of each phase into the config dictionary; the gaps between the
' stamps are the phase timings, and the run's duration is measured from the moment
' configuration loading began.
Dim started As DateTime = Convert.ToDateTime(in_Config("RunStartedUtc"))
Dim durationSeconds As Double = Math.Round((DateTime.UtcNow - started).TotalSeconds, 1)
Dim phaseSpec() As String = {"Init|Read configuration", "ExtractChain|Extract blockchain data", _
    "ExtractErp|Extract ERP data", "Map|Preprocess and map", "Validate|Validation engine (R1-R6)", _
    "Escalate|Human-in-the-loop review", "AuditTrail|Audit log and verification", _
    "Reports|Excel and PDF reports", "Alerts|E-mail, Teams and Slack alerts"}
Dim phases As New JArray()
Dim prevMark As DateTime = started
For Each spec As String In phaseSpec
    Dim parts() As String = spec.Split("|"c)
    If in_Config.ContainsKey("Mark_" & parts(0)) Then
        Dim at As DateTime = Convert.ToDateTime(in_Config("Mark_" & parts(0)))
        Dim ph As New JObject()
        ph("phase") = parts(1)
        ph("seconds") = Math.Round(Math.Max(0.0, (at - prevMark).TotalSeconds), 2)
        phases.Add(ph)
        prevMark = at
    End If
Next
Dim total As Integer = Convert.ToInt32(in_Stats("Total"))
o("startedUtc") = started.ToString("o", CultureInfo.InvariantCulture)
o("durationSeconds") = durationSeconds
o("secondsPerTransaction") = If(total = 0, 0.0, Math.Round(durationSeconds / total, 3))
o("phases") = phases

Dim rulesArr As New JArray()
For Each rule As DataRow In in_dtRules.Rows
    Dim rid As String = Convert.ToString(rule("RuleID"))
    Dim ro As New JObject()
    ro("ruleId") = rid
    ro("name") = Convert.ToString(rule("RuleName"))
    ro("enabled") = (Convert.ToString(rule("Enabled")).Trim().ToUpperInvariant() = "TRUE")
    ro("severity") = Convert.ToString(rule("Severity"))
    ro("findings") = If(in_Stats.ContainsKey(rid & "_Hits"), Convert.ToInt32(in_Stats(rid & "_Hits")), 0)
    rulesArr.Add(ro)
Next
o("rules") = rulesArr

' Top anomaly types: the rules that fired, most findings first.
Dim ranked As New List(Of JObject)
For Each ro As JObject In rulesArr
    If CInt(ro("findings")) > 0 Then ranked.Add(ro)
Next
ranked.Sort(Function(a, b) CInt(b("findings")).CompareTo(CInt(a("findings"))))
Dim top As New JArray()
For Each ro As JObject In ranked
    Dim t As New JObject()
    t("ruleId") = ro("ruleId")
    t("name") = ro("name")
    t("findings") = ro("findings")
    top.Add(t)
Next
o("topAnomalyTypes") = top

' Top anomaly types, most frequent first - the "top anomaly types" tile.
Dim byShipment As New JArray()
For Each r As DataRow In in_dtResults.Rows
    If Convert.ToString(r("ValidationStatus")) <> "PASS" Then
        Dim e As New JObject()
        e("shipmentId") = Convert.ToString(r("ShipmentID"))
        e("event") = Convert.ToString(r("EventName"))
        e("status") = Convert.ToString(r("ValidationStatus"))
        e("severity") = Convert.ToString(r("Severity"))
        e("reasons") = Convert.ToString(r("FailureReasons"))
        byShipment.Add(e)
    End If
Next
o("findings") = byShipment

Dim outPath As String = Path.Combine(Convert.ToString(in_Config("ReportsFolder")), "dashboard_summary.json")
File.WriteAllText(outPath, o.ToString(Newtonsoft.Json.Formatting.Indented), Encoding.UTF8)

' ---- Run history: one row per run, for trend charts ----------------------------------
Dim cfgText As Func(Of String, String, String) = Function(k As String, fallback As String) _
    If(in_Config.ContainsKey(k) AndAlso Convert.ToString(in_Config(k)).Trim() <> "", Convert.ToString(in_Config(k)).Trim(), fallback)
Dim histFile As String = Path.Combine(Convert.ToString(in_Config("AuditLogFolder")), cfgText("RunHistoryFile", "RunHistory.csv"))
Dim HIST_HEADER As String = "RunId,StartedUtc,DurationSeconds,DataSourceMode,RunKind,ErpSource,Total,Pass,Warning,Fail,Escalations,PassPercent,R1,R2,R3,R4,R5,R6,AuditChainValid"
Directory.CreateDirectory(Path.GetDirectoryName(histFile))
If Not File.Exists(histFile) Then File.WriteAllText(histFile, HIST_HEADER & Environment.NewLine, Encoding.UTF8)
Dim hits As New List(Of String)
For Each rid As String In New String() {"R1", "R2", "R3", "R4", "R5", "R6"}
    hits.Add(If(in_Stats.ContainsKey(rid & "_Hits"), Convert.ToString(in_Stats(rid & "_Hits")), "0"))
Next
Dim histRow As New List(Of String)
histRow.Add(Convert.ToString(in_Config("RunId")))
histRow.Add(started.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture))
histRow.Add(durationSeconds.ToString("0.0", CultureInfo.InvariantCulture))
histRow.Add(cfgText("DataSourceMode", "-"))
histRow.Add(cfgText("RunKind", "BATCH"))
histRow.Add(cfgText("LogisticsSourceMode", "EXCEL"))
histRow.Add(Convert.ToString(in_Stats("Total")))
histRow.Add(Convert.ToString(in_Stats("Pass")))
histRow.Add(Convert.ToString(in_Stats("Warning")))
histRow.Add(Convert.ToString(in_Stats("Fail")))
histRow.Add(Convert.ToString(in_Stats("Escalations")))
histRow.Add(Convert.ToDouble(in_Stats("PassPercent")).ToString("0.0", CultureInfo.InvariantCulture))
histRow.AddRange(hits)
histRow.Add(If(in_ChainValid, "TRUE", "FALSE"))
File.AppendAllText(histFile, String.Join(",", histRow) & Environment.NewLine, Encoding.UTF8)

' ---- Optional push to Power BI / Google Sheets --------------------------------------------
' POWERBI: a streaming dataset's push URL takes a JSON array of flat rows.
' JSON: the whole summary document, e.g. for a Google Apps Script web app.
out_PushUrl = cfgText("DashboardPushUrl", "")
If cfgText("DashboardPushFormat", "POWERBI").ToUpperInvariant() = "JSON" Then
    out_PushBody = o.ToString(Newtonsoft.Json.Formatting.None)
Else
    Dim flat As New JObject()
    flat("runId") = Convert.ToString(in_Config("RunId"))
    flat("timestampUtc") = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
    flat("dataSourceMode") = cfgText("DataSourceMode", "-")
    flat("chainId") = cfgText("ChainId", "-")
    flat("total") = Convert.ToInt32(in_Stats("Total"))
    flat("pass") = Convert.ToInt32(in_Stats("Pass"))
    flat("warning") = Convert.ToInt32(in_Stats("Warning"))
    flat("fail") = Convert.ToInt32(in_Stats("Fail"))
    flat("escalations") = Convert.ToInt32(in_Stats("Escalations"))
    flat("passPercent") = Convert.ToDouble(in_Stats("PassPercent"))
    flat("failPercent") = CDbl(o("failPercent"))
    flat("topAnomalyType") = If(top.Count = 0, "none", Convert.ToString(top(0)("name")))
    flat("durationSeconds") = durationSeconds
    For i As Integer = 0 To 5
        flat("r" & (i + 1).ToString() & "Findings") = Convert.ToInt32(hits(i))
    Next
    flat("auditChainValid") = in_ChainValid
    out_PushBody = New JArray(flat).ToString(Newtonsoft.Json.Formatting.None)
End If

out_DashboardPath = outPath
out_Summary = String.Format("dashboard JSON written with {0} finding(s): {1} | processing time {2:0.0}s | run history {3}", _
    byShipment.Count, Path.GetFileName(outPath), durationSeconds, Path.GetFileName(histFile))
'''.strip()

    push = sequence(
        "Push KPIs to the external dashboard",
        f'<ui:HttpClient DisplayName="{lit("HTTP Request - dashboard push")}" '
        f'EndPoint="{vb("pushUrl")}" Method="POST" AcceptFormat="ANY" BodyFormat="application/json" '
        f'Body="{vb("pushBody")}" TimeoutMS="{vb("30000")}" Result="{vb("response")}" '
        f'StatusCode="{vb("statusCode")}" />'
        + if_("statusCode >= 200 AndAlso statusCode < 300",
              log('"[DASHBOARD] KPIs pushed (" & Convert.ToString(in_Config("DashboardPushFormat")) & ") - HTTP " & statusCode.ToString()'),
              log('"[DASHBOARD] Dashboard push rejected - HTTP " & statusCode.ToString() & " " & response', level="Warn"),
              name="If - push accepted"))

    body = sequence(
        "11 - Emit dashboard summary",
        invoke_code(code, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtResults", "sd:DataTable", "in_dtResults"),
            ("In", "in_dtRules", "sd:DataTable", "in_dtRules"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("In", "in_HeadHash", "x:String", "in_HeadHash"),
            ("In", "in_ChainValid", "x:Boolean", "in_ChainValid"),
            ("Out", "out_DashboardPath", "x:String", "out_DashboardPath"),
            ("Out", "out_PushUrl", "x:String", "pushUrl"),
            ("Out", "out_PushBody", "x:String", "pushBody"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - build dashboard JSON and run history")
        + log('"[DASHBOARD] " & summary')
        + if_('pushUrl <> ""', push,
              log('"[DASHBOARD] DashboardPushUrl is blank - no external dashboard push."'),
              name="If - DashboardPushUrl configured"),
        variables(("x:String", "summary"), ("x:String", "pushUrl"), ("x:String", "pushBody"),
                  ("x:String", "response"), ("x:Int32", "statusCode")))

    return workflow("11_Dashboard_Summary", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtResults", "InArgument(sd:DataTable)"),
        ("in_dtRules", "InArgument(sd:DataTable)"),
        ("in_Stats", f"InArgument({DICT_SO})"),
        ("in_HeadHash", "InArgument(x:String)"),
        ("in_ChainValid", "InArgument(x:Boolean)"),
        ("out_DashboardPath", "OutArgument(x:String)"),
    ])


# ==========================================================================
# 08_Report_PDF  -  Phase 4 "PDF Audit Log"
# ==========================================================================
# A formal, paginated audit document: run metadata and bot identity, the outcome, every
# rule, every finding with the reviewer's decision, the full per-transaction results and
# the audit-chain evidence. It is written directly as PDF 1.4 using the standard Helvetica
# and Courier fonts that every reader has built in, so no Word/Office process and no
# third-party library is involved and an unattended robot produces the same file.
# Content is restricted to printable ASCII so the xref byte offsets are exact.
PDF_HELVETICA_WIDTHS = [
    278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278,   # 32-47
    556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556,   # 48-63
    1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778,  # 64-79
    667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556,   # 80-95
    333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556,   # 96-111
    556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584,        # 112-126
]
PDF_HELVETICA_BOLD_WIDTHS = [
    278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278,   # 32-47
    556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611,   # 48-63
    975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778,   # 64-79
    667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556,   # 80-95
    333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611,   # 96-111
    611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584,        # 112-126
]
assert len(PDF_HELVETICA_WIDTHS) == 95 and len(PDF_HELVETICA_BOLD_WIDTHS) == 95

PDF_REPORT_CODE = r'''
' ---- Page geometry, palette and font metrics -------------------------------------
Dim PAGE_W As Double = 842.0   ' A4 landscape, in points
Dim PAGE_H As Double = 595.0
Dim MARGIN As Double = 36.0
Dim CONTENT_W As Double = PAGE_W - 2 * MARGIN

Dim NAVY As String = "0.122 0.220 0.392"
Dim WHITE As String = "1 1 1"
Dim INK As String = "0.086 0.125 0.173"
Dim MUTED As String = "0.408 0.467 0.541"
Dim RULE_LINE As String = "0.875 0.898 0.925"
Dim HEAD_BG As String = "0.933 0.949 0.973"
Dim PASS_BG As String = "0.894 0.961 0.925"
Dim WARN_BG As String = "0.992 0.949 0.847"
Dim FAIL_BG As String = "0.992 0.925 0.918"
Dim PASS_INK As String = "0.102 0.498 0.294"
Dim WARN_INK As String = "0.541 0.380 0.000"
Dim FAIL_INK As String = "0.702 0.149 0.118"

' Glyph widths (1/1000 em) of the standard Helvetica fonts for ASCII 32..126.
Dim helv() As Integer = {@@HELV@@}
Dim helvBold() As Integer = {@@HELVB@@}

' ---- Text helpers ---------------------------------------------------------------------
Dim clean As Func(Of String, String) = Function(s As String) As String
    If s Is Nothing Then Return ""
    Dim sb As New StringBuilder(s.Length)
    For Each ch As Char In s
        Dim code As Integer = AscW(ch)
        If code >= 32 AndAlso code <= 126 Then
            sb.Append(ch)
        ElseIf code = 8211 OrElse code = 8212 Then
            sb.Append("-"c)
        ElseIf code = 9 OrElse code = 10 OrElse code = 13 Then
            sb.Append(" "c)
        Else
            sb.Append("?"c)
        End If
    Next
    Return sb.ToString()
End Function

Dim fmt As Func(Of Double, String) = Function(v As Double) v.ToString("0.##", CultureInfo.InvariantCulture)
Dim pdfString As Func(Of String, String) = Function(s As String) clean(s).Replace("\", "\\").Replace("(", "\(").Replace(")", "\)")

' font: 1 = Helvetica, 2 = Helvetica-Bold, 3 = Courier
Dim textWidth As Func(Of String, Integer, Double, Double) = Function(s As String, font As Integer, size As Double) As Double
    Dim units As Double = 0
    For Each ch As Char In s
        Dim c As Integer = AscW(ch)
        If font = 3 Then
            units += 600
        ElseIf c >= 32 AndAlso c <= 126 Then
            units += If(font = 2, helvBold(c - 32), helv(c - 32))
        Else
            units += 556
        End If
    Next
    Return units * size / 1000.0
End Function

Dim wrapText As Func(Of String, Integer, Double, Double, List(Of String)) = _
    Function(s As String, font As Integer, size As Double, maxW As Double) As List(Of String)
        Dim lines As New List(Of String)
        Dim current As String = ""
        For Each word As String In clean(s).Split(" "c)
            Dim candidate As String = If(current = "", word, current & " " & word)
            If textWidth(candidate, font, size) <= maxW Then
                current = candidate
            Else
                If current <> "" Then lines.Add(current)
                current = word
                ' A single token wider than the column (a hash, a long address) is broken.
                While current.Length > 1 AndAlso textWidth(current, font, size) > maxW
                    Dim cut As Integer = current.Length - 1
                    While cut > 1 AndAlso textWidth(current.Substring(0, cut), font, size) > maxW
                        cut -= 1
                    End While
                    lines.Add(current.Substring(0, cut))
                    current = current.Substring(cut)
                End While
            End If
        Next
        If current <> "" OrElse lines.Count = 0 Then lines.Add(current)
        Return lines
    End Function

' ---- Page and drawing primitives ------------------------------------------------------
Dim pages As New List(Of StringBuilder)
Dim page As StringBuilder = Nothing
Dim y As Double = 0
Dim FOOTER_H As Double = 26.0

Dim newPage As Action = Sub()
    page = New StringBuilder()
    pages.Add(page)
    y = PAGE_H - MARGIN
End Sub

Dim drawText As Action(Of Double, Double, Integer, Double, String, String) = _
    Sub(x As Double, ty As Double, font As Integer, size As Double, color As String, s As String)
        page.Append(color).Append(" rg BT /F").Append(font).Append(" ").Append(fmt(size)).Append(" Tf ") _
            .Append(fmt(x)).Append(" ").Append(fmt(ty)).Append(" Td (").Append(pdfString(s)).Append(") Tj ET").Append(vbLf)
    End Sub

Dim fillRect As Action(Of Double, Double, Double, Double, String) = _
    Sub(x As Double, ry As Double, w As Double, h As Double, color As String)
        page.Append(color).Append(" rg ").Append(fmt(x)).Append(" ").Append(fmt(ry)).Append(" ") _
            .Append(fmt(w)).Append(" ").Append(fmt(h)).Append(" re f").Append(vbLf)
    End Sub

Dim hline As Action(Of Double, Double, Double, String) = _
    Sub(x1 As Double, ly As Double, x2 As Double, color As String)
        page.Append(color).Append(" RG 0.5 w ").Append(fmt(x1)).Append(" ").Append(fmt(ly)).Append(" m ") _
            .Append(fmt(x2)).Append(" ").Append(fmt(ly)).Append(" l S").Append(vbLf)
    End Sub

Dim ensureSpace As Action(Of Double) = Sub(h As Double)
    If page Is Nothing OrElse y - h < MARGIN + FOOTER_H Then newPage()
End Sub

Dim heading As Action(Of String) = Sub(t As String)
    ensureSpace(60)
    y -= 20
    drawText(MARGIN, y, 2, 12, NAVY, t)
    y -= 6
    hline(MARGIN, y, PAGE_W - MARGIN, RULE_LINE)
    y -= 6
End Sub

Dim paragraph As Action(Of String, Double) = Sub(t As String, size As Double)
    For Each ln As String In wrapText(t, 1, size, CONTENT_W)
        ensureSpace(size + 4)
        y -= size + 3
        drawText(MARGIN, y, 1, size, INK, ln)
    Next
End Sub

' A table that wraps cell text, grows rows to fit and repeats its header on a new page.
Dim drawTable As Action(Of String(), Double(), List(Of String()), Func(Of Integer, Integer, String), Integer()) = _
    Sub(headers As String(), widths As Double(), rows As List(Of String()), cellFill As Func(Of Integer, Integer, String), fonts As Integer())
        Dim size As Double = 8
        Dim lead As Double = 10
        Dim pad As Double = 3.5
        Dim tableW As Double = widths.Sum()
        Dim drawHeader As Action = Sub()
            Dim hh As Double = lead + pad * 2
            fillRect(MARGIN, y - hh, tableW, hh, HEAD_BG)
            Dim hx As Double = MARGIN
            For i As Integer = 0 To headers.Length - 1
                drawText(hx + pad, y - pad - size * 0.85, 2, size, MUTED, headers(i))
                hx += widths(i)
            Next
            y -= hh
        End Sub

        ensureSpace(3 * lead + 4 * pad)
        drawHeader()
        For r As Integer = 0 To rows.Count - 1
            Dim cellLines As New List(Of List(Of String))
            Dim maxLines As Integer = 1
            For i As Integer = 0 To headers.Length - 1
                Dim lns As List(Of String) = wrapText(rows(r)(i), fonts(i), size, widths(i) - pad * 2)
                cellLines.Add(lns)
                maxLines = Math.Max(maxLines, lns.Count)
            Next
            Dim rh As Double = maxLines * lead + pad * 2
            If y - rh < MARGIN + FOOTER_H Then
                newPage()
                drawHeader()
            End If
            Dim cx As Double = MARGIN
            For i As Integer = 0 To headers.Length - 1
                Dim fill As String = Nothing
                If cellFill IsNot Nothing Then fill = cellFill(r, i)
                If fill IsNot Nothing Then fillRect(cx, y - rh, widths(i), rh, fill)
                For li As Integer = 0 To cellLines(i).Count - 1
                    drawText(cx + pad, y - pad - size * 0.85 - li * lead, fonts(i), size, INK, cellLines(i)(li))
                Next
                cx += widths(i)
            Next
            y -= rh
            hline(MARGIN, y, MARGIN + tableW, RULE_LINE)
        Next
    End Sub

Dim statusFill As Func(Of String, String) = Function(st As String) As String
    Select Case st
        Case "PASS", "P" : Return PASS_BG
        Case "WARNING", "W" : Return WARN_BG
        Case "FAIL", "F" : Return FAIL_BG
    End Select
    Return Nothing
End Function

Dim cfg As Func(Of String, String) = Function(k As String) If(in_Config.ContainsKey(k), Convert.ToString(in_Config(k)), "")
Dim stat As Func(Of String, String) = Function(k As String) If(in_Stats.ContainsKey(k), Convert.ToString(in_Stats(k)), "0")

Dim runId As String = cfg("RunId")
Dim project As String = cfg("ProjectName")
Dim bot As String = cfg("BotIdentity")
Dim generatedUtc As DateTime = DateTime.UtcNow

' ---- Page 1 header band -------------------------------------------------------------
newPage()
fillRect(0, PAGE_H - 74, PAGE_W, 74, NAVY)
drawText(MARGIN, PAGE_H - 38, 2, 18, WHITE, "Blockchain Transaction Validation - Audit Report")
drawText(MARGIN, PAGE_H - 57, 1, 10, "0.839 0.878 0.933", project & "  |  Run " & runId & "  |  Bot " & bot)
y = PAGE_H - 74 - 6

' ---- 1. Run metadata -------------------------------------------------------------------
heading("1. Run metadata")
Dim enabledRules As Integer = 0
For Each rr As DataRow In in_dtRules.Rows
    If Convert.ToString(rr("Enabled")).Trim().ToUpperInvariant() = "TRUE" Then enabledRules += 1
Next
Dim started As String = ""
If in_Config.ContainsKey("RunStartedUtc") Then
    started = Convert.ToDateTime(in_Config("RunStartedUtc")).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
End If
Dim meta As New List(Of String())
meta.Add(New String() {"Run ID", runId})
meta.Add(New String() {"Bot identity", bot})
meta.Add(New String() {"Machine", cfg("MachineName")})
meta.Add(New String() {"Windows user", cfg("WindowsUser")})
meta.Add(New String() {"Run started (UTC)", started})
meta.Add(New String() {"Report generated (UTC)", generatedUtc.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)})
meta.Add(New String() {"Blockchain source", cfg("DataSourceMode") & " (chain id " & cfg("ChainId") & ")"})
meta.Add(New String() {"Contract audited", cfg("ContractAddress")})
meta.Add(New String() {"Field mapping profile", cfg("FieldMappingSheet")})
meta.Add(New String() {"Shipment mapping", cfg("MappingMode")})
meta.Add(New String() {"ERP / WMS source", cfg("LogisticsSourceMode")})
meta.Add(New String() {"Configuration", cfg("ConfigSource")})
meta.Add(New String() {"Rules enabled", enabledRules.ToString() & " of " & in_dtRules.Rows.Count.ToString()})
meta.Add(New String() {"Escalation mode", cfg("EscalationMode")})
Dim metaRows As New List(Of String())
For i As Integer = 0 To meta.Count - 1 Step 2
    Dim rightPair() As String = If(i + 1 < meta.Count, meta(i + 1), New String() {"", ""})
    metaRows.Add(New String() {meta(i)(0), meta(i)(1), rightPair(0), rightPair(1)})
Next
drawTable(New String() {"Field", "Value", "Field", "Value"}, New Double() {110, 275, 110, 275}, metaRows, Nothing, New Integer() {2, 1, 2, 1})

' ---- 2. Outcome summary -----------------------------------------------------------------
heading("2. Outcome summary")
Dim kpis As New List(Of String())
kpis.Add(New String() {"Transactions validated", stat("Total"), HEAD_BG, INK})
kpis.Add(New String() {"Passed", stat("Pass"), PASS_BG, PASS_INK})
kpis.Add(New String() {"Warnings", stat("Warning"), WARN_BG, WARN_INK})
kpis.Add(New String() {"Failed", stat("Fail"), FAIL_BG, FAIL_INK})
kpis.Add(New String() {"Escalated to a human", stat("Escalations"), HEAD_BG, INK})
kpis.Add(New String() {"Pass rate", stat("PassPercent") & "%", HEAD_BG, INK})
ensureSpace(60)
Dim boxW As Double = (CONTENT_W - 10 * (kpis.Count - 1)) / kpis.Count
For i As Integer = 0 To kpis.Count - 1
    Dim bx As Double = MARGIN + i * (boxW + 10)
    fillRect(bx, y - 50, boxW, 50, kpis(i)(2))
    drawText(bx + 8, y - 16, 1, 8, MUTED, kpis(i)(0))
    drawText(bx + 8, y - 40, 2, 18, kpis(i)(3), kpis(i)(1))
Next
y -= 56
paragraph("Rule findings: " & If(stat("RuleHitSummary") = "0" OrElse stat("RuleHitSummary") = "", "none", stat("RuleHitSummary")), 9)

' ---- 3. Validation rules ------------------------------------------------------------------
heading("3. Validation rules applied")
Dim ruleRows As New List(Of String())
For Each rr As DataRow In in_dtRules.Rows
    Dim rid As String = Convert.ToString(rr("RuleID"))
    Dim params As String = String.Join(" / ", New String() {Convert.ToString(rr("Param1")), Convert.ToString(rr("Param2"))}.Where(Function(p) p.Trim() <> ""))
    ruleRows.Add(New String() {rid, Convert.ToString(rr("RuleName")), Convert.ToString(rr("Enabled")), _
        Convert.ToString(rr("Severity")), Convert.ToString(rr("FailAction")), If(params = "", "-", params), _
        stat(rid & "_Hits"), Convert.ToString(rr("Description"))})
Next
drawTable(New String() {"Rule", "Name", "Enabled", "Severity", "Fail action", "Params", "Findings", "What it checks"}, _
    New Double() {40, 110, 55, 60, 65, 70, 55, 315}, ruleRows, _
    Function(r As Integer, c As Integer) As String
        If c = 6 AndAlso ruleRows(r)(6) <> "0" Then Return WARN_BG
        Return Nothing
    End Function, New Integer() {2, 1, 1, 1, 1, 1, 2, 1})

' ---- 4. Findings -------------------------------------------------------------------------
Dim findingRows As New List(Of String())
For Each r As DataRow In in_dtResults.Rows
    Dim st As String = Convert.ToString(r("ValidationStatus"))
    If st = "PASS" Then Continue For
    Dim decision As String = Convert.ToString(r("HumanDecision"))
    If decision <> "" Then decision = decision & " (" & Convert.ToString(r("HumanDecidedBy")) & ")"
    findingRows.Add(New String() {(findingRows.Count + 1).ToString(), Convert.ToString(r("ShipmentID")), _
        Convert.ToString(r("EventName")), st, Convert.ToString(r("Severity")), _
        Convert.ToString(r("FailureReasons")), If(decision = "", "-", decision)})
Next
heading("4. Findings requiring attention (" & findingRows.Count.ToString() & ")")
If findingRows.Count = 0 Then
    paragraph("Every transaction passed all enabled rules.", 9)
Else
    drawTable(New String() {"#", "Shipment", "Milestone", "Status", "Severity", "Findings", "Human decision"}, _
        New Double() {25, 65, 75, 55, 55, 330, 165}, findingRows, _
        Function(r As Integer, c As Integer) If(c = 3, statusFill(findingRows(r)(3)), Nothing), _
        New Integer() {1, 1, 1, 2, 1, 1, 1})
End If

' ---- 5. Full results -----------------------------------------------------------------------
heading("5. Validation results - every transaction")
Dim shorten As Func(Of String, Integer, Integer, String) = Function(s As String, head As Integer, tail As Integer) As String
    If s.Length <= head + tail + 3 Then Return s
    Return s.Substring(0, head) & "..." & If(tail > 0, s.Substring(s.Length - tail), "")
End Function
Dim letter As Func(Of String, String) = Function(s As String) As String
    Select Case s
        Case "PASS" : Return "P"
        Case "WARNING" : Return "W"
        Case "FAIL" : Return "F"
    End Select
    Return "-"
End Function
Dim resultRows As New List(Of String())
Dim hasInputHash As Boolean = in_dtResults.Columns.Contains("InputDataHash")
For Each r As DataRow In in_dtResults.Rows
    Dim erpQty As String = If(IsDBNull(r("ErpExpectedQty")), "-", Convert.ToString(r("ErpExpectedQty")))
    resultRows.Add(New String() { _
        Convert.ToString(r("RowNo")), Convert.ToString(r("ShipmentID")), Convert.ToString(r("EventName")), _
        shorten(Convert.ToString(r("TxHash")), 10, 4), _
        Convert.ToDateTime(r("ChainTimestampUtc")).ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture), _
        Convert.ToString(r("OnChainQty")), erpQty, _
        letter(Convert.ToString(r("R1_Status"))), letter(Convert.ToString(r("R2_Status"))), _
        letter(Convert.ToString(r("R3_Status"))), letter(Convert.ToString(r("R4_Status"))), _
        letter(Convert.ToString(r("R5_Status"))), letter(Convert.ToString(r("R6_Status"))), _
        Convert.ToString(r("ValidationStatus")), _
        If(hasInputHash, shorten(Convert.ToString(r("InputDataHash")), 22, 0), "-")})
Next
drawTable(New String() {"Row", "Shipment", "Milestone", "Tx hash", "On chain (UTC)", "Qty", "PO qty", _
                        "R1", "R2", "R3", "R4", "R5", "R6", "Status", "Input data hash"}, _
    New Double() {30, 60, 70, 95, 80, 50, 45, 26, 26, 26, 26, 26, 26, 50, 134}, resultRows, _
    Function(r As Integer, c As Integer) If(c >= 7 AndAlso c <= 13, statusFill(resultRows(r)(c)), Nothing), _
    New Integer() {1, 1, 1, 3, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 3})
paragraph("R1-R6 columns: P = pass, W = warning, F = fail, - = not evaluated (rule disabled or no ERP counterpart).", 7.5)

' ---- 6. Audit trail integrity -----------------------------------------------------------------
heading("6. Audit trail integrity")
Dim reportHash As String = "-"
If Not String.IsNullOrWhiteSpace(in_ReportPath) AndAlso File.Exists(in_ReportPath) Then
    Using sha As SHA256 = SHA256.Create()
        Using fs As FileStream = File.OpenRead(in_ReportPath)
            reportHash = BitConverter.ToString(sha.ComputeHash(fs)).Replace("-", "").ToLowerInvariant()
        End Using
    End Using
End If
Dim auditRows As New List(Of String())
auditRows.Add(New String() {"Audit log", in_AuditPath})
auditRows.Add(New String() {"Entries written by this run", in_dtResults.Rows.Count.ToString()})
auditRows.Add(New String() {"Chain head (EntryHash)", If(in_HeadHash, "")})
auditRows.Add(New String() {"Verification", If(in_ChainValid, "PASSED - ", "FAILED - ") & If(in_ChainReport, "")})
auditRows.Add(New String() {"What each entry fixes", "Bot identity, timestamp, run id, transaction hash, shipment, milestone, " & _
    "on-chain quantity, SHA-256 of the transaction's input data, verdict, severity, findings and the human decision - " & _
    "hashed into RecordHash and chained to the previous entry through PrevHash and EntryHash."})
auditRows.Add(New String() {"Excel report SHA-256", reportHash})
drawTable(New String() {"Item", "Evidence"}, New Double() {160, 610}, auditRows, _
    Function(r As Integer, c As Integer) If(r = 3 AndAlso c = 1, If(in_ChainValid, PASS_BG, FAIL_BG), Nothing), _
    New Integer() {2, 3})

' ---- 7. Attestation --------------------------------------------------------------------------
heading("7. Attestation")
paragraph("This document was generated automatically by " & bot & " from the validation results it recorded in " & _
    "the tamper-evident audit log above. Any later alteration of that log is detected by re-running " & _
    "12_Verify_AuditChain, which recomputes every hash and names the first entry that no longer matches.", 9)
ensureSpace(60)
y -= 40
For Each spec As Object() In New Object()() {New Object() {"Reviewed by", MARGIN}, New Object() {"Signature", MARGIN + 270}, New Object() {"Date", MARGIN + 540}}
    Dim sx As Double = CDbl(spec(1))
    hline(sx, y, sx + 220, INK)
    drawText(sx, y - 11, 1, 8, MUTED, CStr(spec(0)))
Next
y -= 16

' ---- Footers, now that the page count is known ----------------------------------------------
For i As Integer = 0 To pages.Count - 1
    page = pages(i)
    hline(MARGIN, 30, PAGE_W - MARGIN, RULE_LINE)
    drawText(MARGIN, 18, 1, 7.5, MUTED, project & "  |  Run " & runId & "  |  " & bot & "  |  generated " & _
        generatedUtc.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) & " UTC")
    Dim label As String = "Page " & (i + 1).ToString() & " of " & pages.Count.ToString()
    drawText(PAGE_W - MARGIN - textWidth(label, 1, 7.5), 18, 1, 7.5, MUTED, label)
Next

' ---- Serialise the PDF ------------------------------------------------------------------------
Dim pdf As New StringBuilder()
Dim offsets As New List(Of Integer)
pdf.Append("%PDF-1.4").Append(vbLf)
Dim addObject As Action(Of String) = Sub(body As String)
    offsets.Add(pdf.Length)
    pdf.Append(offsets.Count).Append(" 0 obj").Append(vbLf).Append(body).Append(vbLf).Append("endobj").Append(vbLf)
End Sub

Dim firstPageObject As Integer = 7
Dim kids As New List(Of String)
For i As Integer = 0 To pages.Count - 1
    kids.Add((firstPageObject + i * 2).ToString() & " 0 R")
Next
addObject("<< /Type /Catalog /Pages 2 0 R >>")
addObject("<< /Type /Pages /Kids [" & String.Join(" ", kids) & "] /Count " & pages.Count.ToString() & " >>")
addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")
addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Courier /Encoding /WinAnsiEncoding >>")
addObject("<< /Title (" & pdfString("Blockchain Validation Audit Report " & runId) & ") /Author (" & pdfString(bot) & _
    ") /Creator (" & pdfString(project) & ") /Producer (08_Report_PDF) /CreationDate (D:" & _
    generatedUtc.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture) & "Z) >>")
For i As Integer = 0 To pages.Count - 1
    Dim content As String = pages(i).ToString().TrimEnd(ChrW(10))
    addObject("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " & fmt(PAGE_W) & " " & fmt(PAGE_H) & "] " & _
        "/Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R >> >> /Contents " & (firstPageObject + i * 2 + 1).ToString() & " 0 R >>")
    addObject("<< /Length " & content.Length.ToString() & " >>" & vbLf & "stream" & vbLf & content & vbLf & "endstream")
Next

Dim xrefAt As Integer = pdf.Length
pdf.Append("xref").Append(vbLf).Append("0 ").Append(offsets.Count + 1).Append(vbLf)
' Every cross-reference entry is exactly 20 bytes: 10-digit offset, generation, type, space, LF.
pdf.Append("0000000000 65535 f ").Append(vbLf)
For Each o As Integer In offsets
    pdf.Append(o.ToString("0000000000", CultureInfo.InvariantCulture)).Append(" 00000 n ").Append(vbLf)
Next
pdf.Append("trailer").Append(vbLf)
pdf.Append("<< /Size ").Append(offsets.Count + 1).Append(" /Root 1 0 R /Info 6 0 R >>").Append(vbLf)
pdf.Append("startxref").Append(vbLf).Append(xrefAt).Append(vbLf).Append("%%EOF").Append(vbLf)

Dim folder As String = cfg("ReportsFolder")
Directory.CreateDirectory(folder)
Dim outPath As String = Path.Combine(folder, "AuditReport_" & runId & ".pdf")
File.WriteAllBytes(outPath, Encoding.ASCII.GetBytes(pdf.ToString()))

out_PdfPath = outPath
out_Summary = String.Format("PDF audit report written: {0} ({1} page(s), {2} finding(s), chain {3})", _
    Path.GetFileName(outPath), pages.Count, findingRows.Count, If(in_ChainValid, "verified", "NOT verified"))
'''.strip().replace("@@HELV@@", ", ".join(str(w) for w in PDF_HELVETICA_WIDTHS)) \
          .replace("@@HELVB@@", ", ".join(str(w) for w in PDF_HELVETICA_BOLD_WIDTHS))


def report_pdf() -> str:
    body = sequence(
        "08 - Generate PDF audit report",
        invoke_code(PDF_REPORT_CODE, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtResults", "sd:DataTable", "in_dtResults"),
            ("In", "in_dtRules", "sd:DataTable", "in_dtRules"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("In", "in_AuditPath", "x:String", "in_AuditPath"),
            ("In", "in_HeadHash", "x:String", "in_HeadHash"),
            ("In", "in_ChainValid", "x:Boolean", "in_ChainValid"),
            ("In", "in_ChainReport", "x:String", "in_ChainReport"),
            ("In", "in_ReportPath", "x:String", "in_ReportPath"),
            ("Out", "out_PdfPath", "x:String", "out_PdfPath"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - write the PDF audit report")
        + log('"[PDF] " & summary'),
        variables(("x:String", "summary")))

    return workflow("08_Report_PDF", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtResults", "InArgument(sd:DataTable)"),
        ("in_dtRules", "InArgument(sd:DataTable)"),
        ("in_Stats", f"InArgument({DICT_SO})"),
        ("in_AuditPath", "InArgument(x:String)"),
        ("in_HeadHash", "InArgument(x:String)"),
        ("in_ChainValid", "InArgument(x:Boolean)"),
        ("in_ChainReport", "InArgument(x:String)"),
        ("in_ReportPath", "InArgument(x:String)"),
        ("out_PdfPath", "OutArgument(x:String)"),
    ])


# ==========================================================================
# 09_Alert_Email
# ==========================================================================
def send_smtp_mail(secure: str) -> str:
    """
    Classic Send SMTP Mail Message (UiPath.Mail.SMTP.Activities.SendMail).

    SecureConnection is a plain enum property, not an argument, so it cannot be bound to
    a config value - the workflow chooses between two otherwise identical activities on
    SmtpUseSsl instead. "Auto" negotiates STARTTLS on 587 or SSL on 465.

    Expressions are written with real double quotes - `vb()` does the XML escaping.
    Pre-escaping them here would double-escape and produce `&amp;quot;`, which VB then
    reads as an undeclared identifier called `quot`.
    """
    def cfg(key: str) -> str:
        return f'Convert.ToString(in_Config("{key}"))'

    to_expr = 'Convert.ToString(in_Config("AlertTo")).Replace(";", ",")'
    port_expr = 'Convert.ToInt32(in_Config("SmtpPort"))'

    return (f'<ui:SendMail DisplayName="{lit("Send SMTP Mail Message - " + secure)}" '
            f'Server="{vb(cfg("SmtpHost"))}" '
            f'Port="{vb(port_expr)}" '
            f'Email="{vb(cfg("SmtpUser"))}" '
            f'Password="{vb(cfg("SmtpPassword"))}" '
            f'From="{vb(cfg("AlertFrom"))}" '
            f'To="{vb(to_expr)}" '
            f'Subject="{vb("subject")}" Body="{vb("htmlBody")}" '
            f'AttachmentsCollection="{vb("attachments")}" '
            f'IsBodyHtml="True" SecureConnection="{secure}" />')


def send_outlook_mail(draft: bool) -> str:
    """Send Outlook Mail Message - the desktop Outlook profile does the delivery."""
    to_expr = 'Convert.ToString(in_Config("AlertTo"))'
    label = "Save Outlook draft" if draft else "Send Outlook Mail Message"
    return (f'<ui:SendOutlookMail DisplayName="{lit(label)}" '
            f'Account="{vb(chr(34) + chr(34) + " & Convert.ToString(in_Config(" + chr(34) + "OutlookAccount" + chr(34) + "))")}" '
            f'To="{vb(to_expr)}" Subject="{vb("subject")}" Body="{vb("htmlBody")}" '
            f'AttachmentsCollection="{vb("attachments")}" '
            f'IsBodyHtml="True" IsDraft="{str(draft)}" />')


def alert_email() -> str:
    code = r'''
' Compose the anomaly alert. AlertMode = EML writes a fully-formed RFC 5322 message to
' disk instead of sending it: the demo is identical, and no SMTP credentials have to
' live in a project folder that gets submitted or shared.
Dim total As Integer = Convert.ToInt32(in_Stats("Total"))
Dim nFail As Integer = Convert.ToInt32(in_Stats("Fail"))
Dim nWarn As Integer = Convert.ToInt32(in_Stats("Warning"))
Dim nEsc As Integer = Convert.ToInt32(in_Stats("Escalations"))

Dim subject As String
If nFail > 0 Then
    subject = String.Format("[ACTION REQUIRED] {0} blockchain validation failure(s) - {1}", nFail, Convert.ToString(in_Config("RunId")))
ElseIf nWarn > 0 Then
    subject = String.Format("[REVIEW] {0} blockchain validation warning(s) - {1}", nWarn, Convert.ToString(in_Config("RunId")))
Else
    subject = String.Format("[OK] All {0} blockchain transactions validated - {1}", total, Convert.ToString(in_Config("RunId")))
End If

Dim sb As New StringBuilder()
sb.AppendLine("<html><body style=""font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#222"">")
sb.AppendLine("<h2 style=""color:#1F3864;margin-bottom:4px"">Blockchain Transaction Validation - " & System.Net.WebUtility.HtmlEncode(Convert.ToString(in_Config("ProjectName"))) & "</h2>")
sb.AppendLine("<p style=""color:#666;margin-top:0"">Run " & Convert.ToString(in_Config("RunId")) & " &middot; bot " & Convert.ToString(in_Config("BotIdentity")) & " on " & Convert.ToString(in_Config("MachineName")) & "</p>")

sb.AppendLine("<table cellpadding=""6"" cellspacing=""0"" style=""border-collapse:collapse;margin:12px 0"">")
sb.AppendLine("<tr><td style=""border:1px solid #ddd""><b>Validated</b></td><td style=""border:1px solid #ddd"">" & total.ToString() & "</td></tr>")
sb.AppendLine("<tr><td style=""border:1px solid #ddd;background:#C6EFCE""><b>Passed</b></td><td style=""border:1px solid #ddd"">" & Convert.ToString(in_Stats("Pass")) & " (" & Convert.ToString(in_Stats("PassPercent")) & "%)</td></tr>")
sb.AppendLine("<tr><td style=""border:1px solid #ddd;background:#FFEB9C""><b>Warnings</b></td><td style=""border:1px solid #ddd"">" & nWarn.ToString() & "</td></tr>")
sb.AppendLine("<tr><td style=""border:1px solid #ddd;background:#FFC7CE""><b>Failed</b></td><td style=""border:1px solid #ddd"">" & nFail.ToString() & "</td></tr>")
sb.AppendLine("<tr><td style=""border:1px solid #ddd""><b>Escalated</b></td><td style=""border:1px solid #ddd"">" & nEsc.ToString() & "</td></tr>")
sb.AppendLine("</table>")

If in_dtExceptions IsNot Nothing AndAlso in_dtExceptions.Rows.Count > 0 Then
    sb.AppendLine("<h3 style=""color:#1F3864"">Findings</h3>")
    sb.AppendLine("<table cellpadding=""6"" cellspacing=""0"" style=""border-collapse:collapse;font-size:12px"">")
    sb.AppendLine("<tr style=""background:#1F3864;color:#fff""><th style=""border:1px solid #ddd"">Shipment</th><th style=""border:1px solid #ddd"">Event</th><th style=""border:1px solid #ddd"">Status</th><th style=""border:1px solid #ddd"">Severity</th><th style=""border:1px solid #ddd;text-align:left"">Reason</th></tr>")
    For Each r As DataRow In in_dtExceptions.Rows
        Dim bg As String = "#FFEB9C"
        If Convert.ToString(r("ValidationStatus")) = "FAIL" Then bg = "#FFC7CE"
        sb.AppendLine("<tr>" & _
            "<td style=""border:1px solid #ddd"">" & Convert.ToString(r("ShipmentID")) & "</td>" & _
            "<td style=""border:1px solid #ddd"">" & Convert.ToString(r("EventName")) & "</td>" & _
            "<td style=""border:1px solid #ddd;background:" & bg & """>" & Convert.ToString(r("ValidationStatus")) & "</td>" & _
            "<td style=""border:1px solid #ddd"">" & Convert.ToString(r("Severity")) & "</td>" & _
            "<td style=""border:1px solid #ddd"">" & Convert.ToString(r("FailureReasons")) & "</td></tr>")
    Next
    sb.AppendLine("</table>")
Else
    sb.AppendLine("<p>No anomalies detected in this run.</p>")
End If

sb.AppendLine("<p style=""color:#666;font-size:11px;margin-top:18px"">Audit chain head: " & in_HeadHash & _
    "<br/>Audit chain check: " & System.Net.WebUtility.HtmlEncode(in_ChainReport) & _
    "<br/>Excel report: " & in_ReportPath & "<br/>PDF audit report: " & in_PdfPath & "</p>")
sb.AppendLine("</body></html>")

Dim body As String = sb.ToString()
Dim mode As String = Convert.ToString(in_Config("AlertMode")).Trim().ToUpperInvariant()
Dim toList As String = Convert.ToString(in_Config("AlertTo")).Replace(";", ", ")

Dim eml As New StringBuilder()
eml.AppendLine("From: " & Convert.ToString(in_Config("AlertFrom")))
eml.AppendLine("To: " & toList)
eml.AppendLine("Subject: " & subject)
eml.AppendLine("Date: " & DateTime.UtcNow.ToString("r", CultureInfo.InvariantCulture))
eml.AppendLine("X-Generated-By: " & Convert.ToString(in_Config("BotIdentity")))
eml.AppendLine("MIME-Version: 1.0")
eml.AppendLine("Content-Type: text/html; charset=utf-8")
eml.AppendLine()
eml.Append(body)

Dim outPath As String = Path.Combine(Convert.ToString(in_Config("ReportsFolder")), _
    "Alert_" & Convert.ToString(in_Config("RunId")) & ".eml")
File.WriteAllText(outPath, eml.ToString(), Encoding.UTF8)

' The reports travel with the alert, so a manager can act on it from their inbox.
Dim files As New List(Of String)
If Convert.ToString(in_Config("AlertAttachReports")).Trim().ToUpperInvariant() = "TRUE" Then
    For Each f As String In New String() {in_ReportPath, in_PdfPath}
        If Not String.IsNullOrWhiteSpace(f) AndAlso File.Exists(f) Then files.Add(f)
    Next
End If

out_AlertPath = outPath
out_Subject = subject
out_Body = body
out_Attachments = files
out_ShouldSend = (nFail > 0 OrElse nWarn > 0)
out_Summary = String.Format("mode={0} subject=""{1}"" written to {2} ({3} attachment(s) for delivery)", _
    mode, subject, Path.GetFileName(outPath), files.Count)
'''.strip()

    body = sequence(
        "09 - Compose and deliver anomaly alert",
        invoke_code(code, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtExceptions", "sd:DataTable", "in_dtExceptions"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("In", "in_ReportPath", "x:String", "in_ReportPath"),
            ("In", "in_PdfPath", "x:String", "in_PdfPath"),
            ("In", "in_HeadHash", "x:String", "in_HeadHash"),
            ("In", "in_ChainReport", "x:String", "in_ChainReport"),
            ("Out", "out_AlertPath", "x:String", "out_AlertPath"),
            ("Out", "out_Subject", "x:String", "subject"),
            ("Out", "out_Body", "x:String", "htmlBody"),
            ("Out", "out_Attachments", "scg:List(x:String)", "attachments"),
            ("Out", "out_ShouldSend", "x:Boolean", "shouldSend"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - compose the alert")
        + switch(
            'Convert.ToString(in_Config("AlertMode")).Trim().ToUpperInvariant()',
            [("SMTP", sequence("SMTP", if_(
                "shouldSend",
                sequence("Send over SMTP",
                         if_('Convert.ToString(in_Config("SmtpUseSsl")).Trim().ToUpperInvariant() = "FALSE"',
                             send_smtp_mail("None"), send_smtp_mail("Auto"),
                             name="If - SmtpUseSsl")
                         + log('"[ALERT] Sent over SMTP via " & Convert.ToString(in_Config("SmtpHost")) '
                               '& " to " & Convert.ToString(in_Config("AlertTo"))')),
                log('"[ALERT] Nothing failed or warned - no SMTP message sent."'),
                name="If - anything to report"))),
             ("OUTLOOK", sequence("OUTLOOK", if_(
                "shouldSend",
                sequence("Deliver through Outlook",
                         if_('Convert.ToString(in_Config("OutlookSaveAsDraft")).Trim().ToUpperInvariant() = "TRUE"',
                             send_outlook_mail(True), send_outlook_mail(False),
                             name="If - OutlookSaveAsDraft")
                         + log('"[ALERT] Delivered through Outlook (draft=" '
                               '& Convert.ToString(in_Config("OutlookSaveAsDraft")) & ")"')),
                log('"[ALERT] Nothing failed or warned - no Outlook message sent."'),
                name="If - anything to report")))],
            sequence("EML", log('"[ALERT] AlertMode is EML - the composed message was written to '
                                'disk instead of sent. Set AlertMode to SMTP or OUTLOOK in Config.xlsx to deliver it."')),
            name="Switch - AlertMode")
        + log('"[ALERT] " & summary')
        # Teams and Slack run alongside e-mail, whatever AlertMode is.
        + invoke_workflow("Workflows\\09b_Alert_Webhooks.xaml", [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtExceptions", "sd:DataTable", "in_dtExceptions"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("In", "in_ShouldSend", "x:Boolean", "shouldSend"),
        ], name="Invoke 09b - Teams and Slack alerts"),
        variables(("x:String", "subject"), ("x:String", "htmlBody"),
                  ("x:Boolean", "shouldSend"), ("x:String", "summary"),
                  ("scg:List(x:String)", "attachments")))

    return workflow("09_Alert_Email", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtExceptions", "InArgument(sd:DataTable)"),
        ("in_Stats", f"InArgument({DICT_SO})"),
        ("in_ReportPath", "InArgument(x:String)"),
        ("in_PdfPath", "InArgument(x:String)"),
        ("in_HeadHash", "InArgument(x:String)"),
        ("in_ChainReport", "InArgument(x:String)"),
        ("out_AlertPath", "OutArgument(x:String)"),
    ])


# ==========================================================================
# 09b_Alert_Webhooks  -  Teams and Slack
# ==========================================================================
# The deck's alerting objective names e-mail, Teams and Slack. Both chat channels take an
# incoming-webhook URL from Config.xlsx. When a URL is blank the composed payload is
# written beside the reports instead, so the message can be inspected - and demonstrated -
# without a tenant, exactly as AlertMode = EML does for e-mail.
WEBHOOK_COMPOSE_CODE = r'''
Dim total As Integer = Convert.ToInt32(in_Stats("Total"))
Dim nFail As Integer = Convert.ToInt32(in_Stats("Fail"))
Dim nWarn As Integer = Convert.ToInt32(in_Stats("Warning"))
Dim nPass As Integer = Convert.ToInt32(in_Stats("Pass"))
Dim nEsc As Integer = Convert.ToInt32(in_Stats("Escalations"))
Dim runId As String = Convert.ToString(in_Config("RunId"))
Dim project As String = Convert.ToString(in_Config("ProjectName"))
Dim headline As String
If nFail > 0 Then
    headline = String.Format("ACTION REQUIRED: {0} blockchain validation failure(s)", nFail)
ElseIf nWarn > 0 Then
    headline = String.Format("REVIEW: {0} blockchain validation warning(s)", nWarn)
Else
    headline = String.Format("OK: all {0} blockchain transactions validated", total)
End If

' The findings list is capped so a large run cannot exceed either service's message size.
Dim findings As New List(Of String())
If in_dtExceptions IsNot Nothing Then
    For Each r As DataRow In in_dtExceptions.Rows
        findings.Add(New String() {Convert.ToString(r("ShipmentID")), Convert.ToString(r("EventName")), _
            Convert.ToString(r("ValidationStatus")), Convert.ToString(r("Severity")), Convert.ToString(r("FailureReasons"))})
    Next
End If
Dim shown As Integer = Math.Min(findings.Count, 15)

' ---- Microsoft Teams: an Adaptive Card, the format Teams Workflows webhooks accept ----
Dim facts As New JArray()
For Each kv As String() In New String()() { _
        New String() {"Run", runId}, New String() {"Bot", Convert.ToString(in_Config("BotIdentity"))}, _
        New String() {"Validated", total.ToString()}, _
        New String() {"Passed", nPass.ToString() & " (" & Convert.ToString(in_Stats("PassPercent")) & "%)"}, _
        New String() {"Warnings", nWarn.ToString()}, New String() {"Failed", nFail.ToString()}, _
        New String() {"Escalated", nEsc.ToString()}}
    Dim fo As New JObject()
    fo("title") = kv(0)
    fo("value") = kv(1)
    facts.Add(fo)
Next
Dim cardBody As New JArray()
Dim t1 As New JObject()
t1("type") = "TextBlock" : t1("size") = "Large" : t1("weight") = "Bolder" : t1("wrap") = True
t1("text") = headline
t1("color") = If(nFail > 0, "Attention", If(nWarn > 0, "Warning", "Good"))
cardBody.Add(t1)
Dim t2 As New JObject()
t2("type") = "TextBlock" : t2("isSubtle") = True : t2("wrap") = True : t2("spacing") = "None"
t2("text") = project & " - blockchain transaction validation"
cardBody.Add(t2)
Dim fs As New JObject()
fs("type") = "FactSet"
fs("facts") = facts
cardBody.Add(fs)
For i As Integer = 0 To shown - 1
    Dim ft As New JObject()
    ft("type") = "TextBlock" : ft("wrap") = True : ft("separator") = (i = 0)
    ft("text") = String.Format("**{0} / {1}** - {2} ({3}): {4}", findings(i)(0), findings(i)(1), findings(i)(2), findings(i)(3), findings(i)(4))
    cardBody.Add(ft)
Next
If findings.Count > shown Then
    Dim more As New JObject()
    more("type") = "TextBlock" : more("isSubtle") = True
    more("text") = String.Format("... and {0} more - see the attached reports.", findings.Count - shown)
    cardBody.Add(more)
End If
Dim card As New JObject()
card("$schema") = "http://adaptivecards.io/schemas/adaptive-card.json"
card("type") = "AdaptiveCard"
card("version") = "1.4"
card("body") = cardBody
Dim attachment As New JObject()
attachment("contentType") = "application/vnd.microsoft.card.adaptive"
attachment("content") = card
Dim teams As New JObject()
teams("type") = "message"
teams("attachments") = New JArray(attachment)

' ---- Slack: Block Kit via an incoming webhook ---------------------------------------
Dim blocks As New JArray()
Dim hdr As New JObject()
hdr("type") = "header"
hdr("text") = New JObject(New JProperty("type", "plain_text"), New JProperty("text", Left(headline, 150)))
blocks.Add(hdr)
Dim fieldsArr As New JArray()
For Each kv As String() In New String()() { _
        New String() {"Validated", total.ToString()}, New String() {"Passed", nPass.ToString()}, _
        New String() {"Warnings", nWarn.ToString()}, New String() {"Failed", nFail.ToString()}, _
        New String() {"Escalated", nEsc.ToString()}, New String() {"Run", runId}}
    fieldsArr.Add(New JObject(New JProperty("type", "mrkdwn"), New JProperty("text", "*" & kv(0) & "*" & vbLf & kv(1))))
Next
blocks.Add(New JObject(New JProperty("type", "section"), New JProperty("fields", fieldsArr)))
If shown > 0 Then
    Dim lines As New StringBuilder()
    For i As Integer = 0 To shown - 1
        lines.AppendLine(String.Format("- *{0} / {1}* {2} ({3}): {4}", findings(i)(0), findings(i)(1), findings(i)(2), findings(i)(3), findings(i)(4)))
    Next
    If findings.Count > shown Then lines.AppendLine(String.Format("_... and {0} more_", findings.Count - shown))
    blocks.Add(New JObject(New JProperty("type", "divider")))
    blocks.Add(New JObject(New JProperty("type", "section"), _
        New JProperty("text", New JObject(New JProperty("type", "mrkdwn"), New JProperty("text", Left(lines.ToString(), 2900))))))
End If
Dim slack As New JObject()
slack("text") = headline & " - " & runId
slack("blocks") = blocks

out_TeamsJson = teams.ToString(Newtonsoft.Json.Formatting.None)
out_SlackJson = slack.ToString(Newtonsoft.Json.Formatting.None)

' Always keep a copy beside the reports: the evidence of what was (or would be) posted.
Dim folder As String = Convert.ToString(in_Config("ReportsFolder"))
Directory.CreateDirectory(folder)
File.WriteAllText(Path.Combine(folder, "Alert_" & runId & "_teams.json"), teams.ToString(Newtonsoft.Json.Formatting.Indented), Encoding.UTF8)
File.WriteAllText(Path.Combine(folder, "Alert_" & runId & "_slack.json"), slack.ToString(Newtonsoft.Json.Formatting.Indented), Encoding.UTF8)
out_Summary = String.Format("Teams and Slack payloads composed ({0} finding(s), {1} listed) and saved beside the reports", findings.Count, shown)
'''.strip()


def _post_webhook(channel: str, url_key: str, body_var: str) -> str:
    url = f'Convert.ToString(in_Config("{url_key}")).Trim()'
    post = sequence(
        f"Post to {channel}",
        f'<ui:HttpClient DisplayName="{lit("HTTP Request - " + channel + " webhook")}" '
        f'EndPoint="{vb(url)}" Method="POST" AcceptFormat="ANY" BodyFormat="application/json" '
        f'Body="{vb(body_var)}" TimeoutMS="{vb("30000")}" Result="{vb("response")}" '
        f'StatusCode="{vb("statusCode")}" />'
        + if_("statusCode >= 200 AndAlso statusCode < 300",
              log(f'"[ALERT/{channel.upper()}] Posted - HTTP " & statusCode.ToString()'),
              log(f'"[ALERT/{channel.upper()}] Webhook rejected the message - HTTP " & statusCode.ToString() & " " & response',
                  level="Warn"),
              name=f"If - {channel} accepted"))
    return if_(
        f'{url} <> ""',
        if_("in_ShouldSend", post,
            log(f'"[ALERT/{channel.upper()}] Nothing failed or warned - no message posted."'),
            name=f"If - anything to report to {channel}"),
        log(f'"[ALERT/{channel.upper()}] {url_key} is blank - payload saved to Reports instead of posted."'),
        name=f"If - {url_key} configured")


def alert_webhooks() -> str:
    body = sequence(
        "09b - Teams and Slack alerts",
        invoke_code(WEBHOOK_COMPOSE_CODE, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtExceptions", "sd:DataTable", "in_dtExceptions"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("Out", "out_TeamsJson", "x:String", "teamsJson"),
            ("Out", "out_SlackJson", "x:String", "slackJson"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - compose Teams and Slack messages")
        + log('"[ALERT/WEBHOOK] " & summary')
        + _post_webhook("Teams", "TeamsWebhookUrl", "teamsJson")
        + _post_webhook("Slack", "SlackWebhookUrl", "slackJson"),
        variables(("x:String", "teamsJson"), ("x:String", "slackJson"), ("x:String", "summary"),
                  ("x:String", "response"), ("x:Int32", "statusCode")))

    return workflow("09b_Alert_Webhooks", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtExceptions", "InArgument(sd:DataTable)"),
        ("in_Stats", f"InArgument({DICT_SO})"),
        ("in_ShouldSend", "InArgument(x:Boolean)"),
    ])


# ==========================================================================
# Queue/Dispatcher + Queue/Performer  -  Phase 5, high-volume processing
# ==========================================================================
# The queue work-unit is a SHIPMENT, not a transaction. Three of the six rules compare
# rows against one another - duplicates (R4), milestone ordering (R5) and required
# events (R6) - so a performer handed a single transaction could not evaluate them at
# all. A shipment is the smallest unit that keeps every rule meaningful, and none of the
# six rules reach across shipments, so nothing is lost by splitting there.
#
# The work queue is a local JSON file - one entry per shipment - so the whole
# dispatch / perform split runs from UiPath Studio with no Orchestrator.
def queue_dispatcher() -> str:
    code = r'''
' Group the joined rows by shipment and serialise each group into one queue payload.
Dim byShipment As New Dictionary(Of String, JArray)(StringComparer.OrdinalIgnoreCase)
Dim colNames As New List(Of String)
For Each c As DataColumn In in_dtJoined.Columns
    colNames.Add(c.ColumnName)
Next

For Each r As DataRow In in_dtJoined.Rows
    Dim sid As String = Convert.ToString(r("ShipmentID"))
    If Not byShipment.ContainsKey(sid) Then byShipment(sid) = New JArray()
    Dim o As New JObject()
    For Each cn As String In colNames
        If IsDBNull(r(cn)) Then
            o(cn) = Nothing
        ElseIf TypeOf r(cn) Is DateTime Then
            o(cn) = Convert.ToDateTime(r(cn)).ToString("o", CultureInfo.InvariantCulture)
        Else
            o(cn) = Convert.ToString(r(cn))
        End If
    Next
    byShipment(sid).Add(o)
Next

' One row per shipment: the queue payload plus the metadata the performer needs.
Dim dt As New DataTable("QueuePayloads")
dt.Columns.Add("Reference", GetType(String))
dt.Columns.Add("ShipmentID", GetType(String))
dt.Columns.Add("EventCount", GetType(Int32))
dt.Columns.Add("EventsJson", GetType(String))

Dim runId As String = Convert.ToString(in_Config("RunId"))
For Each kv As KeyValuePair(Of String, JArray) In byShipment
    dt.Rows.Add(kv.Key & "|" & runId, kv.Key, kv.Value.Count, kv.Value.ToString(Newtonsoft.Json.Formatting.None))
Next

out_dtPayloads = dt
out_Summary = String.Format("{0} shipment payload(s) prepared from {1} transaction(s)", _
    dt.Rows.Count, in_dtJoined.Rows.Count)
'''.strip()

    queue_file_code = r'''
' Write the work queue: one entry per shipment, for the Performer to process.
Dim folder As String = Path.Combine(Convert.ToString(in_Config("ReportsFolder")), "WorkQueue")
Directory.CreateDirectory(folder)
Dim arr As New JArray()
For Each r As DataRow In in_dtPayloads.Rows
    Dim o As New JObject()
    o("Reference") = Convert.ToString(r("Reference"))
    o("ShipmentID") = Convert.ToString(r("ShipmentID"))
    o("EventCount") = Convert.ToInt32(r("EventCount"))
    o("EventsJson") = Convert.ToString(r("EventsJson"))
    arr.Add(o)
Next
Dim outFile As String = Path.Combine(folder, "queue_payloads_" & Convert.ToString(in_Config("RunId")) & ".json")
File.WriteAllText(outFile, arr.ToString(Newtonsoft.Json.Formatting.Indented), Encoding.UTF8)
out_Summary = String.Format("work queue: {0} item(s) written to {1}", arr.Count, Path.GetFileName(outFile))
'''.strip()

    body = sequence(
        "Dispatcher - split the run into per-shipment queue items",
        invoke_code(code, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtJoined", "sd:DataTable", "in_dtJoined"),
            ("Out", "out_dtPayloads", "sd:DataTable", "dtPayloads"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - group transactions by shipment")
        + log('"[DISPATCH] " & summary')
        + invoke_code(queue_file_code, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtPayloads", "sd:DataTable", "dtPayloads"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - write the work queue file")
        + log('"[DISPATCH] " & summary')
        + assign("out_ItemsQueued", "dtPayloads.Rows.Count", type_ref="x:Int32",
                 name="Assign - items queued"),
        variables(("sd:DataTable", "dtPayloads"), ("x:String", "summary")))

    return workflow("Dispatcher", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtJoined", "InArgument(sd:DataTable)"),
        ("out_ItemsQueued", "OutArgument(x:Int32)"),
    ])


STATS_FROM_RESULTS_CODE = r"""
' The job's statistics, recomputed from every shipment it validated - the same figures
' the validation engine produces for a batch run, so reports read identically.
Dim ruleName As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
For Each rr As DataRow In in_dtRules.Rows
    ruleName(Convert.ToString(rr("RuleID")).Trim()) = Convert.ToString(rr("RuleName")).Trim()
Next
Dim nPass As Integer = 0, nWarn As Integer = 0, nFail As Integer = 0, nEsc As Integer = 0
Dim ruleHits As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
For Each r As DataRow In in_dtResults.Rows
    Select Case Convert.ToString(r("ValidationStatus"))
        Case "PASS" : nPass += 1
        Case "WARNING" : nWarn += 1
        Case "FAIL" : nFail += 1
    End Select
    If Convert.ToBoolean(r("RequiresEscalation")) Then nEsc += 1
    For Each rid As String In New String() {"R1", "R2", "R3", "R4", "R5", "R6"}
        Dim st As String = Convert.ToString(r(rid & "_Status"))
        If st = "FAIL" OrElse st = "WARNING" Then
            If Not ruleHits.ContainsKey(rid) Then ruleHits(rid) = 0
            ruleHits(rid) = ruleHits(rid) + 1
        End If
    Next
Next
Dim stats As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
Dim total As Integer = in_dtResults.Rows.Count
stats("Total") = total
stats("Pass") = nPass
stats("Warning") = nWarn
stats("Fail") = nFail
stats("Escalations") = nEsc
stats("PassPercent") = If(total = 0, 0.0, Math.Round(nPass * 100.0 / total, 1))
Dim hits As New List(Of String)
For Each rid As String In New String() {"R1", "R2", "R3", "R4", "R5", "R6"}
    Dim c As Integer = 0
    If ruleHits.ContainsKey(rid) Then c = ruleHits(rid)
    stats(rid & "_Hits") = c
    If c > 0 Then hits.Add(rid & "(" & If(ruleName.ContainsKey(rid), ruleName(rid), rid) & ")=" & c.ToString())
Next
stats("RuleHitSummary") = String.Join(", ", hits)
out_Stats = stats
out_Summary = String.Format("total={0} pass={1} warn={2} fail={3} escalations={4} | {5}", _
    total, nPass, nWarn, nFail, nEsc, String.Join(", ", hits))
""".strip()

MERGE_RESULTS_CODE = r"""
' Accumulate one shipment's verdicts into the job-wide results table.
If io_dtAll Is Nothing Then
    io_dtAll = in_dtResults.Clone()
    io_dtAll.TableName = "ValidationResults"
End If
For Each r As DataRow In in_dtResults.Rows
    io_dtAll.ImportRow(r)
Next
""".strip()

WORK_QUEUE_CODE = r"""
' Read the work queue the Dispatcher wrote. QueueFile picks a file; blank takes the newest.
Dim payloadFile As String = If(in_Config.ContainsKey("QueueFile"), Convert.ToString(in_Config("QueueFile")).Trim(), "")
If payloadFile = "" Then
    Dim folder As String = Path.Combine(Convert.ToString(in_Config("ReportsFolder")), "WorkQueue")
    If Directory.Exists(folder) Then
        Dim newest As FileInfo = New DirectoryInfo(folder).GetFiles("queue_payloads_*.json").OrderByDescending(Function(f) f.LastWriteTimeUtc).FirstOrDefault()
        If newest IsNot Nothing Then payloadFile = newest.FullName
    End If
End If
If payloadFile = "" OrElse Not IO.File.Exists(payloadFile) Then
    Throw New IO.FileNotFoundException("No work queue file was found. Run Main with RunMode = DISPATCH first, or set QueueFile.")
End If
Dim dt As New DataTable("QueuePayloads")
dt.Columns.Add("Reference", GetType(String))
dt.Columns.Add("ShipmentID", GetType(String))
dt.Columns.Add("EventsJson", GetType(String))
For Each tok As JObject In JArray.Parse(IO.File.ReadAllText(payloadFile))
    dt.Rows.Add(Convert.ToString(tok("Reference")), Convert.ToString(tok("ShipmentID")), Convert.ToString(tok("EventsJson")))
Next
out_dtPayloads = dt
out_Summary = String.Format("{0} payload(s) read from {1}", dt.Rows.Count, Path.GetFileName(payloadFile))
""".strip()


def queue_performer() -> str:
    """
    The Performer is its own entry point: open Workflows\Queue\Performer.xaml in Studio and
    run it. It loads the configuration itself when started without arguments, works through
    the Dispatcher's work queue file one shipment at a time, and then runs the same finalise
    pipeline as a batch run - human review, hash-chained audit log, Excel and PDF reports,
    alerts and dashboard - over every shipment it processed.
    """
    rebuild = r"""
' Rehydrate one shipment's transactions back into the schema the validation engine
' expects. The engine is unchanged between batch and queue mode - only the source of
' its DataTable differs.
Dim arr As JArray = JArray.Parse(in_EventsJson)

Dim dt As New DataTable("JoinedTransactions")
Dim typed As New Dictionary(Of String, Type)
typed("RowNo") = GetType(Integer)
typed("BlockNumber") = GetType(Long)
typed("StepOrder") = GetType(Integer)
typed("OnChainQty") = GetType(Long)
typed("ChainTimestampUtc") = GetType(DateTime)
typed("TxFailed") = GetType(Boolean)

If arr.Count > 0 Then
    For Each prop As JProperty In CType(arr(0), JObject).Properties()
        Dim t As Type = GetType(Object)
        If typed.ContainsKey(prop.Name) Then t = typed(prop.Name)
        dt.Columns.Add(prop.Name, t)
    Next
End If

For Each tok As JObject In arr
    Dim rw As DataRow = dt.NewRow()
    For Each c As DataColumn In dt.Columns
        Dim v As JToken = tok(c.ColumnName)
        If v Is Nothing OrElse v.Type = JTokenType.Null Then
            rw(c.ColumnName) = DBNull.Value
        ElseIf c.DataType Is GetType(DateTime) Then
            rw(c.ColumnName) = DateTime.Parse(v.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind)
        ElseIf c.DataType Is GetType(Boolean) Then
            rw(c.ColumnName) = Convert.ToBoolean(v.ToString())
        ElseIf c.DataType Is GetType(Long) Then
            rw(c.ColumnName) = Convert.ToInt64(v.ToString())
        ElseIf c.DataType Is GetType(Integer) Then
            rw(c.ColumnName) = Convert.ToInt32(v.ToString())
        Else
            rw(c.ColumnName) = v.ToString()
        End If
    Next
    dt.Rows.Add(rw)
Next

out_dtJoined = dt
out_Summary = String.Format("{0} transaction(s) rehydrated", dt.Rows.Count)
""".strip()

    def validate_shipment(events_expr: str, shipment_expr: str) -> str:
        return sequence(
            "Validate one shipment",
            assign("shipmentId", shipment_expr, name="Assign - shipment")
            + log('"[PERFORM] " & shipmentId')
            + invoke_code(rebuild, [
                ("In", "in_EventsJson", "x:String", events_expr),
                ("Out", "out_dtJoined", "sd:DataTable", "dtJoined"),
                ("Out", "out_Summary", "x:String", "summary"),
            ], name="Invoke Code - rehydrate the shipment")
            + invoke_workflow("Workflows\\05_Validate_Engine.xaml", [
                ("In", "in_Config", DICT_SO, "Config"),
                ("In", "in_dtJoined", "sd:DataTable", "dtJoined"),
                ("In", "in_dtRules", "sd:DataTable", "dtRules"),
                ("In", "in_dtWallets", "sd:DataTable", "dtWallets"),
                ("Out", "out_dtResults", "sd:DataTable", "dtResults"),
                ("Out", "out_dtExceptions", "sd:DataTable", "dtExceptions"),
                ("Out", "out_Stats", DICT_SO, "Stats"),
            ], name="Invoke 05 - Validation engine")
            + invoke_code(MERGE_RESULTS_CODE, [
                ("In", "in_dtResults", "sd:DataTable", "dtResults"),
                ("InOut", "io_dtAll", "sd:DataTable", "dtAllResults"),
            ], name="Invoke Code - merge into the job's results")
            + assign("itemsProcessed", "itemsProcessed + 1", type_ref="x:Int32", name="Assign - count processed")
            + log('"[PERFORM] " & shipmentId & ": " & Convert.ToString(Stats("Fail")) '
                  '& " failure(s), " & Convert.ToString(Stats("Warning")) & " warning(s)"'))

    queue_loop = sequence(
        "Process the work queue file",
        invoke_code(WORK_QUEUE_CODE, [
            ("In", "in_Config", DICT_SO, "Config"),
            ("Out", "out_dtPayloads", "sd:DataTable", "dtPayloads"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - read the work queue")
        + log('"[PERFORM] Work queue - " & summary')
        + for_each_row("dtPayloads",
                       try_catch(
                           validate_shipment('Convert.ToString(CurrentPayload("EventsJson"))',
                                             'Convert.ToString(CurrentPayload("ShipmentID"))'),
                           sequence("Record the failure",
                                    log('"[PERFORM] " & shipmentId & " failed: " & exception.Message', level="Error")
                                    + assign("itemsFailed", "itemsFailed + 1", type_ref="x:Int32", name="Assign - count failed")),
                           name="Try Catch - payload processing"),
                       row_var="CurrentPayload", name="For Each Row - payloads"))

    finalise = sequence(
        "Finalise the job",
        invoke_code(STATS_FROM_RESULTS_CODE, [
            ("In", "in_dtResults", "sd:DataTable", "dtAllResults"),
            ("In", "in_dtRules", "sd:DataTable", "dtRules"),
            ("Out", "out_Stats", DICT_SO, "jobStats"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - job statistics")
        + log('"[PERFORM] " & summary')
        + _mark("Validate", "Config")
        + invoke_workflow("Workflows\\14_Finalise_Run.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("In", "in_dtRules", "sd:DataTable", "dtRules"),
            ("In", "in_Stats", DICT_SO, "jobStats"),
            ("InOut", "io_dtResults", "sd:DataTable", "dtAllResults"),
            ("Out", "out_dtExceptions", "sd:DataTable", "dtExceptions"),
            ("Out", "out_ReportPath", "x:String", "reportPath"),
            ("Out", "out_PdfPath", "x:String", "pdfPath"),
            ("Out", "out_DashboardPath", "x:String", "dashboardPath"),
            ("Out", "out_ChainValid", "x:Boolean", "chainValid"),
        ], name="Invoke 14 - Finalise run"))

    body = sequence(
        "Performer - validate shipments from the work queue",
        if_("in_Config Is Nothing",
            sequence("Standalone job - load the configuration",
                     assign("projectRoot", "Directory.GetCurrentDirectory()")
                     + invoke_workflow("Workflows\\00_Init_ReadConfig.xaml", [
                         ("In", "in_ConfigPath", "x:String", 'Path.Combine(projectRoot, "Data", "Config.xlsx")'),
                         ("In", "in_ProjectRoot", "x:String", "projectRoot"),
                         ("Out", "out_Config", DICT_SO, "Config"),
                         ("Out", "out_dtRules", "sd:DataTable", "dtRules"),
                         ("Out", "out_dtWallets", "sd:DataTable", "dtWallets"),
                     ], name="Invoke 00 - Initialise")),
            sequence("Configuration supplied by the caller",
                     assign("Config", "in_Config", type_ref=DICT_SO, name="Assign - config")
                     + assign("dtRules", "in_dtRules", type_ref="sd:DataTable", name="Assign - rules")
                     + assign("dtWallets", "in_dtWallets", type_ref="sd:DataTable", name="Assign - wallets")),
            name="If - started without arguments")
        + assign('Config("RunKind")', '"QUEUE"', type_ref="x:Object", name="Assign - run kind")
        + _mark("Init", "Config")
        + queue_loop
        + log('"[PERFORM] " & itemsProcessed.ToString() & " shipment(s) processed, " & itemsFailed.ToString() & " failed"')
        + if_("dtAllResults IsNot Nothing AndAlso dtAllResults.Rows.Count > 0",
              finalise,
              log('"[PERFORM] Nothing was processed - no reports produced."'),
              name="If - anything processed"),
        variables((DICT_SO, "Config"), ("sd:DataTable", "dtRules"), ("sd:DataTable", "dtWallets"),
                  ("x:String", "projectRoot"),
                  ("x:String", "shipmentId"), ("x:String", "summary"),
                  ("sd:DataTable", "dtPayloads"), ("sd:DataTable", "dtJoined"),
                  ("sd:DataTable", "dtResults"), ("sd:DataTable", "dtExceptions"),
                  ("sd:DataTable", "dtAllResults"), (DICT_SO, "Stats"), (DICT_SO, "jobStats"),
                  ("x:Int32", "itemsProcessed"), ("x:Int32", "itemsFailed"),
                  ("x:String", "reportPath"), ("x:String", "pdfPath"),
                  ("x:String", "dashboardPath"), ("x:Boolean", "chainValid")))

    return workflow("Performer", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtRules", "InArgument(sd:DataTable)"),
        ("in_dtWallets", "InArgument(sd:DataTable)"),
    ])


DASHBOARD_CODE = r'''
' Render a self-contained HTML dashboard for this run. No external assets, no server:
' the file opens with a double click and works offline, which is what a demo needs.
Dim st As New StringBuilder()

Dim total As Integer = Convert.ToInt32(in_Stats("Total"))
Dim nPass As Integer = Convert.ToInt32(in_Stats("Pass"))
Dim nWarn As Integer = Convert.ToInt32(in_Stats("Warning"))
Dim nFail As Integer = Convert.ToInt32(in_Stats("Fail"))
Dim nEsc As Integer = Convert.ToInt32(in_Stats("Escalations"))
Dim passPct As Double = Convert.ToDouble(in_Stats("PassPercent"))

' ---- Run history, read back out of the append-only audit log ----------
' The log spans every run the bot has ever done, so this is genuine history
' rather than a record this workflow keeps for itself.
Dim runIds As New List(Of String)
Dim runFirstSeen As New Dictionary(Of String, String)
Dim runTotal As New Dictionary(Of String, Integer)
Dim runPass As New Dictionary(Of String, Integer)
Dim runWarn As New Dictionary(Of String, Integer)
Dim runFail As New Dictionary(Of String, Integer)

If File.Exists(in_AuditPath) Then
    Dim lines() As String = File.ReadAllLines(in_AuditPath)
    For i As Integer = 1 To lines.Length - 1
        Dim p() As String = lines(i).Split(","c)
        If p.Length < 17 Then Continue For
        Dim rid As String = p(2)
        If Not runTotal.ContainsKey(rid) Then
            runIds.Add(rid)
            runFirstSeen(rid) = p(1)
            runTotal(rid) = 0 : runPass(rid) = 0 : runWarn(rid) = 0 : runFail(rid) = 0
        End If
        runTotal(rid) = runTotal(rid) + 1
        Select Case p(10)
            Case "PASS" : runPass(rid) = runPass(rid) + 1
            Case "WARNING" : runWarn(rid) = runWarn(rid) + 1
            Case "FAIL" : runFail(rid) = runFail(rid) + 1
        End Select
    Next
End If

' ---- Run metrics history (processing time and trends) -----------------------
' Written by 11_Dashboard_Summary one row per run, beside the audit log.
Dim histFile As String = Path.Combine(Convert.ToString(in_Config("AuditLogFolder")), _
    If(in_Config.ContainsKey("RunHistoryFile") AndAlso Convert.ToString(in_Config("RunHistoryFile")).Trim() <> "", _
       Convert.ToString(in_Config("RunHistoryFile")).Trim(), "RunHistory.csv"))
Dim hist As New List(Of String())
If File.Exists(histFile) Then
    Dim hl() As String = File.ReadAllLines(histFile)
    For i As Integer = 1 To hl.Length - 1
        Dim hp() As String = hl(i).Split(","c)
        If hp.Length >= 19 Then hist.Add(hp)
    Next
End If
Dim durationById As New Dictionary(Of String, Double)
For Each hp As String() In hist
    Dim dsec As Double
    If Double.TryParse(hp(2), NumberStyles.Float, CultureInfo.InvariantCulture, dsec) Then durationById(hp(0)) = dsec
Next
Dim thisRunId As String = Convert.ToString(in_Config("RunId"))
Dim durationText As String = If(durationById.ContainsKey(thisRunId), _
    durationById(thisRunId).ToString("0.0", CultureInfo.InvariantCulture) & "s", "-")

' Phase timings from the marks Main stamps into the config dictionary.
Dim startedAt As DateTime = Convert.ToDateTime(in_Config("RunStartedUtc"))
Dim phaseSpec() As String = {"Init|Read configuration", "ExtractChain|Extract blockchain data", _
    "ExtractErp|Extract ERP data", "Map|Preprocess and map", "Validate|Validation engine (R1-R6)", _
    "Escalate|Human-in-the-loop review", "AuditTrail|Audit log and verification", _
    "Reports|Excel and PDF reports", "Alerts|E-mail, Teams and Slack alerts"}
Dim phaseNames As New List(Of String)
Dim phaseSecs As New List(Of Double)
Dim prevMark As DateTime = startedAt
For Each spec As String In phaseSpec
    Dim parts() As String = spec.Split("|"c)
    If in_Config.ContainsKey("Mark_" & parts(0)) Then
        Dim at As DateTime = Convert.ToDateTime(in_Config("Mark_" & parts(0)))
        phaseNames.Add(parts(1))
        phaseSecs.Add(Math.Max(0.0, (at - prevMark).TotalSeconds))
        prevMark = at
    End If
Next

' ---- Document -------------------------------------------------------------
st.AppendLine("<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'>")
st.AppendLine("<title>Blockchain Validation Dashboard</title>")
st.AppendLine("<style>")
st.AppendLine(":root{--ink:#16202c;--muted:#68778a;--line:#dfe5ec;--bg:#f4f6f9;--card:#fff;")
st.AppendLine("--navy:#1F3864;--pass:#1a7f4b;--passbg:#e4f5ec;--warn:#8a6100;--warnbg:#fdf2d8;")
st.AppendLine("--fail:#b3261e;--failbg:#fdecea;}")
st.AppendLine("*{box-sizing:border-box}")
st.AppendLine("body{margin:0;background:var(--bg);color:var(--ink);")
st.AppendLine("font-family:'Segoe UI',system-ui,Arial,sans-serif;font-size:14px;line-height:1.5}")
st.AppendLine(".wrap{max-width:1180px;margin:0 auto;padding:26px 22px 60px}")
st.AppendLine("header{background:var(--navy);color:#fff;padding:22px 26px;border-radius:10px}")
st.AppendLine("header h1{margin:0 0 4px;font-size:21px;letter-spacing:.2px}")
st.AppendLine("header .meta{font-size:12.5px;opacity:.85}")
st.AppendLine("header .meta b{opacity:1}")
' Long hex values must wrap rather than run off the edge on a narrow window.
st.AppendLine("header .mono{word-break:break-all}")
st.AppendLine("h2{font-size:15px;margin:30px 0 12px;letter-spacing:.3px;text-transform:uppercase;color:var(--muted)}")
st.AppendLine(".cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-top:16px}")
st.AppendLine(".card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px}")
st.AppendLine(".card .n{font-size:30px;font-weight:600;line-height:1.1}")
st.AppendLine(".card .l{font-size:12px;color:var(--muted);margin-top:3px}")
st.AppendLine(".card.pass .n{color:var(--pass)} .card.warn .n{color:var(--warn)} .card.fail .n{color:var(--fail)}")
st.AppendLine(".bar{height:9px;border-radius:5px;background:var(--failbg);overflow:hidden;display:flex;margin-top:14px}")
st.AppendLine(".bar i{display:block;height:100%}")
st.AppendLine(".bar .p{background:#2e9e63} .bar .w{background:#e8b931} .bar .f{background:#d4453c}")
st.AppendLine(".banner{margin-top:16px;border-radius:10px;padding:13px 16px;font-size:13.5px;border:1px solid}")
st.AppendLine(".banner.ok{background:var(--passbg);border-color:#b7e3ca;color:var(--pass)}")
st.AppendLine(".banner.bad{background:var(--failbg);border-color:#f3c2be;color:var(--fail)}")
st.AppendLine(".banner code{font-size:12px;opacity:.8}")
st.AppendLine(".flow{display:flex;flex-wrap:wrap;gap:7px;align-items:center;background:var(--card);")
st.AppendLine("border:1px solid var(--line);border-radius:10px;padding:15px 16px}")
st.AppendLine(".step{background:#eef2f8;border:1px solid #d6e0ee;border-radius:7px;padding:6px 11px;font-size:12.5px}")
st.AppendLine(".step b{display:block;font-size:10.5px;color:var(--muted);font-weight:600}")
st.AppendLine(".arw{color:#9fb0c4}")
st.AppendLine("table{width:100%;border-collapse:collapse;background:var(--card);")
st.AppendLine("border:1px solid var(--line);border-radius:10px;overflow:hidden}")
st.AppendLine("th{background:#eef2f8;text-align:left;padding:9px 12px;font-size:12px;")
st.AppendLine("text-transform:uppercase;letter-spacing:.4px;color:#4a5b70}")
st.AppendLine("td{padding:9px 12px;border-top:1px solid var(--line);font-size:13px;vertical-align:top}")
st.AppendLine(".pill{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11.5px;font-weight:600}")
st.AppendLine(".pill.PASS{background:var(--passbg);color:var(--pass)}")
st.AppendLine(".pill.WARNING{background:var(--warnbg);color:var(--warn)}")
st.AppendLine(".pill.FAIL{background:var(--failbg);color:var(--fail)}")
st.AppendLine(".pill.off{background:#eceff3;color:var(--muted)}")
st.AppendLine(".mono{font-family:Consolas,monospace;font-size:12px}")
st.AppendLine(".muted{color:var(--muted)}")
st.AppendLine("footer{margin-top:34px;font-size:12px;color:var(--muted);text-align:center}")
st.AppendLine("@media print{body{background:#fff}.card,table,.flow{break-inside:avoid}}")
st.AppendLine("</style></head><body><div class='wrap'>")

' ---- Header ---------------------------------------------------------------
st.AppendLine("<header><h1>Blockchain Transaction Validation &mdash; " & _
    System.Net.WebUtility.HtmlEncode(Convert.ToString(in_Config("ProjectName"))) & "</h1>")
st.AppendLine("<div class='meta'>Run <b>" & Convert.ToString(in_Config("RunId")) & "</b>" & _
    " &middot; bot <b>" & Convert.ToString(in_Config("BotIdentity")) & "</b>" & _
    " on <b>" & Convert.ToString(in_Config("MachineName")) & "</b>" & _
    " &middot; source <b>" & Convert.ToString(in_Config("DataSourceMode")) & "</b>" & _
    " &middot; chain id <b>" & Convert.ToString(in_Config("ChainId")) & "</b></div>")
st.AppendLine("<div class='meta' style='margin-top:6px'>Contract <span class='mono'>" & _
    Convert.ToString(in_Config("ContractAddress")) & "</span> &middot; generated " & _
    DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") & " UTC</div></header>")

' ---- KPI cards ------------------------------------------------------------
st.AppendLine("<div class='cards'>")
st.AppendLine("<div class='card'><div class='n'>" & total & "</div><div class='l'>Transactions validated</div></div>")
st.AppendLine("<div class='card pass'><div class='n'>" & nPass & "</div><div class='l'>Passed</div></div>")
st.AppendLine("<div class='card warn'><div class='n'>" & nWarn & "</div><div class='l'>Warnings</div></div>")
st.AppendLine("<div class='card fail'><div class='n'>" & nFail & "</div><div class='l'>Failed</div></div>")
st.AppendLine("<div class='card'><div class='n'>" & nEsc & "</div><div class='l'>Escalated to a human</div></div>")
st.AppendLine("<div class='card'><div class='n'>" & passPct.ToString("0.0", CultureInfo.InvariantCulture) & _
    "%</div><div class='l'>Pass rate</div></div>")
st.AppendLine("<div class='card'><div class='n'>" & durationText & "</div><div class='l'>Processing time</div></div>")
st.AppendLine("</div>")

If total > 0 Then
    Dim wp As Double = nPass * 100.0 / total
    Dim ww As Double = nWarn * 100.0 / total
    Dim wf As Double = nFail * 100.0 / total
    st.AppendLine("<div class='bar'>" & _
        "<i class='p' style='width:" & wp.ToString("0.##", CultureInfo.InvariantCulture) & "%'></i>" & _
        "<i class='w' style='width:" & ww.ToString("0.##", CultureInfo.InvariantCulture) & "%'></i>" & _
        "<i class='f' style='width:" & wf.ToString("0.##", CultureInfo.InvariantCulture) & "%'></i></div>")
End If

' ---- Audit chain status ---------------------------------------------------
If in_ChainValid Then
    st.AppendLine("<div class='banner ok'><b>Audit trail verified.</b> " & _
        System.Net.WebUtility.HtmlEncode(in_ChainReport) & "</div>")
Else
    st.AppendLine("<div class='banner bad'><b>Audit trail FAILED verification.</b> " & _
        System.Net.WebUtility.HtmlEncode(in_ChainReport) & "</div>")
End If

' ---- Pipeline -------------------------------------------------------------
st.AppendLine("<h2>Automation pipeline</h2><div class='flow'>")
Dim steps() As String = {"00|Read config", "01|Extract chain", "01d|Field mapping", "02|Extract ERP", _
                         "03|Map and join", "05|Validate R1-R6", "06|Escalate", "10|Audit log", _
                         "12|Verify chain", "07|Excel report", "08|PDF audit report", "09|E-mail alert", _
                         "09b|Teams / Slack", "11|KPI summary", "13|Dashboard"}
For i As Integer = 0 To steps.Length - 1
    Dim parts() As String = steps(i).Split("|"c)
    st.Append("<div class='step'><b>" & parts(0) & "</b>" & parts(1) & "</div>")
    If i < steps.Length - 1 Then st.Append("<span class='arw'>&rsaquo;</span>")
Next
st.AppendLine("</div>")

' ---- Rules ----------------------------------------------------------------
st.AppendLine("<h2>Validation rules</h2><table><tr><th>Rule</th><th>Name</th><th>Status</th>" & _
    "<th>Severity</th><th>Findings</th><th>What it checks</th></tr>")
For Each r As DataRow In in_dtRules.Rows
    Dim rid As String = Convert.ToString(r("RuleID"))
    Dim enabled As Boolean = (Convert.ToString(r("Enabled")).Trim().ToUpperInvariant() = "TRUE")
    Dim hits As Integer = 0
    If in_Stats.ContainsKey(rid & "_Hits") Then hits = Convert.ToInt32(in_Stats(rid & "_Hits"))
    Dim pill As String = If(enabled, "<span class='pill PASS'>ENABLED</span>", "<span class='pill off'>DISABLED</span>")
    Dim hitCell As String = If(hits > 0, "<span class='pill FAIL'>" & hits & "</span>", "<span class='muted'>0</span>")
    st.AppendLine("<tr><td class='mono'>" & rid & "</td><td>" & Convert.ToString(r("RuleName")) & _
        "</td><td>" & pill & "</td><td>" & Convert.ToString(r("Severity")) & "</td><td>" & hitCell & _
        "</td><td class='muted'>" & System.Net.WebUtility.HtmlEncode(Convert.ToString(r("Description"))) & "</td></tr>")
Next
st.AppendLine("</table>")

' ---- Findings -------------------------------------------------------------
Dim findings As Integer = 0
Dim fb As New StringBuilder()
For Each r As DataRow In in_dtResults.Rows
    Dim status As String = Convert.ToString(r("ValidationStatus"))
    If status = "PASS" Then Continue For
    findings += 1
    Dim decision As String = Convert.ToString(r("HumanDecision"))
    Dim decCell As String = If(decision = "", "<span class='muted'>&mdash;</span>", _
        System.Net.WebUtility.HtmlEncode(decision) & "<br><span class='muted' style='font-size:11.5px'>" & _
        System.Net.WebUtility.HtmlEncode(Convert.ToString(r("HumanDecidedBy"))) & "</span>")
    fb.AppendLine("<tr><td class='mono'>" & Convert.ToString(r("ShipmentID")) & "</td><td>" & _
        Convert.ToString(r("EventName")) & "</td><td><span class='pill " & status & "'>" & status & _
        "</span></td><td>" & Convert.ToString(r("Severity")) & "</td><td>" & _
        System.Net.WebUtility.HtmlEncode(Convert.ToString(r("FailureReasons"))) & "</td><td>" & decCell & "</td></tr>")
Next

st.AppendLine("<h2>Findings (" & findings & ")</h2>")
If findings = 0 Then
    st.AppendLine("<div class='banner ok'>Every transaction passed all enabled rules.</div>")
Else
    st.AppendLine("<table><tr><th>Shipment</th><th>Milestone</th><th>Status</th><th>Severity</th>" & _
        "<th>Reason</th><th>Human decision</th></tr>")
    st.Append(fb.ToString())
    st.AppendLine("</table>")
End If

' ---- Processing time by phase ----------------------------------------------------
If phaseNames.Count > 0 Then
    Dim maxSec As Double = Math.Max(0.001, phaseSecs.Max())
    st.AppendLine("<h2>Processing time by phase</h2><table><tr><th>Phase</th><th>Seconds</th><th style='width:55%'></th></tr>")
    For i As Integer = 0 To phaseNames.Count - 1
        Dim pct As Double = phaseSecs(i) * 100.0 / maxSec
        st.AppendLine("<tr><td>" & System.Net.WebUtility.HtmlEncode(phaseNames(i)) & "</td><td class='mono'>" & _
            phaseSecs(i).ToString("0.00", CultureInfo.InvariantCulture) & "</td><td><div style='height:9px;border-radius:5px;background:#2e6ad1;width:" & _
            pct.ToString("0.#", CultureInfo.InvariantCulture) & "%'></div></td></tr>")
    Next
    st.AppendLine("</table>")
End If

' ---- Trends across runs -------------------------------------------------------------
If hist.Count > 0 Then
    Dim recent As List(Of String()) = hist.Skip(Math.Max(0, hist.Count - 12)).ToList()
    Dim chartW As Double = 760, chartH As Double = 170, padL As Double = 36, padB As Double = 22, padT As Double = 12
    Dim plotW As Double = chartW - padL - 10
    Dim plotH As Double = chartH - padB - padT
    Dim maxFail As Integer = 1
    For Each h As String() In recent
        maxFail = Math.Max(maxFail, Integer.Parse(h(9), CultureInfo.InvariantCulture) + Integer.Parse(h(8), CultureInfo.InvariantCulture))
    Next
    Dim stepX As Double = plotW / Math.Max(1, recent.Count)
    Dim svg As New StringBuilder()
    svg.Append("<svg viewBox='0 0 " & chartW & " " & chartH & "' width='100%' role='img' aria-label='Pass rate and findings per run'>")
    For Each g As Integer In New Integer() {0, 50, 100}
        Dim gy As Double = padT + plotH - plotH * g / 100.0
        svg.Append("<line x1='" & padL & "' x2='" & (chartW - 10) & "' y1='" & gy.ToString("0.#", CultureInfo.InvariantCulture) & "' y2='" & gy.ToString("0.#", CultureInfo.InvariantCulture) & "' stroke='#dfe5ec'/>")
        svg.Append("<text x='" & (padL - 6) & "' y='" & (gy + 4).ToString("0.#", CultureInfo.InvariantCulture) & "' font-size='10' fill='#68778a' text-anchor='end'>" & g & "%</text>")
    Next
    Dim points As New List(Of String)
    For i As Integer = 0 To recent.Count - 1
        Dim h As String() = recent(i)
        Dim cx As Double = padL + stepX * i + stepX / 2
        Dim nFailH As Integer = Integer.Parse(h(9), CultureInfo.InvariantCulture)
        Dim nWarnH As Integer = Integer.Parse(h(8), CultureInfo.InvariantCulture)
        Dim barW As Double = Math.Min(26, stepX * 0.5)
        Dim failH As Double = plotH * 0.45 * nFailH / maxFail
        Dim warnH As Double = plotH * 0.45 * nWarnH / maxFail
        svg.Append("<rect x='" & (cx - barW / 2).ToString("0.#", CultureInfo.InvariantCulture) & "' y='" & (padT + plotH - failH).ToString("0.#", CultureInfo.InvariantCulture) & "' width='" & barW.ToString("0.#", CultureInfo.InvariantCulture) & "' height='" & failH.ToString("0.#", CultureInfo.InvariantCulture) & "' fill='#d4453c' opacity='0.8'><title>" & nFailH & " failed</title></rect>")
        svg.Append("<rect x='" & (cx - barW / 2).ToString("0.#", CultureInfo.InvariantCulture) & "' y='" & (padT + plotH - failH - warnH).ToString("0.#", CultureInfo.InvariantCulture) & "' width='" & barW.ToString("0.#", CultureInfo.InvariantCulture) & "' height='" & warnH.ToString("0.#", CultureInfo.InvariantCulture) & "' fill='#e8b931' opacity='0.8'><title>" & nWarnH & " warning(s)</title></rect>")
        Dim passRate As Double = Double.Parse(h(11), CultureInfo.InvariantCulture)
        points.Add(cx.ToString("0.#", CultureInfo.InvariantCulture) & "," & (padT + plotH - plotH * passRate / 100.0).ToString("0.#", CultureInfo.InvariantCulture))
        svg.Append("<text x='" & cx.ToString("0.#", CultureInfo.InvariantCulture) & "' y='" & (chartH - 6) & "' font-size='10' fill='#68778a' text-anchor='middle'>#" & (hist.Count - recent.Count + i + 1) & "</text>")
    Next
    svg.Append("<polyline points='" & String.Join(" ", points) & "' fill='none' stroke='#1a7f4b' stroke-width='2.5'/>")
    For Each pt As String In points
        Dim xy() As String = pt.Split(","c)
        svg.Append("<circle cx='" & xy(0) & "' cy='" & xy(1) & "' r='3.5' fill='#1a7f4b'/>")
    Next
    svg.Append("</svg>")

    st.AppendLine("<h2>Trends across the last " & recent.Count & " run(s)</h2>")
    st.AppendLine("<div class='flow' style='display:block'>" & svg.ToString() & _
        "<div class='muted' style='font-size:12px;margin-top:6px'>Line: pass rate. Bars: failed (red) and warning (amber) transactions per run.</div></div>")

    st.AppendLine("<table style='margin-top:12px'><tr><th>Run</th><th>Source</th><th>Duration</th><th>Validated</th><th>Pass %</th>" & _
        "<th>R1</th><th>R2</th><th>R3</th><th>R4</th><th>R5</th><th>R6</th><th>Audit chain</th></tr>")
    For i As Integer = recent.Count - 1 To 0 Step -1
        Dim h As String() = recent(i)
        Dim ruleCells As New StringBuilder()
        For c As Integer = 12 To 17
            ruleCells.Append(If(h(c) = "0", "<td class='muted'>0</td>", "<td><span class='pill FAIL'>" & h(c) & "</span></td>"))
        Next
        st.AppendLine("<tr><td class='mono' style='font-size:11.5px'>" & System.Net.WebUtility.HtmlEncode(h(0)) & "</td><td>" & _
            System.Net.WebUtility.HtmlEncode(h(3) & " / " & h(5)) & "</td><td class='mono'>" & h(2) & "s</td><td>" & h(6) & _
            "</td><td>" & h(11) & "</td>" & ruleCells.ToString() & "<td>" & _
            If(h(18) = "TRUE", "<span class='pill PASS'>VALID</span>", "<span class='pill FAIL'>INVALID</span>") & "</td></tr>")
    Next
    st.AppendLine("</table>")
End If

' ---- Run history ----------------------------------------------------------
st.AppendLine("<h2>Run history (" & runIds.Count & " run(s) in the audit trail)</h2>")
If runIds.Count = 0 Then
    st.AppendLine("<p class='muted'>No audit history yet.</p>")
Else
    st.AppendLine("<table><tr><th>Run</th><th>Started (UTC)</th><th>Duration</th><th>Validated</th>" & _
        "<th>Passed</th><th>Warnings</th><th>Failed</th></tr>")
    Dim shown As Integer = 0
    For i As Integer = runIds.Count - 1 To 0 Step -1
        If shown >= 12 Then Exit For
        shown += 1
        Dim rid As String = runIds(i)
        Dim isCurrent As Boolean = (rid = Convert.ToString(in_Config("RunId")))
        Dim label As String = rid
        If isCurrent Then label = "<b>" & rid & "</b> <span class='pill PASS'>this run</span>"
        Dim started As String = runFirstSeen(rid)
        If started.Length >= 19 Then started = started.Substring(0, 19).Replace("T", " ")
        Dim dur As String = If(durationById.ContainsKey(rid), durationById(rid).ToString("0.0", CultureInfo.InvariantCulture) & "s", "-")
        st.AppendLine("<tr><td class='mono' style='font-size:11.5px'>" & label & "</td><td class='mono'>" & _
            started & "</td><td class='mono'>" & dur & "</td><td>" & runTotal(rid) & "</td><td>" & runPass(rid) & "</td><td>" & _
            runWarn(rid) & "</td><td>" & runFail(rid) & "</td></tr>")
    Next
    st.AppendLine("</table>")
End If

st.AppendLine("<footer>Generated by the BlockchainLogisticsValidator bot &middot; " & _
    "every figure above comes from this run's own outputs</footer>")
st.AppendLine("</div></body></html>")

Dim outPath As String = Path.Combine(Convert.ToString(in_Config("ReportsFolder")), "dashboard.html")
File.WriteAllText(outPath, st.ToString(), Encoding.UTF8)

' A stable copy as well as the per-run one, so a bookmark or a link keeps working.
Dim stamped As String = Path.Combine(Convert.ToString(in_Config("ReportsFolder")), _
    "dashboard_" & Convert.ToString(in_Config("RunId")) & ".html")
File.WriteAllText(stamped, st.ToString(), Encoding.UTF8)

out_DashboardPath = outPath
out_Summary = String.Format("dashboard.html written ({0} finding(s), {1} run(s) of history, processing time {2})", _
    findings, runIds.Count, durationText)
'''.strip()


# ==========================================================================
# 13_Report_Dashboard
# ==========================================================================
# A self-contained HTML dashboard, regenerated on every run. No server, no external
# assets and no CDN: it opens with a double click on any machine, which is what a
# demo actually needs. The run-history table is read back out of the append-only
# audit log rather than kept separately, so it cannot drift from the evidence.
def report_dashboard() -> str:
    body = sequence(
        "13 - Generate HTML dashboard",
        invoke_code(DASHBOARD_CODE, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtResults", "sd:DataTable", "in_dtResults"),
            ("In", "in_dtRules", "sd:DataTable", "in_dtRules"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("In", "in_AuditPath", "x:String", "in_AuditPath"),
            ("In", "in_ChainValid", "x:Boolean", "in_ChainValid"),
            ("In", "in_ChainReport", "x:String", "in_ChainReport"),
            ("Out", "out_DashboardPath", "x:String", "out_DashboardPath"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - render the dashboard")
        + log('"[DASHBOARD] " & summary'),
        variables(("x:String", "summary")))

    return workflow("13_Report_Dashboard", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtResults", "InArgument(sd:DataTable)"),
        ("in_dtRules", "InArgument(sd:DataTable)"),
        ("in_Stats", f"InArgument({DICT_SO})"),
        ("in_AuditPath", "InArgument(x:String)"),
        ("in_ChainValid", "InArgument(x:Boolean)"),
        ("in_ChainReport", "InArgument(x:String)"),
        ("out_DashboardPath", "OutArgument(x:String)"),
    ])


# ==========================================================================
# 14_Finalise_Run  -  everything that happens after the verdicts exist
# ==========================================================================
# Shared by BATCH (Main) and queue mode (Performer), so a queue job produces exactly the
# same escalation, audit trail, reports, alerts and dashboard as a batch run.
REFRESH_EXCEPTIONS_CODE = r"""
' Rebuild the exceptions table from the results AFTER human review, so the reports and
' alerts show the reviewer's decision next to each finding.
Dim ex As DataTable = in_dtResults.Clone()
ex.TableName = "Exceptions"
For Each r As DataRow In in_dtResults.Rows
    If Convert.ToString(r("ValidationStatus")) <> "PASS" Then ex.ImportRow(r)
Next
out_dtExceptions = ex
""".strip()


def _mark(phase: str, config_var: str) -> str:
    return assign(f'{config_var}("Mark_{phase}")', "DateTime.UtcNow", type_ref="x:Object",
                  name=f"Assign - mark end of {phase} phase")


def finalise_run() -> str:
    body = sequence(
        "14 - Finalise the run",
        # ---- Novelty 2: pause for a human on a high-severity anomaly --------------
        # Placed before the audit log on purpose: the reviewer's decision is hashed
        # into the chain, so it is as tamper-evident as the machine's own verdict.
        invoke_workflow("Workflows\\06_HumanInTheLoop_Escalation.xaml", [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("InOut", "io_dtResults", "sd:DataTable", "io_dtResults"),
            ("Out", "out_EscalationCount", "x:Int32", "escalationCount"),
        ], name="Invoke 06 - Human-in-the-loop escalation")
        + _mark("Escalate", "in_Config")
        # ---- Novelty 4: the audit log is written FIRST, then proven intact ---------
        # The chain records what the bot concluded; every report is derived from it.
        + invoke_workflow("Workflows\\10_AuditLog_Chained.xaml", [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtResults", "sd:DataTable", "io_dtResults"),
            ("Out", "out_AuditPath", "x:String", "auditPath"),
            ("Out", "out_HeadHash", "x:String", "headHash"),
        ], name="Invoke 10 - Append audit log")
        + invoke_workflow("Workflows\\12_Verify_AuditChain.xaml", [
            ("In", "in_AuditPath", "x:String", "auditPath"),
            ("Out", "out_IsValid", "x:Boolean", "out_ChainValid"),
            ("Out", "out_Report", "x:String", "chainReport"),
        ], name="Invoke 12 - Verify audit chain")
        + _mark("AuditTrail", "in_Config")
        + invoke_code(REFRESH_EXCEPTIONS_CODE, [
            ("In", "in_dtResults", "sd:DataTable", "io_dtResults"),
            ("Out", "out_dtExceptions", "sd:DataTable", "out_dtExceptions"),
        ], name="Invoke Code - refresh exceptions with review decisions")
        # ---- Phase 4: reports ------------------------------------------------------
        + invoke_workflow("Workflows\\07_Report_Excel.xaml", [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtResults", "sd:DataTable", "io_dtResults"),
            ("In", "in_dtExceptions", "sd:DataTable", "out_dtExceptions"),
            ("In", "in_dtRules", "sd:DataTable", "in_dtRules"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("Out", "out_ReportPath", "x:String", "out_ReportPath"),
        ], name="Invoke 07 - Excel report")
        + invoke_workflow("Workflows\\08_Report_PDF.xaml", [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtResults", "sd:DataTable", "io_dtResults"),
            ("In", "in_dtRules", "sd:DataTable", "in_dtRules"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("In", "in_AuditPath", "x:String", "auditPath"),
            ("In", "in_HeadHash", "x:String", "headHash"),
            ("In", "in_ChainValid", "x:Boolean", "out_ChainValid"),
            ("In", "in_ChainReport", "x:String", "chainReport"),
            ("In", "in_ReportPath", "x:String", "out_ReportPath"),
            ("Out", "out_PdfPath", "x:String", "out_PdfPath"),
        ], name="Invoke 08 - PDF audit report")
        + _mark("Reports", "in_Config")
        # ---- Alerts: e-mail, Teams, Slack ---------------------------------------------
        + invoke_workflow("Workflows\\09_Alert_Email.xaml", [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtExceptions", "sd:DataTable", "out_dtExceptions"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("In", "in_ReportPath", "x:String", "out_ReportPath"),
            ("In", "in_PdfPath", "x:String", "out_PdfPath"),
            ("In", "in_HeadHash", "x:String", "headHash"),
            ("In", "in_ChainReport", "x:String", "chainReport"),
            ("Out", "out_AlertPath", "x:String", "alertPath"),
        ], name="Invoke 09 - Anomaly alerts")
        + _mark("Alerts", "in_Config")
        # ---- Dashboard: KPI JSON + run history, then the HTML page ---------------------
        + invoke_workflow("Workflows\\11_Dashboard_Summary.xaml", [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtResults", "sd:DataTable", "io_dtResults"),
            ("In", "in_dtRules", "sd:DataTable", "in_dtRules"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("In", "in_HeadHash", "x:String", "headHash"),
            ("In", "in_ChainValid", "x:Boolean", "out_ChainValid"),
            ("Out", "out_DashboardPath", "x:String", "summaryPath"),
        ], name="Invoke 11 - KPI summary and run history")
        # Last, so it can show the processing time and history the step above recorded.
        + invoke_workflow("Workflows\\13_Report_Dashboard.xaml", [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtResults", "sd:DataTable", "io_dtResults"),
            ("In", "in_dtRules", "sd:DataTable", "in_dtRules"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("In", "in_AuditPath", "x:String", "auditPath"),
            ("In", "in_ChainValid", "x:Boolean", "out_ChainValid"),
            ("In", "in_ChainReport", "x:String", "chainReport"),
            ("Out", "out_DashboardPath", "x:String", "out_DashboardPath"),
        ], name="Invoke 13 - HTML dashboard")
        + if_("out_ChainValid",
              log('"[FINALISE] Outputs: " & Path.GetFileName(out_ReportPath) & ", " & Path.GetFileName(out_PdfPath) '
                  '& ", " & Path.GetFileName(out_DashboardPath) & " | audit chain verified"'),
              log('"[FINALISE] Audit chain verification FAILED - " & chainReport', level="Error"),
              name="If - audit chain verified"),
        variables(("x:Int32", "escalationCount"), ("x:String", "auditPath"), ("x:String", "headHash"),
                  ("x:String", "chainReport"), ("x:String", "alertPath"), ("x:String", "summaryPath")))

    return workflow("14_Finalise_Run", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtRules", "InArgument(sd:DataTable)"),
        ("in_Stats", f"InArgument({DICT_SO})"),
        ("io_dtResults", "InOutArgument(sd:DataTable)"),
        ("out_dtExceptions", "OutArgument(sd:DataTable)"),
        ("out_ReportPath", "OutArgument(x:String)"),
        ("out_PdfPath", "OutArgument(x:String)"),
        ("out_DashboardPath", "OutArgument(x:String)"),
        ("out_ChainValid", "OutArgument(x:Boolean)"),
    ])


# ==========================================================================
# Main
# ==========================================================================
FAILURE_SCREENSHOT_PATH_CODE = r'''
' Where to put the failure screenshot. Configuration may not have loaded - the failure
' might BE the configuration, as it is when Config.xlsx or a chain feed is missing - so
' this cannot rely on in_Config being populated.
'
' Falling straight back to the current directory is wrong for a packed run: the robot's
' working directory is then the extracted package under %USERPROFILE%\.nuget\packages,
' so the one piece of evidence a failed job leaves behind would be written into a
' disposable cache folder and lost on the next publish. So resolve the output root the
' same way 00_Init does, with the same precedence, without needing 00_Init to have run:
' BLV_OUTPUT_ROOT, then Data\Config.local.json, then the project default.
Dim folder As String = ""
If in_Config IsNot Nothing AndAlso in_Config.ContainsKey("ScreenshotFolder") Then folder = Convert.ToString(in_Config("ScreenshotFolder"))

If folder = "" Then
    Dim outRoot As String = Environment.GetEnvironmentVariable("BLV_OUTPUT_ROOT")

    If String.IsNullOrWhiteSpace(outRoot) Then
        ' Read OutputRoot out of the local override file without a JSON dependency:
        ' this code runs in the failure handler, where the fewer moving parts the better.
        Dim localCfg As String = Path.Combine(Directory.GetCurrentDirectory(), "Data", "Config.local.json")
        If File.Exists(localCfg) Then
            ' ChrW(34) rather than an escaped quote literal: this body is stored inside a
            ' XAML attribute, and nested quote escaping is unreadable there and in Studio.
            Dim dq As Char = ChrW(34)
            Dim txt As String = File.ReadAllText(localCfg)
            Dim k As Integer = txt.IndexOf(dq & "OutputRoot" & dq)
            If k >= 0 Then
                Dim c As Integer = txt.IndexOf(":"c, k)
                Dim q1 As Integer = -1
                If c >= 0 Then q1 = txt.IndexOf(dq, c + 1)
                Dim q2 As Integer = -1
                If q1 >= 0 Then q2 = txt.IndexOf(dq, q1 + 1)
                If q2 > q1 Then outRoot = txt.Substring(q1 + 1, q2 - q1 - 1).Replace("\\", "\")
            End If
        End If
    End If

    If String.IsNullOrWhiteSpace(outRoot) Then outRoot = Path.Combine(Directory.GetCurrentDirectory(), "Data", "Output")
    If Not Path.IsPathRooted(outRoot) Then outRoot = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), outRoot))
    folder = Path.Combine(outRoot, "Screenshots")
End If

Directory.CreateDirectory(folder)
out_Path = Path.Combine(folder, "Failure_" & DateTime.UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture) & ".png")
'''.strip()


def main_workflow() -> str:
    pipeline = sequence(
        "Run the validation pipeline",
        assign("projectRoot", "Directory.GetCurrentDirectory()")
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
        ], name="Invoke 00 - Initialise")
        + _mark("Init", "Config")
        # ---- Phase 1: data acquisition ---------------------------------------------
        + invoke_workflow("Workflows\\01_Extract_Blockchain.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("In", "in_dtMapping", "sd:DataTable", "dtMapping"),
            ("In", "in_dtSignatures", "sd:DataTable", "dtSignatures"),
            ("Out", "out_dtChain", "sd:DataTable", "dtChain"),
        ], name="Invoke 01 - Extract blockchain")
        + _mark("ExtractChain", "Config")
        + invoke_workflow("Workflows\\02_Extract_Logistics.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("Out", "out_dtShipments", "sd:DataTable", "dtShipments"),
            ("Out", "out_dtErpEvents", "sd:DataTable", "dtErpEvents"),
        ], name="Invoke 02 - Extract logistics")
        + _mark("ExtractErp", "Config")
        # ---- Phase 2: preprocessing and mapping --------------------------------------
        + invoke_workflow("Workflows\\03_Preprocess_Map.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("In", "in_dtChain", "sd:DataTable", "dtChain"),
            ("In", "in_dtShipments", "sd:DataTable", "dtShipments"),
            ("In", "in_dtErpEvents", "sd:DataTable", "dtErpEvents"),
            ("In", "in_dtSequence", "sd:DataTable", "dtSequence"),
            ("Out", "out_dtJoined", "sd:DataTable", "dtJoined"),
        ], name="Invoke 03 - Preprocess and map")
        + _mark("Map", "Config")
        + log('"[MAIN] Joined rows ready: " & dtJoined.Rows.Count.ToString()')
        # BATCH validates the whole run in one job. DISPATCH instead splits it into a
        # work queue file, one item per shipment, for the Performer to process.
        + if_('Convert.ToString(Config("RunMode")).Trim().ToUpperInvariant() = "DISPATCH"',
              sequence("Dispatch mode - split into queue items",
                       invoke_workflow("Workflows\\Queue\\Dispatcher.xaml", [
                           ("In", "in_Config", DICT_SO, "Config"),
                           ("In", "in_dtJoined", "sd:DataTable", "dtJoined"),
                           ("Out", "out_ItemsQueued", "x:Int32", "itemsQueued"),
                       ], name="Invoke Dispatcher - enqueue per shipment")
                       + log('"[MAIN] Dispatch complete: " & itemsQueued.ToString() '
                             '& " shipment(s) queued for the Performer."')),
              sequence("Batch mode - validate this run end to end",
                       # ---- Phase 3: validation engine -----------------------------
                       invoke_workflow("Workflows\\05_Validate_Engine.xaml", [
                           ("In", "in_Config", DICT_SO, "Config"),
                           ("In", "in_dtJoined", "sd:DataTable", "dtJoined"),
                           ("In", "in_dtRules", "sd:DataTable", "dtRules"),
                           ("In", "in_dtWallets", "sd:DataTable", "dtWallets"),
                           ("Out", "out_dtResults", "sd:DataTable", "dtResults"),
                           ("Out", "out_dtExceptions", "sd:DataTable", "dtExceptions"),
                           ("Out", "out_Stats", DICT_SO, "Stats"),
                       ], name="Invoke 05 - Validation engine")
                       + _mark("Validate", "Config")
                       + log('"[MAIN] Validation complete: " & Convert.ToString(Stats("Fail")) & " failed, " '
                             '& Convert.ToString(Stats("Warning")) & " warning, " '
                             '& Convert.ToString(Stats("Pass")) & " passed ("'
                             ' & Convert.ToString(Stats("PassPercent")) & "% pass rate)"')
                       # ---- Novelty 2 + 4, Phase 4: review, audit, reports, alerts ---
                       + invoke_workflow("Workflows\\14_Finalise_Run.xaml", [
                           ("In", "in_Config", DICT_SO, "Config"),
                           ("In", "in_dtRules", "sd:DataTable", "dtRules"),
                           ("In", "in_Stats", DICT_SO, "Stats"),
                           ("InOut", "io_dtResults", "sd:DataTable", "dtResults"),
                           ("Out", "out_dtExceptions", "sd:DataTable", "dtExceptions"),
                           ("Out", "out_ReportPath", "x:String", "reportPath"),
                           ("Out", "out_PdfPath", "x:String", "pdfPath"),
                           ("Out", "out_DashboardPath", "x:String", "dashboardHtmlPath"),
                           ("Out", "out_ChainValid", "x:Boolean", "chainValid"),
                       ], name="Invoke 14 - Finalise run")),
              name="If - dispatch or batch"))

    # Any failure is logged with a screenshot of the desktop at that moment, then re-thrown
    # so Studio reports the run as failed.
    screenshot = try_catch(
        sequence("Capture the screen",
                 invoke_code(FAILURE_SCREENSHOT_PATH_CODE, [
                     ("In", "in_Config", DICT_SO, "Config"),
                     ("Out", "out_Path", "x:String", "screenshotPath"),
                 ], name="Invoke Code - failure screenshot path")
                 + f'<ui:TakeScreenshot DisplayName="{lit("Take Screenshot - failure")}" Screenshot="{vb("failureImage")}" />'
                 + f'<ui:SaveImage DisplayName="{lit("Save Image - failure")}" FileName="{vb("screenshotPath")}" Image="{vb("failureImage")}" />'
                 + log('"[MAIN] Failure screenshot saved: " & screenshotPath')),
        log('"[MAIN] Could not capture a failure screenshot: " & shotError.Message', level="Warn"),
        name="Try Catch - failure screenshot", ex_var="shotError")

    on_failure = sequence(
        "Record the failure and re-throw",
        log('"[MAIN] Run failed: " & exception.GetType().Name & ": " & exception.Message', level="Error")
        + screenshot
        + '<Rethrow DisplayName="Rethrow" />')

    body = sequence(
        "Main",
        log('"=== BlockchainLogisticsValidator starting ==="')
        + try_catch(pipeline, on_failure, name="Try Catch - whole run")
        + log('"=== BlockchainLogisticsValidator finished ==="'),
        variables(
            ("x:String", "projectRoot"), ("x:String", "configPath"),
            (DICT_SO, "Config"), (DICT_SO, "Stats"),
            ("sd:DataTable", "dtRules"), ("sd:DataTable", "dtWallets"),
            ("sd:DataTable", "dtMapping"), ("sd:DataTable", "dtSequence"),
            ("sd:DataTable", "dtSignatures"), ("sd:DataTable", "dtChain"),
            ("sd:DataTable", "dtShipments"), ("sd:DataTable", "dtErpEvents"),
            ("sd:DataTable", "dtJoined"), ("sd:DataTable", "dtResults"),
            ("sd:DataTable", "dtExceptions"),
            ("x:String", "reportPath"), ("x:String", "pdfPath"),
            ("x:Boolean", "chainValid"), ("x:Int32", "itemsQueued"),
            ("x:String", "dashboardHtmlPath"), ("x:String", "screenshotPath"),
            ("ui:Image", "failureImage")))

    return workflow("Main", body)


# ==========================================================================
GENERATORS = {
    os.path.join(PROJ, "Main.xaml"): main_workflow,
    os.path.join(WF, "00_Init_ReadConfig.xaml"): init_read_config,
    os.path.join(WF, "01_Extract_Blockchain.xaml"): extract_blockchain,
    os.path.join(WF, "01a_Extract_FromMockJson.xaml"): extract_from_mock,
    os.path.join(WF, "01b_Extract_FromEtherscanApi.xaml"): extract_from_api,
    os.path.join(WF, "01c_Extract_FromExplorerUI.xaml"): extract_from_ui,
    os.path.join(WF, "01d_Apply_FieldMapping.xaml"): apply_field_mapping,
    os.path.join(WF, "02_Extract_Logistics.xaml"): extract_logistics,
    os.path.join(WF, "03_Preprocess_Map.xaml"): preprocess_map,
    os.path.join(WF, "05_Validate_Engine.xaml"): validate_engine,
    os.path.join(WF, "06_HumanInTheLoop_Escalation.xaml"): human_in_the_loop,
    os.path.join(WF, "07_Report_Excel.xaml"): report_excel,
    os.path.join(WF, "08_Report_PDF.xaml"): report_pdf,
    os.path.join(WF, "09_Alert_Email.xaml"): alert_email,
    os.path.join(WF, "09b_Alert_Webhooks.xaml"): alert_webhooks,
    os.path.join(WF, "10_AuditLog_Chained.xaml"): audit_log_chained,
    os.path.join(WF, "11_Dashboard_Summary.xaml"): dashboard_summary,
    os.path.join(WF, "12_Verify_AuditChain.xaml"): verify_audit_chain,
    os.path.join(WF, "13_Report_Dashboard.xaml"): report_dashboard,
    os.path.join(WF, "14_Finalise_Run.xaml"): finalise_run,
    os.path.join(QUEUE, "Dispatcher.xaml"): queue_dispatcher,
    os.path.join(QUEUE, "Performer.xaml"): queue_performer,
    os.path.join(RULES, "R1_TimestampCheck.xaml"): r1_timestamp,
    os.path.join(RULES, "R2_QuantityMatch.xaml"): r2_quantity,
    os.path.join(RULES, "R3_AddressWhitelist.xaml"): r3_whitelist,
    os.path.join(RULES, "R4_DuplicateDetection.xaml"): r4_duplicate,
    os.path.join(RULES, "R5_SequenceValidation.xaml"): r5_sequence,
    os.path.join(RULES, "R6_SmartContractEvent.xaml"): r6_contract_event,
}


# Entry points that can be run on their own from Studio. The Performer processes the
# work queue file; 12_Verify_AuditChain lets an auditor check the trail without running a
# validation. Test-case entry points are managed by build_tests.py.
PROCESS_ENTRY_POINTS = [
    "Main.xaml",
    "Workflows\\Queue\\Performer.xaml",
    "Workflows\\12_Verify_AuditChain.xaml",
]


def register_entry_points() -> None:
    import json
    import uuid

    path = os.path.join(PROJ, "project.json")
    with open(path, encoding="utf-8-sig") as f:
        data = json.load(f)

    eps = data.setdefault("entryPoints", [])
    known = {e.get("filePath") for e in eps}
    process = [e for e in eps if e.get("filePath") in PROCESS_ENTRY_POINTS]
    others = [e for e in eps if e.get("filePath") not in PROCESS_ENTRY_POINTS]
    for rel in PROCESS_ENTRY_POINTS:
        if rel not in known:
            process.append({
                "filePath": rel,
                "uniqueId": str(uuid.uuid5(uuid.NAMESPACE_URL, "blv-process/" + rel)),
                "input": [],
                "output": [],
            })
    process.sort(key=lambda e: PROCESS_ENTRY_POINTS.index(e["filePath"]))
    data["entryPoints"] = process + others

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
        f.write("\n")
    print("registered process entry points:", ", ".join(PROCESS_ENTRY_POINTS))


def main() -> None:
    os.makedirs(WF, exist_ok=True)
    os.makedirs(RULES, exist_ok=True)
    os.makedirs(QUEUE, exist_ok=True)
    for path, gen in GENERATORS.items():
        with open(path, "w", encoding="utf-8") as f:
            f.write(gen())
        print("wrote:", os.path.relpath(path, ROOT))
    register_entry_points()


if __name__ == "__main__":
    main()

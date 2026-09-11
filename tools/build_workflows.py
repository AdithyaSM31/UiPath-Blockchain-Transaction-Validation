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

QUEUE_NAME_EXPR = 'Convert.ToString(in_Config("QueueName"))'
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

' ---- Resolve relative paths against the project root --------------------
' Paths are stored relative in Config.xlsx so the project stays portable; at
' runtime the working directory is the deployed package's content folder.
Dim pathKeys() As String = {"MockChainFile", "LogisticsFile", "TxShipmentMapFile"}
For Each key As String In pathKeys
    If d.ContainsKey(key) Then
        Dim v As String = Convert.ToString(d(key)).Trim()
        If v <> "" AndAlso Not Path.IsPathRooted(v) Then
            d(key) = Path.GetFullPath(Path.Combine(in_ProjectRoot, v))
        End If
    End If
Next

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
d("ReportsFolder") = Path.Combine(outputRoot, "Reports")
d("AuditLogFolder") = Path.Combine(outputRoot, "AuditLogs")
d("ScreenshotFolder") = Path.Combine(outputRoot, "Screenshots")

' ---- Make sure output folders exist before anything tries to write ------
For Each key As String In New String() {"ReportsFolder", "AuditLogFolder", "ScreenshotFolder"}
    If d.ContainsKey(key) Then
        Dim v As String = Convert.ToString(d(key)).Trim()
        If v <> "" Then Directory.CreateDirectory(v)
    End If
Next

' ---- Fail fast on missing or contradictory configuration ----------------
Dim missing As New List(Of String)
For Each key As String In New String() {"DataSourceMode", "LogisticsFile", "BotIdentity", "ContractAddress"}
    If Not d.ContainsKey(key) OrElse Convert.ToString(d(key)).Trim() = "" Then
        missing.Add(key)
    End If
Next
If missing.Count > 0 Then
    Throw New InvalidOperationException("Config.xlsx is missing required setting(s): " & String.Join(", ", missing))
End If

Dim mode As String = Convert.ToString(d("DataSourceMode")).Trim().ToUpperInvariant()
d("DataSourceMode") = mode
If Array.IndexOf(New String() {"MOCK", "API", "WEB"}, mode) < 0 Then
    Throw New InvalidOperationException("DataSourceMode must be MOCK, API or WEB. Found: " & mode)
End If
If mode = "API" AndAlso Convert.ToString(d("EtherscanApiKey")).Trim() = "" Then
    Throw New InvalidOperationException("DataSourceMode is API but EtherscanApiKey is blank in Config.xlsx.")
End If

Dim srcFile As String = ""
If mode = "MOCK" Then srcFile = Convert.ToString(d("MockChainFile"))
If mode = "MOCK" AndAlso Not File.Exists(srcFile) Then
    Throw New IO.FileNotFoundException("Mock chain file not found: " & srcFile)
End If
If Not File.Exists(Convert.ToString(d("LogisticsFile"))) Then
    Throw New IO.FileNotFoundException("Logistics workbook not found: " & Convert.ToString(d("LogisticsFile")))
End If

' ---- Stamp this run -----------------------------------------------------
d("RunId") = DateTime.UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'") & "-" & Convert.ToString(d("BotIdentity"))
d("RunStartedUtc") = DateTime.UtcNow
d("MachineName") = Environment.MachineName
d("WindowsUser") = Environment.UserName
d("ProjectRoot") = in_ProjectRoot

out_Config = d
out_Summary = String.Format("mode={0} rules={1} wallets={2} milestones={3} runId={4}", _
    mode, in_dtRules.Rows.Count, in_dtWallets.Rows.Count, in_dtSequence.Rows.Count, d("RunId"))
'''.strip()

    excel_body = sequence("Read configuration sheets",
        excel_read("Settings", "dtSettings")
        + excel_read("ValidationRules", "out_dtRules")
        + excel_read("ApprovedWallets", "out_dtWallets")
        + excel_read("FieldMapping", "out_dtMapping")
        + excel_read("EventSequence", "out_dtSequence")
        + excel_read("EventSignatures", "out_dtSignatures"))

    body = sequence(
        "00 - Initialise and read configuration",
        log('"[INIT] Reading configuration from " & in_ConfigPath')
        + excel_scope("in_ConfigPath", excel_body, name="Excel Application Scope - Config.xlsx")
        + invoke_code(code, [
            ("In", "in_dtSettings", "sd:DataTable", "dtSettings"),
            ("In", "in_dtRules", "sd:DataTable", "out_dtRules"),
            ("In", "in_dtWallets", "sd:DataTable", "out_dtWallets"),
            ("In", "in_dtSequence", "sd:DataTable", "out_dtSequence"),
            ("In", "in_ProjectRoot", "x:String", "in_ProjectRoot"),
            ("Out", "out_Config", DICT_SO, "out_Config"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - build config dictionary and validate")
        + log('"[INIT] " & summary'),
        variables(("sd:DataTable", "dtSettings"), ("x:String", "summary")))

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
# 01a_Extract_FromMockJson
# ==========================================================================
# The ABI decode lives here and is reused verbatim by the API reader, because both
# receive an identical Etherscan envelope. That shared shape is what makes the bot
# blockchain-agnostic rather than merely configurable.
ABI_DECODE_CODE = r'''
' Flatten an Etherscan `txlist` result array into the canonical chain DataTable.
' Every extraction mode (MOCK, API, WEB) produces exactly this schema, so nothing
' downstream needs to know where the data came from.
Dim arr As JArray = CType(in_Json("result"), JArray)

Dim dt As New DataTable("ChainTransactions")
dt.Columns.Add("TxHash", GetType(String))
dt.Columns.Add("BlockNumber", GetType(Long))
dt.Columns.Add("EventTimestampUtc", GetType(DateTime))
dt.Columns.Add("SenderAddress", GetType(String))
dt.Columns.Add("ContractAddress", GetType(String))
dt.Columns.Add("EventSelector", GetType(String))
dt.Columns.Add("ShipmentID", GetType(String))
dt.Columns.Add("OnChainQty", GetType(Long))
dt.Columns.Add("PONumber", GetType(String))
dt.Columns.Add("TxFailed", GetType(Boolean))
dt.Columns.Add("SourceMode", GetType(String))
dt.Columns.Add("RawInput", GetType(String))

Dim epoch As New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)
Dim skipped As Integer = 0

For Each tok As JObject In arr
    Dim inputHex As String = Convert.ToString(tok("input"))

    ' A plain value transfer carries no call data - nothing to validate, so skip it.
    If inputHex Is Nothing OrElse inputHex.Length < 10 Then
        skipped += 1
        Continue For
    End If

    ' ---- ABI decode: f(string, uint256, string) -------------------------
    ' Call data is the 4-byte selector followed by 32-byte words. The first and
    ' third words hold byte offsets to the dynamic strings; the second is the
    ' uint256 quantity inline.
    Dim body As String = inputHex.Substring(10)
    Dim words As New List(Of String)
    Dim i As Integer = 0
    While i + 64 <= body.Length
        words.Add(body.Substring(i, 64))
        i += 64
    End While

    If words.Count < 3 Then
        skipped += 1
        Continue For
    End If

    Dim qty As Long = Convert.ToInt64(words(1), 16)
    Dim offs() As Integer = {Convert.ToInt32(words(0), 16) \ 32, Convert.ToInt32(words(2), 16) \ 32}
    Dim strs(1) As String

    For n As Integer = 0 To 1
        Dim idx As Integer = offs(n)
        Dim slen As Integer = Convert.ToInt32(words(idx), 16)
        Dim sb As New StringBuilder()
        Dim w As Integer = idx + 1
        While sb.Length < slen * 2 AndAlso w < words.Count
            sb.Append(words(w))
            w += 1
        End While
        Dim hx As String = sb.ToString()
        Dim raw(Math.Max(slen - 1, 0)) As Byte
        For k As Integer = 0 To slen - 1
            raw(k) = Convert.ToByte(hx.Substring(k * 2, 2), 16)
        Next
        strs(n) = Encoding.UTF8.GetString(raw, 0, slen)
    Next

    ' ---- Normalisation happens on the way in ----------------------------
    ' Addresses are lower-cased here so EIP-55 checksummed values from the chain
    ' reconcile against the lower-case forms the ERP stores. Comparing raw would
    ' fail on every single row.
    Dim rw As DataRow = dt.NewRow()
    rw("TxHash") = Convert.ToString(tok("hash")).ToLowerInvariant()
    rw("BlockNumber") = Convert.ToInt64(Convert.ToString(tok("blockNumber")))
    rw("EventTimestampUtc") = epoch.AddSeconds(Convert.ToDouble(Convert.ToString(tok("timeStamp"))))
    rw("SenderAddress") = Convert.ToString(tok("from")).ToLowerInvariant()
    rw("ContractAddress") = Convert.ToString(tok("to")).ToLowerInvariant()
    rw("EventSelector") = inputHex.Substring(0, 10).ToLowerInvariant()
    rw("ShipmentID") = strs(0)
    rw("OnChainQty") = qty
    rw("PONumber") = strs(1)
    rw("TxFailed") = (Convert.ToString(tok("isError")) = "1")
    rw("SourceMode") = in_SourceMode
    rw("RawInput") = inputHex
    dt.Rows.Add(rw)
Next

out_dtChain = dt
out_Summary = String.Format("{0} transactions decoded, {1} skipped (no call data)", dt.Rows.Count, skipped)
'''.strip()


def extract_from_mock() -> str:
    body = sequence(
        "01a - Extract from mock chain feed",
        log('"[EXTRACT/MOCK] Reading " & in_MockFilePath')
        + read_text("in_MockFilePath", "jsonText", name="Read Text File - mock Etherscan response")
        + deserialize_json("jsonText", "jsonObj")
        + if_('Convert.ToString(jsonObj("status")) <> "1"',
              throw('New InvalidOperationException("Explorer returned status=" '
                    '& Convert.ToString(jsonObj("status")) & " message=" '
                    '& Convert.ToString(jsonObj("message")))'),
              name="If - explorer reported an error")
        + invoke_code(ABI_DECODE_CODE, [
            ("In", "in_Json", "njl:JObject", "jsonObj"),
            ("In", "in_SourceMode", "x:String", '"MOCK"'),
            ("Out", "out_dtChain", "sd:DataTable", "out_dtChain"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - ABI decode and normalise")
        + log('"[EXTRACT/MOCK] " & summary'),
        variables(("x:String", "jsonText"), ("njl:JObject", "jsonObj"), ("x:String", "summary")))

    return workflow("01a_Extract_FromMockJson", body, members=[
        ("in_MockFilePath", "InArgument(x:String)"),
        ("out_dtChain", "OutArgument(sd:DataTable)"),
    ])


# ==========================================================================
# 01b_Extract_FromEtherscanApi
# ==========================================================================
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
        + if_('Convert.ToString(jsonObj("status")) <> "1"',
              throw('New InvalidOperationException("Etherscan returned status=" '
                    '& Convert.ToString(jsonObj("status")) & " message=" '
                    '& Convert.ToString(jsonObj("message")) & ". A rate limit or an '
                    'invalid key is the usual cause.")'),
              name="If - Etherscan reported an error")
        + invoke_code(ABI_DECODE_CODE, [
            ("In", "in_Json", "njl:JObject", "jsonObj"),
            ("In", "in_SourceMode", "x:String", '"API"'),
            ("Out", "out_dtChain", "sd:DataTable", "out_dtChain"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - ABI decode and normalise")
        + log('"[EXTRACT/API] " & summary'),
        variables(("x:String", "requestUrl"), ("x:String", "safeUrl"),
                  ("x:String", "responseText"), ("njl:JObject", "jsonObj"),
                  ("x:String", "summary")))

    return workflow("01b_Extract_FromEtherscanApi", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("out_dtChain", "OutArgument(sd:DataTable)"),
    ])


# ==========================================================================
# 01_Extract_Blockchain  (dispatcher)
# ==========================================================================
def extract_blockchain() -> str:
    mock_branch = sequence("MOCK", invoke_workflow(
        "Workflows\\01a_Extract_FromMockJson.xaml", [
            ("In", "in_MockFilePath", "x:String", 'Convert.ToString(in_Config("MockChainFile"))'),
            ("Out", "out_dtChain", "sd:DataTable", "out_dtChain"),
        ], name="Invoke 01a - mock chain feed"))

    api_branch = sequence("API", invoke_workflow(
        "Workflows\\01b_Extract_FromEtherscanApi.xaml", [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("Out", "out_dtChain", "sd:DataTable", "out_dtChain"),
        ], name="Invoke 01b - Etherscan API"))

    web_branch = sequence("WEB", invoke_workflow(
        "Workflows\\01c_Extract_FromExplorerUI.xaml", [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("Out", "out_dtChain", "sd:DataTable", "out_dtChain"),
        ], name="Invoke 01c - explorer UI scraping"))

    dispatch = switch(
        'Convert.ToString(in_Config("DataSourceMode")).ToUpperInvariant()',
        [("MOCK", mock_branch), ("API", api_branch), ("WEB", web_branch)],
        sequence("Default", throw(
            'New InvalidOperationException("Unsupported DataSourceMode: " '
            '& Convert.ToString(in_Config("DataSourceMode")))')),
        name="Switch - DataSourceMode")

    body = sequence(
        "01 - Extract blockchain transactions",
        log('"[EXTRACT] Source mode: " & Convert.ToString(in_Config("DataSourceMode"))')
        + dispatch
        + if_("out_dtChain Is Nothing OrElse out_dtChain.Rows.Count = 0",
              log('"[EXTRACT] No transactions returned for contract "'
                  ' & Convert.ToString(in_Config("ContractAddress"))', level="Warn"),
              name="If - nothing returned"))

    return workflow("01_Extract_Blockchain", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("out_dtChain", "OutArgument(sd:DataTable)"),
    ])


# ==========================================================================
# 01c_Extract_FromExplorerUI  (placeholder - real browser automation lands in Stage 2b)
# ==========================================================================
MAP_01C_CODE = r'''
' Parse the explorer's transaction table, then ABI-decode each row's call data exactly as
' the API reader does. Cells are read positionally, matching COLUMNS in
' tools/generate_explorer_page.py:
'   0 TxHash, 1 Method, 2 Block, 3 DateTimeUtc, 4 From, 5 To, 6 Value, 7 Input

' Narrow to the transaction table's body so unrelated tables on the page are ignored.
Dim tbodyStart As Integer = in_PageHtml.IndexOf("<tbody>", StringComparison.OrdinalIgnoreCase)
Dim tbodyEnd As Integer = in_PageHtml.IndexOf("</tbody>", StringComparison.OrdinalIgnoreCase)
If tbodyStart < 0 OrElse tbodyEnd < 0 Then
    Throw New InvalidOperationException("No transaction table found on the explorer page. " & _
        "If a live explorer is being targeted, its markup differs from the expected layout.")
End If
Dim tbody As String = in_PageHtml.Substring(tbodyStart, tbodyEnd - tbodyStart)

Dim rowRx As New Regex("<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
Dim cellRx As New Regex("<td[^>]*>(.*?)</td>", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
Dim tagRx As New Regex("<[^>]+>")

Dim dt As New DataTable("ChainTransactions")
dt.Columns.Add("TxHash", GetType(String))
dt.Columns.Add("BlockNumber", GetType(Long))
dt.Columns.Add("EventTimestampUtc", GetType(DateTime))
dt.Columns.Add("SenderAddress", GetType(String))
dt.Columns.Add("ContractAddress", GetType(String))
dt.Columns.Add("EventSelector", GetType(String))
dt.Columns.Add("ShipmentID", GetType(String))
dt.Columns.Add("OnChainQty", GetType(Long))
dt.Columns.Add("PONumber", GetType(String))
dt.Columns.Add("TxFailed", GetType(Boolean))
dt.Columns.Add("SourceMode", GetType(String))
dt.Columns.Add("RawInput", GetType(String))

Dim scrapedRows As Integer = 0
Dim skipped As Integer = 0

For Each rowMatch As Match In rowRx.Matches(tbody)
    Dim cells As New List(Of String)
    For Each cellMatch As Match In cellRx.Matches(rowMatch.Groups(1).Value)
        Dim rawCell As String = tagRx.Replace(cellMatch.Groups(1).Value, "")
        cells.Add(System.Net.WebUtility.HtmlDecode(rawCell).Trim())
    Next
    If cells.Count < 8 Then
        Continue For
    End If
    scrapedRows += 1

    Dim inputHex As String = cells(7)
    If inputHex.Length < 10 Then
        skipped += 1
        Continue For
    End If

    ' ---- ABI decode: f(string, uint256, string) -------------------------
    Dim payload As String = inputHex.Substring(10)
    Dim words As New List(Of String)
    Dim i As Integer = 0
    While i + 64 <= payload.Length
        words.Add(payload.Substring(i, 64))
        i += 64
    End While
    If words.Count < 3 Then
        skipped += 1
        Continue For
    End If

    Dim qty As Long = Convert.ToInt64(words(1), 16)
    Dim offs() As Integer = {Convert.ToInt32(words(0), 16) \ 32, Convert.ToInt32(words(2), 16) \ 32}
    Dim strs(1) As String
    For n As Integer = 0 To 1
        Dim idx As Integer = offs(n)
        Dim slen As Integer = Convert.ToInt32(words(idx), 16)
        Dim sb As New StringBuilder()
        Dim w As Integer = idx + 1
        While sb.Length < slen * 2 AndAlso w < words.Count
            sb.Append(words(w))
            w += 1
        End While
        Dim hx As String = sb.ToString()
        Dim rawBytes(Math.Max(slen - 1, 0)) As Byte
        For k As Integer = 0 To slen - 1
            rawBytes(k) = Convert.ToByte(hx.Substring(k * 2, 2), 16)
        Next
        strs(n) = Encoding.UTF8.GetString(rawBytes, 0, slen)
    Next

    Dim rw As DataRow = dt.NewRow()
    rw("TxHash") = cells(0).ToLowerInvariant()
    rw("BlockNumber") = Convert.ToInt64(cells(2))
    rw("EventTimestampUtc") = DateTime.SpecifyKind( _
        DateTime.ParseExact(cells(3), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), _
        DateTimeKind.Utc)
    rw("SenderAddress") = cells(4).ToLowerInvariant()
    rw("ContractAddress") = cells(5).ToLowerInvariant()
    rw("EventSelector") = inputHex.Substring(0, 10).ToLowerInvariant()
    rw("ShipmentID") = strs(0)
    rw("OnChainQty") = qty
    rw("PONumber") = strs(1)
    rw("TxFailed") = False
    rw("SourceMode") = "WEB"
    rw("RawInput") = inputHex
    dt.Rows.Add(rw)
Next

out_dtChain = dt
out_Summary = String.Format("{0} table row(s) scraped, {1} decoded, {2} skipped", _
    scrapedRows, dt.Rows.Count, skipped)
'''.strip()

RESOLVE_01C_CODE = r'''
' The setting may be a project-relative path or a live http(s) URL.
Dim v As String = Convert.ToString(in_Config("ExplorerPageUrl")).Trim()
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
    table, mapping onto the same canonical schema MOCK and API produce.

    What this is, precisely: it parses the explorer's HTML. It does NOT drive a browser
    with UiPath's Data Scraping wizard. Two concrete reasons:

      * UIAutomation 25.10 removed the classic `ExtractStructuredData` activity. Its
        replacement, NExtractData, is built around descriptors the Studio recorder
        generates against a live page - not something to hand-author and call done.
      * Browser automation needs the UiPath browser extension, which is not installed on
        this machine, so a recorded version could be neither executed nor verified.

    Parsing the page is a genuine scrape and it is fully testable, which a hand-written
    descriptor would not be. To move to true browser automation: install the extension
    (Studio - Home - Tools - UiPath Extensions), record the table with the Data Scraping
    wizard, and feed its DataTable into the mapping code below. The mapping is what carries
    the value here and does not change.

    ExplorerPageUrl accepts a project-relative path or an http(s) URL.
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
            ("Out", "out_dtChain", "sd:DataTable", "out_dtChain"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - parse the table and ABI decode")
        + log('"[EXTRACT/WEB] " & summary'),
        variables(("x:Boolean", "isRemote"), ("x:String", "location"),
                  ("x:String", "pageHtml"), ("x:String", "summary")))

    return workflow("01c_Extract_FromExplorerUI", body,
                    members=[
                        ("in_Config", f"InArgument({DICT_SO})"),
                        ("out_dtChain", "OutArgument(sd:DataTable)"),
                    ],
                    extra_imports=["System.Text.RegularExpressions", "System.Net"],
                    extra_refs=["System.Text.RegularExpressions"])


# ==========================================================================
# 02_Extract_Logistics
# ==========================================================================
def extract_logistics() -> str:
    code = r'''
' ERP addresses are already lower-case, but normalise defensively - a hand-edited
' export is exactly the kind of thing that reintroduces mixed case.
For Each r As DataRow In in_dtEvents.Rows
    r("HandlerAddress") = Convert.ToString(r("HandlerAddress")).Trim().ToLowerInvariant()
Next
For Each r As DataRow In in_dtShipments.Rows
    r("ReceiverAddress") = Convert.ToString(r("ReceiverAddress")).Trim().ToLowerInvariant()
Next

out_Summary = String.Format("{0} shipments, {1} expected milestones", _
    in_dtShipments.Rows.Count, in_dtEvents.Rows.Count)
'''.strip()

    excel_body = sequence("Read ERP sheets",
        excel_read("Shipments", "out_dtShipments", name="Read Range - Shipments")
        + excel_read("ShipmentEvents", "out_dtErpEvents", name="Read Range - ShipmentEvents"))

    body = sequence(
        "02 - Extract logistics (ERP/WMS) records",
        log('"[ERP] Reading " & Convert.ToString(in_Config("LogisticsFile"))')
        + excel_scope('Convert.ToString(in_Config("LogisticsFile"))', excel_body,
                      name="Excel Application Scope - LogisticsRecords.xlsx")
        + invoke_code(code, [
            ("In", "in_dtShipments", "sd:DataTable", "out_dtShipments"),
            ("In", "in_dtEvents", "sd:DataTable", "out_dtErpEvents"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - normalise ERP addresses")
        + log('"[ERP] " & summary'),
        variables(("x:String", "summary")))

    return workflow("02_Extract_Logistics", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("out_dtShipments", "OutArgument(sd:DataTable)"),
        ("out_dtErpEvents", "OutArgument(sd:DataTable)"),
    ])


# ==========================================================================
# 03_Preprocess_Map
# ==========================================================================
def preprocess_map() -> str:
    code = r'''
' Join the chain feed to the ERP expectation for the same (shipment, milestone).
' The milestone name is resolved from the function selector via the EventSequence
' sheet, so adding a new milestone is a config change, not a workflow change.
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

Dim unmatched As Integer = 0
Dim n As Integer = 0

For Each c As DataRow In in_dtChain.Rows
    n += 1
    Dim sel As String = Convert.ToString(c("EventSelector")).Trim().ToLowerInvariant()
    Dim sid As String = Convert.ToString(c("ShipmentID")).Trim()
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
    rw("PONumber") = c("PONumber")
    rw("TxFailed") = c("TxFailed")

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
'''.strip()

    body = sequence(
        "03 - Preprocess and map chain to ERP",
        log('"[MAP] Joining " & in_dtChain.Rows.Count.ToString() & " chain rows to ERP milestones"')
        + invoke_code(code, [
            ("In", "in_dtChain", "sd:DataTable", "in_dtChain"),
            ("In", "in_dtShipments", "sd:DataTable", "in_dtShipments"),
            ("In", "in_dtErpEvents", "sd:DataTable", "in_dtErpEvents"),
            ("In", "in_dtSequence", "sd:DataTable", "in_dtSequence"),
            ("Out", "out_dtJoined", "sd:DataTable", "out_dtJoined"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - map and join")
        + log('"[MAP] " & summary'),
        variables(("x:String", "summary")))

    return workflow("03_Preprocess_Map", body, members=[
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

For Each j As DataRow In in_dtJoined.Rows
    Dim rw As DataRow = dt.NewRow()
    For Each c As DataColumn In in_dtJoined.Columns
        If dt.Columns.Contains(c.ColumnName) Then rw(c.ColumnName) = j(c.ColumnName)
    Next
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

    def rule_case(rid: str, fname: str, extra=()) -> tuple[str, str]:
        args = list(rule_args) + list(extra)
        return rid, sequence(rid, invoke_workflow(
            f"Workflows\\Rules\\{fname}.xaml", args, name=f"Invoke {rid} - {fname}"))

    cases = [
        rule_case("R1", "R1_TimestampCheck"),
        rule_case("R2", "R2_QuantityMatch"),
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
            ("Out", "out_dtExceptions", "sd:DataTable", "out_dtExceptions"),
            ("Out", "out_Stats", DICT_SO, "out_Stats"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - roll up statuses and build exceptions")
        + log('"[VALIDATE] " & summary'),
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
#   PROMPT   - show the reviewer a dialog and wait. Attended runs only.
#   SIMULATE - apply a pre-set decision. For rehearsals and automated tests.
#   AUTO_LOG - record DEFERRED and carry on. The correct behaviour unattended, where
#              there is nobody at the machine and blocking would hang the job.
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
    prompt_body = sequence(
        "Ask the reviewer",
        if_('Convert.ToBoolean(CurrentRow("RequiresEscalation"))',
            sequence(
                "Escalate this transaction",
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
                + assign('CurrentRow("HumanDecision")', "decision",
                         type_ref="x:Object", name="Assign - record decision")
                + assign('CurrentRow("HumanDecidedBy")', "Environment.UserName",
                         type_ref="x:Object", name="Assign - record reviewer")
                + assign('CurrentRow("HumanDecidedAtUtc")',
                         'DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)',
                         type_ref="x:Object", name="Assign - record decision time")
                + log('"[ESCALATE] " & Convert.ToString(CurrentRow("ShipmentID")) '
                      '& " -> " & decision & " (by " & Environment.UserName & ")"')),
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
              name="If - anything to escalate"),
        variables(("x:Int32", "escalationCount"), ("x:String", "summary"),
                  ("x:String", "mode"), ("x:String", "decision")))

    return workflow("06_HumanInTheLoop_Escalation", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("io_dtResults", "InOutArgument(sd:DataTable)"),
        ("out_EscalationCount", "OutArgument(x:Int32)"),
    ])


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
# Split(","c) therefore recovers fields 0..14 correctly no matter what the free text
# contains, and the remainder rejoins to give field 15. No CSV state machine needed.
AUDIT_HEADER = ("Seq,TimestampUtc,RunId,BotIdentity,MachineName,TxHash,ShipmentID,EventName,"
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
If Not File.Exists(outPath) Then outLines.Add("''' + AUDIT_HEADER + r'''")

Dim runId As String = Convert.ToString(in_Config("RunId"))
Dim bot As String = Convert.ToString(in_Config("BotIdentity"))
Dim machine As String = Convert.ToString(in_Config("MachineName"))

For Each r As DataRow In in_dtResults.Rows
    seq += 1
    Dim stamp As String = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
    Dim chainTs As String = Convert.ToDateTime(r("ChainTimestampUtc")).ToString("s", CultureInfo.InvariantCulture)
    Dim reasons As String = Convert.ToString(r("FailureReasons"))
    ' Blank placeholders keep every hash-relevant field comma-free and non-empty.
    Dim humanDecision As String = Convert.ToString(r("HumanDecision")).Replace(",", ";")
    Dim humanBy As String = Convert.ToString(r("HumanDecidedBy")).Replace(",", ";")
    If humanDecision = "" Then humanDecision = "-"
    If humanBy = "" Then humanBy = "-"

    ' RecordHash fixes the verdict and the evidence behind it.
    Dim recordInput As String = String.Join("|", New String() { _
        runId, Convert.ToString(r("TxHash")), Convert.ToString(r("ShipmentID")), _
        Convert.ToString(r("EventName")), chainTs, Convert.ToString(r("OnChainQty")), _
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
        humanDecision, humanBy, recordHash, prevHash, entryHash, quoted}))

    prevHash = entryHash
Next

File.AppendAllLines(outPath, outLines, Encoding.UTF8)

out_AuditPath = outPath
out_HeadHash = prevHash
out_Summary = String.Format("{0} entries appended, chain head {1}...", in_dtResults.Rows.Count, prevHash.Substring(0, 16))
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
    If p.Length < 18 Then
        problems.Add("line " & (i + 1).ToString() & ": malformed (only " & p.Length.ToString() & " fields)")
        If bad = 0 Then bad = i + 1
        Continue For
    End If

    ' Fields 0..14 are comma-free by construction; field 15 onwards is the quoted
    ' free-text reason, which is rejoined and unescaped here.
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

    Dim reasons As String = String.Join(",", p, 17, p.Length - 17)
    If reasons.StartsWith("""") AndAlso reasons.EndsWith("""") AndAlso reasons.Length >= 2 Then
        reasons = reasons.Substring(1, reasons.Length - 2).Replace("""""", """")
    End If

    expectedSeq += 1
    If Convert.ToInt64(seq) <> expectedSeq Then
        problems.Add("line " & (i + 1).ToString() & ": sequence jumped, expected " & expectedSeq.ToString() & " found " & seq)
        If bad = 0 Then bad = i + 1
    End If

    Dim recomputedRecord As String = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes( _
        String.Join("|", New String() {runId, txHash, shipment, evName, chainTs, qty, status, severity, reasons, humanDecision, humanBy}) _
        ))).Replace("-", "").ToLowerInvariant()

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
# identically in an unattended Orchestrator run on a machine with no Office installed.
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
ws.Cell(1, 1).Value = "Blockchain Transaction Validation Report"
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
o("auditChainHead") = in_HeadHash

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

out_DashboardPath = outPath
out_Summary = String.Format("dashboard JSON written with {0} finding(s): {1}", byShipment.Count, Path.GetFileName(outPath))
'''.strip()

    body = sequence(
        "11 - Emit dashboard summary",
        invoke_code(code, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtResults", "sd:DataTable", "in_dtResults"),
            ("In", "in_dtRules", "sd:DataTable", "in_dtRules"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("In", "in_HeadHash", "x:String", "in_HeadHash"),
            ("Out", "out_DashboardPath", "x:String", "out_DashboardPath"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - build dashboard JSON")
        + log('"[DASHBOARD] " & summary'),
        variables(("x:String", "summary")))

    return workflow("11_Dashboard_Summary", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtResults", "InArgument(sd:DataTable)"),
        ("in_dtRules", "InArgument(sd:DataTable)"),
        ("in_Stats", f"InArgument({DICT_SO})"),
        ("in_HeadHash", "InArgument(x:String)"),
        ("out_DashboardPath", "OutArgument(x:String)"),
    ])


# ==========================================================================
# 09_Alert_Email
# ==========================================================================
def send_smtp_mail() -> str:
    """
    Classic Send SMTP Mail Message (UiPath.Mail.SMTP.Activities.SendMail).

    Expressions are written with real double quotes - `vb()` does the XML escaping.
    Pre-escaping them here would double-escape and produce `&amp;quot;`, which VB then
    reads as an undeclared identifier called `quot`.
    """
    def cfg(key: str) -> str:
        return f'Convert.ToString(in_Config("{key}"))'

    to_expr = 'Convert.ToString(in_Config("AlertTo")).Replace(";", ",")'
    port_expr = 'Convert.ToInt32(in_Config("SmtpPort"))'

    return (f'<ui:SendMail DisplayName="{lit("Send SMTP Mail Message")}" '
            f'Server="{vb(cfg("SmtpHost"))}" '
            f'Port="{vb(port_expr)}" '
            f'Email="{vb(cfg("SmtpUser"))}" '
            f'Password="{vb(cfg("SmtpPassword"))}" '
            f'From="{vb(cfg("AlertFrom"))}" '
            f'To="{vb(to_expr)}" '
            f'Subject="{vb("subject")}" Body="{vb("htmlBody")}" '
            f'IsBodyHtml="True" SecureConnection="Auto" />')



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
sb.AppendLine("<h2 style=""color:#1F3864;margin-bottom:4px"">Blockchain Transaction Validation</h2>")
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

sb.AppendLine("<p style=""color:#666;font-size:11px;margin-top:18px"">Audit chain head: " & in_HeadHash & "<br/>Report: " & in_ReportPath & "</p>")
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

out_AlertPath = outPath
out_Subject = subject
out_Body = body
out_ShouldSend = (nFail > 0 OrElse nWarn > 0)
out_Summary = String.Format("mode={0} subject=""{1}"" written to {2}", mode, subject, Path.GetFileName(outPath))
'''.strip()

    body = sequence(
        "09 - Compose and deliver anomaly alert",
        invoke_code(code, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtExceptions", "sd:DataTable", "in_dtExceptions"),
            ("In", "in_Stats", DICT_SO, "in_Stats"),
            ("In", "in_ReportPath", "x:String", "in_ReportPath"),
            ("In", "in_HeadHash", "x:String", "in_HeadHash"),
            ("Out", "out_AlertPath", "x:String", "out_AlertPath"),
            ("Out", "out_Subject", "x:String", "subject"),
            ("Out", "out_Body", "x:String", "htmlBody"),
            ("Out", "out_ShouldSend", "x:Boolean", "shouldSend"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - compose the alert")
        + if_('Convert.ToString(in_Config("AlertMode")).Trim().ToUpperInvariant() = "SMTP" AndAlso shouldSend',
              sequence("Send over SMTP", send_smtp_mail()),
              sequence("Written to disk",
                       log('"[ALERT] AlertMode is not SMTP - the composed message was written to '
                           'disk instead of sent. Set AlertMode=SMTP in Config.xlsx to deliver it."')),
              name="If - deliver over SMTP")
        + log('"[ALERT] " & summary'),
        variables(("x:String", "subject"), ("x:String", "htmlBody"),
                  ("x:Boolean", "shouldSend"), ("x:String", "summary")))

    return workflow("09_Alert_Email", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtExceptions", "InArgument(sd:DataTable)"),
        ("in_Stats", f"InArgument({DICT_SO})"),
        ("in_ReportPath", "InArgument(x:String)"),
        ("in_HeadHash", "InArgument(x:String)"),
        ("out_AlertPath", "OutArgument(x:String)"),
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
# QueueMode = DRYRUN writes the payloads to disk instead of Orchestrator, so the
# grouping and serialisation can be tested without a tenant.
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

    dryrun_code = r'''
' Write what would have been queued, so the split can be inspected without a tenant.
Dim folder As String = Path.Combine(Convert.ToString(in_Config("ReportsFolder")), "QueueDryRun")
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
out_Summary = String.Format("dry run: {0} payload(s) written to {1}", arr.Count, Path.GetFileName(outFile))
'''.strip()

    # ItemInformation is a KEYED dictionary of arguments, not one argument holding a
    # dictionary - binding a prebuilt Dictionary fails with "Missing key value on
    # 'InArgument' object". Listing the fields individually also makes them visible in
    # the designer and in Orchestrator's queue-item view.
    item_fields = [
        ("ShipmentID", 'Convert.ToString(CurrentPayload("ShipmentID"))'),
        ("EventCount", 'Convert.ToString(CurrentPayload("EventCount"))'),
        ("EventsJson", 'Convert.ToString(CurrentPayload("EventsJson"))'),
        ("RunId", 'Convert.ToString(in_Config("RunId"))'),
    ]
    fields_xml = "".join(
        f'<InArgument x:TypeArguments="x:String" x:Key="{k}">{vb(e)}</InArgument>'
        for k, e in item_fields)

    ref_expr = 'Convert.ToString(CurrentPayload("Reference"))'

    add_item = sequence(
        "Enqueue this shipment",
        # The property is QueueType, not QueueName - its DISPLAY name is "Queue name",
        # which is why every obvious guess fails with "unknown member".
        f'<ui:AddQueueItem DisplayName="{lit("Add Queue Item")}" '
        f'QueueType="{vb(QUEUE_NAME_EXPR)}" '
        f'Reference="{vb(ref_expr)}">'
        f'<ui:AddQueueItem.ItemInformation>'
        f'<scg:Dictionary x:TypeArguments="x:String, InArgument">{fields_xml}</scg:Dictionary>'
        f'</ui:AddQueueItem.ItemInformation></ui:AddQueueItem>')

    build_info = ""

    body = sequence(
        "Dispatcher - split the run into per-shipment queue items",
        invoke_code(code, [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtJoined", "sd:DataTable", "in_dtJoined"),
            ("Out", "out_dtPayloads", "sd:DataTable", "dtPayloads"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - group transactions by shipment")
        + log('"[DISPATCH] " & summary')
        + if_('Convert.ToString(in_Config("QueueMode")).Trim().ToUpperInvariant() = "ORCHESTRATOR"',
              sequence("Push to Orchestrator",
                       for_each_row("dtPayloads",
                                    sequence("Queue one shipment", build_info + add_item),
                                    row_var="CurrentPayload",
                                    name="For Each Row - shipment payloads")
                       + log('"[DISPATCH] " & dtPayloads.Rows.Count.ToString() & " item(s) added to BlockchainValidationQueue"')),
              sequence("Dry run",
                       invoke_code(dryrun_code, [
                           ("In", "in_Config", DICT_SO, "in_Config"),
                           ("In", "in_dtPayloads", "sd:DataTable", "dtPayloads"),
                           ("Out", "out_Summary", "x:String", "summary"),
                       ], name="Invoke Code - write payloads to disk")
                       + log('"[DISPATCH] " & summary')),
              name="If - Orchestrator or dry run")
        + assign("out_ItemsQueued", "dtPayloads.Rows.Count", type_ref="x:Int32",
                 name="Assign - items queued"),
        variables(("sd:DataTable", "dtPayloads"), ("x:String", "summary"),
                  (DICT_SO, "itemInfo")))

    return workflow("Dispatcher", body, members=[
        ("in_Config", f"InArgument({DICT_SO})"),
        ("in_dtJoined", "InArgument(sd:DataTable)"),
        ("out_ItemsQueued", "OutArgument(x:Int32)"),
    ])


def queue_performer() -> str:
    rebuild = r'''
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
'''.strip()

    process_one = sequence(
        "Process one queue item",
        assign("eventsJson", 'Convert.ToString(txItem.SpecificContent("EventsJson"))',
               name="Assign - payload")
        + assign("shipmentId", 'Convert.ToString(txItem.SpecificContent("ShipmentID"))',
                 name="Assign - shipment")
        + log('"[PERFORM] " & shipmentId')
        + invoke_code(rebuild, [
            ("In", "in_EventsJson", "x:String", "eventsJson"),
            ("Out", "out_dtJoined", "sd:DataTable", "dtJoined"),
            ("Out", "out_Summary", "x:String", "summary"),
        ], name="Invoke Code - rehydrate the shipment")
        + invoke_workflow("Workflows\\05_Validate_Engine.xaml", [
            ("In", "in_Config", DICT_SO, "in_Config"),
            ("In", "in_dtJoined", "sd:DataTable", "dtJoined"),
            ("In", "in_dtRules", "sd:DataTable", "in_dtRules"),
            ("In", "in_dtWallets", "sd:DataTable", "in_dtWallets"),
            ("Out", "out_dtResults", "sd:DataTable", "dtResults"),
            ("Out", "out_dtExceptions", "sd:DataTable", "dtExceptions"),
            ("Out", "out_Stats", DICT_SO, "Stats"),
        ], name="Invoke 05 - Validation engine")
        + f'<ui:SetTransactionStatus DisplayName="{lit("Set Transaction Status - Successful")}" '
          f'TransactionItem="{vb("txItem")}" Status="Successful" />'
        + log('"[PERFORM] " & shipmentId & ": " & Convert.ToString(Stats("Fail")) '
              '& " failure(s), " & Convert.ToString(Stats("Warning")) & " warning(s)"'))

    body = sequence(
        "Performer - validate one shipment per queue transaction",
        f'<ui:GetQueueItem DisplayName="{lit("Get Transaction Item")}" '
        f'QueueType="{vb(QUEUE_NAME_EXPR)}" '
        f'TransactionItem="{vb("txItem")}" />'
        + if_("txItem IsNot Nothing",
              try_catch(
                  process_one,
                  sequence("Report the failure to Orchestrator",
                           log('"[PERFORM] " & shipmentId & " failed: " & exception.Message', level="Error")
                           + f'<ui:SetTransactionStatus DisplayName="{lit("Set Transaction Status - Failed")}" '
                             f'TransactionItem="{vb("txItem")}" Status="Failed" ErrorType="Application" '
                             f'Reason="{vb("exception.Message")}" />'),
                  name="Try Catch - transaction processing"),
              sequence("Queue empty", log('"[PERFORM] Queue is empty - nothing to process."')),
              name="If - a transaction was dequeued"),
        variables(("ui:QueueItem", "txItem"), ("x:String", "eventsJson"),
                  ("x:String", "shipmentId"), ("x:String", "summary"),
                  ("sd:DataTable", "dtJoined"), ("sd:DataTable", "dtResults"),
                  ("sd:DataTable", "dtExceptions"), (DICT_SO, "Stats")))

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
st.AppendLine("<header><h1>Blockchain Transaction Validation &mdash; Logistics</h1>")
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
Dim steps() As String = {"00|Read config", "01|Extract chain", "02|Extract ERP", "03|Map and join", _
                         "05|Validate R1-R6", "06|Escalate", "10|Audit log", "07|Excel report", _
                         "09|Alert", "12|Verify chain", "13|Dashboard"}
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

' ---- Run history ----------------------------------------------------------
st.AppendLine("<h2>Run history (" & runIds.Count & " run(s) in the audit trail)</h2>")
If runIds.Count = 0 Then
    st.AppendLine("<p class='muted'>No audit history yet.</p>")
Else
    st.AppendLine("<table><tr><th>Run</th><th>Started (UTC)</th><th>Validated</th>" & _
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
        st.AppendLine("<tr><td class='mono' style='font-size:11.5px'>" & label & "</td><td class='mono'>" & _
            started & "</td><td>" & runTotal(rid) & "</td><td>" & runPass(rid) & "</td><td>" & _
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
out_Summary = String.Format("dashboard.html written ({0} finding(s), {1} run(s) of history)", _
    findings, runIds.Count)
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
# Main (extended in later stages)
# ==========================================================================
def main_workflow() -> str:
    body = sequence(
        "Main",
        log('"=== BlockchainLogisticsValidator starting ==="')
        + assign("projectRoot", "Directory.GetCurrentDirectory()")
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
        + invoke_workflow("Workflows\\01_Extract_Blockchain.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("Out", "out_dtChain", "sd:DataTable", "dtChain"),
        ], name="Invoke 01 - Extract blockchain")
        + invoke_workflow("Workflows\\02_Extract_Logistics.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("Out", "out_dtShipments", "sd:DataTable", "dtShipments"),
            ("Out", "out_dtErpEvents", "sd:DataTable", "dtErpEvents"),
        ], name="Invoke 02 - Extract logistics")
        + invoke_workflow("Workflows\\03_Preprocess_Map.xaml", [
            ("In", "in_dtChain", "sd:DataTable", "dtChain"),
            ("In", "in_dtShipments", "sd:DataTable", "dtShipments"),
            ("In", "in_dtErpEvents", "sd:DataTable", "dtErpEvents"),
            ("In", "in_dtSequence", "sd:DataTable", "dtSequence"),
            ("Out", "out_dtJoined", "sd:DataTable", "dtJoined"),
        ], name="Invoke 03 - Preprocess and map")
        + log('"[MAIN] Joined rows ready: " & dtJoined.Rows.Count.ToString()')
        # BATCH validates the whole run in one job. DISPATCH instead splits it into
        # one Orchestrator queue item per shipment for the Performer to pick up.
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
          invoke_workflow("Workflows\\05_Validate_Engine.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("In", "in_dtJoined", "sd:DataTable", "dtJoined"),
            ("In", "in_dtRules", "sd:DataTable", "dtRules"),
            ("In", "in_dtWallets", "sd:DataTable", "dtWallets"),
            ("Out", "out_dtResults", "sd:DataTable", "dtResults"),
            ("Out", "out_dtExceptions", "sd:DataTable", "dtExceptions"),
            ("Out", "out_Stats", DICT_SO, "Stats"),
        ], name="Invoke 05 - Validation engine")
        + log('"[MAIN] Validation complete: " & Convert.ToString(Stats("Fail")) & " failed, " '
              '& Convert.ToString(Stats("Warning")) & " warning, " '
              '& Convert.ToString(Stats("Pass")) & " passed ("'
              ' & Convert.ToString(Stats("PassPercent")) & "% pass rate)"')
        # ---- Novelty 2: pause for a human on a CRITICAL anomaly ----------
        # Placed before the audit log on purpose: the reviewer's decision is hashed
        # into the chain, so it is as tamper-evident as the machine's own verdict.
        + invoke_workflow("Workflows\\06_HumanInTheLoop_Escalation.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("InOut", "io_dtResults", "sd:DataTable", "dtResults"),
            ("Out", "out_EscalationCount", "x:Int32", "escalationCount"),
        ], name="Invoke 06 - Human-in-the-loop escalation")
        # ---- Phase 4: audit trail, reporting, alerting -------------------
        # The audit log is written FIRST, before any report. The chain records what
        # the bot concluded; reports are derived artefacts and can be regenerated.
        + invoke_workflow("Workflows\\10_AuditLog_Chained.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("In", "in_dtResults", "sd:DataTable", "dtResults"),
            ("Out", "out_AuditPath", "x:String", "auditPath"),
            ("Out", "out_HeadHash", "x:String", "headHash"),
        ], name="Invoke 10 - Append audit log")
        + invoke_workflow("Workflows\\07_Report_Excel.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("In", "in_dtResults", "sd:DataTable", "dtResults"),
            ("In", "in_dtExceptions", "sd:DataTable", "dtExceptions"),
            ("In", "in_dtRules", "sd:DataTable", "dtRules"),
            ("In", "in_Stats", DICT_SO, "Stats"),
            ("Out", "out_ReportPath", "x:String", "reportPath"),
        ], name="Invoke 07 - Excel report")
        + invoke_workflow("Workflows\\11_Dashboard_Summary.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("In", "in_dtResults", "sd:DataTable", "dtResults"),
            ("In", "in_dtRules", "sd:DataTable", "dtRules"),
            ("In", "in_Stats", DICT_SO, "Stats"),
            ("In", "in_HeadHash", "x:String", "headHash"),
            ("Out", "out_DashboardPath", "x:String", "dashboardPath"),
        ], name="Invoke 11 - Dashboard summary")
        + invoke_workflow("Workflows\\09_Alert_Email.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("In", "in_dtExceptions", "sd:DataTable", "dtExceptions"),
            ("In", "in_Stats", DICT_SO, "Stats"),
            ("In", "in_ReportPath", "x:String", "reportPath"),
            ("In", "in_HeadHash", "x:String", "headHash"),
            ("Out", "out_AlertPath", "x:String", "alertPath"),
        ], name="Invoke 09 - Anomaly alert")
        # ---- Self-check: prove the audit trail we just wrote is intact ----
        + invoke_workflow("Workflows\\12_Verify_AuditChain.xaml", [
            ("In", "in_AuditPath", "x:String", "auditPath"),
            ("Out", "out_IsValid", "x:Boolean", "chainValid"),
            ("Out", "out_Report", "x:String", "chainReport"),
        ], name="Invoke 12 - Verify audit chain")
        # Last, so it can report the verification outcome the step above produced.
        + invoke_workflow("Workflows\\13_Report_Dashboard.xaml", [
            ("In", "in_Config", DICT_SO, "Config"),
            ("In", "in_dtResults", "sd:DataTable", "dtResults"),
            ("In", "in_dtRules", "sd:DataTable", "dtRules"),
            ("In", "in_Stats", DICT_SO, "Stats"),
            ("In", "in_AuditPath", "x:String", "auditPath"),
            ("In", "in_ChainValid", "x:Boolean", "chainValid"),
            ("In", "in_ChainReport", "x:String", "chainReport"),
            ("Out", "out_DashboardPath", "x:String", "dashboardHtmlPath"),
        ], name="Invoke 13 - HTML dashboard")),
              name="If - dispatch or batch")
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
            ("x:String", "auditPath"), ("x:String", "headHash"),
            ("x:String", "reportPath"), ("x:String", "dashboardPath"),
            ("x:String", "alertPath"), ("x:Boolean", "chainValid"),
            ("x:String", "chainReport"), ("x:Int32", "escalationCount"),
            ("x:Int32", "itemsQueued"), ("x:String", "dashboardHtmlPath")))

    return workflow("Main", body)


# ==========================================================================
GENERATORS = {
    os.path.join(PROJ, "Main.xaml"): main_workflow,
    os.path.join(WF, "00_Init_ReadConfig.xaml"): init_read_config,
    os.path.join(WF, "01_Extract_Blockchain.xaml"): extract_blockchain,
    os.path.join(WF, "01a_Extract_FromMockJson.xaml"): extract_from_mock,
    os.path.join(WF, "01b_Extract_FromEtherscanApi.xaml"): extract_from_api,
    os.path.join(WF, "01c_Extract_FromExplorerUI.xaml"): extract_from_ui,
    os.path.join(WF, "02_Extract_Logistics.xaml"): extract_logistics,
    os.path.join(WF, "03_Preprocess_Map.xaml"): preprocess_map,
    os.path.join(WF, "05_Validate_Engine.xaml"): validate_engine,
    os.path.join(WF, "06_HumanInTheLoop_Escalation.xaml"): human_in_the_loop,
    os.path.join(WF, "07_Report_Excel.xaml"): report_excel,
    os.path.join(WF, "09_Alert_Email.xaml"): alert_email,
    os.path.join(WF, "10_AuditLog_Chained.xaml"): audit_log_chained,
    os.path.join(WF, "11_Dashboard_Summary.xaml"): dashboard_summary,
    os.path.join(WF, "12_Verify_AuditChain.xaml"): verify_audit_chain,
    os.path.join(WF, "13_Report_Dashboard.xaml"): report_dashboard,
    os.path.join(QUEUE, "Dispatcher.xaml"): queue_dispatcher,
    os.path.join(QUEUE, "Performer.xaml"): queue_performer,
    os.path.join(RULES, "R1_TimestampCheck.xaml"): r1_timestamp,
    os.path.join(RULES, "R2_QuantityMatch.xaml"): r2_quantity,
    os.path.join(RULES, "R3_AddressWhitelist.xaml"): r3_whitelist,
    os.path.join(RULES, "R4_DuplicateDetection.xaml"): r4_duplicate,
    os.path.join(RULES, "R5_SequenceValidation.xaml"): r5_sequence,
    os.path.join(RULES, "R6_SmartContractEvent.xaml"): r6_contract_event,
}


def main() -> None:
    os.makedirs(WF, exist_ok=True)
    os.makedirs(RULES, exist_ok=True)
    os.makedirs(QUEUE, exist_ok=True)
    for path, gen in GENERATORS.items():
        with open(path, "w", encoding="utf-8") as f:
            f.write(gen())
        print("wrote:", os.path.relpath(path, ROOT))


if __name__ == "__main__":
    main()

<#
    build_erp_database.ps1 - build the SQL flavour of the ERP data.

    LogisticsSourceMode = SQL reads the shipment master and the expected milestones with
    UiPath's Database Connect + Execute Query activities. To make that path testable on any
    Windows machine with Microsoft Office, the database is a Microsoft Access file queried
    through ODBC - the same activities, provider and SQL would run unchanged against SQL
    Server or Oracle by editing DbProviderName / DbConnectionString in Config.xlsx.

    The tables are loaded from Data\Input\ERP_CSV\*.csv (written by generate_sample_data.py)
    with proper column types - DATETIME for timestamps, LONG for quantities - so the bot
    sees typed values exactly as it would from a real ERP database, not strings.

    Run:  powershell -File tools\build_erp_database.ps1
#>
$ErrorActionPreference = "Stop"

$Root     = Split-Path $PSScriptRoot -Parent
$InputDir = Join-Path $Root "BlockchainLogisticsValidator\Data\Input"
$CsvDir   = Join-Path $InputDir "ERP_CSV"
$DbPath   = Join-Path $InputDir "LogisticsERP.accdb"

$tables = @(
    @{
        Name    = "Shipments"
        Csv     = Join-Path $CsvDir "Shipments.csv"
        Columns = [ordered]@{
            ShipmentID = "TEXT(20) NOT NULL"; PONumber = "TEXT(20)"; Product = "TEXT(80)"
            ExpectedQty = "LONG"; UoM = "TEXT(20)"; Origin = "TEXT(60)"; Destination = "TEXT(60)"
            SupplierName = "TEXT(80)"; ReceiverAddress = "TEXT(42)"
            DispatchTimestampUtc = "DATETIME"; ExpectedDeliveryUtc = "DATETIME"
            Carrier = "TEXT(80)"; TempRange = "TEXT(20)"; Status = "TEXT(20)"
        }
        Key     = "ShipmentID"
    },
    @{
        Name    = "ShipmentEvents"
        Csv     = Join-Path $CsvDir "ShipmentEvents.csv"
        Columns = [ordered]@{
            ShipmentID = "TEXT(20) NOT NULL"; EventName = "TEXT(40) NOT NULL"
            ExpectedEventTimestampUtc = "DATETIME"; HandlerPartner = "TEXT(80)"
            HandlerAddress = "TEXT(42)"; ExpectedQty = "LONG"; RecordedBy = "TEXT(20)"
        }
        Key     = "ShipmentID, EventName"
    }
)

foreach ($t in $tables) {
    if (-not (Test-Path $t.Csv)) { throw "Missing $($t.Csv) - run tools\generate_sample_data.py first." }
}

# Rebuild from scratch so the file always matches the CSVs exactly.
if (Test-Path $DbPath) { Remove-Item $DbPath -Force }

$connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$DbPath"
$catalog = New-Object -ComObject ADOX.Catalog
[void]$catalog.Create($connStr)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($catalog)

$conn = New-Object -ComObject ADODB.Connection
$conn.Open($connStr)
try {
    foreach ($t in $tables) {
        $defs = ($t.Columns.GetEnumerator() | ForEach-Object { "[$($_.Key)] $($_.Value)" }) -join ", "
        [void]$conn.Execute("CREATE TABLE [$($t.Name)] ($defs, CONSTRAINT PK_$($t.Name) PRIMARY KEY ($($t.Key)))")

        # A parameterised command per row: values are bound, never concatenated into SQL.
        $rows = Import-Csv $t.Csv
        $names = @($t.Columns.Keys)
        $colList = ($names | ForEach-Object { "[$_]" }) -join ", "
        $marks = ($names | ForEach-Object { "?" }) -join ", "

        foreach ($r in $rows) {
            $cmd = New-Object -ComObject ADODB.Command
            $cmd.ActiveConnection = $conn
            $cmd.CommandText = "INSERT INTO [$($t.Name)] ($colList) VALUES ($marks)"
            foreach ($n in $names) {
                $type = $t.Columns[$n]
                $raw = $r.$n
                if ($type -like "DATETIME*") {
                    $v = [datetime]::ParseExact($raw, "yyyy-MM-dd HH:mm:ss", [Globalization.CultureInfo]::InvariantCulture)
                    $p = $cmd.CreateParameter($n, 7, 1, 0, $v)            # adDate
                } elseif ($type -like "LONG*") {
                    $p = $cmd.CreateParameter($n, 3, 1, 0, [int]$raw)     # adInteger
                } else {
                    $p = $cmd.CreateParameter($n, 202, 1, 255, [string]$raw)  # adVarWChar
                }
                [void]$cmd.Parameters.Append($p)
            }
            [void]$cmd.Execute()
        }
        $count = $conn.Execute("SELECT COUNT(*) FROM [$($t.Name)]").Fields.Item(0).Value
        Write-Host ("  {0,-15} {1} row(s)" -f $t.Name, $count)
    }
}
finally {
    $conn.Close()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($conn)
}

Write-Host "wrote: $DbPath" -ForegroundColor Green

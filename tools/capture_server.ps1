<#
    capture_server.ps1 - a local stand-in for an SMTP relay and for webhook endpoints.

    The alerting paths (AlertMode = SMTP, TeamsWebhookUrl, SlackWebhookUrl,
    DashboardPushUrl) deliver to external services. To test them end to end without
    sending anything off the machine, point those settings at this server:

        SmtpHost = 127.0.0.1   SmtpPort = 2525   SmtpUseSsl = False
        TeamsWebhookUrl  = http://127.0.0.1:18080/teams
        SlackWebhookUrl  = http://127.0.0.1:18080/slack
        DashboardPushUrl = http://127.0.0.1:18080/powerbi

    Every SMTP message is saved as smtp_<n>.eml and every HTTP request as
    http_<n>_<path>.txt (request line, headers, blank line, body) in -CaptureDir. The
    server accepts what it receives and answers 250 / 200, and stops when -StopFile
    appears or after -TimeoutSeconds.

    It listens on 127.0.0.1 only, uses plain TcpListener (no admin rights or URL ACLs),
    and handles one connection at a time, which is all a single bot run needs.

    Run:  powershell -File tools\capture_server.ps1 -CaptureDir <dir>
#>
param(
    [Parameter(Mandatory = $true)][string]$CaptureDir,
    [int]$SmtpPort = 2525,
    [int]$HttpPort = 18080,
    [int]$TimeoutSeconds = 900,
    [string]$StopFile = ""
)
$ErrorActionPreference = "Stop"

New-Item -ItemType Directory -Force -Path $CaptureDir | Out-Null
if (-not $StopFile) { $StopFile = Join-Path $CaptureDir "stop" }
$loopback = [System.Net.IPAddress]::Loopback

$smtp = New-Object System.Net.Sockets.TcpListener($loopback, $SmtpPort)
$http = New-Object System.Net.Sockets.TcpListener($loopback, $HttpPort)
$smtp.Start()
$http.Start()
Set-Content -Path (Join-Path $CaptureDir "ready") -Value "smtp=$SmtpPort http=$HttpPort" -Encoding ascii

$count = 0
$deadline = (Get-Date).AddSeconds($TimeoutSeconds)

function Read-Line([System.IO.Stream]$stream) {
    $buf = New-Object System.Collections.Generic.List[byte]
    while ($true) {
        $b = $stream.ReadByte()
        if ($b -lt 0) { break }
        if ($b -eq 10) { break }
        if ($b -ne 13) { $buf.Add([byte]$b) }
    }
    return [System.Text.Encoding]::UTF8.GetString($buf.ToArray())
}

function Send-Text([System.IO.Stream]$stream, [string]$text) {
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($text)
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Flush()
}

function Invoke-SmtpSession([System.Net.Sockets.TcpClient]$client) {
    $s = $client.GetStream()
    $s.ReadTimeout = 30000
    Send-Text $s "220 localhost capture ESMTP`r`n"
    $message = New-Object System.Text.StringBuilder
    while ($true) {
        $line = Read-Line $s
        $verb = ($line -split ' ')[0].ToUpperInvariant()
        switch -Regex ($verb) {
            '^(EHLO|HELO)$' { Send-Text $s "250-localhost`r`n250-8BITMIME`r`n250 SIZE 52428800`r`n"; break }
            '^MAIL$'        { [void]$message.AppendLine("X-Capture-" + $line); Send-Text $s "250 OK`r`n"; break }
            '^RCPT$'        { [void]$message.AppendLine("X-Capture-" + $line); Send-Text $s "250 OK`r`n"; break }
            '^DATA$' {
                Send-Text $s "354 End data with <CR><LF>.<CR><LF>`r`n"
                while ($true) {
                    $d = Read-Line $s
                    if ($d -eq ".") { break }
                    if ($d.StartsWith("..")) { $d = $d.Substring(1) }
                    [void]$message.AppendLine($d)
                }
                $script:count++
                $file = Join-Path $CaptureDir ("smtp_{0}.eml" -f $script:count)
                [System.IO.File]::WriteAllText($file, $message.ToString(), [System.Text.Encoding]::UTF8)
                $message.Clear() | Out-Null
                Send-Text $s "250 OK queued`r`n"
                break
            }
            '^(RSET|NOOP)$' { Send-Text $s "250 OK`r`n"; break }
            '^QUIT$'        { Send-Text $s "221 Bye`r`n"; return }
            default         { if ($line -eq "") { return }; Send-Text $s "250 OK`r`n" }
        }
    }
}

function Invoke-HttpSession([System.Net.Sockets.TcpClient]$client) {
    $s = $client.GetStream()
    $s.ReadTimeout = 30000
    $requestLine = Read-Line $s
    $headers = New-Object System.Text.StringBuilder
    $length = 0
    while ($true) {
        $h = Read-Line $s
        if ($h -eq "") { break }
        [void]$headers.AppendLine($h)
        if ($h -match '^Content-Length:\s*(\d+)') { $length = [int]$Matches[1] }
    }
    $body = New-Object byte[] $length
    $read = 0
    while ($read -lt $length) {
        $n = $s.Read($body, $read, $length - $read)
        if ($n -le 0) { break }
        $read += $n
    }
    $path = (($requestLine -split ' ')[1] -replace '[^A-Za-z0-9]', '_').Trim('_')
    $script:count++
    $file = Join-Path $CaptureDir ("http_{0}_{1}.txt" -f $script:count, $path)
    $text = $requestLine + "`r`n" + $headers.ToString() + "`r`n" + [System.Text.Encoding]::UTF8.GetString($body, 0, $read)
    [System.IO.File]::WriteAllText($file, $text, [System.Text.Encoding]::UTF8)
    Send-Text $s "HTTP/1.1 200 OK`r`nContent-Type: text/plain`r`nContent-Length: 2`r`nConnection: close`r`n`r`nok"
}

try {
    while ((Get-Date) -lt $deadline -and -not (Test-Path $StopFile)) {
        if ($smtp.Pending()) {
            $c = $smtp.AcceptTcpClient()
            try { Invoke-SmtpSession $c } catch { } finally { $c.Close() }
        } elseif ($http.Pending()) {
            $c = $http.AcceptTcpClient()
            try { Invoke-HttpSession $c } catch { } finally { $c.Close() }
        } else {
            Start-Sleep -Milliseconds 100
        }
    }
}
finally {
    $smtp.Stop()
    $http.Stop()
}

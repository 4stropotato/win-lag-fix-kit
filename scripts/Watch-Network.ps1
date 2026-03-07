param(
    [string]$Ref1    = "1.1.1.1",
    [string]$Ref2    = "8.8.8.8",
    [int]$IntervalMs = 1000,
    [int]$SpikeMs    = 50,
    [switch]$KeepRawLog,
    [switch]$DeepCapture,
    [int]$DeepCaptureMaxMB = 512,
    [switch]$KeepOutput,
    [switch]$NoAutoFix
)
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host ""
    Write-Host "  ERROR: Run this script as Administrator." -ForegroundColor Red
    Write-Host "  Right-click PowerShell -> Run as Administrator, then run again." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

if ($DeepCaptureMaxMB -lt 64) { $DeepCaptureMaxMB = 64 }
if ($DeepCaptureMaxMB -gt 4096) { $DeepCaptureMaxMB = 4096 }

$script:OutputRoot = Join-Path $env:TEMP "ushie"
$script:SessionDir = Join-Path $script:OutputRoot ("run_{0}" -f (Get-Date).ToString("yyyyMMdd_HHmmss"))
try {
    New-Item -ItemType Directory -Path $script:OutputRoot -Force | Out-Null
    Get-ChildItem -Path $script:OutputRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like 'run_*' -and $_.LastWriteTime -lt (Get-Date).AddHours(-6) } |
        ForEach-Object { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $script:SessionDir -Force | Out-Null
} catch {}

$r1 = [System.Collections.Generic.List[PSCustomObject]]::new()
$r2 = [System.Collections.Generic.List[PSCustomObject]]::new()
$rD = [System.Collections.Generic.List[PSCustomObject]]::new()
$startTime = Get-Date
$script:dotaIP = $null
$script:dotaPop = $null
$script:dotaRegion = "Unknown"
$script:dotaState = "UNKNOWN"
$script:dotaSdrFrontMs = $null
$script:dotaSdrBackMs = $null
$script:dotaSdrTotalMs = $null
$script:dotaLastDetect = [datetime]::MinValue
$script:dotaMisses = 0
$script:CurrentSpikeMs = $SpikeMs
$script:lastAutoSpikeMs = $SpikeMs
$script:dotaConsoleLogPath = $null
$script:eventTimeline = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:PrivateIPv4Pattern = '^(127\.|0\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[01])\.|169\.254\.)'
$script:DeepCaptureEnabled = $DeepCapture.IsPresent
$script:DeepCaptureActive = $false
$script:DeepCaptureDir = $null
$script:DeepCaptureEtlPath = $null
$script:DeepCapturePcapPath = $null
$script:DeepCaptureError = $null
$script:AutoFixEnabled = -not $NoAutoFix
$script:AutoFixActions = [System.Collections.Generic.List[string]]::new()
$script:AutoFixStatus = "Disabled"
$script:AutoFixReason = "-"

# Enable ANSI/VT escape processing so cursor positioning works in modern consoles.
try {
    $sig = '[DllImport("kernel32.dll")] public static extern bool GetConsoleMode(IntPtr h, out uint m);
             [DllImport("kernel32.dll")] public static extern bool SetConsoleMode(IntPtr h, uint m);
             [DllImport("kernel32.dll")] public static extern IntPtr GetStdHandle(int h);'
    $k32 = Add-Type -MemberDefinition $sig -Name K32VT -Namespace Win32 -PassThru -ErrorAction Stop
    $hOut = [Win32.K32VT]::GetStdHandle(-11)
    $mode = 0
    [Win32.K32VT]::GetConsoleMode($hOut, [ref]$mode) | Out-Null
    [Win32.K32VT]::SetConsoleMode($hOut, $mode -bor 4) | Out-Null
} catch {}

function Ping-Once([string]$Address, [int]$TimeoutMs = 2000) {
    $p = [System.Net.NetworkInformation.Ping]::new()
    try {
        $r = $p.Send($Address, $TimeoutMs)
        if ($r.Status -eq "Success") { return $r.RoundtripTime }
    } catch {}
    return $null
}

function Get-PublicIPv4FromEndpoint([string]$Endpoint) {
    if ([string]::IsNullOrWhiteSpace($Endpoint)) { return $null }
    $ep = $Endpoint.Trim()
    if ($ep -eq '*:*') { return $null }
    if ($ep.StartsWith('[')) { return $null }  # ignore IPv6 for relay ping display

    $idx = $ep.LastIndexOf(':')
    if ($idx -lt 1) { return $null }

    $ip = $ep.Substring(0, $idx)
    if ($ip -notmatch '^\d+\.\d+\.\d+\.\d+$') { return $null }
    if ($ip -match $script:PrivateIPv4Pattern) { return $null }
    return $ip
}

function Get-DotaProcess {
    try {
        return (Get-Process -Name dota2,dota2beta,dota2x64 -ErrorAction SilentlyContinue |
            Sort-Object StartTime -Descending |
            Select-Object -First 1
        )
    } catch {}
    return $null
}

function Get-DotaUdpPorts {
    $proc = Get-DotaProcess
    if ($null -eq $proc) { return @() }

    try {
        return @(Get-NetUDPEndpoint -OwningProcess $proc.Id -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty LocalPort -Unique)
    } catch {}

    return @()
}

function Get-DotaConsoleLogPath {
    if ($script:dotaConsoleLogPath -and (Test-Path $script:dotaConsoleLogPath)) {
        return $script:dotaConsoleLogPath
    }

    $candidates = [System.Collections.Generic.List[string]]::new()

    try {
        $proc = Get-DotaProcess
        if ($null -ne $proc -and $proc.Path) {
            $binDir = Split-Path -Path $proc.Path -Parent
            $gameDir = Split-Path -Path (Split-Path -Path $binDir -Parent) -Parent
            if ($gameDir) {
                $candidates.Add((Join-Path $gameDir 'dota\\console.log'))
            }
        }
    } catch {}

    $steamDefault = "$env:ProgramFiles(x86)\\Steam\\steamapps\\common\\dota 2 beta\\game\\dota\\console.log"
    $steamAlt = "$env:ProgramFiles\\Steam\\steamapps\\common\\dota 2 beta\\game\\dota\\console.log"
    $candidates.Add($steamDefault)
    $candidates.Add($steamAlt)

    foreach ($cand in $candidates) {
        if ([string]::IsNullOrWhiteSpace($cand)) { continue }
        if (Test-Path $cand) {
            $script:dotaConsoleLogPath = $cand
            return $cand
        }
    }

    return $null
}

function Add-TimelineEvent([string]$Type, [string]$Detail) {
    if ([string]::IsNullOrWhiteSpace($Type)) { return }
    if ([string]::IsNullOrWhiteSpace($Detail)) { $Detail = "-" }

    if ($script:eventTimeline.Count -gt 0) {
        $last = $script:eventTimeline[$script:eventTimeline.Count - 1]
        if ($last.Type -eq $Type -and $last.Detail -eq $Detail) { return }
    }

    $script:eventTimeline.Add([PSCustomObject]@{
        Time   = (Get-Date).ToString("HH:mm:ss")
        Type   = $Type
        Detail = $Detail
    })

    while ($script:eventTimeline.Count -gt 25) {
        $script:eventTimeline.RemoveAt(0)
    }
}

function Start-DeepCapture {
    if (-not $script:DeepCaptureEnabled) { return }

    try {
        $script:DeepCaptureDir = Join-Path $script:SessionDir "deepcap"
        if (Test-Path $script:DeepCaptureDir) {
            Remove-Item $script:DeepCaptureDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        New-Item -ItemType Directory -Path $script:DeepCaptureDir -Force | Out-Null

        $script:DeepCaptureEtlPath = Join-Path $script:DeepCaptureDir "network_trace.etl"
        $script:DeepCapturePcapPath = Join-Path $script:DeepCaptureDir "network_trace.pcapng"

        try { netsh trace stop | Out-Null } catch {}

        netsh trace start capture=yes report=no persistent=no overwrite=yes maxsize=$DeepCaptureMaxMB tracefile="$($script:DeepCaptureEtlPath)" | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "netsh trace start failed (exit code $LASTEXITCODE)."
        }

        $script:DeepCaptureActive = $true
        $script:DeepCaptureError = $null
        Add-TimelineEvent "DeepCapture" ("started ({0}MB max)" -f $DeepCaptureMaxMB)
    } catch {
        $script:DeepCaptureActive = $false
        $script:DeepCaptureEnabled = $false
        $script:DeepCaptureError = $_.Exception.Message
        Add-TimelineEvent "DeepCapture" ("start failed: {0}" -f $script:DeepCaptureError)
    }
}

function Stop-DeepCapture {
    if (-not $script:DeepCaptureEnabled) { return }
    if (-not $script:DeepCaptureActive) { return }

    $cabPath = if ($script:DeepCaptureEtlPath) { [System.IO.Path]::ChangeExtension($script:DeepCaptureEtlPath, ".cab") } else { $null }

    try {
        netsh trace stop | Out-Null
    } catch {
        if ([string]::IsNullOrWhiteSpace($script:DeepCaptureError)) {
            $script:DeepCaptureError = $_.Exception.Message
        }
    }

    $script:DeepCaptureActive = $false

    if (-not (Test-Path $script:DeepCaptureEtlPath)) {
        if ([string]::IsNullOrWhiteSpace($script:DeepCaptureError)) {
            $script:DeepCaptureError = "Capture stopped but ETL was not found."
        }
        $script:DeepCapturePcapPath = $null
        Add-TimelineEvent "DeepCapture" "stopped (etl missing)"
        return
    }

    $converted = $false
    $etl2pcapng = Get-Command etl2pcapng.exe -ErrorAction SilentlyContinue
    if ($null -ne $etl2pcapng) {
        try {
            & $etl2pcapng.Source $script:DeepCaptureEtlPath $script:DeepCapturePcapPath | Out-Null
            if ($LASTEXITCODE -eq 0 -and (Test-Path $script:DeepCapturePcapPath)) {
                $converted = $true
            }
        } catch {}
    }

    if (-not $converted) {
        $pktmon = Get-Command pktmon.exe -ErrorAction SilentlyContinue
        if ($null -ne $pktmon) {
            try {
                & $pktmon.Source etl2pcap $script:DeepCaptureEtlPath --out $script:DeepCapturePcapPath | Out-Null
                if ($LASTEXITCODE -eq 0 -and (Test-Path $script:DeepCapturePcapPath)) {
                    $converted = $true
                }
            } catch {}
        }
    }

    if ($converted) {
        Add-TimelineEvent "DeepCapture" "pcapng exported"
    } else {
        $script:DeepCapturePcapPath = $null
        if ([string]::IsNullOrWhiteSpace($script:DeepCaptureError)) {
            $script:DeepCaptureError = "PCAP export tool not found or conversion failed."
        }
        Add-TimelineEvent "DeepCapture" "stopped (etl only)"
    }

    if ($cabPath -and (Test-Path $cabPath)) {
        try { Remove-Item $cabPath -Force -ErrorAction SilentlyContinue } catch {}
    }
}

function Get-RegionFromPop([string]$Pop) {
    if ([string]::IsNullOrWhiteSpace($Pop)) { return "Unknown" }
    $p = $Pop.ToLowerInvariant()
    if ($p -match '^(tyo|jpn)#') { return "Japan" }
    if ($p -match '^(seo|icn|kor)#') { return "Korea" }
    if ($p -match '^(sgp|sin)#') { return "Singapore" }
    if ($p -match '^hkg#') { return "HongKong" }
    if ($p -match '^(man|sea)#') { return "SEA" }
    return "Unknown"
}

function Get-UiSimpleState([string]$UiToken) {
    if ([string]::IsNullOrWhiteSpace($UiToken)) { return "UNKNOWN" }
    switch ($UiToken) {
        'DOTA_GAME_UI_DOTA_INGAME' { return "INGAME" }
        'DOTA_GAME_UI_STATE_DASHBOARD' { return "LOBBY" }
        'DOTA_GAME_UI_STATE_LOADING_SCREEN' { return "CONNECTING" }
        default { return "UNKNOWN" }
    }
}

function Get-SdrMetricsFromLine([string]$Line) {
    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }

    # Example: Ping = 6 = 6+0 (front+back)
    if ($Line -match 'Ping\s*=\s*(\d+)\s*=\s*(\d+)\+(\d+)\s+\(front\+back\)') {
        return [PSCustomObject]@{
            TotalMs = [int]$matches[1]
            FrontMs = [int]$matches[2]
            BackMs  = [int]$matches[3]
        }
    }

    # Example: Ping = 6+0=6 (front+back=total)
    if ($Line -match 'Ping\s*=\s*(\d+)\+(\d+)\s*=\s*(\d+)\s+\(front\+back=total\)') {
        return [PSCustomObject]@{
            TotalMs = [int]$matches[3]
            FrontMs = [int]$matches[1]
            BackMs  = [int]$matches[2]
        }
    }

    # Example: Ping = 66 = 36+30+0 (front+interior+remote)
    if ($Line -match 'Ping\s*=\s*(\d+)\s*=\s*(\d+)\+(\d+)\+(\d+)\s+\(front\+interior\+remote\)') {
        return [PSCustomObject]@{
            TotalMs = [int]$matches[1]
            FrontMs = [int]$matches[2]
            BackMs  = ([int]$matches[3] + [int]$matches[4])
        }
    }

    return $null
}

function Get-DotaServerIPFromConsoleLog {
    $proc = Get-DotaProcess
    if ($null -eq $proc) {
        return [PSCustomObject]@{
            UiState = "OFFLINE"; UiToken = $null
            IP = $null; Pop = $null; Region = "Unknown"
            SdrFrontMs = $null; SdrBackMs = $null; SdrTotalMs = $null
            Source = "ConsoleLog"
        }
    }

    $logPath = Get-DotaConsoleLogPath
    if (-not $logPath) {
        return [PSCustomObject]@{
            UiState = "UNKNOWN"; UiToken = $null
            IP = $null; Pop = $null; Region = "Unknown"
            SdrFrontMs = $null; SdrBackMs = $null; SdrTotalMs = $null
            Source = "ConsoleLog"
        }
    }

    try {
        $meta = Get-Item -Path $logPath -ErrorAction SilentlyContinue
        if ($null -eq $meta) {
            return [PSCustomObject]@{
                UiState = "UNKNOWN"; UiToken = $null
                IP = $null; Pop = $null; Region = "Unknown"
                SdrFrontMs = $null; SdrBackMs = $null; SdrTotalMs = $null
                Source = "ConsoleLog"
            }
        }

        $lines = Get-Content -Path $logPath -Tail 12000 -ErrorAction SilentlyContinue
        if (-not $lines) {
            return [PSCustomObject]@{
                UiState = "UNKNOWN"; UiToken = $null
                IP = $null; Pop = $null; Region = "Unknown"
                SdrFrontMs = $null; SdrBackMs = $null; SdrTotalMs = $null
                Source = "ConsoleLog"
            }
        }

        $uiToken = $null
        $uiState = "UNKNOWN"

        # If latest UI state is dashboard, do not keep stale relay from previous matches.
        for ($i = $lines.Count - 1; $i -ge 0; $i--) {
            $line = $lines[$i]
            if ($line -match 'ChangeGameUIState:\s+\S+\s+->\s+(\S+)') {
                $uiToken = $matches[1]
                $uiState = Get-UiSimpleState $uiToken
                break
            }
        }

        if ($uiState -eq "LOBBY") {
            return [PSCustomObject]@{
                UiState = $uiState; UiToken = $uiToken
                IP = $null; Pop = $null; Region = "Unknown"
                SdrFrontMs = $null; SdrBackMs = $null; SdrTotalMs = $null
                Source = "ConsoleLog"
            }
        }

        for ($i = $lines.Count - 1; $i -ge 0; $i--) {
            $line = $lines[$i]
            if ($line -match 'SteamNetSockets.*Switched primary to\s+(\S+)\s+\((\d{1,3}(?:\.\d{1,3}){3}):\d+\)') {
                $pop = $matches[1]
                $ip = $matches[2]
                if ($ip -notmatch $script:PrivateIPv4Pattern) {
                    $sdr = Get-SdrMetricsFromLine $line
                    return [PSCustomObject]@{
                        UiState = $uiState; UiToken = $uiToken
                        IP = $ip; Pop = $pop; Region = (Get-RegionFromPop $pop)
                        SdrFrontMs = if ($sdr) { $sdr.FrontMs } else { $null }
                        SdrBackMs  = if ($sdr) { $sdr.BackMs } else { $null }
                        SdrTotalMs = if ($sdr) { $sdr.TotalMs } else { $null }
                        Source = "ConsoleLog"
                    }
                }
            }
            if ($line -match 'SteamNetSockets.*Selecting\s+(\S+)\s+\((\d{1,3}(?:\.\d{1,3}){3}):\d+\)\s+as\s+primary') {
                $pop = $matches[1]
                $ip = $matches[2]
                if ($ip -notmatch $script:PrivateIPv4Pattern) {
                    $sdr = Get-SdrMetricsFromLine $line
                    return [PSCustomObject]@{
                        UiState = $uiState; UiToken = $uiToken
                        IP = $ip; Pop = $pop; Region = (Get-RegionFromPop $pop)
                        SdrFrontMs = if ($sdr) { $sdr.FrontMs } else { $null }
                        SdrBackMs  = if ($sdr) { $sdr.BackMs } else { $null }
                        SdrTotalMs = if ($sdr) { $sdr.TotalMs } else { $null }
                        Source = "ConsoleLog"
                    }
                }
            }
            if ($line -match 'SteamNetSockets.*Requesting\s+session\s+from\s+(\S+)\s+\((\d{1,3}(?:\.\d{1,3}){3}):\d+\).*Rank=1') {
                $pop = $matches[1]
                $ip = $matches[2]
                if ($ip -notmatch $script:PrivateIPv4Pattern) {
                    $sdr = Get-SdrMetricsFromLine $line
                    return [PSCustomObject]@{
                        UiState = $uiState; UiToken = $uiToken
                        IP = $ip; Pop = $pop; Region = (Get-RegionFromPop $pop)
                        SdrFrontMs = if ($sdr) { $sdr.FrontMs } else { $null }
                        SdrBackMs  = if ($sdr) { $sdr.BackMs } else { $null }
                        SdrTotalMs = if ($sdr) { $sdr.TotalMs } else { $null }
                        Source = "ConsoleLog"
                    }
                }
            }
            if ($line -match '\[Networking\]\s+Primary router:\s+(\S+)\s+\((\d{1,3}(?:\.\d{1,3}){3}):\d+\)') {
                $pop = $matches[1]
                $ip = $matches[2]
                if ($ip -notmatch $script:PrivateIPv4Pattern) {
                    $sdr = Get-SdrMetricsFromLine $line
                    return [PSCustomObject]@{
                        UiState = $uiState; UiToken = $uiToken
                        IP = $ip; Pop = $pop; Region = (Get-RegionFromPop $pop)
                        SdrFrontMs = if ($sdr) { $sdr.FrontMs } else { $null }
                        SdrBackMs  = if ($sdr) { $sdr.BackMs } else { $null }
                        SdrTotalMs = if ($sdr) { $sdr.TotalMs } else { $null }
                        Source = "ConsoleLog"
                    }
                }
            }
        }
    } catch {}
    return [PSCustomObject]@{
        UiState = "UNKNOWN"; UiToken = $null
        IP = $null; Pop = $null; Region = "Unknown"
        SdrFrontMs = $null; SdrBackMs = $null; SdrTotalMs = $null
        Source = "ConsoleLog"
    }
}

function Get-DotaServerIPFromNetstat([int[]]$LocalPorts) {
    if ($LocalPorts.Count -eq 0) { return $null }

    try {
        $portSet = [System.Collections.Generic.HashSet[int]]::new()
        foreach ($p in $LocalPorts) { [void]$portSet.Add([int]$p) }

        $lines = netstat -ano -p udp 2>$null
        foreach ($line in $lines) {
            # Example: UDP 192.168.1.10:54811 45.121.184.26:27020 1234
            if ($line -notmatch '^\s*UDP\s+(\S+):(\d+)\s+(\S+)\s+(\d+)\s*$') { continue }

            $localPort = [int]$matches[2]
            if (-not $portSet.Contains($localPort)) { continue }

            $remoteIP = Get-PublicIPv4FromEndpoint $matches[3]
            if ($null -ne $remoteIP) { return $remoteIP }
        }
    } catch {}

    return $null
}

function Get-CaptureCandidateIPs {
    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $out = [System.Collections.Generic.List[string]]::new()

    $add = {
        param([string]$ip)
        if ([string]::IsNullOrWhiteSpace($ip)) { return }
        if ($ip -notmatch '^\d+\.\d+\.\d+\.\d+$') { return }
        if ($ip -match '^(127\.|169\.254\.)') { return }
        if ($set.Add($ip)) { [void]$out.Add($ip) }
    }

    try {
        $defRoute = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
            Sort-Object -Property @{Expression = { $_.RouteMetric } }, @{Expression = { if ($_.InterfaceMetric -ne $null) { $_.InterfaceMetric } else { $_.ifMetric } } } |
            Select-Object -First 1

        if ($null -ne $defRoute) {
            $ifIndex = if ($defRoute.InterfaceIndex) { $defRoute.InterfaceIndex } else { $defRoute.ifIndex }
            if ($ifIndex) {
                Get-NetIPAddress -AddressFamily IPv4 -InterfaceIndex $ifIndex -ErrorAction SilentlyContinue |
                    ForEach-Object { & $add $_.IPAddress }
            }
        }
    } catch {}

    try {
        Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
            Where-Object { $_.Status -eq 'Up' } |
            ForEach-Object {
                Get-NetIPAddress -AddressFamily IPv4 -InterfaceIndex $_.ifIndex -ErrorAction SilentlyContinue |
                    ForEach-Object { & $add $_.IPAddress }
            }
    } catch {}

    try {
        Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            ForEach-Object { & $add $_.IPAddress }
    } catch {}

    return @($out)
}

function Get-DotaServerIPFromRawSniff([int[]]$LocalPorts) {
    if ($LocalPorts.Count -eq 0) { return $null }

    $portSet = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($p in $LocalPorts) { [void]$portSet.Add([int]$p) }

    $candidateIPs = Get-CaptureCandidateIPs
    if ($candidateIPs.Count -eq 0) { return $null }

    foreach ($localIP in $candidateIPs) {
        $sock = $null
        try {
            $sock = [System.Net.Sockets.Socket]::new(
                [System.Net.Sockets.AddressFamily]::InterNetwork,
                [System.Net.Sockets.SocketType]::Raw,
                [System.Net.Sockets.ProtocolType]::IP
            )
            $sock.Bind([System.Net.IPEndPoint]::new([System.Net.IPAddress]::Parse($localIP), 0))
            $sock.SetSocketOption('IP', 'HeaderIncluded', $true)
            $sock.IOControl([System.Net.Sockets.IOControlCode]::ReceiveAll, [byte[]](1,0,0,0), [byte[]](1,0,0,0)) | Out-Null
            $sock.ReceiveTimeout = 200

            $buf = New-Object byte[] 65535
            $limit = (Get-Date).AddMilliseconds(900)

            while ((Get-Date) -lt $limit) {
                try {
                    $len = $sock.Receive($buf)
                    if ($len -lt 28) { continue }
                    if ($buf[9] -ne 17) { continue }  # UDP only

                    $ihl = ($buf[0] -band 0x0F) * 4
                    if ($len -lt ($ihl + 8)) { continue }

                    $srcPort = [uint16](($buf[$ihl] -shl 8) -bor $buf[$ihl + 1])
                    $dstPort = [uint16](($buf[$ihl + 2] -shl 8) -bor $buf[$ihl + 3])

                    $fromDota = $portSet.Contains([int]$srcPort)
                    $toDota = $portSet.Contains([int]$dstPort)
                    if (-not $fromDota -and -not $toDota) { continue }

                    $srcIP = "$($buf[12]).$($buf[13]).$($buf[14]).$($buf[15])"
                    $dstIP = "$($buf[16]).$($buf[17]).$($buf[18]).$($buf[19])"
                    $remoteIP = if ($fromDota) { $dstIP } else { $srcIP }

                    if ($remoteIP -match '^\d+\.\d+\.\d+\.\d+$' -and $remoteIP -notmatch $script:PrivateIPv4Pattern) {
                        return $remoteIP
                    }
                } catch {}
            }
        } catch {}
        finally {
            if ($null -ne $sock) {
                try { $sock.Dispose() } catch {}
            }
        }
    }

    return $null
}

function Get-DotaServerIP {
    $localPorts = Get-DotaUdpPorts

    # Most reliable for Source 2 SDR: parse latest relay and state from Dota console.log.
    $fromLog = Get-DotaServerIPFromConsoleLog
    if ($null -ne $fromLog) {
        $script:dotaState = $fromLog.UiState
        $script:dotaSdrFrontMs = $fromLog.SdrFrontMs
        $script:dotaSdrBackMs = $fromLog.SdrBackMs
        $script:dotaSdrTotalMs = $fromLog.SdrTotalMs
    }

    if ($null -ne $fromLog -and $fromLog.IP) {
        $script:dotaPop = $fromLog.Pop
        $script:dotaRegion = $fromLog.Region
        return $fromLog.IP
    }

    # Log can say LOBBY / OFFLINE. In that case do not force old relay.
    if ($null -ne $fromLog -and ($fromLog.UiState -in @("LOBBY", "OFFLINE"))) {
        $script:dotaPop = $null
        $script:dotaRegion = "Unknown"
        return $null
    }

    if ($localPorts.Count -eq 0) { return $null }

    # Fast path: connected UDP endpoints visible in netstat.
    $remote = Get-DotaServerIPFromNetstat -LocalPorts $localPorts
    if ($null -ne $remote) {
        $script:dotaPop = $null
        $script:dotaRegion = "Unknown"
        return $remote
    }

    # Fallback: raw sniff on active interfaces.
    $raw = Get-DotaServerIPFromRawSniff -LocalPorts $localPorts
    if ($null -ne $raw) {
        $script:dotaPop = $null
        $script:dotaRegion = "Unknown"
    }
    return $raw
}

function Get-AutoSpikeMs([string]$Pop, [int]$FallbackMs) {
    if ([string]::IsNullOrWhiteSpace($Pop)) { return $FallbackMs }
    $p = $Pop.ToLowerInvariant()

    if ($p -match '^(tyo|jpn)#') { return 50 }
    if ($p -match '^(seo|icn|kor)#') { return 70 }
    if ($p -match '^(sgp|sin|hkg|man|sea)#') { return 90 }
    return 80
}

function Get-PercentileValue([double[]]$Values, [double]$Percentile) {
    if ($null -eq $Values -or $Values.Count -eq 0) { return $null }
    if ($Percentile -le 0) { return $Values[0] }
    if ($Percentile -ge 100) { return $Values[$Values.Count - 1] }

    $rank = ($Percentile / 100.0) * ($Values.Count - 1)
    $low = [int][Math]::Floor($rank)
    $high = [int][Math]::Ceiling($rank)
    if ($low -eq $high) { return $Values[$low] }

    $weight = $rank - $low
    return ($Values[$low] + (($Values[$high] - $Values[$low]) * $weight))
}

function Get-SessionGrade($DotaStats, [int]$ThresholdMs) {
    if ($null -eq $DotaStats -or $DotaStats.Total -eq 0) {
        return [PSCustomObject]@{ Grade = "N/A"; Score = 0; Reason = "No Dota samples"; }
    }

    $score = 100
    $reasons = [System.Collections.Generic.List[string]]::new()
    $lossPct = [double]$DotaStats.LossPct
    $spikePct = if ($DotaStats.Total -gt 0) { ([double]$DotaStats.Spikes / [double]$DotaStats.Total) * 100.0 } else { 0.0 }

    if ($lossPct -gt 0) {
        $pen = [Math]::Min(55, [Math]::Ceiling($lossPct * 12))
        $score -= $pen
        $reasons.Add("loss $([Math]::Round($lossPct,1))%")
    }
    if ($DotaStats.JitterMs -gt 1.5) {
        $pen = [Math]::Min(20, [Math]::Ceiling(($DotaStats.JitterMs - 1.5) * 4))
        $score -= $pen
        $reasons.Add("jitter $([Math]::Round($DotaStats.JitterMs,1))ms")
    }
    if ($spikePct -gt 5) {
        $pen = [Math]::Min(30, [Math]::Ceiling(($spikePct - 5) * 0.6))
        $score -= $pen
        $reasons.Add("spikes $([Math]::Round($spikePct,1))%")
    }
    if ($DotaStats.P99Ms -gt ($ThresholdMs + 18)) {
        $pen = [Math]::Min(18, [Math]::Ceiling(($DotaStats.P99Ms - ($ThresholdMs + 18)) * 0.5))
        $score -= $pen
        $reasons.Add("high p99 $([Math]::Round($DotaStats.P99Ms,1))ms")
    }
    if ($DotaStats.LongestBurst -ge 10) {
        $pen = [Math]::Min(15, [Math]::Floor($DotaStats.LongestBurst / 3))
        $score -= $pen
        $reasons.Add("burst x$($DotaStats.LongestBurst)")
    }

    $score = [Math]::Max(0, [Math]::Min(100, $score))
    $grade =
        if ($score -ge 93) { "S" }
        elseif ($score -ge 85) { "A" }
        elseif ($score -ge 75) { "B" }
        elseif ($score -ge 65) { "C" }
        elseif ($score -ge 50) { "D" }
        else { "F" }

    return [PSCustomObject]@{
        Grade = $grade
        Score = $score
        Reason = if ($reasons.Count -gt 0) { ($reasons -join ", ") } else { "stable" }
    }
}

function Write-PinnedLine([string]$text, [string]$color) {
    $width = 80
    try {
        $width = [Math]::Max([Console]::WindowWidth - 1, 60)
    } catch {}

    if ($script:UsePinnedConsole) {
        Write-Host $text.PadRight($width) -ForegroundColor $color
    } else {
        Write-Host $text -ForegroundColor $color
    }
}

function Write-PingPinned([string]$ts, [string]$label, $ms, [int]$spike) {
    $line = "  [{0}] {1,-22}" -f $ts, $label
    if ($null -eq $ms) {
        Write-PinnedLine ($line + "  TIMEOUT") "Red"
    } elseif ($ms -ge $spike) {
        Write-PinnedLine ($line + ("  {0,4}ms  *** SPIKE ***" -f $ms)) "Yellow"
    } else {
        Write-PinnedLine ($line + ("  {0,4}ms" -f $ms)) "Green"
    }
}

function Get-Stats([System.Collections.Generic.List[PSCustomObject]]$data, [int]$ThresholdMs) {
    $total = $data.Count
    if ($total -eq 0) { return $null }

    $lost = ($data | Where-Object { $null -eq $_.Ms }).Count
    $ok = @($data | Where-Object { $null -ne $_.Ms })

    $avg = $min = $max = $jitter = $null
    $p95 = $p99 = $null
    $spikes = 0
    $burstCount = 0
    $longestBurst = 0

    if ($ok.Count -gt 0) {
        $vals = @($ok | ForEach-Object { [double]$_.Ms })
        $avg = [math]::Round(($vals | Measure-Object -Average).Average, 2)
        $min = [double](($vals | Measure-Object -Minimum).Minimum)
        $max = [double](($vals | Measure-Object -Maximum).Maximum)
        $spikes = ($vals | Where-Object { $_ -ge $ThresholdMs }).Count

        if ($vals.Count -gt 1) {
            $diffs = for ($i = 1; $i -lt $vals.Count; $i++) {
                [math]::Abs($vals[$i] - $vals[$i - 1])
            }
            $jitter = [math]::Round(($diffs | Measure-Object -Average).Average, 2)
        } else {
            $jitter = 0.0
        }

        $sortedVals = @($vals | Sort-Object)
        $p95 = [math]::Round((Get-PercentileValue -Values $sortedVals -Percentile 95), 2)
        $p99 = [math]::Round((Get-PercentileValue -Values $sortedVals -Percentile 99), 2)

        $run = 0
        foreach ($v in $vals) {
            if ($v -ge $ThresholdMs) {
                $run++
                if ($run -eq 1) { $burstCount++ }
                if ($run -gt $longestBurst) { $longestBurst = $run }
            } else {
                $run = 0
            }
        }
    }

    return [PSCustomObject]@{
        Total = $total
        Lost = $lost
        LossPct = [math]::Round($lost / $total * 100, 1)
        Avg = $avg
        Min = $min
        Max = $max
        Jitter = $jitter
        P95 = $p95
        P99 = $p99
        AvgMs = $avg
        MinMs = $min
        MaxMs = $max
        JitterMs = $jitter
        P95Ms = $p95
        P99Ms = $p99
        Spikes = $spikes
        BurstCount = $burstCount
        LongestBurst = $longestBurst
    }
}

function Format-MsValue($v) {
    if ($null -eq $v) { return "n/a" }
    return ("{0}ms" -f ([Math]::Round([double]$v, 2)))
}

function Save-SessionHistory(
    [string]$Path,
    [datetime]$StartTime,
    [timespan]$Duration,
    [string]$RelayPop,
    [string]$RelayIP,
    [string]$RelayRegion,
    $DotaStats,
    [int]$ThresholdMs,
    $GradeObj
) {
    try {
        $row = [PSCustomObject]@{
            StartedAt      = $StartTime.ToString("yyyy-MM-dd HH:mm:ss")
            DurationSec    = [int]$Duration.TotalSeconds
            RelayPop       = $(if ($RelayPop) { $RelayPop } else { "" })
            RelayIP        = $(if ($RelayIP) { $RelayIP } else { "" })
            RelayRegion    = $(if ($RelayRegion) { $RelayRegion } else { "Unknown" })
            AvgMs          = $(if ($DotaStats) { $DotaStats.AvgMs } else { $null })
            P95Ms          = $(if ($DotaStats) { $DotaStats.P95Ms } else { $null })
            P99Ms          = $(if ($DotaStats) { $DotaStats.P99Ms } else { $null })
            JitterMs       = $(if ($DotaStats) { $DotaStats.JitterMs } else { $null })
            LossPct        = $(if ($DotaStats) { $DotaStats.LossPct } else { $null })
            Spikes         = $(if ($DotaStats) { $DotaStats.Spikes } else { $null })
            BurstCount     = $(if ($DotaStats) { $DotaStats.BurstCount } else { $null })
            LongestBurst   = $(if ($DotaStats) { $DotaStats.LongestBurst } else { $null })
            SpikeThreshold = $ThresholdMs
            Grade          = $GradeObj.Grade
            Score          = $GradeObj.Score
            Reason         = $GradeObj.Reason
        }

        if (Test-Path $Path) {
            $row | Export-Csv -Path $Path -Append -NoTypeInformation -Encoding UTF8
        } else {
            $row | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        }
    } catch {}
}

function Add-AutoFixAction([string]$Message) {
    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    $script:AutoFixActions.Add($Message)
}

function Get-AutoFixDnsWinner {
    $ifc = Get-NetIPConfiguration -ErrorAction SilentlyContinue |
        Where-Object { $_.IPv4DefaultGateway -and $_.NetAdapter -and $_.NetAdapter.Status -eq 'Up' } |
        Select-Object -First 1
    if ($null -eq $ifc) { return $null }

    $current = @()
    try {
        $current = @(Get-DnsClientServerAddress -InterfaceIndex $ifc.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses
        $current = @($current | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' })
    } catch {}

    $candidates = [System.Collections.Generic.List[PSCustomObject]]::new()
    if ($current.Count -gt 0) {
        $candidates.Add([PSCustomObject]@{ Name = "CurrentAdapterDNS"; Servers = @($current) })
    }
    $candidates.Add([PSCustomObject]@{ Name = "Cloudflare"; Servers = @("1.1.1.1","1.0.0.1") })
    $candidates.Add([PSCustomObject]@{ Name = "Google"; Servers = @("8.8.8.8","8.8.4.4") })
    $candidates.Add([PSCustomObject]@{ Name = "Quad9"; Servers = @("9.9.9.9","149.112.112.112") })
    $candidates.Add([PSCustomObject]@{ Name = "OpenDNS"; Servers = @("208.67.222.222","208.67.220.220") })
    $candidates.Add([PSCustomObject]@{ Name = "AdGuard"; Servers = @("94.140.14.14","94.140.15.15") })

    $best = $null
    foreach ($cand in $candidates) {
        $primary = $cand.Servers[0]
        $samples = [System.Collections.Generic.List[double]]::new()
        for ($i = 0; $i -lt 2; $i++) {
            $ms = Ping-Once -Address $primary -TimeoutMs 1000
            if ($null -ne $ms) {
                $samples.Add([double]$ms)
            } else {
                $samples.Add(2500.0)
            }
        }

        $score = [Math]::Round((($samples | Measure-Object -Average).Average), 2)
        if ($null -eq $best -or $score -lt $best.Score) {
            $best = [PSCustomObject]@{
                Name = $cand.Name
                Score = $score
                Servers = $cand.Servers
                InterfaceIndex = $ifc.InterfaceIndex
                InterfaceAlias = $ifc.InterfaceAlias
                Current = @($current)
            }
        }
    }
    return $best
}

function Apply-AutoFixDns {
    try {
        $winner = Get-AutoFixDnsWinner
        if ($null -eq $winner) {
            Add-AutoFixAction "DNS auto-select skipped (no active gateway adapter)."
            return
        }

        $currentCsv = ($winner.Current -join ",")
        $targetCsv = ($winner.Servers -join ",")
        $changed = ($currentCsv -ne $targetCsv)

        if ($changed) {
            Set-DnsClientServerAddress -InterfaceIndex $winner.InterfaceIndex -ServerAddresses $winner.Servers -ErrorAction Stop
            Add-AutoFixAction ("DNS set to {0} on {1} ({2}ms)" -f $winner.Name, $winner.InterfaceAlias, $winner.Score)
        } else {
            Add-AutoFixAction ("DNS already best ({0}) on {1} ({2}ms)" -f $winner.Name, $winner.InterfaceAlias, $winner.Score)
        }
    } catch {
        Add-AutoFixAction ("DNS auto-select failed: {0}" -f $_.Exception.Message)
    }
}

function Apply-AutoFixCommon {
    try { Clear-DnsClientCache -ErrorAction SilentlyContinue } catch {}
    try { ipconfig /flushdns | Out-Null } catch {}
    Add-AutoFixAction "Flushed DNS cache."

    try { netsh int tcp set global rss=enabled | Out-Null } catch {}
    try { netsh int tcp set global autotuninglevel=normal | Out-Null } catch {}
    try { netsh int tcp set global ecncapability=disabled | Out-Null } catch {}
    Add-AutoFixAction "Applied TCP sanity defaults (RSS on, AutoTune normal, ECN off)."
}

function Invoke-NetworkAutoFix($Ref1Stats, $Ref2Stats, $DotaStats, [int]$ThresholdMs) {
    if (-not $script:AutoFixEnabled) {
        $script:AutoFixStatus = "Disabled"
        $script:AutoFixReason = "-"
        return
    }

    $script:AutoFixStatus = "Skipped"
    $script:AutoFixReason = "No action needed"
    Add-TimelineEvent "AutoFix" "evaluate"

    if ($null -eq $DotaStats -or $DotaStats.Total -lt 30) {
        $script:AutoFixReason = "Not enough Dota samples (<30)"
        Add-AutoFixAction "AutoFix skipped: gather more samples first."
        Add-TimelineEvent "AutoFix" "skip-not-enough-samples"
        return
    }

    $ref1Bad = ($null -eq $Ref1Stats -or $Ref1Stats.LossPct -gt 1.0 -or $Ref1Stats.JitterMs -gt 3.0 -or $Ref1Stats.P99Ms -gt 45)
    $ref2Bad = ($null -eq $Ref2Stats -or $Ref2Stats.LossPct -gt 1.0 -or $Ref2Stats.JitterMs -gt 3.0 -or $Ref2Stats.P99Ms -gt 45)
    $refsBad = ($ref1Bad -or $ref2Bad)

    $dotaBad = (
        $DotaStats.LossPct -gt 0.5 -or
        $DotaStats.JitterMs -gt 2.0 -or
        $DotaStats.P99Ms -gt ($ThresholdMs + 15) -or
        $DotaStats.Spikes -gt [Math]::Max(3, [int]($DotaStats.Total * 0.02))
    )

    if (-not $dotaBad -and -not $refsBad) {
        $script:AutoFixReason = "Session already stable"
        Add-AutoFixAction "AutoFix skipped: metrics already stable."
        Add-TimelineEvent "AutoFix" "skip-stable"
        return
    }

    if ($dotaBad -and -not $refsBad) {
        $script:AutoFixReason = "Detected game-path/route issue"
    } elseif ($refsBad) {
        $script:AutoFixReason = "Detected local/ISP path issue"
    } else {
        $script:AutoFixReason = "Detected mixed latency issue"
    }

    Apply-AutoFixCommon
    Apply-AutoFixDns

    $script:AutoFixStatus = "Applied"
    Add-TimelineEvent "AutoFix" "applied"
}

$script:UsePinnedConsole = $true
$w = 80
$pingRow = -1

Write-Host ""
Write-Host "  ==========================================" -ForegroundColor Cyan
Write-Host "         USHIE NETWORK MONITOR" -ForegroundColor Cyan
Write-Host "  ==========================================" -ForegroundColor Cyan
Write-Host "  Ref1     : $Ref1" -ForegroundColor Gray
Write-Host "  Ref2     : $Ref2" -ForegroundColor Gray
Write-Host "  Dota     : auto-detecting (queue a match first)" -ForegroundColor Gray
Write-Host "  Interval : ${IntervalMs}ms   Spike: auto (base >=${SpikeMs}ms)" -ForegroundColor Gray
if ($script:DeepCaptureEnabled) {
    Write-Host "  DeepCap  : ON (packet capture + optional pcapng export)" -ForegroundColor Gray
} else {
    Write-Host "  DeepCap  : OFF (add -DeepCapture for forensic packet capture)" -ForegroundColor Gray
}
Write-Host ("  AutoFix  : {0}" -f $(if ($script:AutoFixEnabled) { "ON (post-test automatic remediation)" } else { "OFF (-NoAutoFix)" })) -ForegroundColor Gray
Write-Host "  Output   : TEMP (auto-clean after run; use -KeepOutput to retain)" -ForegroundColor Gray
Write-Host "  Started  : $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
Write-Host "  Ctrl+C to stop and generate report." -ForegroundColor Gray
Write-Host ""

# Print 3 placeholder lines, then record their start row for in-place overwrite.
try {
    $w = [Math]::Max([Console]::WindowWidth - 1, 60)
    Write-Host "  [--:--:--] initializing...".PadRight($w)
    Write-Host "  [--:--:--] initializing...".PadRight($w)
    Write-Host "  [--:--:--] initializing...".PadRight($w)
    $pingRow = [Console]::CursorTop - 3
} catch {
    $script:UsePinnedConsole = $false
}

try {
    Add-TimelineEvent "WatcherStart" ("Interval={0}ms BaseSpike={1}ms" -f $IntervalMs, $SpikeMs)

    if ($script:DeepCaptureEnabled) {
        Start-DeepCapture
        if ($script:DeepCaptureActive) {
            Write-Host ("  DeepCap started: $($script:DeepCaptureEtlPath)") -ForegroundColor DarkGray
        } else {
            Write-Host ("  DeepCap warning: $($script:DeepCaptureError)") -ForegroundColor Yellow
        }
    }

    $prevState = $script:dotaState
    $prevRelayKey = ""

    while ($true) {
        $ts = (Get-Date).ToString("HH:mm:ss")

        $ms1 = Ping-Once $Ref1
        $ms2 = Ping-Once $Ref2

        $r1.Add([PSCustomObject]@{ Time = $ts; Ms = $ms1 })
        $r2.Add([PSCustomObject]@{ Time = $ts; Ms = $ms2 })

        if ($script:dotaIP -and $script:dotaIP -notmatch '^\d+\.\d+\.\d+\.\d+$') {
            $script:dotaIP = $null
            $script:dotaPop = $null
        }

        # Refresh detection every 10s so it can follow relay changes between matches.
        if (((Get-Date) - $script:dotaLastDetect).TotalSeconds -ge 10 -or $null -eq $script:dotaIP) {
            $detected = Get-DotaServerIP
            if ($null -ne $detected) {
                $script:dotaIP = $detected
                $script:dotaMisses = 0
            } else {
                $script:dotaMisses++
                # Two consecutive misses ~= 20s => clear stale relay after game ends.
                if ($script:dotaMisses -ge 2) {
                    $script:dotaIP = $null
                    $script:dotaPop = $null
                    $script:dotaMisses = 0
                }
            }
            $script:dotaLastDetect = Get-Date
        }

        $script:CurrentSpikeMs = Get-AutoSpikeMs -Pop $script:dotaPop -FallbackMs $SpikeMs

        if ($script:dotaState -ne $prevState) {
            Add-TimelineEvent "State" ("{0} -> {1}" -f $prevState, $script:dotaState)
            $prevState = $script:dotaState
        }

        $relayKey = if ($script:dotaIP) { "$($script:dotaPop)|$($script:dotaIP)" } else { "" }
        if ($relayKey -ne $prevRelayKey) {
            if ($script:dotaIP) {
                $relayDetail = if ($script:dotaPop) {
                    "{0} {1} ({2})" -f $script:dotaPop, $script:dotaIP, $script:dotaRegion
                } else {
                    "{0}" -f $script:dotaIP
                }
                Add-TimelineEvent "Relay" $relayDetail
            } else {
                Add-TimelineEvent "Relay" "cleared"
            }
            $prevRelayKey = $relayKey
        }

        if ($script:CurrentSpikeMs -ne $script:lastAutoSpikeMs) {
            Add-TimelineEvent "Threshold" ("{0}ms -> {1}ms" -f $script:lastAutoSpikeMs, $script:CurrentSpikeMs)
            $script:lastAutoSpikeMs = $script:CurrentSpikeMs
        }

        $msD = $null
        if ($null -ne $script:dotaIP) {
            $msD = Ping-Once $script:dotaIP
            $rD.Add([PSCustomObject]@{ Time = $ts; Ms = $msD })
        }

        if ($script:UsePinnedConsole -and $pingRow -ge 0) {
            try {
                [Console]::SetCursorPosition(0, $pingRow)
                Write-PingPinned $ts $Ref1 $ms1 $script:CurrentSpikeMs
                Write-PingPinned $ts $Ref2 $ms2 $script:CurrentSpikeMs
                if ($null -eq $script:dotaIP) {
                    Write-PinnedLine "  [$ts] Dota server           -- detecting... queue a match" "Gray"
                } else {
                    $relayLabel = if ($script:dotaPop) { "ValveRelay:$($script:dotaPop) $($script:dotaIP)" } else { "ValveRelay:$($script:dotaIP)" }
                    if ($null -ne $script:dotaSdrTotalMs) {
                        $relayLabel = "{0} SDR:{1}+{2}={3}" -f $relayLabel, $script:dotaSdrFrontMs, $script:dotaSdrBackMs, $script:dotaSdrTotalMs
                    }
                    Write-PingPinned $ts $relayLabel $msD $script:CurrentSpikeMs
                }
            } catch {
                $script:UsePinnedConsole = $false
            }
        }

        if (-not $script:UsePinnedConsole) {
            Write-PingPinned $ts $Ref1 $ms1 $script:CurrentSpikeMs
            Write-PingPinned $ts $Ref2 $ms2 $script:CurrentSpikeMs
            if ($null -eq $script:dotaIP) {
                Write-PinnedLine "  [$ts] Dota server           -- detecting... queue a match" "Gray"
            } else {
                $relayLabel = if ($script:dotaPop) { "ValveRelay:$($script:dotaPop) $($script:dotaIP)" } else { "ValveRelay:$($script:dotaIP)" }
                if ($null -ne $script:dotaSdrTotalMs) {
                    $relayLabel = "{0} SDR:{1}+{2}={3}" -f $relayLabel, $script:dotaSdrFrontMs, $script:dotaSdrBackMs, $script:dotaSdrTotalMs
                }
                Write-PingPinned $ts $relayLabel $msD $script:CurrentSpikeMs
            }
        }

        Start-Sleep -Milliseconds $IntervalMs
    }
}
finally {
    Write-Host ""

    Stop-DeepCapture

    $endTime = Get-Date
    $duration = $endTime - $startTime

    $finalSpikeMs = $script:CurrentSpikeMs
    $s1 = Get-Stats $r1 $finalSpikeMs
    $s2 = Get-Stats $r2 $finalSpikeMs
    $sD = if ($rD.Count -gt 0) { Get-Stats $rD $finalSpikeMs } else { $null }
    $grade = Get-SessionGrade -DotaStats $sD -ThresholdMs $finalSpikeMs
    Invoke-NetworkAutoFix -Ref1Stats $s1 -Ref2Stats $s2 -DotaStats $sD -ThresholdMs $finalSpikeMs

    $relayLabel = if ($script:dotaIP) {
        if ($script:dotaPop) { "$($script:dotaPop) $($script:dotaIP)" } else { $script:dotaIP }
    } else {
        "not detected"
    }

    Write-Host "  ==================== SUMMARY ====================" -ForegroundColor Cyan
    Write-Host ("  Duration     : {0:hh\:mm\:ss}" -f $duration) -ForegroundColor White
    Write-Host ("  Match State  : {0}" -f $script:dotaState) -ForegroundColor White
    Write-Host ("  Relay        : {0}" -f $relayLabel) -ForegroundColor White
    Write-Host ("  Relay Region : {0}" -f $script:dotaRegion) -ForegroundColor White
    Write-Host ("  Spike Rule   : >= {0}ms (base {1}ms)" -f $finalSpikeMs, $SpikeMs) -ForegroundColor White
    Write-Host ("  AutoFix      : {0} ({1})" -f $script:AutoFixStatus, $script:AutoFixReason) -ForegroundColor White
    Write-Host ("  Output Mode  : {0}" -f $(if ($KeepOutput) { "keep temp files" } else { "auto-clean temp files" })) -ForegroundColor White
    if ($DeepCapture) {
        Write-Host ("  DeepCap      : {0}" -f $(if ($script:DeepCaptureEtlPath) { "ON" } else { "FAILED" })) -ForegroundColor White
        if ($script:DeepCaptureEtlPath) {
            Write-Host ("  DeepCap ETL  : {0}" -f $script:DeepCaptureEtlPath) -ForegroundColor White
        }
        if ($script:DeepCapturePcapPath) {
            Write-Host ("  DeepCap PCAP : {0}" -f $script:DeepCapturePcapPath) -ForegroundColor White
        } else {
            Write-Host "  DeepCap PCAP : not exported (etl2pcapng not available or conversion failed)" -ForegroundColor DarkYellow
        }
        if ($script:DeepCaptureError) {
            Write-Host ("  DeepCap Note : {0}" -f $script:DeepCaptureError) -ForegroundColor DarkYellow
        }
    }
    if ($null -ne $script:dotaSdrTotalMs) {
        Write-Host ("  SDR Route    : front {0}ms + back {1}ms = {2}ms" -f $script:dotaSdrFrontMs, $script:dotaSdrBackMs, $script:dotaSdrTotalMs) -ForegroundColor White
    }
    Write-Host ""

    foreach ($t in @(
        @{ Label = $Ref1; S = $s1 },
        @{ Label = $Ref2; S = $s2 },
        @{ Label = "$(if ($script:dotaPop) { "Dota:$($script:dotaPop) $($script:dotaIP)" } else { "Dota:$($script:dotaIP)" })"; S = $sD }
    )) {
        if ($null -eq $t.S) { continue }
        $s = $t.S
        Write-Host ("  [ {0} ]" -f $t.Label) -ForegroundColor Yellow
        Write-Host ("    Loss     : {0}/{1} ({2}%)" -f $s.Lost, $s.Total, $s.LossPct)
        Write-Host ("    Avg/Min/Max : {0} / {1} / {2}" -f (Format-MsValue $s.AvgMs), (Format-MsValue $s.MinMs), (Format-MsValue $s.MaxMs))
        Write-Host ("    Jitter/P95/P99 : {0} / {1} / {2}" -f (Format-MsValue $s.JitterMs), (Format-MsValue $s.P95Ms), (Format-MsValue $s.P99Ms))
        Write-Host ("    Spikes   : {0} (>= {1}ms)   Bursts: {2}   Longest: {3}" -f $s.Spikes, $finalSpikeMs, $s.BurstCount, $s.LongestBurst)
        Write-Host ""
    }

    $gradeColor = switch ($grade.Grade) {
        "S" { "Green" }
        "A" { "Green" }
        "B" { "Yellow" }
        "C" { "Yellow" }
        "D" { "DarkYellow" }
        "F" { "Red" }
        default { "Gray" }
    }
    Write-Host ("  Session Grade: {0} ({1}/100)  {2}" -f $grade.Grade, $grade.Score, $grade.Reason) -ForegroundColor $gradeColor
    Write-Host ""

    if ($script:AutoFixActions.Count -gt 0) {
        Write-Host "  AutoFix Actions:" -ForegroundColor Cyan
        foreach ($a in $script:AutoFixActions) {
            Write-Host ("    - {0}" -f $a) -ForegroundColor Gray
        }
        Write-Host ""
    }

    if ($script:eventTimeline.Count -gt 0) {
        Write-Host "  Recent Events:" -ForegroundColor Cyan
        foreach ($ev in ($script:eventTimeline | Select-Object -Last 8)) {
            Write-Host ("    [{0}] {1}: {2}" -f $ev.Time, $ev.Type, $ev.Detail) -ForegroundColor Gray
        }
        Write-Host ""
    }

    $historyPath = Join-Path $script:OutputRoot "ushie_net_history.csv"
    if ($KeepOutput) {
        Save-SessionHistory -Path $historyPath -StartTime $startTime -Duration $duration -RelayPop $script:dotaPop -RelayIP $script:dotaIP -RelayRegion $script:dotaRegion -DotaStats $sD -ThresholdMs $finalSpikeMs -GradeObj $grade
    }

    $logPath = Join-Path $script:SessionDir "netlog.txt"
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("USHIE NETWORK MONITOR - LOG")
    $lines.Add("============================")
    $lines.Add("Started  : $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))")
    $lines.Add("Ended    : $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))")
    $lines.Add("Duration : $($duration.ToString('hh\:mm\:ss'))")
    $lines.Add("Ref1     : $Ref1")
    $lines.Add("Ref2     : $Ref2")
    $lines.Add("State    : $($script:dotaState)")
    $lines.Add("Dota     : $relayLabel")
    $lines.Add("Region   : $($script:dotaRegion)")
    $lines.Add("AutoFix  : $($script:AutoFixStatus) ($($script:AutoFixReason))")
    if ($DeepCapture) {
        $lines.Add("DeepCap  : $(if ($script:DeepCaptureEtlPath) { 'ON' } else { 'FAILED' })")
        if ($script:DeepCaptureEtlPath) { $lines.Add("DeepCapETL  : $($script:DeepCaptureEtlPath)") }
        if ($script:DeepCapturePcapPath) { $lines.Add("DeepCapPCAP : $($script:DeepCapturePcapPath)") }
        if ($script:DeepCaptureError) { $lines.Add("DeepCapNote : $($script:DeepCaptureError)") }
    }
    if ($null -ne $script:dotaSdrTotalMs) {
        $lines.Add("SDR      : front $($script:dotaSdrFrontMs) + back $($script:dotaSdrBackMs) = total $($script:dotaSdrTotalMs)")
    }
    $lines.Add("Spike threshold (final): ${finalSpikeMs}ms (base ${SpikeMs}ms)")
    $lines.Add("Session Grade: $($grade.Grade) ($($grade.Score)/100) $($grade.Reason)")
    $lines.Add("")
    if ($script:AutoFixActions.Count -gt 0) {
        $lines.Add("--- AUTOFIX ACTIONS ---")
        foreach ($a in $script:AutoFixActions) {
            $lines.Add("- $a")
        }
        $lines.Add("")
    }
    $lines.Add("--- SUMMARY ---")

    foreach ($t in @(
        @{ Label = $Ref1; S = $s1 },
        @{ Label = $Ref2; S = $s2 },
        @{ Label = "$(if ($script:dotaPop) { "Dota:$($script:dotaPop) $($script:dotaIP)" } else { "Dota:$($script:dotaIP)" })"; S = $sD }
    )) {
        if ($null -eq $t.S) { continue }
        $s = $t.S
        $lines.Add("[$($t.Label)]")
        $lines.Add("  Loss     : $($s.Lost)/$($s.Total) ($($s.LossPct)%)")
        $lines.Add("  Avg/Min/Max : $(Format-MsValue $s.AvgMs) / $(Format-MsValue $s.MinMs) / $(Format-MsValue $s.MaxMs)")
        $lines.Add("  Jitter/P95/P99 : $(Format-MsValue $s.JitterMs) / $(Format-MsValue $s.P95Ms) / $(Format-MsValue $s.P99Ms)")
        $lines.Add("  Spikes   : $($s.Spikes) (>= ${finalSpikeMs}ms)  Bursts: $($s.BurstCount)  Longest: $($s.LongestBurst)")
        $lines.Add("")
    }

    if ($script:eventTimeline.Count -gt 0) {
        $lines.Add("--- TIMELINE ---")
        foreach ($ev in ($script:eventTimeline | Select-Object -Last 20)) {
            $lines.Add("[$($ev.Time)] $($ev.Type): $($ev.Detail)")
        }
        $lines.Add("")
    }

    if ($KeepRawLog) {
        $lines.Add("--- RAW PINGS $Ref1 ---")
        $lines.Add(($r1 | Format-Table -AutoSize | Out-String))
        $lines.Add("--- RAW PINGS $Ref2 ---")
        $lines.Add(($r2 | Format-Table -AutoSize | Out-String))
        if ($rD.Count -gt 0) {
            $lines.Add("--- RAW PINGS Dota:$($script:dotaIP) ---")
            $lines.Add(($rD | Format-Table -AutoSize | Out-String))
        }
    } else {
        $lines.Add("Raw ping list omitted (use -KeepRawLog to include).")
    }

    $lines | Set-Content -Path $logPath -Encoding UTF8

    if ($KeepOutput) {
        Write-Host "  Log saved: $logPath" -ForegroundColor Cyan
        Write-Host "  History saved: $historyPath" -ForegroundColor Cyan
        if ($DeepCapture -and $script:DeepCaptureEtlPath) {
            Write-Host "  DeepCap ETL saved: $($script:DeepCaptureEtlPath)" -ForegroundColor Cyan
            if ($script:DeepCapturePcapPath) {
                Write-Host "  DeepCap PCAP saved: $($script:DeepCapturePcapPath)" -ForegroundColor Cyan
            }
        }
        Write-Host "  Output folder: $($script:SessionDir)" -ForegroundColor Cyan
    } else {
        try {
            Remove-Item $script:SessionDir -Recurse -Force -ErrorAction SilentlyContinue
        } catch {}
        Write-Host "  Temp output cleaned (use -KeepOutput to retain logs/capture files)." -ForegroundColor DarkGray
    }
    Write-Host ""
}


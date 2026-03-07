param(
    [string]$Ref1    = "1.1.1.1",
    [string]$Ref2    = "8.8.8.8",
    [int]$IntervalMs = 1000,
    [int]$SpikeMs    = 50,
    [switch]$KeepRawLog
)
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host ""
    Write-Host "  ERROR: Run this script as Administrator." -ForegroundColor Red
    Write-Host "  Right-click PowerShell -> Run as Administrator, then run again." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

$r1 = [System.Collections.Generic.List[PSCustomObject]]::new()
$r2 = [System.Collections.Generic.List[PSCustomObject]]::new()
$rD = [System.Collections.Generic.List[PSCustomObject]]::new()
$startTime = Get-Date
$script:dotaIP = $null
$script:dotaPop = $null
$script:dotaLastDetect = [datetime]::MinValue
$script:dotaMisses = 0
$script:dotaConsoleLogPath = $null
$script:PrivateIPv4Pattern = '^(127\.|0\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[01])\.|169\.254\.)'

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

function Get-DotaServerIPFromConsoleLog {
    $proc = Get-DotaProcess
    if ($null -eq $proc) { return $null }

    $logPath = Get-DotaConsoleLogPath
    if (-not $logPath) { return $null }

    try {
        $meta = Get-Item -Path $logPath -ErrorAction SilentlyContinue
        if ($null -eq $meta) { return $null }

        $lines = Get-Content -Path $logPath -Tail 12000 -ErrorAction SilentlyContinue
        if (-not $lines) { return $null }

        # If latest UI state is dashboard, do not keep stale relay from previous matches.
        for ($i = $lines.Count - 1; $i -ge 0; $i--) {
            $line = $lines[$i]
            if ($line -match 'ChangeGameUIState:\s+\S+\s+->\s+(\S+)') {
                $uiDest = $matches[1]
                if ($uiDest -eq 'DOTA_GAME_UI_STATE_DASHBOARD') { return $null }
                break
            }
        }

        for ($i = $lines.Count - 1; $i -ge 0; $i--) {
            $line = $lines[$i]
            if ($line -match 'SteamNetSockets.*Switched primary to\s+(\S+)\s+\((\d{1,3}(?:\.\d{1,3}){3}):\d+\)') {
                $pop = $matches[1]
                $ip = $matches[2]
                if ($ip -notmatch $script:PrivateIPv4Pattern) {
                    return [PSCustomObject]@{ IP = $ip; Pop = $pop; Source = "ConsoleLog" }
                }
            }
            if ($line -match 'SteamNetSockets.*Selecting\s+(\S+)\s+\((\d{1,3}(?:\.\d{1,3}){3}):\d+\)\s+as\s+primary') {
                $pop = $matches[1]
                $ip = $matches[2]
                if ($ip -notmatch $script:PrivateIPv4Pattern) {
                    return [PSCustomObject]@{ IP = $ip; Pop = $pop; Source = "ConsoleLog" }
                }
            }
            if ($line -match 'SteamNetSockets.*Requesting\s+session\s+from\s+(\S+)\s+\((\d{1,3}(?:\.\d{1,3}){3}):\d+\).*Rank=1') {
                $pop = $matches[1]
                $ip = $matches[2]
                if ($ip -notmatch $script:PrivateIPv4Pattern) {
                    return [PSCustomObject]@{ IP = $ip; Pop = $pop; Source = "ConsoleLog" }
                }
            }
            if ($line -match '\[Networking\]\s+Primary router:\s+(\S+)\s+\((\d{1,3}(?:\.\d{1,3}){3}):\d+\)') {
                $pop = $matches[1]
                $ip = $matches[2]
                if ($ip -notmatch $script:PrivateIPv4Pattern) {
                    return [PSCustomObject]@{ IP = $ip; Pop = $pop; Source = "ConsoleLog" }
                }
            }
        }
    } catch {}

    return $null
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

    # Most reliable for Source 2 SDR: parse latest relay from Dota console.log.
    $fromLog = Get-DotaServerIPFromConsoleLog
    if ($null -ne $fromLog) {
        $script:dotaPop = $fromLog.Pop
        return $fromLog.IP
    }

    if ($localPorts.Count -eq 0) { return $null }

    # Fast path: connected UDP endpoints visible in netstat.
    $remote = Get-DotaServerIPFromNetstat -LocalPorts $localPorts
    if ($null -ne $remote) {
        $script:dotaPop = $null
        return $remote
    }

    # Fallback: raw sniff on active interfaces.
    $raw = Get-DotaServerIPFromRawSniff -LocalPorts $localPorts
    if ($null -ne $raw) { $script:dotaPop = $null }
    return $raw
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

function Get-Stats([System.Collections.Generic.List[PSCustomObject]]$data) {
    $total = $data.Count
    if ($total -eq 0) { return $null }

    $lost = ($data | Where-Object { $null -eq $_.Ms }).Count
    $ok = @($data | Where-Object { $null -ne $_.Ms })

    $avg = $min = $max = $jitter = "n/a"
    $spikes = 0

    if ($ok.Count -gt 0) {
        $vals = @($ok | ForEach-Object { $_.Ms })
        $avg = "$([math]::Round(($vals | Measure-Object -Average).Average, 1))ms"
        $min = "$(($vals | Measure-Object -Minimum).Minimum)ms"
        $max = "$(($vals | Measure-Object -Maximum).Maximum)ms"
        $spikes = ($vals | Where-Object { $_ -ge $SpikeMs }).Count

        if ($vals.Count -gt 1) {
            $diffs = for ($i = 1; $i -lt $vals.Count; $i++) {
                [math]::Abs($vals[$i] - $vals[$i - 1])
            }
            $jitter = "$([math]::Round(($diffs | Measure-Object -Average).Average, 1))ms"
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
        Spikes = $spikes
    }
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
Write-Host "  Interval : ${IntervalMs}ms   Spike: >=${SpikeMs}ms" -ForegroundColor Gray
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

        $msD = $null
        if ($null -ne $script:dotaIP) {
            $msD = Ping-Once $script:dotaIP
            $rD.Add([PSCustomObject]@{ Time = $ts; Ms = $msD })
        }

        if ($script:UsePinnedConsole -and $pingRow -ge 0) {
            try {
                [Console]::SetCursorPosition(0, $pingRow)
                Write-PingPinned $ts $Ref1 $ms1 $SpikeMs
                Write-PingPinned $ts $Ref2 $ms2 $SpikeMs
                if ($null -eq $script:dotaIP) {
                    Write-PinnedLine "  [$ts] Dota server           -- detecting... queue a match" "Gray"
                } else {
                    $relayLabel = if ($script:dotaPop) { "ValveRelay:$($script:dotaPop) $($script:dotaIP)" } else { "ValveRelay:$($script:dotaIP)" }
                    Write-PingPinned $ts $relayLabel $msD $SpikeMs
                }
            } catch {
                $script:UsePinnedConsole = $false
            }
        }

        if (-not $script:UsePinnedConsole) {
            Write-PingPinned $ts $Ref1 $ms1 $SpikeMs
            Write-PingPinned $ts $Ref2 $ms2 $SpikeMs
            if ($null -eq $script:dotaIP) {
                Write-PinnedLine "  [$ts] Dota server           -- detecting... queue a match" "Gray"
            } else {
                $relayLabel = if ($script:dotaPop) { "ValveRelay:$($script:dotaPop) $($script:dotaIP)" } else { "ValveRelay:$($script:dotaIP)" }
                Write-PingPinned $ts $relayLabel $msD $SpikeMs
            }
        }

        Start-Sleep -Milliseconds $IntervalMs
    }
}
finally {
    Write-Host ""

    $endTime = Get-Date
    $duration = $endTime - $startTime

    $s1 = Get-Stats $r1
    $s2 = Get-Stats $r2
    $sD = if ($rD.Count -gt 0) { Get-Stats $rD } else { $null }

    Write-Host "  ==================== SUMMARY ====================" -ForegroundColor Cyan
    Write-Host ("  Duration : {0:hh\:mm\:ss}" -f $duration) -ForegroundColor White
    Write-Host ""

    foreach ($t in @(
        @{ Label = $Ref1; S = $s1 },
        @{ Label = $Ref2; S = $s2 },
        @{ Label = "$(if ($script:dotaPop) { "Dota:$($script:dotaPop) $($script:dotaIP)" } else { "Dota:$($script:dotaIP)" })"; S = $sD }
    )) {
        if ($null -eq $t.S) { continue }
        $s = $t.S
        Write-Host ("  [ {0} ]" -f $t.Label) -ForegroundColor Yellow
        Write-Host ("    Loss   : {0}/{1} ({2}%)" -f $s.Lost, $s.Total, $s.LossPct)
        Write-Host ("    Avg    : {0}   Min: {1}   Max: {2}" -f $s.Avg, $s.Min, $s.Max)
        Write-Host ("    Jitter : {0}" -f $s.Jitter)
        Write-Host ("    Spikes : {0} (>= ${SpikeMs}ms)" -f $s.Spikes)
        Write-Host ""
    }

    $logPath = "$env:USERPROFILE\Desktop\netlog_$($startTime.ToString('yyyyMMdd_HHmmss')).txt"
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("USHIE NETWORK MONITOR - LOG")
    $lines.Add("============================")
    $lines.Add("Started  : $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))")
    $lines.Add("Ended    : $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))")
    $lines.Add("Duration : $($duration.ToString('hh\:mm\:ss'))")
    $lines.Add("Ref1     : $Ref1")
    $lines.Add("Ref2     : $Ref2")
    $lines.Add("Dota     : $(if ($script:dotaIP) { if ($script:dotaPop) { "$($script:dotaPop) $($script:dotaIP)" } else { $script:dotaIP } } else { 'not detected' })")
    $lines.Add("Spike threshold: ${SpikeMs}ms")
    $lines.Add("")
    $lines.Add("--- SUMMARY ---")

    foreach ($t in @(
        @{ Label = $Ref1; S = $s1 },
        @{ Label = $Ref2; S = $s2 },
        @{ Label = "$(if ($script:dotaPop) { "Dota:$($script:dotaPop) $($script:dotaIP)" } else { "Dota:$($script:dotaIP)" })"; S = $sD }
    )) {
        if ($null -eq $t.S) { continue }
        $s = $t.S
        $lines.Add("[$($t.Label)]")
        $lines.Add("  Loss   : $($s.Lost)/$($s.Total) ($($s.LossPct)%)")
        $lines.Add("  Avg    : $($s.Avg)   Min: $($s.Min)   Max: $($s.Max)")
        $lines.Add("  Jitter : $($s.Jitter)")
        $lines.Add("  Spikes : $($s.Spikes) (>= ${SpikeMs}ms)")
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
    Write-Host "  Log saved: $logPath" -ForegroundColor Cyan
    Write-Host ""
}


# ushie


## Files
- `scripts/Run-AllInOne.ps1` - single all-in-one script (apply + verify)
- `scripts/Watch-Network.ps1` - live Dota latency monitor (auto relay detect + session grading)

## Run (Administrator PowerShell)
```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
.\scripts\Run-AllInOne.ps1 -m Safe
```

`Run-AllInOne.ps1` now auto-runs verify in the same execution by default.

Useful switches:
- `-m Safe|Extreme` - profile mode (`Safe` default)
- `-Dns Auto|Cloudflare|Google|Quad9|OpenDNS|AdGuard|ControlD|DNSSB|Comodo` - DNS mode (`Auto` default)
- `-DnsServers <ipv4,ipv4,...>` - manual DNS override (highest priority)
- `-v` - full-view output (no pinned header)
- `-KeepWSL` - skip disabling WSL/VirtualMachinePlatform
- `-SkipVerify` - apply only, no auto-verify
- `-VerifyOnly` - run built-in verification only
- `-NoRestore` - skip restore-point creation in `Extreme`
- `-h` / `-help` / `-man` - show manual usage

## One-Run From GitHub
Default one-liner (Safe mode):
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([ScriptBlock]::Create((irm 'https://raw.githubusercontent.com/4stropotato/ushie/main/scripts/Run-AllInOne.ps1'))) -m Safe"
```

One-liner with mode/verbosity:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([ScriptBlock]::Create((irm 'https://raw.githubusercontent.com/4stropotato/ushie/main/scripts/Run-AllInOne.ps1'))) -m Extreme -v"
```

One-liner with DNS override:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([ScriptBlock]::Create((irm 'https://raw.githubusercontent.com/4stropotato/ushie/main/scripts/Run-AllInOne.ps1'))) -m Safe -Dns Auto"
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([ScriptBlock]::Create((irm 'https://raw.githubusercontent.com/4stropotato/ushie/main/scripts/Run-AllInOne.ps1'))) -m Safe -Dns Google"
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([ScriptBlock]::Create((irm 'https://raw.githubusercontent.com/4stropotato/ushie/main/scripts/Run-AllInOne.ps1'))) -m Safe -Dns 'Cloudflare,Google'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([ScriptBlock]::Create((irm 'https://raw.githubusercontent.com/4stropotato/ushie/main/scripts/Run-AllInOne.ps1'))) -m Safe -DnsServers 1.1.1.1,8.8.8.8,9.9.9.9"
```

Verify only:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([ScriptBlock]::Create((irm 'https://raw.githubusercontent.com/4stropotato/ushie/main/scripts/Run-AllInOne.ps1'))) -VerifyOnly -v"
```

Manual / Help:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([ScriptBlock]::Create((irm 'https://raw.githubusercontent.com/4stropotato/ushie/main/scripts/Run-AllInOne.ps1'))) -h"
```

## Network Watch (Dota)
Run locally (Admin PowerShell):
```powershell
.\scripts\Watch-Network.ps1
```

Wireshark-like deep capture mode (ETL + optional PCAP export):
```powershell
.\scripts\Watch-Network.ps1 -DeepCapture
```

Test + auto-remediate (post-session, default behavior):
```powershell
.\scripts\Watch-Network.ps1
```

Notes:
- `-DeepCapture` uses built-in Windows trace capture (`netsh trace`) and auto-stops when you stop the watcher (`Ctrl+C`).
- If `etl2pcapng.exe` or `pktmon.exe` conversion succeeds, a `.pcapng` is also saved.
- Default output is temporary (`%TEMP%\ushie\run_*`) and auto-cleaned after run.
- Add `-KeepOutput` if you want to keep logs/history/capture files.
- AutoFix is enabled by default and runs targeted remediation after analysis (DNS auto-select + DNS flush + TCP sanity defaults).
- Use `-NoAutoFix` if you want monitor-only mode.
- Startup auto-cleans stale trace state (`netsh trace stop`, `pktmon stop/unload`) and old temp run folders (when not using `-KeepOutput`).

## What It Changes
- Removes `OverlayTestMode` from DWM.
- Resets IPv6 `DisabledComponents` to `0`.
- Sets DNS on active adapters using auto benchmark (default) or manual override.
  - Auto mode benchmarks public resolvers plus current adapter DNS and picks the fastest score.
- Restores keyboard responsiveness values.
- Re-enables HPET device (`ACPI\PNP0103\0`) if disabled.
- Enables HAGS (`HwSchMode=2`) for supported GPUs.
- Enables Game Mode and disables Game DVR capture.
- Forces Dota 2 (`dota2.exe`) to High Performance GPU preference.
- Sets power plan to Ultimate Performance (fallback: High Performance).
- Forces processor min/max state to `100%` (AC/DC) on active plan.
- Disables hypervisor launch + VBS flags for low-latency gaming.
- Disables WSL/VirtualMachinePlatform (optional skip switch).
- Restores key servicing/update service startup modes.
- Applies key debloat/privacy toggles (Copilot/activity/background-app related).
- Applies profile-based shell tuning:
  - `Safe`: restores default visual/shell responsiveness values
  - `Extreme`: aggressive visual/shell performance values (display/perf focused toggles)
  - `Extreme`: removes menu/minimize/window animations (`MinAnimate`, `DWM Animations`, taskbar animations) while keeping ClearType font smoothing enabled
- `Extreme` creates a System Restore Point (`Before-Ushie-Extreme`) before applying tweaks.
- `Extreme` adds system-wide responsiveness tuning:
  - lower startup shell delay
  - disable transparency effects
  - reduce background service load (`SysMain`, `WSearch`, telemetry-related services)
  - low-latency scheduler/network stack tuning (`MMCSS`, `Win32PrioritySeparation`, TCP globals)
  - interface-level Nagle delay reduction on active adapters
  - aggressive CPU power behavior and memory policy (`MemoryCompression`, `PageCombining`, `DisablePagingExecutive`, prefetch policy)
  - boot timer policy tuning (`useplatformclock`, `disabledynamictick`, `tscsyncpolicy`)
- Writes profile marker at `HKCU:\Software\ushie\WinLagFix` for verify tracking.
- Writes DNS marker at `HKCU:\Software\ushie\WinLagFix\DnsApplied` for verify tracking.
- Cleans temp + cache paths (temp, shader caches, browser/system cache paths).
- `Extreme` also runs deeper cleanup (`SoftwareDistribution` cache + component cleanup).

## Notes
- This is one-shot only (no persistent background task is created).
- Some changes require reboot.
- `Extreme` can reduce indexing/background features to prioritize responsiveness.
- If lag remains even after this and clean driver install, use a clean official non-debloated Windows image.

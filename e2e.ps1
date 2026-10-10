# JinJuApp end-to-end smoke test (pure PowerShell, zero dependencies).
# Programmatic verification only — no visual judgment.
# Exit 0 = all checks green.

$ErrorActionPreference = 'Stop'
$repo = $PSScriptRoot
$exe = Join-Path $repo 'target\release\JinJuApp.exe'
$failures = 0
$checks = 0

function Check([string]$name, [scriptblock]$assert) {
    $script:checks++
    $ok = $false
    try { $ok = & $assert } catch { $ok = $false; Write-Output "    exception: $($_.Exception.Message)" }
    if ($ok) { Write-Output "PASS : $name" } else { Write-Output "FAIL : $name"; $script:failures++ }
}

Add-Type -Namespace Native -Name User32 -MemberDefinition '
[DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindowW(string cls, string title);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
'

# --- setup -----------------------------------------------------------------
$dir = Join-Path $env:TEMP ("JinJuApp-e2e-" + [Guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $dir | Out-Null
Copy-Item $exe (Join-Path $dir 'JinJuApp.exe')
$runExe = Join-Path $dir 'JinJuApp.exe'

try {
    # 1. binary exists and is a PE executable
    Check 'release binary exists' { Test-Path $runExe }
    $bytes = [IO.File]::ReadAllBytes($runExe) | Select-Object -First 2
    Check 'binary is a PE (MZ header)' { ($bytes[0] -eq 0x4D) -and ($bytes[1] -eq 0x5A) }

    # 2. first run: exports quotes.json next to the exe
    $p = Start-Process -FilePath $runExe -WorkingDirectory $dir -PassThru
    Start-Sleep -Milliseconds 1500
    Check 'process is alive after 1.5 s' { -not $p.HasExited }
    $exported = Join-Path $dir 'quotes.json'
    Check 'quotes.json exported on first run' { Test-Path $exported }
    Check 'exported corpus matches repo corpus' {
        (Get-FileHash $exported).Hash -eq (Get-FileHash (Join-Path $repo 'quotes.json')).Hash
    }

    # 3. window exists and is visible
    # (title passed explicitly: PowerShell marshals $null as "" not as a
    # null pointer, which would fail the title comparison)
    $hwnd = [Native.User32]::FindWindowW('JinJuApp', 'JinJuApp')
    Check 'window class JinJuApp exists' { $hwnd -ne [IntPtr]::Zero }
    Check 'window is visible' { [Native.User32]::IsWindowVisible($hwnd) }

    # 4. autostart is opt-in: nothing written without user consent
    Check 'autostart NOT registered by default' {
        -not (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'JinJuApp' -ErrorAction SilentlyContinue)
    }

    # 5. single instance: a second launch defers and exits immediately
    $p2 = Start-Process -FilePath $runExe -WorkingDirectory $dir -PassThru
    Start-Sleep -Milliseconds 800
    Check 'second instance exits (mutex)' { $p2.HasExited -and ($p2.ExitCode -eq 0) }
    Check 'first instance still alive' { -not $p.HasExited }

    # 6. graceful shutdown via DestroyWindow (the tray exit path)
    $null = & taskkill /PID $p.Id /T 2>&1
    Start-Sleep -Milliseconds 800
    Check 'process terminated' { $p.HasExited }
}
finally {
    # cleanup: no orphan processes, no residue in the run key
    Get-Process -Name 'JinJuApp' -ErrorAction SilentlyContinue | Stop-Process -Force
    Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'JinJuApp' -ErrorAction SilentlyContinue
}

Write-Output ''
Write-Output "checks: $checks, failures: $failures"
if ($failures -gt 0) { exit 1 } else { exit 0 }

# JinJuApp end-to-end smoke test (pure PowerShell, zero dependencies).
# Programmatic verification only - no visual judgment.
# Pure ASCII on purpose: Windows PowerShell 5.1 misreads UTF-8 without BOM.
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
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int left, top, right, bottom; }
[StructLayout(LayoutKind.Sequential)] public struct APPBARDATA { public int cbSize; public IntPtr hWnd; public uint uCallbackMessage; public uint uEdge; public RECT rc; public IntPtr lParam; }
[DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindowW(string cls, string title);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT r);
[DllImport("shell32.dll")] public static extern IntPtr SHAppBarMessage(uint m, ref APPBARDATA d);
[DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
[DllImport("user32.dll")] public static extern IntPtr GetTopWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr hWnd, uint cmd);
[DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
'
Add-Type -AssemblyName System.Drawing

# Physical coordinates for all geometry queries: a DPI-unaware host would
# virtualize SHAppBarMessage/GetWindowRect on scaled displays and mix
# coordinate spaces, breaking the position assertions below.
if (-not [Native.User32]::SetProcessDpiAwarenessContext([IntPtr](-4))) {
    [Native.User32]::SetProcessDPIAware() | Out-Null
}

# True when $hwnd sits above $tray in the z-order (walk down from the top;
# whichever we meet first is higher). Both live in the topmost band.
function Test-Above([IntPtr]$hwnd, [IntPtr]$tray) {
    $cur = [Native.User32]::GetTopWindow([IntPtr]::Zero)
    $guard = 0
    while (($cur -ne [IntPtr]::Zero) -and ($guard -lt 2000)) {
        if ($cur -eq $hwnd) { return $true }
        if ($cur -eq $tray) { return $false }
        $cur = [Native.User32]::GetWindow($cur, 2) # GW_HWNDNEXT
        $guard++
    }
    return $false
}

# MD5 hex of the screen rectangle's pixels. Layered windows are composed
# by DWM, so CopyFromScreen sees the strip exactly as the user does.
function Capture-Hash([int]$x, [int]$y, [int]$w, [int]$h) {
    $bmp = New-Object System.Drawing.Bitmap $w, $h
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($x, $y, 0, 0, (New-Object System.Drawing.Size $w, $h))
    $g.Dispose()
    $rect = New-Object System.Drawing.Rectangle 0, 0, $w, $h
    $data = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $bytes = New-Object byte[] ($w * $h * 4)
    [Runtime.InteropServices.Marshal]::Copy($data.Scan0, $bytes, 0, $bytes.Length)
    $bmp.UnlockBits($data)
    $bmp.Dispose()
    $md5 = [System.Security.Cryptography.MD5]::Create()
    return ($md5.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) -join ''
}

# --- setup -----------------------------------------------------------------
# A running instance would win the mutex and our test instance would exit
# silently, so clear the field first (and again in the finally block).
Get-Process -Name 'JinJuApp' -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 300

# Snapshot the user's real autostart entry (if any): the app must not
# modify it without consent, and cleanup must restore it - never delete it.
$runKeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$runKeyBefore = (Get-ItemProperty -Path $runKeyPath -Name 'JinJuApp' -ErrorAction SilentlyContinue).JinJuApp

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

    # 3b. position: for a normally docked bottom taskbar the strip must
    # sit inside the taskbar band (screen corner anchor, v1 parity), not
    # floating above it - regression guard for the placement math.
    $wr = New-Object Native.User32+RECT
    [Native.User32]::GetWindowRect($hwnd, [ref]$wr) | Out-Null
    $abd = New-Object Native.User32+APPBARDATA
    $abd.cbSize = [Runtime.InteropServices.Marshal]::SizeOf([type][Native.User32+APPBARDATA])
    [Native.User32]::SHAppBarMessage(5, [ref]$abd) | Out-Null  # ABM_GETTASKBARPOS
    if (($abd.uEdge -eq 3) -and (($abd.rc.bottom - $abd.rc.top) -ge 40)) {
        Check 'strip sits inside the bottom taskbar band (v1 parity)' {
            ($wr.bottom -gt $abd.rc.top) -and ($wr.bottom -le $abd.rc.bottom)
        }
        Check 'strip anchors to the left edge of the screen' {
            ($wr.left -ge 0) -and ($wr.left -le 30)
        }
    } else {
        Write-Output "SKIP : taskbar not bottom-docked (edge=$($abd.uEdge)); position check not applicable"
    }

    # 4. autostart is opt-in: running the app must not create or modify
    # the Run entry without the explicit tray-menu toggle. Compare against
    # the snapshot so a developer's real entry makes this pass, not fail.
    Check 'autostart entry untouched by running the app' {
        $now = (Get-ItemProperty -Path $runKeyPath -Name 'JinJuApp' -ErrorAction SilentlyContinue).JinJuApp
        $now -eq $runKeyBefore
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

    # --- phase B: marquee + z-order regression guards ----------------------
    # A single ASCII quote far wider than the 470 px resting width forces
    # the scroll state machine (13 pt UI font puts this at ~900 px).
    $longQuote = 'The quick brown fox jumps over the lazy dog 0123456789 ' +
                 'and keeps running far beyond the edge of the strip ABCDEF'
    Set-Content -Path $exported -Encoding Ascii -Value ('[{"quote":"' + $longQuote + '"}]')

    $p = Start-Process -FilePath $runExe -WorkingDirectory $dir -PassThru
    Start-Sleep -Milliseconds 700
    $hwnd = [Native.User32]::FindWindowW('JinJuApp', 'JinJuApp')
    Check 'phase B window exists' { $hwnd -ne [IntPtr]::Zero }
    $wr = New-Object Native.User32+RECT
    [Native.User32]::GetWindowRect($hwnd, [ref]$wr) | Out-Null
    $cx = $wr.left; $cy = $wr.top
    $cw = $wr.right - $wr.left; $ch = $wr.bottom - $wr.top

    # 7. hold phase: a fresh long quote rests readable for ~2 s before it
    # slides. Two frames inside the hold must be pixel-identical.
    $holdA = Capture-Hash $cx $cy $cw $ch
    Start-Sleep -Milliseconds 700
    $holdB = Capture-Hash $cx $cy $cw $ch
    Check 'long quote rests readable during the hold phase' { $holdA -eq $holdB }

    # 8. marquee motion: after the hold the text must actually move -
    # regression guard for "long quotes are truncated to ~29 chars".
    # (45 logical px/s; even 1 frame interval changes the pixels.)
    Start-Sleep -Milliseconds 1600   # let the hold expire
    $frames = @()
    1..3 | ForEach-Object {
        $frames += Capture-Hash $cx $cy $cw $ch
        Start-Sleep -Milliseconds 800
    }
    Check 'marquee scrolls long quotes (pixels change over time)' {
        ($frames[0] -ne $frames[1]) -or ($frames[1] -ne $frames[2])
    }

    # 9. z-order recovery: when explorer raises the taskbar above the strip
    # (icon clicks, new buttons, flyouts), the 500 ms re-assert timer must
    # bring us back - regression guard for "text vanishes for seconds".
    $tray = [Native.User32]::FindWindowW('Shell_TrayWnd', [NullString]::Value)
    Check 'taskbar window found' { $tray -ne [IntPtr]::Zero }
    if ($tray -ne [IntPtr]::Zero) {
        # Push the tray to the top of the topmost band, covering the strip.
        # SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE = 0x13
        [Native.User32]::SetWindowPos($tray, [IntPtr](-1), 0, 0, 0, 0, 0x13) | Out-Null
        Start-Sleep -Milliseconds 120
        $recovered = $false
        1..15 | ForEach-Object {
            if (Test-Above $hwnd $tray) { $script:recovered = $true; return }
            Start-Sleep -Milliseconds 200
        }
        Check 'strip re-asserts above the taskbar within 3 s' { $recovered }
    }

    # final shutdown
    $null = & taskkill /PID $p.Id /T 2>&1
    Start-Sleep -Milliseconds 500
}
finally {
    # cleanup: no orphan processes; restore the user's autostart entry
    # exactly as it was (delete only if it did not exist beforehand)
    Get-Process -Name 'JinJuApp' -ErrorAction SilentlyContinue | Stop-Process -Force
    Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue
    if ($null -ne $runKeyBefore) {
        Set-ItemProperty -Path $runKeyPath -Name 'JinJuApp' -Value $runKeyBefore
    } else {
        Remove-ItemProperty -Path $runKeyPath -Name 'JinJuApp' -ErrorAction SilentlyContinue
    }
}

Write-Output ''
Write-Output "checks: $checks, failures: $failures"
if ($failures -gt 0) { exit 1 } else { exit 0 }

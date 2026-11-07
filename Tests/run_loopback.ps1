<#
.SYNOPSIS
Runs the real serial modules against a real or virtual COM port, using loopback_tests.bas.

.DESCRIPTION
Sends test data on the ports you name. Exits 1 if any result is FAIL or a compile fails, else 0 (2 for a usage
error or a refused port). Run with -DryRun first: it uses port 0, which both modules reject before any native
call, so every test ends as SKIP and no port is opened. See README.md in this folder.

.PARAMETER PortA
COM port number, for example 3 for COM3. Required unless -DryRun.

.PARAMETER PortB
Second port of a pair (com0com, null-modem cable). Defaults to PortA, which means loopback-plug mode with TX
wired to RX.

.PARAMETER Module
Main, NoDebug or Both (default). Each runs in its own throwaway workbook.

.PARAMETER RtsCtsLoop
Plug mode only: RTS is wired to CTS, so the flow-control test can run.

.PARAMETER OutFile
Also write the report to this file.

.PARAMETER DryRun
Uses port 0, skips the port check, runs everything. No port is opened.
#>
param(
    [int]$PortA = -1,
    [int]$PortB = -1,
    [ValidateSet('Main', 'NoDebug', 'Both')][string]$Module = 'Both',
    [switch]$RtsCtsLoop,
    [string]$OutFile,
    [switch]$DryRun
)
$ErrorActionPreference = 'Stop'
$T = $PSScriptRoot
$Repo = Split-Path $T -Parent
$report = New-Object System.Collections.Generic.List[string]
function Say($s) { $report.Add([string]$s); Write-Host $s }
function Finish($code) {
    if ($OutFile) { try { [IO.File]::WriteAllLines($OutFile, $report.ToArray()) } catch { Write-Host "could not write ${OutFile}: $_" } }
    exit $code
}

# ---- arguments
if ($DryRun) {
    if ($PortA -ge 0 -and $PortA -ne 0) { Say "-DryRun ignores -PortA $PortA and uses port 0." }
    $PortA = 0; $PortB = 0
} else {
    if ($PortA -lt 0) {
        Say "ERROR: -PortA is required unless -DryRun is given. Example: run_loopback.ps1 -PortA 3   (or -DryRun to check the setup first)"
        Finish 2
    }
    if ($PortB -lt 0) { $PortB = $PortA }
    foreach ($p in @($PortA, $PortB)) {
        if ($p -lt 1 -or $p -gt 255) { Say "ERROR: port number $p is outside 1 to 255."; Finish 2 }
    }
}
$Plug = ($PortA -eq $PortB)
if ($RtsCtsLoop -and -not $Plug) { Say "Note: -RtsCtsLoop applies to plug mode only and is ignored for a pair." }
$LoopFlag = [bool]($RtsCtsLoop -and $Plug)

Say "Loopback test: mode=$(if ($DryRun) { 'DRY RUN (port 0, nothing is opened)' } elseif ($Plug) { "plug on COM$PortA" + $(if ($LoopFlag) { ' with RTS wired to CTS' } else { '' }) } else { "pair COM$PortA <-> COM$PortB" }); module=$Module"

# ---- 1. registry check (read-only)
if (-not $DryRun) {
    $key = 'HKLM:\HARDWARE\DEVICEMAP\SERIALCOMM'
    $listed = @{}
    if (Test-Path $key) {
        $k = Get-Item $key
        foreach ($n in $k.GetValueNames()) { $listed[[string]$k.GetValue($n)] = $n }
    }
    $listing = if ($listed.Count) { ($listed.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)" }) -join ', ' } else { 'none' }
    $bad = @(@($PortA, $PortB) | Select-Object -Unique | Where-Object { -not $listed.ContainsKey("COM$_") })
    if ($bad.Count) {
        Say "ERROR: COM$($bad -join ', COM') not listed in $key. Ports present: $listing. Nothing was opened and Excel was not started."
        Finish 2
    }
    foreach ($p in @($PortA, $PortB) | Select-Object -Unique) { Say "COM$p device: $($listed["COM$p"])" }
}

# ---- 2. warning
Say "WARNING: test data (all byte values, text, parity-error bytes, flow-control stalls) will be sent on $(if ($DryRun) { 'no port (dry run)' } elseif ($Plug) { "COM$PortA" } else { "COM$PortA and COM$PortB" }). Nothing that reacts to arbitrary bytes may be connected to it."

# ---- 3. temp copies with the port range widened
$work = Join-Path ([IO.Path]::GetTempPath()) ('serial_vba_loopback_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $work | Out-Null
$latin1 = [Text.Encoding]::GetEncoding(28591)       # byte-for-byte, the modules are not UTF-8
$lo = if ($DryRun) { 1 } else { [Math]::Min($PortA, $PortB) }
$hi = if ($DryRun) { 2 } else { [Math]::Max($PortA, $PortB) }
$kinds = @()
if ($Module -in 'Main', 'Both') { $kinds += , @('Main', (Join-Path $Repo 'SERIAL_PORT_VBA.bas')) }
if ($Module -in 'NoDebug', 'Both') { $kinds += , @('NoDebug', (Join-Path $Repo 'No-Debug\SERIAL_PORT_VBA_NO_DEBUG.bas')) }
$copies = @{}
foreach ($kd in $kinds) {
    $name = $kd[0]; $text = [IO.File]::ReadAllText($kd[1], $latin1)
    foreach ($c in @(@('MIN', $lo), @('MAX', $hi))) {
        $re = [regex]"(?m)^(Private Const COM_PORT_$($c[0]) As Integer = )\d+"
        $n = $re.Matches($text).Count
        if ($n -ne 1) { Say "ERROR: COM_PORT_$($c[0]) line matched $n times in $($kd[1]), expected exactly 1."; Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue; Finish 2 }
        $text = $re.Replace($text, ('${1}' + $c[1]))
    }
    $dest = Join-Path $work "$name.bas"
    [IO.File]::WriteAllText($dest, $text, $latin1)
    $a = [IO.File]::ReadAllText($kd[1], $latin1) -split "`n"; $b = $text -split "`n"
    $diff = 0; for ($i = 0; $i -lt $a.Count; $i++) { if ($a[$i] -ne $b[$i]) { $diff++ } }
    $copies[$name] = $dest
    Say "[$name] temp copy: COM_PORT_MIN=$lo COM_PORT_MAX=$hi. This is the only change; $diff line(s) differ from the repo file, which is not edited."
}

# ---- 4-7. Excel
$results = New-Object System.Collections.Generic.List[string]
$compileFailed = $false
$tests = @(
    'LB_SAVE_SETTINGS', 'LB_SETTINGS_KEYVALUE', 'LB_SETTINGS_COMMA', 'LB_SETTINGS_BAUD_ONLY', 'LB_SETTINGS_BAD',
    'LB_BYTES_A_TO_B', 'LB_BYTES_B_TO_A', 'LB_TEXT', 'LB_MAXBYTES_BYTES', 'LB_MAXBYTES_TEXT',
    'LB_LATE_REPLY', 'LB_LATE_TIMING', 'LB_CHECK_COM_PORT', 'LB_LAST_ERROR', 'LB_FLOW_STOP', 'LB_PARITY'
)
$excel = $null; $watchdog = $null; $pidOut = [uint32]0
try {
    Add-Type -Namespace Native -Name Win32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
    $excel = New-Object -ComObject Excel.Application
    [void][Native.Win32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidOut)
    $watchdog = Start-Process pwsh -WindowStyle Hidden -PassThru -ArgumentList '-NoProfile', '-Command', "Start-Sleep 900; Stop-Process -Id $pidOut -Force"
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach ($kd in $kinds) {
        $name = $kd[0]
        $wb = $null
        try {
            $wb = $excel.Workbooks.Add()
            [void]$wb.VBProject.VBComponents.Import($copies[$name])
            [void]$wb.VBProject.VBComponents.Import((Join-Path $T 'loopback_tests.bas'))
            $excel.VBE.ActiveVBProject = $wb.VBProject
            $compile = $excel.VBE.CommandBars.FindControl(1, 578)
            $compile.Execute()
            if ($compile.Enabled) {
                $compileFailed = $true
                $results.Add("FAIL|$name/compile|loopback_tests.bas with the $name module did not compile (Compile button still enabled)")
                continue
            }
            Say "[$name] compiled clean with loopback_tests.bas"
            $wbName = $wb.Name
            $run = {
                param($fn, [object[]]$a)
                try { [string]([System.__ComObject].InvokeMember('Run', 'InvokeMethod', $null, $excel, (@("'$wbName'!$fn") + @($a)))) }
                catch { "FAIL|$fn|runner error: " + $_.Exception.Message.Split("`n")[0] }
            }
            $add = {
                param($line, $fn)
                $line = "$line".Trim()
                if ($line -notmatch '^(PASS|FAIL|INFO|SKIP)\|') { $line = "FAIL|$fn|unexpected result: $line" }
                $parts = $line -split '\|', 3
                $results.Add("$($parts[0])|$name/$($parts[1])|$($parts[2])")
            }
            try {
                foreach ($fn in $tests) {
                    $args2 = @($PortA, $PortB)
                    if ($fn -eq 'LB_FLOW_STOP') { $args2 += $LoopFlag }
                    & $add (& $run $fn $args2) $fn
                }
            } finally {
                # always, even after an error: stop anything left open, then put the settings back
                & $add (& $run 'LB_CLEANUP' @($PortA, $PortB)) 'LB_CLEANUP'
                & $add (& $run 'LB_RESTORE_SETTINGS' @($PortA, $PortB)) 'LB_RESTORE_SETTINGS'
            }
        } catch {
            $results.Add("FAIL|$name/runner|$($_.Exception.Message.Split("`n")[0])")
        } finally {
            if ($wb) { try { $wb.Close($false) } catch {} }
        }
    }
} catch {
    $results.Add("FAIL|runner|$($_.Exception.Message.Split("`n")[0])")
} finally {
    if ($excel) {
        try { $excel.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    if ($watchdog) { Stop-Process -Id $watchdog.Id -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Milliseconds 500
    if ($pidOut -ne 0 -and (Get-Process -Id $pidOut -ErrorAction SilentlyContinue)) { Stop-Process -Id $pidOut -Force }
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
}

# ---- 8. report
Say ''
Say 'Results'
foreach ($r in $results) { Say $r }
$count = @{ PASS = 0; FAIL = 0; INFO = 0; SKIP = 0 }
foreach ($r in $results) { $count[($r -split '\|', 2)[0]]++ }
Say ''
Say "Summary: $($count.PASS) PASS, $($count.FAIL) FAIL, $($count.INFO) INFO, $($count.SKIP) SKIP$(if ($compileFailed) { '; COMPILE FAILED' })"
if ($count.FAIL -gt 0 -or $compileFailed) { Say 'FAILED'; Finish 1 }
Say 'OK (no FAIL)'
Finish 0

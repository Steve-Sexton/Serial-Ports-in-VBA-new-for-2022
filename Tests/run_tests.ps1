<#
.SYNOPSIS
Runs the serial module tests in Excel and compares the results with expected_results.txt.

.DESCRIPTION
Exits 0 if every result matches, 1 if not. See README.md in this folder for what is real and what is stubbed.

.PARAMETER Baseline
Runs an older commit from git instead of the working tree and only prints the results, for example
-Baseline cbf8166 to see the original bugs. Use a commit from before cefd14c.

.PARAMETER UpdateExpected
Rewrites expected_results.txt from this run. Only use it after checking that the changed results are right.
#>
param(
    [string]$Baseline,
    [switch]$UpdateExpected
)
$ErrorActionPreference = 'Stop'
$T = $PSScriptRoot
$Repo = Split-Path $T -Parent
$Expected = Join-Path $T 'expected_results.txt'
$work = Join-Path ([IO.Path]::GetTempPath()) ('serial_vba_tests_' + [guid]::NewGuid().ToString('N'))
$data = Join-Path $work 'data'
New-Item -ItemType Directory -Force $data | Out-Null
$Which = if ($Baseline) { 'head' } else { 'new' }

$files = @{ M = 'SERIAL_PORT_VBA.bas'; ND = 'No-Debug/SERIAL_PORT_VBA_NO_DEBUG.bas' }
$src = @{}
foreach ($k in $files.Keys) {
    if ($Baseline) {
        $src[$k] = Join-Path $work "baseline_$k.bas"
        cmd /c "git -C `"$Repo`" show ${Baseline}:$($files[$k]) > `"$($src[$k])`""     # cmd keeps the bytes unchanged
        if ($LASTEXITCODE -ne 0) { throw "git show failed for ${Baseline}:$($files[$k])" }
    } else {
        $src[$k] = Join-Path $Repo $files[$k]
    }
}

try {
$results = & {
# data files
[IO.File]::WriteAllBytes("$data\d256.bin", [byte[]](0..255))
$p = New-Object byte[] 5000; for ($i = 0; $i -lt 5000; $i++) { $p[$i] = $i % 251 }
[IO.File]::WriteAllBytes("$data\d5000.bin", $p)
[IO.File]::WriteAllBytes("$data\d3.bin", [byte[]](65, 66, 67))
[IO.File]::WriteAllBytes("$data\d5.bin", [byte[]](1, 2, 3, 4, 5))
[IO.File]::WriteAllBytes("$data\start.bin", [byte[]](0))

Add-Type -Namespace Native -Name Win32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$excel = New-Object -ComObject Excel.Application
$pidOut = [uint32]0
[void][Native.Win32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidOut)
$watchdog = Start-Process pwsh -WindowStyle Hidden -PassThru -ArgumentList '-NoProfile', '-Command', "Start-Sleep 300; Stop-Process -Id $pidOut -Force"
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Hex-File($path) { if (Test-Path $path) { ([IO.File]::ReadAllBytes($path) | ForEach-Object { '{0:X2}' -f $_ }) -join '' } else { '<nofile>' } }

try {
foreach ($kind in @('M', 'ND')) {
    # the module as shipped, with the Ribbon module
    $wb = $excel.Workbooks.Add()
    [void]$wb.VBProject.VBComponents.Import($src[$kind])
    [void]$wb.VBProject.VBComponents.Import("$Repo\Ribbon\SERIAL_PORT_RIBBON.bas")
    $excel.VBE.ActiveVBProject = $wb.VBProject
    $compile = $excel.VBE.CommandBars.FindControl(1, 578)
    $compile.Execute()
    "[$Which/$kind] module + Ribbon compile enabled after compile (False=clean): $($compile.Enabled)"
    $wb.Close($false)

    # the test copy
    $copy = Join-Path $work ("test_${Which}_$kind.bas")
    python "$T\mk_test.py" $src[$kind] $copy $kind $Which "$T\tests.bas" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "mk_test.py failed for $kind" }
    $wb = $excel.Workbooks.Add()
    [void]$wb.VBProject.VBComponents.Import($copy)
    $excel.VBE.ActiveVBProject = $wb.VBProject
    $compile = $excel.VBE.CommandBars.FindControl(1, 578)
    $compile.Execute()
    "[$Which/$kind] compile enabled after compile (False=clean): $($compile.Enabled)"
    $name = $wb.Name
    function RunT($fn, [object[]]$a) {
        try { return [string]([System.__ComObject].InvokeMember('Run', 'InvokeMethod', $null, $excel, (@("'$name'!$fn") + @($a)))) } catch { return "ERR: " + $_.Exception.Message.Split("`n")[0] }
    }
    function Do-T($label, $fn, [object[]]$a) {
        $r = RunT $fn $a
        "[$Which/$kind] $label => $r"
    }
    $f = { param($n) "$data\$n" }

    # --- settings / cleaner
    foreach ($s in @('9600', '9600,N,8,1', 'baud=9600,parity=e', 'BAUD = 9600')) { Do-T "CLEAN '$s'" 'T_CLEAN' @($s) }
    Do-T 'SPC "9600" seed baud 1200' 'T_SPC' @('9600', 0x4800, 0, 0, 1200)
    Do-T 'SPC "9600,N,8,1" seed baud 1200' 'T_SPC' @('9600,N,8,1', 0x4800, 0, 0, 1200)
    Do-T 'SPC "baud=9600 parity=E" seed 4800' 'T_SPC' @('baud=9600 parity=E', 0x4800, 0, 0, 1200)
    Do-T 'SPC "baud=9600 parity=N" seed 4800 xon=5 xoff=7' 'T_SPC' @('baud=9600 parity=N', 0x4800, 5, 7, 1200)
    Do-T 'SPC "abc"' 'T_SPC' @('abc', 0x4800, 17, 19, 1200)
    # --- start paths
    Do-T 'START no-settings seed 0x4800' 'T_START_NOSET' @((& $f 'start.bin'), 0x4800, $false)
    Do-T 'START no-settings seed 0x4800 applyFails' 'T_START_NOSET' @((& $f 'start.bin'), 0x4800, $true)
    Do-T 'START no-settings seed 0 (no apply expected)' 'T_START_NOSET' @((& $f 'start.bin'), 0, $false)
    Do-T 'START trapped error after open' 'T_START_RAISE' @((& $f 'start.bin'))
    Do-T 'START configure fail (dll error kept?)' 'T_START_CFGFAIL' @((& $f 'start.bin'))
    Do-T 'START with stale caller Err' 'T_START_STALE_ERR' @((& $f 'start.bin'))
    # --- guard
    Do-T 'trapped CHECK (guard released, vba err kept)' 'T_TRAP_CHECK' @((& $f 'd5.bin'))
    Do-T 'trapped CHECK + print raises' 'T_TRAP_PRINT' @((& $f 'd5.bin'))
    Do-T 'STOP closed-handle' 'T_STOP_BOGUS' @((& $f 'd5.bin'))
    Do-T 'DRAIN 10ms queue busy once' 'T_DRAIN' @((& $f 'd5.bin'))
    Do-T 'DRAIN deadline 500ms' 'T_DRAIN_TIME' @((& $f 'd5.bin'), 500)
    Do-T 'STOP drain times out' 'T_STOP_DRAIN' @((& $f 'd5.bin'), 100000, $false, $false)
    Do-T 'STOP drain empties' 'T_STOP_DRAIN' @((& $f 'd5.bin'), 1, $false, $false)
    Do-T 'STOP drain cancelled' 'T_STOP_DRAIN' @((& $f 'd5.bin'), 100000, $true, $false)
    Do-T 'STOP drain status unreadable' 'T_STOP_DRAIN' @((& $f 'd5.bin'), 0, $false, $true)
    Do-T 'RECEIVE reply arrives late' 'T_RECV_LATE' @("$data\late_$kind.bin", $true)
    Do-T 'RECEIVE idle port' 'T_RECV_LATE' @("$data\idle_$kind.bin", $false)
    Do-T 'DEBUG_COM_PORT(99) stale Err' 'T_DEBUG_STALE' @()
    # --- send / transmit
    foreach ($k in @('error', 'bytes', 'range1', 'range2', 'text')) {
        $o = "$data\out_send_$k.bin"; Remove-Item $o -ErrorAction SilentlyContinue
        $r = RunT 'T_SEND' @($o, $k)
        "[$Which/$kind] SEND $k => $r ; file=" + (Hex-File $o)
    }
    foreach ($k in @('transmit', 'put', 'bytes', 'bytes_big', 'bytes_changed')) {
        $o = "$data\out_tx_$k.bin"; Remove-Item $o -ErrorAction SilentlyContinue
        $r = RunT 'T_TX' @($o, $k)
        $h = Hex-File $o
        if ($h.Length -gt 120) { $h = "len=$($h.Length / 2) head=" + $h.Substring(0, 40) + " ... " + ($h.Length / 2) + " bytes; count55=" + ([regex]::Matches($h, '55')).Count + " countAA=" + ([regex]::Matches($h, 'AA')).Count }
        "[$Which/$kind] TX $k => $r ; file=$h"
    }
    # --- read
    $r = RunT 'T_READ256' @((& $f 'd256.bin')); "[$Which/$kind] READ256 => $r"
    $hash = [BitConverter]::ToString([Security.Cryptography.SHA1]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($r))) ; "[$Which/$kind] READ256 sha1 => $hash"
    # --- receive limits
    foreach ($w in @('str', 'bytes')) {
        Do-T "RECV $w max=1500" 'T_RECV' @((& $f 'd5000.bin'), $w, 1500, 0, 0)
        Do-T "RECV $w max=7" 'T_RECV' @((& $f 'd5000.bin'), $w, 7, 0, 0)
        Do-T "RECV $w unlimited" 'T_RECV' @((& $f 'd5000.bin'), $w, 0, 0, 0)
        Do-T "RECV $w 3 bytes charWait=1000 total=100ms" 'T_RECV' @((& $f 'd3.bin'), $w, 0, 100, 1000)
        Do-T "RECV $w 3 bytes charWait=1000 max=3" 'T_RECV' @((& $f 'd3.bin'), $w, 3, 0, 1000)
    }
    Do-T 'RECEIVE_BYTES fixed array' 'T_RB_FIXED' @((& $f 'd5.bin'))
    Do-T 'RECEIVE_BYTES busy rejection' 'T_RB_BUSY' @((& $f 'd5.bin'))
    Do-T 'RECEIVE_BYTES locked array' 'T_RB_LOCKED' @((& $f 'd5.bin'))
    Do-T 'START no-settings seed 0x0400' 'T_START_NOSET' @((& $f 'start.bin'), 0x0400, $false)
    Do-T 'RECEIVE_BYTES max byval (hook zeroes caller var)' 'T_BYVAL' @((& $f 'd5000.bin'))
    # --- waits
    Do-T 'WAITS 9600' 'T_WAITS' @((& $f 'd5.bin'), 9600)
    Do-T 'WAITS 115200' 'T_WAITS' @((& $f 'd5.bin'), 115200)
    Do-T 'WAITS 921600' 'T_WAITS' @((& $f 'd5.bin'), 921600)
    # --- diagnostics
    Do-T 'DLL_ERROR_TEXT 21|15' 'T_DLLTEXT' @()
    Do-T 'SHOW_DCB bits=0x4000' 'T_SHOW_DCB' @(0x4000)
    Do-T 'SHOW_DCB bits=0x8000' 'T_SHOW_DCB' @(0x8000)
    Do-T 'SHOW_PORT_ERRORS clear fails' 'T_SHOW_ERR' @()
    if ($Which -eq 'new') {
        Do-T 'ALIASES' 'T_ALIASES' @()
        Do-T 'PORT_LAST_ERROR' 'T_LAST_ERROR' @("$data\lasterr_$kind.bin", "$data\does_not_exist.bin")
        Do-T 'HELPER bits 0x4801 xon=xoff=0' 'T_HELPER' @(0x4801, 0, 0)
        Do-T 'HELPER bits 0xFFFF xon=5 xoff=7' 'T_HELPER' @(0xFFFF, 5, 7)
        Do-T 'HELPER xon=xoff=3' 'T_HELPER' @(0, 3, 3)
        Do-T 'HELPER bits 0x0401' 'T_HELPER' @(0x0401, 17, 19)
    }
    $wb.Close($false)
}
} finally {
    try { $excel.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    Stop-Process -Id $watchdog.Id -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    if (Get-Process -Id $pidOut -ErrorAction SilentlyContinue) { Stop-Process -Id $pidOut -Force }
}
}
} finally {
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
}

$results = @($results | ForEach-Object { "$_" })
$results

if ($Baseline) { exit 0 }

# Timings vary from run to run, so they are compared by limit rather than exactly.
$actual = @($results | ForEach-Object { $_ -replace 'ms=\d+', 'ms=*' })

if ($UpdateExpected) {
    [IO.File]::WriteAllLines($Expected, [string[]]$actual)
    "Expected results updated: $($actual.Count) lines"
    exit 0
}

$failures = 0
# Compared by test name, so adding one test reports one difference rather than shifting every later line.
$want = @(Get-Content $Expected)
$nameOf = { param($line) ($line -split ' => ', 2)[0] }
$wantByName = @{}
foreach ($line in $want) { $wantByName[(& $nameOf $line)] = $line }
$actualNames = @{}
foreach ($line in $actual) {
    $name = & $nameOf $line
    $actualNames[$name] = $true
    if (-not $wantByName.ContainsKey($name)) {
        $failures++
        "FAIL new result with no expected value"
        "  actual:   $line"
    } elseif ($wantByName[$name] -cne $line) {
        $failures++
        "FAIL $name"
        "  expected: $($wantByName[$name])"
        "  actual:   $line"
    }
}
foreach ($line in $want) {
    if (-not $actualNames.ContainsKey((& $nameOf $line))) {
        $failures++
        "FAIL expected result is missing"
        "  expected: $line"
    }
}
foreach ($line in $results) {
    if ($line -match 'DRAIN deadline 500ms => r=False;ms=(\d+)' -and ([int]$Matches[1] -lt 450 -or [int]$Matches[1] -ge 650)) {
        $failures++
        "FAIL timing, a 500 mS drain should take about 500 mS: $line"
    }
    if ($line -match '3 bytes charWait=1000 (total=100ms|max=3) => n=3;ms=(\d+)' -and [int]$Matches[2] -ge 500) {
        $failures++
        "FAIL timing, should stop well before the 1000 mS wait: $line"
    }
}
if ($failures -gt 0) {
    "FAILED: $failures problem(s) in $($actual.Count) results"
    exit 1
}
"PASSED: all $($actual.Count) results match"
exit 0

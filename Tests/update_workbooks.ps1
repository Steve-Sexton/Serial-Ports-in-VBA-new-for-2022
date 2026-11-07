# Replaces the serial module inside SERIAL_PORT_TEST_MAIN.xlsm and SERIAL_PORT_TEST_NODEBUG.xlsm with the current
# source, then saves, reopens and compile-checks each one. It overwrites both workbooks in place: git holds the old copies.
$Repo = Split-Path $PSScriptRoot -Parent
$jobs = @(
    @{ Book = "$Repo\SERIAL_PORT_TEST_MAIN.xlsm"; Module = "$Repo\SERIAL_PORT_VBA.bas" },
    @{ Book = "$Repo\SERIAL_PORT_TEST_NODEBUG.xlsm"; Module = "$Repo\No-Debug\SERIAL_PORT_VBA_NO_DEBUG.bas" }
)
$ribbon = "$Repo\Ribbon\SERIAL_PORT_RIBBON.bas"

Add-Type -Namespace Native3 -Name Win32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$excel = New-Object -ComObject Excel.Application
$pidOut = [uint32]0
[void][Native3.Win32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidOut)
$watchdog = Start-Process pwsh -WindowStyle Hidden -PassThru -ArgumentList '-NoProfile', '-Command', "Start-Sleep 240; Stop-Process -Id $pidOut -Force"
$excel.Visible = $false; $excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

function List-Comps($wb, $tag) {
    foreach ($c in $wb.VBProject.VBComponents) { "$tag component: $($c.Name) type=$($c.Type) lines=$($c.CodeModule.CountOfLines)" }
}
function Compile-Check($wb) {
    $excel.VBE.ActiveVBProject = $wb.VBProject
    $compile = $excel.VBE.CommandBars.FindControl(1, 578)
    $before = $compile.Enabled
    try { $compile.Execute() } catch { return "compile Execute threw ($($_.Exception.Message.Trim())); button enabled before=$before after=$($compile.Enabled)" }
    return "compile button enabled before=$before after compile (False = clean): $($compile.Enabled)"
}
try {
    foreach ($j in $jobs) {
        $leaf = Split-Path $j.Book -Leaf
        "=== $leaf"
        $wb = $excel.Workbooks.Open($j.Book)
        List-Comps $wb 'before'
        $names = @($wb.VBProject.VBComponents | ForEach-Object { $_.Name })
        if ($names -contains 'SERIAL_PORT_VBA') {
            $wb.VBProject.VBComponents.Remove($wb.VBProject.VBComponents.Item('SERIAL_PORT_VBA'))
            [void]$wb.VBProject.VBComponents.Import($j.Module)
            "replaced SERIAL_PORT_VBA"
        } else { "NO SERIAL_PORT_VBA component found - nothing replaced" }
        if ($names -contains 'SERIAL_PORT_RIBBON') {
            $wb.VBProject.VBComponents.Remove($wb.VBProject.VBComponents.Item('SERIAL_PORT_RIBBON'))
            [void]$wb.VBProject.VBComponents.Import($ribbon)
            "replaced SERIAL_PORT_RIBBON"
        } else { "no SERIAL_PORT_RIBBON component in this workbook - not added" }
        List-Comps $wb 'after'
        Compile-Check $wb
        $wb.Save()
        $wb.Close($false)
        "--- reopen"
        $wb = $excel.Workbooks.Open($j.Book)
        List-Comps $wb 'reopened'
        Compile-Check $wb
        $name = $wb.Name
        "ABI_SELF_TEST = " + $excel.Run("'$name'!ABI_SELF_TEST")
        "START invalid port = " + $excel.Run("'$name'!START_COM_PORT", 99)
        $wb.Close($false)
    }
} finally {
    try { $excel.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    Stop-Process -Id $watchdog.Id -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    if (Get-Process -Id $pidOut -ErrorAction SilentlyContinue) { Stop-Process -Id $pidOut -Force }
}
"done"

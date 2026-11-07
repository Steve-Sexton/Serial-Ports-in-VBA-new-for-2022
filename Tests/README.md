# Serial module tests

These tests run the VBA modules in Excel without a serial port. They catch logic regressions. They do not show
that the modules work with real hardware.

**Do not import anything from this folder into a workbook you use.** `tests.bas` only works inside the test
builds that `mk_test.py` makes, which replace the Windows port calls with stubs.

## Requirements

- Windows with desktop Excel
- In Excel, **File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA
  project object model**
- Python 3 and PowerShell 7 (`pwsh`)

## Running

```powershell
pwsh -File Tests\run_tests.ps1
```

It prints every result, then `PASSED` or `FAILED`, and exits 0 or 1. It opens no COM port, starts its own hidden
Excel, and does not change the modules or the test workbooks. A watchdog stops that Excel after 5 minutes.

| Command | What it does |
|---|---|
| `run_tests.ps1 -Baseline cbf8166` | Runs the code from before the 2026-11-07 fixes and prints the results, so you can see the original bugs. Use a commit from before `cefd14c`. |
| `run_tests.ps1 -UpdateExpected` | Rewrites `expected_results.txt`. Check the changed results first, because this accepts whatever the code does now. |
| `update_workbooks.ps1` | Copies the current modules into `SERIAL_PORT_TEST_MAIN.xlsm` and `SERIAL_PORT_TEST_NODEBUG.xlsm`, overwriting both. |

## What is real and what is stubbed

Each module is copied to a temp folder and changed there. The file in the repo is never edited.

| Real | Stubbed or changed |
|---|---|
| `ReadFile`, `WriteFile`, `CreateFile`, `CloseHandle` on a temp file | `GetCommState`, `SetCommState`, `SetCommTimeouts`, `ClearCommError`, `EscapeCommFunction` |
| `BuildCommDCB` | `PurgeComm` is stubbed too; `OPEN_COM_PORT` opens the temp file instead of `\\.\COMn` |
| All VBA logic | Hooks inject errors and run code at yield points |

## What these tests cannot show

- Anything about a real port: flag changes reaching the driver, parity, handshaking, whether a graceful stop
  actually waits (the queue count is stubbed, and Microsoft documents it as zero for non-overlapped writes), or
  timing on the wire. Run a loopback test on a real or virtual port pair for those.
- Cancellation under live traffic, or the write retry path after a stall.
- Rejection of text that expands on a double-byte code page, unless you run on such a system.
- The Ribbon buttons. They are compiled but not clicked, because they show message boxes.

## Expected results

`expected_results.txt` came from 64-bit Office 2021 on a code page 1252 system. On a different code page the
256-byte read results can differ legitimately. Results should not depend on bitness, so a difference on 32-bit
Office is worth investigating. 32-bit Office has not been run.

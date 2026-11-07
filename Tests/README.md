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

## Loopback test

`run_loopback.ps1` and `loopback_tests.bas` run the real, unstubbed modules against a real or virtual COM port.
**It has not been run on real ports yet.** Only `-DryRun` and the Excel compile check have been run.

**It sends test data on the ports you give it** (all 256 byte values, text, parity-error bytes, flow-control
stalls). Do not connect anything that reacts to arbitrary bytes. Do not import `loopback_tests.bas` into a
workbook you use.

### Wiring

| Mode | Command | Wiring |
|---|---|---|
| com0com pair | `-PortA 3 -PortB 4` | Create a pair in com0com. Nothing else. Parity is not modelled, and baud rate is not paced. |
| Null-modem cable | `-PortA 3 -PortB 4` | Two real ports joined by a null-modem cable (TX-RX crossed, RTS-CTS crossed). |
| Loopback plug | `-PortA 3` | One port, DB9 pins 2 and 3 joined (TX to RX). |
| Loopback plug with flow control | `-PortA 3 -RtsCtsLoop` | Pins 2-3 and pins 7-8 joined (RTS to CTS). Needed for the flow-control test. |

### Commands

```powershell
pwsh -File Tests\run_loopback.ps1 -DryRun                       # first: checks the setup, opens no port
pwsh -File Tests\run_loopback.ps1 -PortA 3                      # plug
pwsh -File Tests\run_loopback.ps1 -PortA 3 -RtsCtsLoop          # plug with RTS wired to CTS
pwsh -File Tests\run_loopback.ps1 -PortA 3 -PortB 4             # pair
pwsh -File Tests\run_loopback.ps1 -PortA 3 -Module Main -OutFile loopback.txt
```

`-Module` is `Main`, `NoDebug` or `Both` (default). Each module runs in its own throwaway workbook. `-DryRun`
uses port 0, which both modules reject before any native call, so every test ends as SKIP. It exits 0 if no
result is FAIL and the compile worked, 1 if not, and 2 for a usage error or a port that is not listed in
`HKLM:\HARDWARE\DEVICEMAP\SERIALCOMM`. The run takes a few minutes at most. A watchdog stops Excel after 15 minutes.
Unless it is a dry run, it refuses a port that is not in the registry list and prints each port's device name,
which shows com0com or a USB/serial driver.

### Results

| Result | Meaning |
|---|---|
| PASS | The check held. |
| FAIL | The check did not hold, or a VBA error occurred, or the module did not compile. |
| INFO | A measurement that depends on the hardware. Read it, it is not pass or fail. |
| SKIP | The test could not run: a port did not start (always so in `-DryRun`), or the wiring needed was not there. |

### What it changes

- A temp copy of each module has its `COM_PORT_MIN` and `COM_PORT_MAX` lines changed to cover the chosen ports.
  Nothing else in the copy changes, and the files in the repo are never edited. The temp folder is deleted.
- Port settings. Each port's baud, parity, data and stop bits are recorded first and put back at the end, and the
  restore is reported as a result. Handshaking can't be read through the module, so the restore sets it to off
  (no CTS, DSR or XON/XOFF handshaking, DTR and RTS on). The tests change it along the way: the comma-form
  settings test rewrites DTR, RTS and handshaking, and the flow-control test turns on CTS handshaking. **If your
  port normally uses hardware or XON/XOFF handshaking, set it again after the test.**
- Starting a port also clears the error-character, discard-NUL and abort-on-error flags, as the modules always do.

### What it still cannot test

- Cancel under live traffic.
- Unplugging a USB adapter and recovering.
- 32-bit Office.
- On a virtual pair: parity detection, baud-rate timing, and hardware flow control may not behave as on real
  hardware.

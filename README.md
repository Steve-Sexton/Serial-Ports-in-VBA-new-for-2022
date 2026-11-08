# VBA Serial Port routines for Microsoft Office
## Windows 10/11, Office 2016-2021 (Excel, Word, Access) — hardened 2026 revision

Getting Serial (COM) Ports working as intended in VBA can be surprisingly difficult in certain usage scenarios. 

New VBA routines here will help resolve these issues in Excel, Word and Access (Windows PC versions only).

Functions are straightforward to use with coding style to support infrequent VBA users and developers.

Intended to help implement ad-hoc projects for serial data acquisition or transfer.

No plug-ins, DLLs, ActiveX, licences, payments or registrations are required.  

**2026 hardening revision** — following an independent verification & validation review, both modules were
reworked for correctness and robustness: exact Win32 declarations for 32/64-bit Office, a per-port operation
guard against re-entrant calls, cooperative cancellation, runtime error handling on every public entry point,
partial-write recovery, bounded receives, binary-safe `Byte()` transport, and corrected timing/error decoding.
See [CHANGELOG.md](CHANGELOG.md) for the full list.

> **Not yet validated on a real serial port.** The 2026 revisions have been compiled, and tested against a stubbed
> harness, on 64-bit Office 2021 only. Before relying on them, run the [loopback test](Tests/README.md#loopback-test) on
> the ports you will use, and test 32-bit Office separately if you use it. Until then, treat this revision as a
> release candidate.
>
> **Upgrading from an earlier version?** Some behaviour changed. Check your macros for:
> - short settings strings such as `"8N1"` or `"COM1"`, which used to be ignored and now make `start_com_port` fail
> - `stop_com_port(port, drain)`, which can now return `False` and leave the port open
> - `receive_com_port` calls on an idle port, which now wait for the read timer (about 100-500 mS) before
>   returning empty
>
> The [CHANGELOG](CHANGELOG.md) lists every change.

<details><summary>More Information</summary>
<p>
   
<details><summary>VBA Issues</summary>
<p>

The in-built VBA functions for COM Port data can suffer from the following issues :- 
   
1. Setting port parameters with the VBA open command may not work in some Windows versions e.g.

   `Open "COM1:9600,N,8,1" For Read Access As #1`       \
     _(command line workaround known, settings can revert after reboot)_

2. Attempting to read data when there is none waiting will cause VBA to hang with a 'not responding' message.  
  
   `Get #1, , Read_Data_Byte`  
   
   The new functions address both of these issues, and also where data transfers take longer than the 5-6 second VBA timeout.
   
</p>
</details>   

<details><summary>Background</summary>  
<p>

The legacy of serial comms means that many online solution searches are now time-expired with links to defunct web sites etc.    

New functions here are therefore a fresh start and are based largely on Microsoft's Win32 API calls and documentation. 

Originally developed on Windows 10 (64-Bit) with Microsoft Office 2019 Professional (32-Bit VBA7).

The original 2022 code was tested on Office 2016 Professional (64-bit VBA7) and Office 2019 Professional (32-Bit VBA7).

The 2026 revisions compile clean on Office 16.0 (VBA7). The 2026-11 changes have only been compiled and tested on
64-bit Office 2021, and not yet on a real port. Win32 `BOOL` results are declared `As Long`,
native NULL pointers are passed as explicit `LongPtr` zeros, and a built-in `ABI_SELF_TEST` verifies structure
sizes (`DCB=28, COMSTAT=12, COMMTIMEOUTS=20`) before the first port is opened on any host.

</p>
</details>

<details><summary>COM Ports</summary>
<p>

Multiple com ports are supported, including physical hardware ports and synthetic virtual software ports.

Each module accepts COM1 and COM2 only until you change `COM_PORT_MIN` and `COM_PORT_MAX` near the top of the
module. A port outside that range is rejected as invalid.

All read and write functions are synchronous, in part because not all serial port types support overlapped operation.

Performance on a modern PC is good, with software timing delays required to allow the relatively slow serial com ports to catch up. 

Reading, Writing and Waiting are 'timesliced' to ensure that VBA remains responsive during any extended data transfers or waiting times. 

</p>
</details>

<details><summary>Text vs Binary data</summary>
<p>

* The classic string functions (`read/receive/transmit/send/get/put_com_port`) are for **SBCS/ANSI text**:
  VBA strings pass through an ANSI conversion at the native boundary, so non-ASCII, DBCS and arbitrary
  binary payloads can be transformed.
* For binary protocols and full 0–255 byte values use **`transmit_bytes` / `receive_bytes`**, which move raw
  `Byte()` arrays with no conversion — embedded NUL and all byte values are preserved.

</p>
</details>

<details><summary>Concurrency, cancellation and recovery</summary>
<p>

Every port operation yields to Windows (`DoEvents`) so VBA stays responsive — which also means a timer, event
handler or Ribbon callback can call back into the module while an operation is still running. Each port
therefore has a non-blocking operation guard:

* A second call on the **same** port while one is active is rejected cleanly (returns `False`, an empty
  string, or `-1`) instead of corrupting the operation in progress. The guard is independent per port, but a long
  nested operation on one port (for example a callback that runs a long receive) suspends the outer operation
  on another port until it returns.
* `port_busy(1)` tests the guard; `cancel_com_port(1)` asks a long receive/transmit/wait to unwind cleanly
  at its next check, keeping any partial data.
* `stop_com_port(1)` on a busy port automatically requests cancellation and returns `False` — call it again
  once the operation has unwound: retry after returning from the callback, for example with `Application.OnTime`,
  never in a loop inside the callback. `stop_com_port(1, 2000)` drains the transmit queue (graceful close) before
  closing. If the driver still reports queued data after that time it returns `False` and leaves the port open
  with nothing discarded, so you can retry or abort. The default remains an abort close. `True` does not prove
  delivery: the driver may report an empty queue at once on these non-overlapped handles, so have the device
  acknowledge if delivery matters.
* `port_last_error(1)` returns the Win32 error from the most recent operation on the port, so a failed read can
  be told apart from no data. Read it straight after the call, because the next operation replaces it.
* `clear_port_busy(1)` is a recovery-only escape hatch after a VBE Ctrl-Break (break mode only). A VBE Reset or
  `End` loses the stored port handle while the OS port stays open, so the host application must be closed to
  release it.
* `port_error_history(1)` returns hardware error bits (overflow/overrun/parity/framing/break) accumulated by
  status polling, so no error evidence is lost between checks.

</p>
</details>

<details><summary>Debugging</summary>
<p>

* Debugging can be set on/off per port with results shown in the VBA immediate window. 

* Extensive debug functionality makes several modules quite verbose. 

* Runtime errors are trapped at every public entry point, recorded per port, and reported as the function's
  normal failure value. The one exception is a failure while building the error report itself, such as running
  out of memory. Win32 errors are available from `port_last_error` in both modules; the per-port `VBA_Error`
  record is in the main module only.

</p>
</details>  

<details><summary>Testing</summary>
<p>

| Test | How to run | What it shows |
|---|---|---|
| Compile test workbooks | Open `SERIAL_PORT_TEST_MAIN.xlsm` or `SERIAL_PORT_TEST_NODEBUG.xlsm` | Each module compiles in Office. Their smoke tests used invalid port numbers only, so no port is touched. |
| Unit harness | `pwsh -File Tests\run_tests.ps1` | Logic regressions. The Windows port calls are stubbed and the modules run against a temporary file. |
| Loopback test | `pwsh -File Tests\run_loopback.ps1 -PortA 3` (add `-PortB 4` for a pair) | The real modules on a real or virtual port: settings, all 256 byte values, receive limits, a late reply, flow control and parity. |

**The loopback test is the only one that touches a real port, and it has not been run yet.** It is needed
because the other two cannot show what a driver does: whether settings reach the port, whether a graceful stop
really waits for queued data, whether parity errors are detected, or how timing behaves on the wire. Run it on
each kind of port you use (a USB adapter can behave differently from a virtual pair), and on 32-bit Office if you
use it. [Tests/README.md](Tests/README.md) has the wiring, the commands and how to read the results.

The tests need desktop Excel with **Trust access to the VBA project object model** turned on, PowerShell 7, and
Python 3 for the unit harness. So far they have only been run on 64-bit Office 2021.

**Never import both modules into the same VBA project** — they share public names by design.

</p>
</details>

<details><summary>Other Versions</summary>
<p> 

* [No-Debug](No-Debug)      (more compact)

* [Simplified](https://github.com/Serialcomms/Serial-Ports-in-VBA-Simple-2022)   (single com port)

* [Minimal](https://github.com/Serialcomms/Serial-Ports-in-VBA-Extra-Simple-2022)   (single com port, no settings)

* [VBA6 / 32-Bit](https://github.com/Serialcomms/Serial-Ports-in-VBA6-legacy-for-2022)  (legacy Windows/Office)

</p>
</details> 
   
<details><summary>Alternatives (Excel Only)</summary>
<p> 

* [Microsoft Data Streamer for Excel](https://learn.microsoft.com/en-us/microsoft-365/education/data-streamer/)   
 
</p>
</details>   
     
<details><summary>Optional steps for Excel only</summary>
<p>  

- Functions can be used directly in Worksheet cells as formulas where appropriate.  
- Remove comment mark before `Option Private Module` to prevent function names appearing in cell formula drop-down lists.  
- Remove comment mark before `Application.Volatile` where indicated to refresh results when functions are used in cells and the worksheet is recalculated (e.g. with F9 key).
- Port status functions (`device_ready`, `clear_to_send`, etc.) are safe to use in cells while transfers run —
  they never disturb an operation in progress.

</p>
</details>

<details><summary>Optional Ribbon Customisation</summary>
<p>

[Office 2010 XML](/Ribbon/RIBBON_2010.xml) and [SERIAL_PORT_RIBBON](/Ribbon/SERIAL_PORT_RIBBON.bas) example files are available in the [Ribbon](/Ribbon) folder. 
   
</p>
</details>   
   
<details><summary>Function List</summary>
<p>   

[COM Port Control](Functions/Function_List_Control.md)
   
[Read/Write/Check Data](Functions/Function_List_Data.md)
   
[Port Signalling Functions](Functions/Function_List_Signalling.md)

[Show Functions](Functions/Function_List_Show.md)

[Helper Functions](Functions/Function_List_Helpers.md)

Private functions are not intended to be called directly by users.
  
</p>
</details>   
   
</p>
</details>   

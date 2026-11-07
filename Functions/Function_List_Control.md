# VBA Serial Port Functions

## start, stop and debug COM ports

First parameter (1) is a valid[^1] COM Port number on host PC

| VBA Function                         | Description                                                                                                   |
| ------------------------------------ | --------------------------------------------------------------------------------------------------------------|
| `debug_com_port(1)`                  | Toggles debug messaging on/off                                                                                |
| `debug_com_port(1,True)`             | Set port debug messaging on                                                                                   |
| `debug_com_port(1,False)`            | Set port debug messaging off                                                                                  |
| `start_com_port(1)`                  | Starts port with existing settings                                                                            |
| `start_com_port(1,"Baud=1200")`      | Starts port with settings as supplied [^5]                                                                    |
| `start_com_port(1,SCANNER)`          | Starts port with settings defined in string constant or variable e.g. SCANNER                                 |
| `get_port_settings(1)`               | Returns string [^2] with port settings or error text                                                          |
| `stop_com_port(1)`                   | Stops port and hands its control back to Windows. Queued transmit/receive data is purged (abort close)        |
| `stop_com_port(1,2000)`              | Waits up to 2000mS for queued transmit data to be sent, then stops. If data is still queued, returns `False` and leaves the port open [^6] |
| `port_busy(1)`                       | Returns `True` while an operation on the port is in progress [^3]                                             |
| `cancel_com_port(1)`                 | Requests cooperative cancel of the running operation; it unwinds cleanly at its next check [^4]               |
| `clear_port_busy(1)`                 | Recovery only - clears the port operation guard after a VBE Ctrl-Break (break mode only) [^3]                  |
| `abi_self_test()`                    | Verifies Win32 structure sizes (DCB=28, COMSTAT=12, COMMTIMEOUTS=20); runs automatically before first start   |
| `port_last_error(1)`                 | Returns the Win32 error from the most recent native call on the port, 0 if it succeeded [^7]                  |
| `port_error_history(1)`              | Returns OR-accumulated comms-error bits captured by status polls (1=overflow 2=overrun 4=parity 8=frame 16=break) |
| `port_error_history(1,True)`         | As above, and clears the accumulated history after reading                                                    |

* Debug results are shown in the VBA Immediate Window (Control-G)
* Debug functions return `True` or `False` to indicate debug state
* Other functions return `True` or `False` to indicate success or failure

[^1]: Valid Minimum and Maximum port numbers should be defined in declarations section at the start of the module. 
[^2]: e.g. 9600-8-N-1 , PORT-NOT-STARTED, INVALID-PORT
[^3]: Each port has one operation guard. A port function called again for the same port (for example from a timer,
      event handler or Ribbon callback while an earlier call is still running) is rejected rather than allowed to
      interfere with the operation in progress, and returns its usual failure value (`False`, empty string, or -1).
      `clear_port_busy` helps in break mode only: a VBE Reset or `End` loses the stored handle while the OS port
      stays open, so the host application must be closed to release it. The guard is independent per port, but a
      long nested operation on one port suspends the outer operation on another.
[^4]: `stop_com_port` on a busy port automatically requests cancellation and returns `False`. Retry after returning
      from the callback, for example with `Application.OnTime`, never in a loop inside the callback. Data received or
      sent before the cancel point is preserved.
[^5]: Settings may be given in comma form (`9600,N,8,1`). `key=value` forms are cleaned first, with commas
      becoming spaces only when the string contains `=`. A short string such as `9600` is accepted, and one
      `BuildCommDCB` rejects, such as `abc`, fails the start. The bare-baud and comma forms also rewrite line
      control: they set one stop bit and turn on DTR and RTS, and a trailing `,P` turns on the CTS/DSR handshake.
      The `key=value` form leaves DTR, RTS and handshaking alone but still resets stop bits, so include `Stop=` if
      it matters. Starting with settings also clears the error-character, discard-NUL and abort-on-error flags.
[^6]: If the queue does not empty in time, or its status cannot be read, the result is `False`, the port stays open
      and nothing is discarded. Retry, or call `stop_com_port(1)` to abort. A cancel request during the wait forces
      the abort close and the stop then returns `True`. `Drain_MilliSeconds` is `Optional ByVal`.
[^7]: Returns -1 for an invalid port number. Use it to tell a failed operation from an empty result, for example a
      failed read from no data. A call rejected because the port is busy or stopped makes no native call and leaves
      the value unchanged. Trapped VBA runtime errors are not included; the main module records those separately
      in `VBA_Error`, which is not exposed.

# VBA Helper Functions

## Port state, timing and conversion helpers

These are public so your own code and the test scripts can use them. Apart from `example`, which sends a request
and reads the reply, they send nothing on the port and take no port guard. *Main only* means the function exists only in `SERIAL_PORT_VBA.bas`, not in the No-Debug module.

| VBA Function | Module | Description |
| ------------ | ------ | ----------- |
| `port_started(1)` | Both | Returns `True` if the port number is valid and the port is started |
| `port_ready(1)` | Both | Same as `port_started` (the No-Debug module's original name) |
| `port_number_valid(1)` | Both | Returns `True` if the number is within `COM_PORT_MIN` to `COM_PORT_MAX` [^1] |
| `port_valid(1)` | Both | Same as `port_number_valid` (the No-Debug module's original name) |
| `get_host_milliseconds()` | Both | Milliseconds from the high-resolution performance counter, as `Currency` [^2] |
| `get_host_microseconds()` | Both | Microseconds from the high-resolution performance counter, as `Currency` [^2] |
| `convert_parity(p)` | Both | Converts a DCB parity value (0-4) to `N`, `O`, `E`, `M` or `S` |
| `convert_stopbits(s)` | Both | Converts a DCB stop-bits value (0-2) to `1`, `1.5` or `2` |
| `convert_line_error(e)` | Main only | Converts a single communications-error value to `BREAK`, `FRAME`, `OVERFLOW`, `OVERRUN`, `PARITY` or `UNKNOWN` |
| `decode_port_errors(mask)` | Main only | Prints each bit of an error mask, for example from `port_error_history`, to the Immediate Window |
| `dll_error_text(code)` | Main only | Returns a description of a Win32 error code, for example from `port_last_error` |
| `timestamp()` | Main only | Returns the local time with milliseconds, as used at the start of each debug line |
| `print_debug_text(...)` | Main only | Prints one formatted debug line. Used internally; it never raises an error |
| `print_show_text(...)` | Main only | Prints one formatted line for the show functions. Used internally; it never raises an error |
| `template(1)` | Main only | A skeleton to copy when writing your own port function. Returns a short text if the port is started |
| `example(1)` | Main only | A worked example that sends a request and reads the reply. See the comments in the module |

* `port_busy`, `port_last_error`, `port_error_history`, `cancel_com_port`, `clear_port_busy` and `abi_self_test`
  are in the [COM Port Control](Function_List_Control.md) list.
* The two modules share these names and must never be loaded into the same VBA project.

[^1]: Both modules default to 1 and 2, which is COM1 and COM2. Change the constants near the top of the module to
      use other ports.
[^2]: The counter starts when Windows starts, so use the difference between two readings. `Currency` is used
      because a millisecond count overflows a `Long` after about 24.9 days.

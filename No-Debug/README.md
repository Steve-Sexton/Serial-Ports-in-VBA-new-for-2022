# Non-Debug Version
This is the 'lightweight' version of SERIAL_PORTS_VBA with all debug and show functionality removed.

Main VBA port functions remain functionally equivalent. 

Port debug function remains for compatibility and always returns false.

`port_last_error` returns the most recent Win32 error here too. The per-port `VBA_Error` record exists in the main module only.

Suggestion is to use full debug version initially for testing, and then change to this version for regular use if appropriate.

The automated tests in [Tests](../Tests) run this module as well as the main one. The 2026 revision has not yet been
validated on a real serial port: run the [loopback test](../Tests/README.md#loopback-test) on your ports before
relying on it.

Like the main module, this module accepts COM1 and COM2 only until you change `COM_PORT_MIN` and `COM_PORT_MAX`.

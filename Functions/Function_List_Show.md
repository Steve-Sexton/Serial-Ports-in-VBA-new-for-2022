# VBA Show Functions

## Show COM port information

First parameter (1) is a valid[^1] and started COM Port number on host PC

| VBA Function                         | Description                                                                                                   |
| ------------------------------------ | --------------------------------------------------------------------------------------------------------------|
| `show_port_dcb(1)`                   | Show decoded contents of the port's Device Control Block                                                      |
| `show_port_errors(1)`                | Show port overflow/overrun/parity/framing/break conditions                                                    |          
| `show_port_modem(1)`                 | Show port modem signals DSR/CTS/RING/CD                                                                       |  
| `show_port_queues(1)`                | Show port input and output data queues                                                                        |  
| `show_port_status(1)`                | Show various transmission waiting conditions etc.                                                             |
| `show_port_timers(1)`                | Show read and write timer values                                                                              |
| `show_port_values(1)`                | Show various values used by receive_com_port etc.                                                             |
| `show_port_all(1)`                   | Runs all of the show functions above, and returns `True` only if every one succeeded                         |

* Show functions not included in No-Debug version.
* Show results are in the VBA Immediate Window (Control-G)

[^1]: Valid Minimum and Maximum port numbers should be defined in declarations section at the start of the module. Both modules default to 1 and 2 (COM1 and COM2), so change `COM_PORT_MIN` and `COM_PORT_MAX` to use other ports. 
  

Attribute VB_Name = "SERIAL_PORT_VBA"
'
' https://github.com/Serialcomms/Serial-Ports-in-VBA-new-for-2022
' https://github.com/Serialcomms/Serial-Ports-in-VBA-new-for-2022/tree/main/No-Debug
'
  Option Explicit
'
' Option Private Module
'
'-------------------------------------------------------------------------
' Change min/max values below to match your com ports and intended usage.
' Data functions should work with most hardware and software port types.
' Signalling functions should be tested individually if required.
' Functions work with port numbers greater than 10 if specified.
'
Private Const COM_PORT_MIN As Integer = 1               ' = COM1
Private Const COM_PORT_MAX As Integer = 2               ' = COM2

'-------------------------------------------------------------------------
' Optional - can define port settings for your devices here.
' Use constant to start com port instead of settings string.
'
' Public Const BARCODE As String = "Baud=9600 Data=8 Parity=N Stop=1"
' Public Const SAT_NAV As String = "Baud=1200 Data=7 Parity=E Stop=1"
'-------------------------------------------------------------------------

Private Const HANDLE_INVALID As LongPtr = -1
Private Const NULL_POINTER As LongPtr = 0               ' explicit NULL for native pointer/handle arguments
Private Const MAXDWORD As Long = &HFFFFFFFF
Private Const VBA_TIMEOUT As Long = 5200                ' VBA "Not Responding" time in MilliSeconds (approximate)
Private Const LONG_NEG_1 As Long = -1

Private Const LONG_0  As Long = 0                       ' some predefined constants for minor performance gain.
Private Const LONG_1  As Long = 1
Private Const LONG_2  As Long = 2
Private Const LONG_3  As Long = 3
Private Const LONG_4  As Long = 4
Private Const LONG_5  As Long = 5

Private Const LONG_10 As Long = 10
Private Const LONG_21 As Long = 21
Private Const LONG_50 As Long = 50

Private Const LONG_100 As Long = 100
Private Const LONG_333 As Long = 333
Private Const LONG_1000 As Long = 1000
Private Const LONG_2048 As Long = 2048
Private Const LONG_3000 As Long = 3000
Private Const LONG_1E6  As Long = 1000000

Private Const HEX_10 As Byte = &H10                      ' some hexadecimal constants for minor readability gain.
Private Const HEX_20 As Byte = &H20
Private Const HEX_40 As Byte = &H40
Private Const HEX_80 As Byte = &H80

Private Const HEX_400 As Long = &H400
Private Const HEX_800 As Long = &H800
Private Const HEX_4000 As Long = &H4000

Private Type DEVICE_CONTROL_BLOCK

             LENGTH_DCB As Long
             BAUD_RATE  As Long
             BIT_FIELD  As Long
             RESERVED   As Integer
             LIMIT_XON  As Integer
             LIMIT_XOFF As Integer
             BYTE_SIZE  As Byte
             PARITY     As Byte
             STOP_BITS  As Byte
             CHAR_XON   As Byte
             CHAR_XOFF  As Byte
             CHAR_ERROR As Byte
             CHAR_EOF   As Byte
             CHAR_EVENT As Byte
             RESERVED_1 As Integer
End Type

Private Type COM_PORT_STATUS

             BIT_FIELD As Long                            ' 32 bits = waiting for CTS, DRS etc. Top 25 bits not used.
             QUEUE_IN  As Long
             QUEUE_OUT As Long
End Type

Private Type COM_PORT_TIMEOUTS                            ' Check latest Microsoft documentation before changing

             Read_Interval_Timeout          As Long
             Read_Total_Timeout_Multiplier  As Long
             Read_Total_Timeout_Constant    As Long
             Write_Total_Timeout_Multiplier As Long
             Write_Total_Timeout_Constant   As Long
End Type

Private Type COM_PORT_TIMERS
            
             Char_Loop_Wait As Long                        ' Arbitrary loop wait time before next read (assuming single characters)
             Data_Loop_Wait As Long                        ' Arbitrary loop wait time before next read (assuming multiple characters)
             Line_Loop_Wait As Long                        ' Arbitrary loop wait time before next read (assuming lines)
             Exit_Loop_Wait As Long                        ' Arbitrary loop wait time before read exit (allow minimum 1 character time)
             Read_Timeout As Boolean
             Timeslice_Bytes As Long                       ' Approximate bytes per second for timesliced synchronous read/write
             Bytes_Per_Second As Long
             Port_Data_Time As Currency                    ' Currency-scaled QPC MicroSeconds of the last read of > 0 bytes, or of the start of RECEIVE_COM_PORT
             Last_Data_Time As Currency                    ' Currency-scaled time in QPC MicroSeconds since Port_Data_Time
             Read_Wait_Time As Currency                    ' Currency-scaled time in QPC MicroSeconds of read wait before timeout
             Timing_QPC_Now As Currency                    ' Currency-scaled time in QPC MicroSeconds for timing data start
             Timing_QPC_End As Currency                    ' Currency-scaled time in QPC MicroSeconds for timing data end
             Frame_MilliSeconds As Single                  ' Approximate time in MilliSeconds required to send or receive a character
             Frame_MicroSeconds As Single                  ' Approximate time in MicroSeconds required to send or receive a character
End Type

Private Type COM_PORT_BUFFERS
            
             Read_Result As String
             Read_Buffer As String * 4096                  ' fixed size buffer for synchronous port read (maximum timeslice bytes)
             Write_Result As String
             Write_Buffer As String
             Receive_Result As String
             Receive_Buffer As String
             Receive_Length As Long
             Transmit_Length As Long
             Transmit_Result As String
             Transmit_Buffer As String
             Read_Buffer_Empty As Boolean
             Read_Buffer_Length As Long
             Synchronous_Bytes_Read As Long
             Synchronous_Bytes_Sent As Long
End Type

Private Type COM_PORT_PROFILE

             Handle As LongPtr
             DLL_Error As Long                            ' Win32 error from the most recent native call on this port (see PORT_LAST_ERROR)
             Errors As Long
             Error_History As Long                        ' OR-accumulated CE_ bits from every ClearCommError poll, so
                                                          ' status polling cannot silently destroy error evidence
             Signals As Long
             Port_Busy As Boolean                         ' per-port operation guard, set while a port operation is in progress
             Cancel_Requested As Boolean                  ' cooperative cancel - the active operation unwinds at its next check
             Status As COM_PORT_STATUS
             Timers As COM_PORT_TIMERS
             Buffers As COM_PORT_BUFFERS
             Timeouts As COM_PORT_TIMEOUTS
             DCB As DEVICE_CONTROL_BLOCK
End Type

Private COM_PORT(COM_PORT_MIN To COM_PORT_MAX) As COM_PORT_PROFILE

Private Host_Ticks_Per_Second As Currency                 ' cached QueryPerformanceFrequency, Currency-scaled (0 = not yet read)

Private ABI_Checked As Boolean                             ' structure sizes verified once per session before first port open
Private ABI_Valid As Boolean

' Win32 BOOL is a 32-bit value, so native BOOL results are declared As Long and tested with <> LONG_0.
' Native pointer/handle arguments are declared As LongPtr and passed as explicit NULL_POINTER - not as untyped Optional Variants.

Private Declare PtrSafe Sub Kernel_Sleep_MilliSeconds Lib "Kernel32.dll" Alias "Sleep" (ByVal Sleep_MilliSeconds As Long)
Private Declare PtrSafe Function QPC Lib "Kernel32.dll" Alias "QueryPerformanceCounter" (ByRef Query_PerfCounter As Currency) As Long
Private Declare PtrSafe Function QPF Lib "Kernel32.dll" Alias "QueryPerformanceFrequency" (ByRef Query_Frequency As Currency) As Long

Private Declare PtrSafe Function Query_Port_DCB Lib "Kernel32.dll" Alias "GetCommState" (ByVal Port_Handle As LongPtr, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Long
Private Declare PtrSafe Function Apply_Port_DCB Lib "Kernel32.dll" Alias "SetCommState" (ByVal Port_Handle As LongPtr, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Long
Private Declare PtrSafe Function Build_Port_DCB Lib "Kernel32.dll" Alias "BuildCommDCBA" (ByVal Config_Text As String, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Long
Private Declare PtrSafe Function Get_Com_Timers Lib "Kernel32.dll" Alias "GetCommTimeouts" (ByVal Port_Handle As LongPtr, ByRef TIMEOUT As COM_PORT_TIMEOUTS) As Long
Private Declare PtrSafe Function Set_Com_Timers Lib "Kernel32.dll" Alias "SetCommTimeouts" (ByVal Port_Handle As LongPtr, ByRef TIMEOUT As COM_PORT_TIMEOUTS) As Long
Private Declare PtrSafe Function Set_Com_Signal Lib "Kernel32.dll" Alias "EscapeCommFunction" (ByVal Port_Handle As LongPtr, ByVal Signal_Function As Long) As Long
Private Declare PtrSafe Function Get_Port_Modem Lib "Kernel32.dll" Alias "GetCommModemStatus" (ByVal Port_Handle As LongPtr, ByRef Modem_Status As Long) As Long
Private Declare PtrSafe Function Com_Port_Purge Lib "Kernel32.dll" Alias "PurgeComm" (ByVal Port_Handle As LongPtr, ByVal Port_Purge_Flags As Long) As Long
Private Declare PtrSafe Function Com_Port_Close Lib "Kernel32.dll" Alias "CloseHandle" (ByVal Port_Handle As LongPtr) As Long

Private Declare PtrSafe Function Com_Port_Clear Lib "Kernel32.dll" Alias "ClearCommError" _
(ByVal Port_Handle As LongPtr, ByRef Port_Error_Mask As Long, ByRef Port_Status As COM_PORT_STATUS) As Long

Private Declare PtrSafe Function Com_Port_Create Lib "Kernel32.dll" Alias "CreateFileA" _
(ByVal Port_Name As String, ByVal PORT_ACCESS As Long, ByVal SHARE_MODE As Long, ByVal SECURITY_ATTRIBUTES_NULL As LongPtr, _
 ByVal CREATE_DISPOSITION As Long, ByVal FLAGS_AND_ATTRIBUTES As Long, ByVal TEMPLATE_FILE_HANDLE_NULL As LongPtr) As LongPtr

Private Declare PtrSafe Function Synchronous_Read Lib "Kernel32.dll" Alias "ReadFile" _
(ByVal Port_Handle As LongPtr, ByVal Buffer_Data As String, ByVal Bytes_Requested As Long, ByRef Bytes_Processed As Long, ByVal Overlapped_Null As LongPtr) As Long

Private Declare PtrSafe Function Synchronous_Write Lib "Kernel32.dll" Alias "WriteFile" _
(ByVal Port_Handle As LongPtr, ByVal Buffer_Data As String, ByVal Bytes_Requested As Long, ByRef Bytes_Processed As Long, ByVal Overlapped_Null As LongPtr) As Long

' Byte-oriented aliases of the same APIs for binary-safe transport (TRANSMIT_BYTES / RECEIVE_BYTES).
' ByRef ... As Any passes the address of the first byte - no ANSI/Unicode string conversion occurs.

Private Declare PtrSafe Function Synchronous_Read_Bytes Lib "Kernel32.dll" Alias "ReadFile" _
(ByVal Port_Handle As LongPtr, ByRef Buffer_Byte As Any, ByVal Bytes_Requested As Long, ByRef Bytes_Processed As Long, ByVal Overlapped_Null As LongPtr) As Long

Private Declare PtrSafe Function Synchronous_Write_Bytes Lib "Kernel32.dll" Alias "WriteFile" _
(ByVal Port_Handle As LongPtr, ByRef Buffer_Byte As Any, ByVal Bytes_Requested As Long, ByRef Bytes_Processed As Long, ByVal Overlapped_Null As LongPtr) As Long
'

Public Function START_COM_PORT(ByVal Port_Number As Long, Optional Port_Setttings As String) As Boolean

' Port_Settings if supplied should have the same structure as the equivalent command-line Mode arguments for a COM Port:
' [baud=b][parity=p][data=d][stop=s][to={on|off}][xon={on|off}][odsr={on|off}][octs={on|off}][dtr={on|off|hs}][rts={on|off|hs|tg}][idsr={on|off}]
' For example, to configure a baud rate of 1200, no parity, 8 data bits, and 1 stop bit, Port_Settings text is "baud=1200 parity=N data=8 stop=1"

Dim Temp_Result As Boolean, Port_Entered As Boolean, Opened_Here As Boolean
Dim Saved_DLL_Error As Long

Err.Clear                                                      ' a stale error from the caller must not be reported as a start failure

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                       ' non-blocking per-port operation guard

  If Port_Entered Then

    On Error GoTo Clean_Exit

    If Not STRUCTURES_VALID() Then

        Temp_Result = False                                ' never marshal structures the compiler laid out wrongly

    ElseIf COM_PORT_CLOSED(Port_Number) Then

        If OPEN_COM_PORT(Port_Number) Then

            Opened_Here = True                                 ' this call owns the handle until the start succeeds or fails

            If CONFIGURE_COM_PORT(Port_Number, Port_Setttings) Then

                Temp_Result = True
                PURGE_BUFFERS Port_Number

            Else

                Temp_Result = False
                Saved_DLL_Error = COM_PORT(Port_Number).DLL_Error   ' keep the configure error - closing the handle overwrites it
                STOP_PORT_CORE Port_Number, LONG_0             ' close com port if configure failed
                COM_PORT(Port_Number).DLL_Error = Saved_DLL_Error

            End If

        End If

    End If

  End If

End If

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then

    Temp_Result = False                                        ' a trapped runtime error is reported as failure

    If Opened_Here Then ABANDON_OPENED_PORT Port_Number        ' this call opened the port - close it after a failed start

End If

START_COM_PORT = Temp_Result

End Function

Private Function PORT_ENTER(Port_Number As Long) As Boolean

' Non-blocking per-port operation guard.
'
' Every public port operation yields to Windows through DoEvents, so a timer, event handler, Ribbon
' callback or nested macro can call back into this module while an earlier operation is still active.
' Re-entering the same port would let the second call overwrite buffers, timers, counts and handles
' belonging to the first. PORT_ENTER claims the port for one operation and returns False if it is
' already claimed - the caller then rejects the operation instead of corrupting the one in progress.
'
' The guard is independent per port: a busy port never blocks a call on another port. A long nested
' operation on one port does, however, suspend an outer operation on any port while it yields.
' VBA yields cooperatively (only at DoEvents), so this test-and-set cannot be interrupted part way.
'
' Public guarded functions must NOT call other public guarded functions for the same port -
' they call the private core routines (STOP_PORT_CORE, TRANSMIT_PORT_CORE) instead.

Dim Temp_Result As Boolean

If Port_Valid(Port_Number) Then

    If Not COM_PORT(Port_Number).Port_Busy Then

        COM_PORT(Port_Number).Port_Busy = True
        COM_PORT(Port_Number).Cancel_Requested = False     ' each operation starts with a clean cancel state
        Temp_Result = True

    End If

End If

PORT_ENTER = Temp_Result

End Function

Private Sub PORT_LEAVE(Port_Number As Long)

' Releases the per-port operation guard. Called from the single cleanup path of every guarded function.

If Port_Valid(Port_Number) Then

    COM_PORT(Port_Number).Cancel_Requested = False
    COM_PORT(Port_Number).Port_Busy = False

End If

End Sub

Public Function CANCEL_COM_PORT(Port_Number As Long) As Boolean

' Requests cooperative cancellation of the operation currently running on this port.
' The active operation unwinds cleanly at its next check and returns a failure or partial result.
' A transmit that had already written every byte when the request arrived still returns True.
' Returns True if an operation was active and the request was recorded; False if the port was idle.

If Port_Valid(Port_Number) Then

    If COM_PORT(Port_Number).Port_Busy Then

        COM_PORT(Port_Number).Cancel_Requested = True
        CANCEL_COM_PORT = True

    End If

End If

End Function

Private Function OPERATION_INTERRUPTED(Port_Number As Long, Expected_Handle As LongPtr) As Boolean

' Post-yield state verification, checked after every DoEvents inside a long operation:
' cancel requested, or the port handle no longer matches the one the operation started with.

If COM_PORT(Port_Number).Cancel_Requested Then

    OPERATION_INTERRUPTED = True

ElseIf COM_PORT(Port_Number).Handle <> Expected_Handle Then

    OPERATION_INTERRUPTED = True

End If

End Function

Public Function ABI_SELF_TEST(Optional ByRef Detail As String) As Boolean

' Verifies that the private structures marshalled to Win32 have the exact documented sizes on this
' host and Office bitness: DCB = 28, COMSTAT = 12, COMMTIMEOUTS = 20 bytes.
' Run automatically before the first START_COM_PORT of the session.

Dim Test_DCB As DEVICE_CONTROL_BLOCK
Dim Test_STAT As COM_PORT_STATUS
Dim Test_TIME As COM_PORT_TIMEOUTS

Const EXPECTED_DCB As Long = 28
Const EXPECTED_STAT As Long = 12
Const EXPECTED_TIME As Long = 20

Const Detail_Pass As String = "PASS: DCB=28, COMSTAT=12, COMMTIMEOUTS=20"
Const Detail_Fail As String = "FAIL: structure size mismatch - "

If LenB(Test_DCB) <> EXPECTED_DCB Then
    Detail = Detail_Fail & "DCB LenB = " & LenB(Test_DCB) & ", expected " & EXPECTED_DCB
ElseIf LenB(Test_STAT) <> EXPECTED_STAT Then
    Detail = Detail_Fail & "COMSTAT LenB = " & LenB(Test_STAT) & ", expected " & EXPECTED_STAT
ElseIf LenB(Test_TIME) <> EXPECTED_TIME Then
    Detail = Detail_Fail & "COMMTIMEOUTS LenB = " & LenB(Test_TIME) & ", expected " & EXPECTED_TIME
Else
    Detail = Detail_Pass
    ABI_SELF_TEST = True
End If

End Function

Private Function STRUCTURES_VALID() As Boolean

' Cached wrapper around ABI_SELF_TEST - the structure layout cannot change within a session.

If Not ABI_Checked Then

    ABI_Valid = ABI_SELF_TEST()
    ABI_Checked = True

End If

STRUCTURES_VALID = ABI_Valid

End Function

Public Function Port_Busy(Port_Number As Long) As Boolean

' True while an operation on this port is in progress, e.g. when tested from an event or timer callback.

Dim Temp_Result As Boolean

If Port_Valid(Port_Number) Then Temp_Result = COM_PORT(Port_Number).Port_Busy

Port_Busy = Temp_Result

End Function

Public Function CLEAR_PORT_BUSY(Port_Number As Long) As Boolean

' Recovery escape hatch: clears the per-port operation guard. It helps in break mode only, when an
' operation interrupted by the VBE (Ctrl-Break) will not be resumed: call this from the Immediate window,
' then stop the port. To resume the operation instead, just continue - it releases the guard itself.
' A VBE Reset or End discards the module-level state, losing the stored handle while the operating
' system keeps the port open - close the host application to release it. Never call this if the
' interrupted operation may still resume.

If Port_Valid(Port_Number) Then

    COM_PORT(Port_Number).Port_Busy = False
    CLEAR_PORT_BUSY = True

End If

End Function

Public Function PORT_LAST_ERROR(Port_Number As Long) As Long

' Returns the Win32 error code recorded by the most recent operation on this port, 0 if it succeeded, or -1
' for an invalid port number. Use it to tell a failed operation from an empty result, for example a failed
' read from no data. Read it straight after the call: the next operation replaces it, including a modem
' status query such as DEVICE_READY on an idle port. A status query while the port is busy records nothing.
' A call rejected because the port is busy or stopped makes no native call and leaves this value unchanged.
' Trapped VBA runtime errors are not included.

If Port_Valid(Port_Number) Then

    PORT_LAST_ERROR = COM_PORT(Port_Number).DLL_Error

Else

    PORT_LAST_ERROR = LONG_NEG_1

End If

End Function

Public Function PORT_ERROR_HISTORY(Port_Number As Long, Optional Clear_History As Boolean) As Long

' Returns the OR-accumulated communications-error mask (CE_ bits: 1=overflow, 2=overrun, 4=parity,
' 8=frame, 16=break) collected by every status poll since the history was last cleared. Status
' polling (wait/check functions) necessarily calls ClearCommError, which consumes the driver's
' current error flags - this history preserves that evidence. Returns -1 for an invalid port number.

If Port_Valid(Port_Number) Then

    PORT_ERROR_HISTORY = COM_PORT(Port_Number).Error_History

    If Clear_History Then COM_PORT(Port_Number).Error_History = LONG_0

Else

    PORT_ERROR_HISTORY = LONG_NEG_1

End If

End Function

Public Function TRANSMIT_BYTES(ByVal Port_Number As Long, Transmit_Data() As Byte) As Boolean

' Binary-safe transmit: sends every byte of Transmit_Data exactly as supplied - no string
' conversion, no code-page translation, embedded NUL and all values 0-255 are preserved.
' Returns True only if the entire array was written. An empty (unallocated) array returns False.

Dim Temp_Result As Boolean, Api_Result As Boolean
Dim Port_Entered As Boolean, Interrupted As Boolean
Dim Transmit_Handle As LongPtr
Dim Data_Length As Long, Total_Bytes_Sent As Long, Bytes_Remaining As Long
Dim Chunk_Bytes As Long, Chunk_Bytes_Sent As Long, Stalled_Writes As Long

Const MAX_STALLED_WRITES As Long = 3

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                   ' non-blocking per-port operation guard

  If Port_Entered Then

    On Error GoTo Clean_Exit

    If Port_Ready(Port_Number) Then

        Data_Length = BYTE_ARRAY_LENGTH(Transmit_Data)

        If Data_Length > LONG_0 Then

        Transmit_Handle = COM_PORT(Port_Number).Handle     ' verified after every yield

        Do

            Bytes_Remaining = Data_Length - Total_Bytes_Sent
            Chunk_Bytes = IIf(Bytes_Remaining < COM_PORT(Port_Number).Timers.Timeslice_Bytes Or COM_PORT(Port_Number).Timers.Timeslice_Bytes < LONG_1, Bytes_Remaining, COM_PORT(Port_Number).Timers.Timeslice_Bytes)

            Chunk_Bytes_Sent = LONG_0                      ' never assess a write using a stale byte count

            Api_Result = (Synchronous_Write_Bytes(Transmit_Handle, Transmit_Data(LBound(Transmit_Data) + Total_Bytes_Sent), Chunk_Bytes, Chunk_Bytes_Sent, NULL_POINTER) <> LONG_0)
            COM_PORT(Port_Number).DLL_Error = IIf(Api_Result, LONG_0, Err.LastDllError)

            If Not Api_Result Then Exit Do                 ' fail fast - a later chunk must not mask an earlier failure

            If Chunk_Bytes_Sent < LONG_0 Or Chunk_Bytes_Sent > Chunk_Bytes Then Exit Do  ' native count out of contract

            Total_Bytes_Sent = Total_Bytes_Sent + Chunk_Bytes_Sent

            If Chunk_Bytes_Sent > LONG_0 Then Stalled_Writes = LONG_0 Else Stalled_Writes = Stalled_Writes + LONG_1

            If Total_Bytes_Sent >= Data_Length Then Exit Do
            If Stalled_Writes >= MAX_STALLED_WRITES Then Exit Do

            DoEvents

            Interrupted = OPERATION_INTERRUPTED(Port_Number, Transmit_Handle)

            If Interrupted Then Exit Do                    ' cancelled or port restarted - stop touching the port

            If BYTE_ARRAY_LENGTH(Transmit_Data) <> Data_Length Then Interrupted = True: Exit Do   ' caller's array was resized during the yield

        Loop

        Temp_Result = Api_Result And (Total_Bytes_Sent = Data_Length) And Not Interrupted

        COM_PORT(Port_Number).Buffers.Transmit_Length = Total_Bytes_Sent

        End If

    End If

  End If

End If

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then Temp_Result = False           ' a trapped runtime error is reported as failure

TRANSMIT_BYTES = Temp_Result

End Function

Public Function RECEIVE_BYTES(ByVal Port_Number As Long, ByRef Receive_Data() As Byte, Optional ByVal Max_Bytes As Long = LONG_0, Optional ByVal Total_MilliSeconds As Long = LONG_0) As Long

' Binary-safe receive: accumulates raw bytes into Receive_Data (redimensioned to the exact count)
' and returns the number of bytes received. No string conversion occurs - all values 0-255 and
' embedded NUL bytes are preserved. Returns 0 with an unallocated array when nothing was received;
' returns -1 for an invalid, stopped or busy port, or when the read failed before any data arrived.
' Ends on read inactivity, the optional Max_Bytes / total deadline bounds, a cooperative cancel,
' or read failure. Bytes already received are kept on all of these; after a trapped VBA runtime error
' the array is erased and -1 is returned. A busy port rejects the call and leaves the caller's array
' untouched. A fixed-size array is rejected with -1 before any data is consumed, and keeps its contents.
' Each read request is
' capped at the bytes still allowed, so Max_Bytes is exact.

Dim Api_Result As Boolean, Failed As Boolean, Array_Resizable As Boolean
Dim Port_Entered As Boolean, Interrupted As Boolean, Limit_Reached As Boolean
Dim Receive_Handle As LongPtr
Dim Scratch(LONG_0 To 4095) As Byte                        ' fixed read buffer, one timeslice maximum
Dim Receive_Start_Time As Currency, Last_Data_Time As Currency
Dim Read_Request As Long, This_Request As Long, Bytes_Read As Long, Total_Bytes As Long, Capacity As Long
Dim Copy_Index As Long, Result_Count As Long

Const INITIAL_CAPACITY As Long = 4096

Result_Count = LONG_NEG_1

On Error GoTo Clean_Exit                                   ' armed first - the array handling below can raise

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                   ' non-blocking per-port operation guard

  If Port_Entered Then

    ReDim Receive_Data(LONG_0 To LONG_0)                   ' probe: a fixed-size caller array raises here, before any data is consumed
    Erase Receive_Data                                     ' caller's array starts empty (a busy port leaves it untouched)
    Array_Resizable = True

    If Port_Ready(Port_Number) Then

        Receive_Handle = COM_PORT(Port_Number).Handle      ' verified after every yield

        Receive_Start_Time = GET_HOST_MICROSECONDS
        Last_Data_Time = Receive_Start_Time

        Read_Request = COM_PORT(Port_Number).Timers.Timeslice_Bytes
        If Read_Request < LONG_1 Then Read_Request = LONG_1
        If Read_Request > INITIAL_CAPACITY Then Read_Request = INITIAL_CAPACITY

        Result_Count = LONG_0

        Do

            Bytes_Read = LONG_0                            ' never assess a read using a stale byte count

            This_Request = Read_Request                    ' cap the request at the bytes still allowed
            If Max_Bytes > LONG_0 Then
                If Max_Bytes - Total_Bytes < This_Request Then This_Request = Max_Bytes - Total_Bytes
            End If

            Api_Result = (Synchronous_Read_Bytes(Receive_Handle, Scratch(LONG_0), This_Request, Bytes_Read, NULL_POINTER) <> LONG_0)
            COM_PORT(Port_Number).DLL_Error = IIf(Api_Result, LONG_0, Err.LastDllError)

            If Not Api_Result Then Failed = True: Exit Do

            If Bytes_Read < LONG_0 Or Bytes_Read > This_Request Then Failed = True: Exit Do   ' native count out of contract

            If Bytes_Read > LONG_0 Then

                Last_Data_Time = GET_HOST_MICROSECONDS

                If Total_Bytes + Bytes_Read > Capacity Then           ' grow by doubling, never discard received data

                    Capacity = IIf(Capacity < INITIAL_CAPACITY, INITIAL_CAPACITY, Capacity * LONG_2)
                    If Capacity < Total_Bytes + Bytes_Read Then Capacity = Total_Bytes + Bytes_Read
                    ReDim Preserve Receive_Data(LONG_0 To Capacity - LONG_1)

                End If

                For Copy_Index = LONG_0 To Bytes_Read - LONG_1
                    Receive_Data(Total_Bytes + Copy_Index) = Scratch(Copy_Index)
                Next Copy_Index

                Total_Bytes = Total_Bytes + Bytes_Read

            Else

                ' no data this pass - stop when the inactivity window expires
                If (GET_HOST_MICROSECONDS - Last_Data_Time) > COM_PORT(Port_Number).Timers.Read_Wait_Time Then Exit Do

                RECEIVE_SLEEP COM_PORT(Port_Number).Timers.Exit_Loop_Wait, Total_MilliSeconds, Receive_Start_Time

            End If

            If Max_Bytes > LONG_0 Then Limit_Reached = (Total_Bytes >= Max_Bytes)

            If Total_MilliSeconds > LONG_0 And Not Limit_Reached Then
                Limit_Reached = ((GET_HOST_MICROSECONDS - Receive_Start_Time) >= CCur(Total_MilliSeconds) * LONG_1000)
            End If

            If Limit_Reached Then Exit Do

            DoEvents

            Interrupted = OPERATION_INTERRUPTED(Port_Number, Receive_Handle)

            If Interrupted Then Exit Do                    ' cancelled or port restarted - keep what was received

        Loop

        If Failed And Total_Bytes = LONG_0 Then

            Result_Count = LONG_NEG_1                      ' failed before any data arrived

        Else

            Result_Count = Total_Bytes

        End If

        If Total_Bytes > LONG_0 Then

            ReDim Preserve Receive_Data(LONG_0 To Total_Bytes - LONG_1)   ' exact length for the caller

        Else

            Erase Receive_Data

        End If

    End If

  End If

Else

    Erase Receive_Data                                     ' invalid port - caller's array is still emptied

End If

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then

    Result_Count = LONG_NEG_1
    If Array_Resizable Then ERASE_BYTES Receive_Data       ' a rejected fixed-size array keeps its contents

End If

RECEIVE_BYTES = Result_Count

End Function

Private Sub ERASE_BYTES(Data() As Byte)

' Erase from inside an active error handler. A locked caller array would raise again, and that second
' error would escape to the caller, so this has its own handler.

On Error Resume Next

Erase Data

End Sub

Private Function BYTE_ARRAY_LENGTH(Data() As Byte) As Long

' Length of a Byte array, or 0 if the array is unallocated (LBound/UBound raise error 9 on an
' empty dynamic array - trapped here so callers never need On Error around a length check).

On Error GoTo Empty_Array

BYTE_ARRAY_LENGTH = UBound(Data) - LBound(Data) + LONG_1

Exit Function

Empty_Array:

BYTE_ARRAY_LENGTH = LONG_0

End Function

Private Function OPEN_COM_PORT(Port_Number As Long) As Boolean

Dim Temp_Name As String
Dim Temp_Handle As LongPtr
Dim Temp_Result As Boolean
Dim Device_Path As String

Const OPEN_EXISTING As Long = LONG_3
Const OPEN_EXCLUSIVE As Long = LONG_0
Const SYNCHRONOUS_MODE As Long = LONG_0

Const GENERIC_RW As Long = &HC0000000
Const DEVICE_PREFIX As String = "\\.\COM"
        
Device_Path = DEVICE_PREFIX & CStr(Port_Number)

Temp_Handle = Com_Port_Create(Device_Path, GENERIC_RW, OPEN_EXCLUSIVE, NULL_POINTER, OPEN_EXISTING, SYNCHRONOUS_MODE, NULL_POINTER)

COM_PORT(Port_Number).DLL_Error = IIf(Temp_Handle = HANDLE_INVALID, Err.LastDllError, LONG_0)

Select Case Temp_Handle

Case HANDLE_INVALID

    Temp_Result = False
    COM_PORT(Port_Number).Handle = LONG_0

Case Else

    Temp_Result = True
    COM_PORT(Port_Number).Handle = Temp_Handle
    COM_PORT(Port_Number).Error_History = LONG_0           ' a fresh session starts with clean error evidence

End Select

OPEN_COM_PORT = Temp_Result

End Function

Private Function CONFIGURE_COM_PORT(Port_Number As Long, Optional Port_Settings As String) As Boolean

Dim Temp_Result As Boolean
Dim Clean_Settings As String

Clean_Settings = CLEAN_PORT_SETTINGS(Port_Settings)

If SET_PORT_CONFIG(Port_Number, Clean_Settings) Then
    
    If SET_PORT_TIMERS(Port_Number) Then
        
        Temp_Result = SET_PORT_VALUES(Port_Number)
      
    End If
    
End If
     
CONFIGURE_COM_PORT = Temp_Result

End Function

Private Function SET_PORT_CONFIG(Port_Number As Long, Optional Port_Settings As String) As Boolean

Dim Temp_Build As Boolean
Dim Temp_Result As Boolean

With COM_PORT(Port_Number)

.DCB.LENGTH_DCB = LenB(.DCB)                      ' DCBlength must be set before first use

If Query_Port_DCB(.Handle, .DCB) <> LONG_0 Then

    .DLL_Error = LONG_0

    If Len(Port_Settings) > LONG_0 Then

        Temp_Build = (Build_Port_DCB(Port_Settings, .DCB) <> LONG_0)
        .DLL_Error = IIf(Temp_Build, LONG_0, Err.LastDllError)

        ' BuildCommDCB does not set DCBlength - SetCommState requires it to be the size of the structure.

        If Temp_Build Then .DCB.LENGTH_DCB = LenB(.DCB)

        If Temp_Build Then

            SET_DCB_SAFE_FLAGS Port_Number                         ' binary-safe: no NUL discard, no abort on error

            .DCB.BIT_FIELD = .DCB.BIT_FIELD Or LONG_1              ' fBinary - always on for Windows
            If .DCB.PARITY <> LONG_0 Then
                .DCB.BIT_FIELD = .DCB.BIT_FIELD Or LONG_2          ' fParity - check parity when parity is in use
            Else
                .DCB.BIT_FIELD = .DCB.BIT_FIELD And Not LONG_2
            End If

            Temp_Result = (Apply_Port_DCB(.Handle, .DCB) <> LONG_0)
            .DLL_Error = IIf(Temp_Result, LONG_0, Err.LastDllError)

        End If

    Else

        Temp_Result = True

        ' Existing settings are kept, but fErrorChar / fNull / fAbortOnError left by a previous user of the port
        ' are removed. A failure here must not fail the start.
        ' The DCB is re-read after a failure so the stored copy still matches what the driver holds.

        If (.DCB.BIT_FIELD And (HEX_400 Or HEX_800 Or HEX_4000)) <> LONG_0 Then

            SET_DCB_SAFE_FLAGS Port_Number

            If Apply_Port_DCB(.Handle, .DCB) = LONG_0 Then Query_Port_DCB .Handle, .DCB

        End If

    End If

Else

    .DLL_Error = Err.LastDllError
    Temp_Result = False
   
End If

End With

SET_PORT_CONFIG = Temp_Result

End Function

Private Sub ABANDON_OPENED_PORT(Port_Number As Long)

' Closes a port after a runtime error in the call that opened it. Not STOP_PORT_CORE - that yields.
' Has its own handler because an On Error statement inside a handler that is already active has no effect.
' The handle is kept if the close failed, unless it was already invalid, so a later stop can still try.

Dim Closed As Boolean

Const INVALID_HANDLE_ERROR As Long = 6

On Error Resume Next

If COM_PORT(Port_Number).Handle <> LONG_0 Then

    Closed = (Com_Port_Close(COM_PORT(Port_Number).Handle) <> LONG_0)
    COM_PORT(Port_Number).DLL_Error = IIf(Closed, LONG_0, Err.LastDllError)

    If Closed Or COM_PORT(Port_Number).DLL_Error = INVALID_HANDLE_ERROR Then COM_PORT(Port_Number).Handle = LONG_0

End If

End Sub

Private Sub SET_DCB_SAFE_FLAGS(Port_Number As Long)

' Clears fErrorChar (replace bytes with parity errors), fNull (discard NUL bytes) and fAbortOnError
' (abort reads and writes on a line error), which would otherwise silently corrupt binary data or abort
' every transfer after one line error. fErrorChar matters because the settings path turns on fParity.
' XON/XOFF characters are only defaulted (17/19) when the two are equal, e.g. both zero.

Const XON_DEFAULT As Byte = 17
Const XOFF_DEFAULT As Byte = 19

With COM_PORT(Port_Number).DCB

    .BIT_FIELD = .BIT_FIELD And Not (HEX_400 Or HEX_800 Or HEX_4000)

    If .CHAR_XON = .CHAR_XOFF Then

        .CHAR_XON = XON_DEFAULT
        .CHAR_XOFF = XOFF_DEFAULT

    End If

End With

End Sub

Private Function SET_PORT_VALUES(Port_Number As Long) As Boolean

' ------------------------------------------------------------------------
'  optional - can change wait_characters_nnnn to suit local requirements.
'
Const WAIT_CHARACTERS_EXIT As Long = 2                  ' characters
Const WAIT_CHARACTERS_CHAR As Long = 5
Const WAIT_CHARACTERS_DATA As Long = 20
Const WAIT_CHARACTERS_LINE As Long = 100
'
'  optional - can change read exit wait timers to suit local requirements.
'
Const READ_EXIT_TIMER_FAST As Long = 100000             ' MicroSeconds
Const READ_EXIT_TIMER_SLOW As Long = 500000
Const READ_EXIT_TIMER_ELSE As Long = 125000
'
' ------------------------------------------------------------------------

Dim Temp_Result As Boolean
Dim Timeslice_Bytes As Long, Wait_Cap As Long
Dim Bytes_Per_Second As Long
Dim Read_Buffer_Length As Long
Dim Frame_MicroSeconds As Single

Frame_MicroSeconds = GET_FRAME_TIME(Port_Number)

' GET_FRAME_TIME returns zero if the DCB has no usable baud rate - do not divide by it.

If Frame_MicroSeconds > 0 Then Bytes_Per_Second = Int(LONG_1 / Frame_MicroSeconds * LONG_1E6)

Read_Buffer_Length = Len(COM_PORT(Port_Number).Buffers.Read_Buffer)
Timeslice_Bytes = IIf(Bytes_Per_Second < Read_Buffer_Length, Bytes_Per_Second, Read_Buffer_Length)

If Read_Buffer_Length > LONG_0 And Bytes_Per_Second > LONG_0 Then

Temp_Result = True

With COM_PORT(Port_Number)

.Timers.Port_Data_Time = LONG_0
.Timers.Last_Data_Time = LONG_0
.Timers.Timeslice_Bytes = Timeslice_Bytes
.Timers.Bytes_Per_Second = Bytes_Per_Second
.Timers.Frame_MicroSeconds = Frame_MicroSeconds
.Timers.Frame_MilliSeconds = Frame_MicroSeconds / LONG_1000
.Buffers.Read_Buffer_Length = Read_Buffer_Length

.Timers.Exit_Loop_Wait = Int(LONG_1 + .Timers.Frame_MilliSeconds) * WAIT_CHARACTERS_EXIT
.Timers.Char_Loop_Wait = Int(LONG_1 + .Timers.Frame_MilliSeconds) * WAIT_CHARACTERS_CHAR
.Timers.Data_Loop_Wait = Int(LONG_1 + .Timers.Frame_MilliSeconds) * WAIT_CHARACTERS_DATA
.Timers.Line_Loop_Wait = Int(LONG_1 + .Timers.Frame_MilliSeconds) * WAIT_CHARACTERS_LINE

If .Timers.Exit_Loop_Wait > VBA_TIMEOUT / LONG_5 Then .Timers.Exit_Loop_Wait = LONG_1000
If .Timers.Char_Loop_Wait > VBA_TIMEOUT / LONG_5 Then .Timers.Char_Loop_Wait = LONG_1000
If .Timers.Data_Loop_Wait > VBA_TIMEOUT / LONG_5 Then .Timers.Data_Loop_Wait = LONG_1000
If .Timers.Line_Loop_Wait > VBA_TIMEOUT / LONG_5 Then .Timers.Line_Loop_Wait = LONG_1000

Wait_Cap = Int(LONG_2048 * LONG_1000 / Bytes_Per_Second)           ' about 2048 bytes of line time - keeps a yield close at high baud rates
If Wait_Cap < LONG_1 Then Wait_Cap = LONG_1

If .Timers.Exit_Loop_Wait > Wait_Cap Then .Timers.Exit_Loop_Wait = Wait_Cap
If .Timers.Char_Loop_Wait > Wait_Cap Then .Timers.Char_Loop_Wait = Wait_Cap
If .Timers.Data_Loop_Wait > Wait_Cap Then .Timers.Data_Loop_Wait = Wait_Cap
If .Timers.Line_Loop_Wait > Wait_Cap Then .Timers.Line_Loop_Wait = Wait_Cap

Select Case Bytes_Per_Second

    Case Is > LONG_1000: .Timers.Read_Wait_Time = READ_EXIT_TIMER_FAST
    Case Is < LONG_100:  .Timers.Read_Wait_Time = READ_EXIT_TIMER_SLOW
    Case Else:           .Timers.Read_Wait_Time = READ_EXIT_TIMER_ELSE

End Select

End With

Else   ' read buffer size not > 0

Temp_Result = False

End If

SET_PORT_VALUES = Temp_Result

End Function

Public Function STOP_COM_PORT(ByVal Port_Number As Long, Optional ByVal Drain_MilliSeconds As Long = LONG_0) As Boolean

' Drain_MilliSeconds of 0 or less (default 0) is an abort close - queued transmit and unread receive data are purged.
' Drain_MilliSeconds > 0 is a graceful close - waits up to that time for the driver's transmit queue to empty.
' If the driver still reports queued data, nothing is discarded: the port stays open and the result is False.
' Retry, or stop with no drain time to abort. A cancel request during the wait forces the abort close.
' The wait only covers what the driver reports. The count is documented as zero for a non-overlapped write,
' which is what this module uses, so the wait can end at once. After the wait only receive data is purged.
' True therefore does not prove delivery - have the device acknowledge if that matters.

Dim Temp_Result As Boolean, Port_Entered As Boolean

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                       ' non-blocking per-port operation guard

  If Port_Entered Then

    On Error GoTo Clean_Exit

    Temp_Result = STOP_PORT_CORE(Port_Number, Drain_MilliSeconds)

  Else

    ' Port busy: request cooperative cancellation so the owning operation unwinds. The caller retries
    ' the stop after returning from its callback (for example with Application.OnTime) - never in a loop
    ' inside the callback, which would starve the operation being cancelled. The port is never purged or
    ' closed underneath an active operation.
    CANCEL_COM_PORT Port_Number

  End If

End If

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then Temp_Result = False               ' a trapped runtime error is reported as failure

STOP_COM_PORT = Temp_Result

End Function

Private Function STOP_PORT_CORE(Port_Number As Long, Drain_MilliSeconds As Long) As Boolean

' unguarded stop - the caller must already hold the per-port operation guard.

Dim Temp_Result As Boolean, Close_Result As Boolean, Drained As Boolean

Const INVALID_HANDLE_ERROR As Long = 6

If Port_Ready(Port_Number) Then

    If Drain_MilliSeconds > LONG_0 Then Drained = DRAIN_PORT_OUTPUT(Port_Number, Drain_MilliSeconds)

    If Drain_MilliSeconds > LONG_0 And Not Drained And Not COM_PORT(Port_Number).Cancel_Requested Then

        Temp_Result = False                                    ' keep the queued data and the port - the caller retries or aborts

    Else

        If Drained Then PURGE_RECEIVE_QUEUE Port_Number Else PURGE_COM_PORT Port_Number

        Close_Result = (Com_Port_Close(COM_PORT(Port_Number).Handle) <> LONG_0)
        COM_PORT(Port_Number).DLL_Error = IIf(Close_Result, LONG_0, Err.LastDllError)

        If Close_Result Then

            Temp_Result = True

            COM_PORT(Port_Number).Handle = LONG_0

            PURGE_BUFFERS Port_Number

        ElseIf COM_PORT(Port_Number).DLL_Error = INVALID_HANDLE_ERROR Then

            COM_PORT(Port_Number).Handle = LONG_0              ' the handle is already gone - do not wedge the entry

        End If

    End If

End If

STOP_PORT_CORE = Temp_Result

End Function

Private Function PURGE_RECEIVE_QUEUE(Port_Number As Long) As Boolean

' Graceful close: clears unread receive data only. The transmit purge flags would discard any bytes the
' driver still holds after reporting an empty transmit queue.

Dim Temp_Result As Boolean

Const PURGE_RX_ABORT_AND_CLEAR As Long = LONG_10           ' PURGE_RXABORT (2) Or PURGE_RXCLEAR (8)

Temp_Result = (Com_Port_Purge(COM_PORT(Port_Number).Handle, PURGE_RX_ABORT_AND_CLEAR) <> LONG_0)
COM_PORT(Port_Number).DLL_Error = IIf(Temp_Result, LONG_0, Err.LastDllError)

PURGE_RECEIVE_QUEUE = Temp_Result

End Function

Private Function DRAIN_PORT_OUTPUT(Port_Number As Long, Drain_MilliSeconds As Long) As Boolean

' waits up to Drain_MilliSeconds for queued transmit data to leave the driver output queue.
' returns True if the output queue emptied within the deadline.

Dim Queue_Empty As Boolean, Clear_Result As Boolean
Dim Sleep_Time As Long
Dim Deadline As Currency, Time_Left As Currency

Const Loop_Time As Long = LONG_10                            ' MilliSeconds

Deadline = GET_HOST_MICROSECONDS + CCur(Drain_MilliSeconds) * LONG_1000   ' elapsed time - Sleep overruns and yields both count

With COM_PORT(Port_Number)

Do

    Clear_Result = (Com_Port_Clear(.Handle, .Errors, .Status) <> LONG_0)
    .DLL_Error = IIf(Clear_Result, LONG_0, Err.LastDllError)

    If Not Clear_Result Then Exit Do

    .Error_History = .Error_History Or .Errors             ' preserve error evidence consumed by this poll

    Queue_Empty = (.Status.QUEUE_OUT < LONG_1)

    If Queue_Empty Then Exit Do

    If .Cancel_Requested Then Exit Do          ' cancel promotes the graceful close to an abort

    Time_Left = Deadline - GET_HOST_MICROSECONDS

    If Time_Left <= LONG_0 Then Exit Do        ' deadline reached - the queue was polled once more after the last sleep

    Sleep_Time = Loop_Time
    If Time_Left < CCur(Loop_Time) * LONG_1000 Then Sleep_Time = Int(Time_Left / LONG_1000) + LONG_1

    Kernel_Sleep_MilliSeconds Sleep_Time

    DoEvents                                               ' yield so a cancel can arrive and Office stays responsive

Loop

End With

DRAIN_PORT_OUTPUT = Queue_Empty

End Function

Public Function WAIT_COM_PORT(ByVal Port_Number As Long, Optional ByVal Wait_MilliSeconds As Long = LONG_333) As Boolean

Dim Wait_Result As Boolean, Port_Entered As Boolean

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                       ' non-blocking per-port operation guard

  If Port_Entered Then

    On Error GoTo Clean_Exit

    If Port_Ready(Port_Number) Then Wait_Result = SYNCHRONOUS_WAIT_COM_PORT(Port_Number, Wait_MilliSeconds)

  End If

End If

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then Wait_Result = False               ' a trapped runtime error is reported as failure

WAIT_COM_PORT = Wait_Result

End Function

Private Function SYNCHRONOUS_WAIT_COM_PORT(Port_Number As Long, Wait_MilliSeconds As Long) As Boolean

Dim Wait_Remaining As Long, Sleep_Time As Long
Dim Data_Waiting As Boolean, Loop_Wait_Time As Long
Dim Wait_Expired As Boolean, Clear_Result As Boolean
Dim Wait_Handle As LongPtr, Interrupted As Boolean

Const Loop_Time As Long = LONG_100                          ' MilliSeconds

Wait_Remaining = IIf(Wait_MilliSeconds < LONG_1, LONG_1, Wait_MilliSeconds)
Loop_Wait_Time = IIf(Wait_MilliSeconds < Loop_Time, Wait_Remaining, Loop_Time)

Wait_Handle = COM_PORT(Port_Number).Handle                  ' verified after every yield

With COM_PORT(Port_Number)

Do

Clear_Result = (Com_Port_Clear(.Handle, .Errors, .Status) <> LONG_0)
.DLL_Error = IIf(Clear_Result, LONG_0, Err.LastDllError)

If Clear_Result Then

    .Error_History = .Error_History Or .Errors             ' preserve error evidence consumed by this poll

    Data_Waiting = .Status.QUEUE_IN > LONG_0

    If Not Data_Waiting Then

        Wait_Expired = Wait_Remaining < LONG_1

        If Not Wait_Expired Then

            Sleep_Time = IIf(Wait_Remaining < Loop_Wait_Time, Wait_Remaining, Loop_Wait_Time)

            Kernel_Sleep_MilliSeconds Sleep_Time

            Wait_Remaining = Wait_Remaining - Sleep_Time

        End If

    End If

Else

    Wait_Expired = True
    Data_Waiting = False

End If

DoEvents

Interrupted = OPERATION_INTERRUPTED(Port_Number, Wait_Handle)

Loop Until Data_Waiting Or Wait_Expired Or Not Clear_Result Or Interrupted

End With

If Interrupted Then Data_Waiting = False                    ' cancelled or port restarted - report no data

SYNCHRONOUS_WAIT_COM_PORT = Data_Waiting

End Function

Public Function READ_COM_PORT(ByVal Port_Number As Long, Optional ByVal Number_Characters As Long) As String

Dim Temp_Result As Boolean
Dim Read_Limit_Check As Boolean
Dim Read_Character_Count As Long
Dim Read_Character_String As String
Dim Port_Entered As Boolean

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                       ' non-blocking per-port operation guard

  If Port_Entered Then

    On Error GoTo Clean_Exit

    If Port_Ready(Port_Number) Then

    With COM_PORT(Port_Number)

        Read_Limit_Check = Number_Characters < LONG_1 Or Number_Characters > .Timers.Timeslice_Bytes

        Read_Character_Count = IIf(Read_Limit_Check, .Timers.Timeslice_Bytes, Number_Characters)

        Temp_Result = SYNCHRONOUS_READ_COM_PORT(Port_Number, Read_Character_Count)

        If Temp_Result And Not .Buffers.Read_Buffer_Empty Then Read_Character_String = .Buffers.Read_Result

    End With

    End If

  End If

End If

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then Read_Character_String = vbNullString

READ_COM_PORT = Read_Character_String

End Function

Public Function RECEIVE_COM_PORT(ByVal Port_Number As Long, Optional ByVal Max_Bytes As Long = LONG_0, Optional ByVal Total_MilliSeconds As Long = LONG_0) As String

' Max_Bytes = 0 (default) and Total_MilliSeconds = 0 (default) keep the original unbounded behaviour, where
' receive ends only on read inactivity (Read_Wait_Time) or read failure.
' Max_Bytes > 0 stops receiving once that many bytes have been accumulated. Each read request is capped at the
' bytes still allowed, so the limit is exact for single-byte character sets.
' Total_MilliSeconds > 0 stops receiving once that much time has elapsed since the call started.

Dim Full_Buffer As Long, Read_Request As Long
Dim Temp_Result As Boolean, Limit_Reached As Boolean, Port_Entered As Boolean, Interrupted As Boolean
Dim Receive_Start_Time As Currency
Dim Receive_Handle As LongPtr
Dim Receive_String As String

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                       ' non-blocking per-port operation guard

  If Port_Entered Then

    On Error GoTo Clean_Exit

    If Port_Ready(Port_Number) Then

    With COM_PORT(Port_Number)

    Full_Buffer = .Timers.Timeslice_Bytes

    Receive_Handle = .Handle                                ' verified after every yield

    Receive_Start_Time = GET_HOST_MICROSECONDS

    ' The inactivity window starts at this call, so a reply arriving shortly after a transmit is not missed.
    COM_PORT(Port_Number).Timers.Port_Data_Time = Receive_Start_Time

    .Buffers.Receive_Result = vbNullString

    Do
        Do
            Read_Request = Full_Buffer                              ' cap the request at the bytes still allowed
            If Max_Bytes > LONG_0 Then
                If Max_Bytes - Len(.Buffers.Receive_Result) < Read_Request Then Read_Request = Max_Bytes - Len(.Buffers.Receive_Result)
            End If
            If Read_Request < LONG_1 Then Read_Request = LONG_1

            Temp_Result = SYNCHRONOUS_READ_COM_PORT(Port_Number, Read_Request)

            If Temp_Result And Not .Buffers.Read_Buffer_Empty Then

                .Buffers.Receive_Result = .Buffers.Receive_Result & .Buffers.Read_Result

                Limit_Reached = RECEIVE_LIMIT_REACHED(Port_Number, Max_Bytes, Total_MilliSeconds, Receive_Start_Time)

                If Not Limit_Reached Then                            ' limit hit - skip the wait

                    Select Case .Buffers.Synchronous_Bytes_Read

                        Case Is < LONG_4:       RECEIVE_SLEEP .Timers.Char_Loop_Wait, Total_MilliSeconds, Receive_Start_Time
                        Case Is < LONG_21:      RECEIVE_SLEEP .Timers.Data_Loop_Wait, Total_MilliSeconds, Receive_Start_Time
                        Case Is = Full_Buffer ' Timeslice buffer full, no waiting, more expected
                        Case Else:              RECEIVE_SLEEP .Timers.Line_Loop_Wait, Total_MilliSeconds, Receive_Start_Time

                    End Select

                    DoEvents

                    Interrupted = OPERATION_INTERRUPTED(Port_Number, Receive_Handle)

                    Limit_Reached = RECEIVE_LIMIT_REACHED(Port_Number, Max_Bytes, Total_MilliSeconds, Receive_Start_Time)

                End If

            End If

        Loop Until .Buffers.Read_Buffer_Empty Or Not Temp_Result Or Limit_Reached Or Interrupted

        If Not .Timers.Read_Timeout And Not Limit_Reached And Not Interrupted Then RECEIVE_SLEEP .Timers.Exit_Loop_Wait, Total_MilliSeconds, Receive_Start_Time

        Limit_Reached = RECEIVE_LIMIT_REACHED(Port_Number, Max_Bytes, Total_MilliSeconds, Receive_Start_Time)

        If Not Interrupted Then Interrupted = OPERATION_INTERRUPTED(Port_Number, Receive_Handle)

     Loop Until .Timers.Read_Timeout Or Not Temp_Result Or Limit_Reached Or Interrupted

     Receive_String = .Buffers.Receive_Result

     End With

     End If

  End If

End If

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

RECEIVE_COM_PORT = Receive_String              ' local result - never index COM_PORT with an unvalidated port number

End Function

Private Function RECEIVE_LIMIT_REACHED(Port_Number As Long, Max_Bytes As Long, Total_MilliSeconds As Long, Start_Time As Currency) As Boolean

' returns True when an optional receive bound (maximum bytes or total elapsed time) has been reached.

Dim Limit_Hit As Boolean

If Max_Bytes > LONG_0 Then Limit_Hit = (Len(COM_PORT(Port_Number).Buffers.Receive_Result) >= Max_Bytes)

If Total_MilliSeconds > LONG_0 And Not Limit_Hit Then

    Limit_Hit = ((GET_HOST_MICROSECONDS - Start_Time) >= CCur(Total_MilliSeconds) * LONG_1000)

End If

RECEIVE_LIMIT_REACHED = Limit_Hit

End Function

Private Sub RECEIVE_SLEEP(ByVal Wait_MilliSeconds As Long, ByVal Total_MilliSeconds As Long, Start_Time As Currency)

' Receive loop wait. With a Total_MilliSeconds bound the wait is clamped to the time left, so a long
' loop wait can no longer overshoot the bound; without a bound the wait is exactly as requested.

Dim Time_Left As Currency

If Total_MilliSeconds > LONG_0 Then

    Time_Left = CCur(Total_MilliSeconds) - (GET_HOST_MICROSECONDS - Start_Time) / LONG_1000

    If Time_Left < LONG_1 Then
        Wait_MilliSeconds = LONG_0
    ElseIf Time_Left < Wait_MilliSeconds Then
        Wait_MilliSeconds = Int(Time_Left)
    End If

End If

Kernel_Sleep_MilliSeconds Wait_MilliSeconds

End Sub

Public Function TRANSMIT_COM_PORT(ByVal Port_Number As Long, Transmit_Text As String) As Boolean

Dim Write_Result As Boolean, Port_Entered As Boolean

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                       ' non-blocking per-port operation guard

  If Port_Entered Then

    On Error GoTo Clean_Exit

    Write_Result = TRANSMIT_PORT_CORE(Port_Number, Transmit_Text)

  End If

End If

DoEvents                                                   ' yields while the guard is held, so a status query cannot reset the last error before the caller reads it

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then Write_Result = False              ' a trapped runtime error is reported as failure

TRANSMIT_COM_PORT = Write_Result

End Function

Private Function TRANSMIT_PORT_CORE(Port_Number As Long, ByVal Transmit_Text As String) As Boolean

' unguarded transmit - the caller must already hold the per-port operation guard.
' the result is True only if every chunk was written and the total bytes sent equals the requested length.

Dim Loop_Counter As Long, Transmit_Length As Long
Dim Timeslice_Bytes As Long, Total_Bytes_Sent As Long
Dim Write_Result As Boolean, Chunk_Result As Boolean, Interrupted As Boolean
Dim Transmit_Handle As LongPtr

If Port_Ready(Port_Number) Then

  With COM_PORT(Port_Number)

    Transmit_Length = Len(Transmit_Text)
    Timeslice_Bytes = .Timers.Timeslice_Bytes
    Transmit_Handle = .Handle                  ' verified after every yield

    If Timeslice_Bytes > LONG_0 And ANSI_SAFE(Transmit_Text) Then   ' a zero step would loop forever; text must be one ANSI byte per character

    For Loop_Counter = LONG_1 To Transmit_Length Step Timeslice_Bytes

        .Buffers.Write_Buffer = Mid$(Transmit_Text, Loop_Counter, Timeslice_Bytes)

        Chunk_Result = SYNCHRONOUS_WRITE_COM_PORT(Port_Number)

        Total_Bytes_Sent = Total_Bytes_Sent + .Buffers.Synchronous_Bytes_Sent

        If Not Chunk_Result Then Exit For      ' fail fast - a later chunk must not mask an earlier failure

        DoEvents

        Interrupted = OPERATION_INTERRUPTED(Port_Number, Transmit_Handle)

        If Interrupted Then Exit For           ' cancelled or port restarted - stop touching the port

    Next Loop_Counter

    ' A cancel that arrives after the last byte was written does not make a complete write a failure.
    Write_Result = Chunk_Result And (Total_Bytes_Sent = Transmit_Length) And (Transmit_Length > LONG_0)

    .Buffers.Transmit_Length = Total_Bytes_Sent

    End If

  End With

End If

TRANSMIT_PORT_CORE = Write_Result

End Function

Private Function ANSI_SAFE(Text_Value As String) As Boolean

' True if the text converts to one ANSI byte per character. The write loops count characters as bytes,
' so on a double-byte code page text that converts to more bytes than characters would be truncated.
' This does not detect substitution: on a single-byte code page an unmappable character still becomes
' one byte (for example "?"), so it passes. Use TRANSMIT_BYTES when exact bytes matter.

ANSI_SAFE = (LenB(StrConv(Text_Value, vbFromUnicode)) = Len(Text_Value))

End Function

Private Function GET_FRAME_TIME(Port_Number As Long) As Single

' Returns the frame (character) duration in MicroSeconds, or zero if the DCB has no usable baud rate.
' Stop bits are held as a Double because the DCB value 1 means 1.5 stop bits, not 2.

Dim Length_Data As Long
Dim Length_Stop As Double
Dim Length_Start As Long
Dim Frame_Length As Double
Dim Length_Parity As Long
Dim Frame_Duration As Single

With COM_PORT(Port_Number)

    Length_Start = LONG_1
    Length_Data = .DCB.BYTE_SIZE
    Length_Parity = IIf(.DCB.PARITY = LONG_0, LONG_0, LONG_1)

    Select Case .DCB.STOP_BITS

    Case LONG_0:    Length_Stop = 1#          ' 1 stop bit
    Case LONG_1:    Length_Stop = 1.5         ' 1.5 stop bits
    Case LONG_2:    Length_Stop = 2#          ' 2 stop bits

    Case Else:      Length_Stop = 1#

    End Select

    Frame_Length = Length_Start + Length_Data + Length_Parity + Length_Stop

    If .DCB.BAUD_RATE > LONG_0 Then Frame_Duration = Frame_Length / .DCB.BAUD_RATE * LONG_1E6

End With

GET_FRAME_TIME = Frame_Duration

End Function

Public Function GET_PORT_SETTINGS(Port_Number As Long) As String

Dim Port_Settings As String

Const TEXT_DASH As String = "-"
Const TEXT_PORT_INVALID As String = "INVALID-PORT"
Const TEXT_NOT_STARTED As String = "PORT-NOT-STARTED"
Const TEXT_SETTINGS_ERROR As String = "ERROR-READING-SETTINGS"

On Error GoTo Clean_Exit

If Port_Valid(Port_Number) Then

    With COM_PORT(Port_Number)

        If .Handle > LONG_0 Then

            Port_Settings = vbNullString
            Port_Settings = Port_Settings & .DCB.BAUD_RATE & TEXT_DASH
            Port_Settings = Port_Settings & .DCB.BYTE_SIZE & TEXT_DASH
            Port_Settings = Port_Settings & CONVERT_PARITY(.DCB.PARITY) & TEXT_DASH
            Port_Settings = Port_Settings & CONVERT_STOPBITS(.DCB.STOP_BITS)

        Else

            Port_Settings = TEXT_NOT_STARTED

        End If
    
    End With

Else

    Port_Settings = TEXT_PORT_INVALID

End If

Clean_Exit:

If Err.Number <> LONG_0 Then Port_Settings = TEXT_SETTINGS_ERROR

GET_PORT_SETTINGS = Port_Settings

End Function

Private Function SYNCHRONOUS_READ_COM_PORT(Port_Number As Long, Read_Bytes_Requested As Long) As Boolean

Dim Temp_Result As Boolean, Read_Raw As String
Dim Read_Bytes(LONG_0 To 4095) As Byte                    ' same size as the fixed read buffer (maximum timeslice bytes)
Dim Read_Count As Long

With COM_PORT(Port_Number)

.Buffers.Synchronous_Bytes_Read = LONG_0        ' never assess a read using a stale byte count

Read_Count = Read_Bytes_Requested
If Read_Count > 4096 Then Read_Count = 4096

' Read raw bytes (no ANSI marshalling of a fixed string), then convert only the bytes received.

Temp_Result = (Synchronous_Read_Bytes(.Handle, Read_Bytes(LONG_0), Read_Count, .Buffers.Synchronous_Bytes_Read, NULL_POINTER) <> LONG_0)
.DLL_Error = IIf(Temp_Result, LONG_0, Err.LastDllError)

If Temp_Result Then

    If .Buffers.Synchronous_Bytes_Read = LONG_0 Then
     
        .Timers.Last_Data_Time = GET_HOST_MICROSECONDS - .Timers.Port_Data_Time
        .Timers.Read_Timeout = (.Timers.Last_Data_Time > .Timers.Read_Wait_Time)
        .Buffers.Read_Result = vbNullString
        .Buffers.Read_Buffer_Empty = True
    
    Else
        
        .Timers.Port_Data_Time = GET_HOST_MICROSECONDS
        .Timers.Last_Data_Time = LONG_0
        .Timers.Read_Timeout = False
        Read_Raw = Read_Bytes
        .Buffers.Read_Result = StrConv(LeftB$(Read_Raw, .Buffers.Synchronous_Bytes_Read), vbUnicode)
        .Buffers.Read_Buffer_Empty = False
        
    End If

Else

     Temp_Result = False
    .Timers.Read_Timeout = True
    .Buffers.Read_Buffer_Empty = True
    .Buffers.Read_Result = vbNullString
      
End If

End With

DoEvents

SYNCHRONOUS_READ_COM_PORT = Temp_Result

End Function

Private Function SYNCHRONOUS_WRITE_COM_PORT(Port_Number As Long) As Boolean

' The result is True only if the native call reported success AND every requested byte was written.
' A short write is retried for the remaining bytes, so a partial write is never reported as success.

Dim Api_Result As Boolean
Dim Write_Complete As Boolean
Dim Write_Chunk As String
Dim Write_Buffer_Length As Long, Bytes_Remaining As Long
Dim Total_Bytes_Sent As Long, Chunk_Bytes_Sent As Long, Stalled_Writes As Long
Dim Write_Handle As LongPtr

Const MAX_STALLED_WRITES As Long = 3

With COM_PORT(Port_Number)

    Write_Buffer_Length = Len(.Buffers.Write_Buffer)
    Write_Handle = .Handle                      ' verified after every yield

    If Write_Buffer_Length > LONG_0 Then

    Do

        Bytes_Remaining = Write_Buffer_Length - Total_Bytes_Sent

        Write_Chunk = Mid$(.Buffers.Write_Buffer, Total_Bytes_Sent + LONG_1, Bytes_Remaining)

        Chunk_Bytes_Sent = LONG_0               ' never assess a write using a stale byte count

        Api_Result = (Synchronous_Write(.Handle, Write_Chunk, Bytes_Remaining, Chunk_Bytes_Sent, NULL_POINTER) <> LONG_0)
        .DLL_Error = IIf(Api_Result, LONG_0, Err.LastDllError)

        If Not Api_Result Then Exit Do

        Total_Bytes_Sent = Total_Bytes_Sent + Chunk_Bytes_Sent

        If Chunk_Bytes_Sent > LONG_0 Then Stalled_Writes = LONG_0 Else Stalled_Writes = Stalled_Writes + LONG_1

        If Total_Bytes_Sent >= Write_Buffer_Length Or Stalled_Writes >= MAX_STALLED_WRITES Then Exit Do

        DoEvents                                ' each stalled write can block for the full write timeout

        If OPERATION_INTERRUPTED(Port_Number, Write_Handle) Then Exit Do

    Loop

    Write_Complete = Api_Result And (Total_Bytes_Sent = Write_Buffer_Length)

    End If

    .Buffers.Synchronous_Bytes_Sent = Total_Bytes_Sent

End With

DoEvents

SYNCHRONOUS_WRITE_COM_PORT = Write_Complete

End Function

Public Function SEND_COM_PORT(ByVal Port_Number As Long, Send_Variable As Variant) As Boolean

Dim Send_Result As Boolean, Port_Entered As Boolean, Send_Value As Variant

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                       ' non-blocking per-port operation guard

  If Port_Entered Then

    On Error GoTo Clean_Exit

    If Port_Ready(Port_Number) Then

        Send_Value = Send_Variable                             ' a Range yields its value; arrays and error values cannot be sent as text

        If Not (IsError(Send_Value) Or IsArray(Send_Value)) Then Send_Result = TRANSMIT_PORT_CORE(Port_Number, CStr(Send_Value))

    End If

  End If

End If

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then Send_Result = False               ' a trapped runtime error is reported as failure

SEND_COM_PORT = Send_Result

End Function

Public Function PUT_COM_PORT(ByVal Port_Number As Long, Put_Character As String) As Boolean

' Writes exactly one byte. The result is True only if the native call succeeded and one byte was written.
' An empty Put_Character is rejected - it must not be sent to the port as a single NUL byte.

Dim Write_Result As Boolean, Api_Result As Boolean, Port_Entered As Boolean
Dim Write_Byte_Count As Long

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                       ' non-blocking per-port operation guard

  If Port_Entered Then

    On Error GoTo Clean_Exit

    If Port_Ready(Port_Number) And Len(Put_Character) > LONG_0 Then

      If ANSI_SAFE(Left$(Put_Character, LONG_1)) Then

        Api_Result = (Synchronous_Write(COM_PORT(Port_Number).Handle, Left$(Put_Character, LONG_1), LONG_1, Write_Byte_Count, NULL_POINTER) <> LONG_0)
        COM_PORT(Port_Number).DLL_Error = IIf(Api_Result, LONG_0, Err.LastDllError)

        Write_Result = Api_Result And (Write_Byte_Count = LONG_1)

      End If

    End If

  End If

End If

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then Write_Result = False              ' a trapped runtime error is reported as failure

PUT_COM_PORT = Write_Result

End Function

Public Function GET_COM_PORT(ByVal Port_Number As Long) As String

' Reads at most one byte. An empty result means no data OR a failed read - PORT_LAST_ERROR tells the two
' cases apart (0 means the read succeeded and there was no data). The byte is only returned when the
' native call succeeded and reported exactly one byte.

Dim Read_Byte_Count As Long
Dim Api_Result As Boolean, Port_Entered As Boolean
Dim Get_Character As String
Dim Read_Buffer As String * LONG_1  ' must be fixed length 1

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                       ' non-blocking per-port operation guard

  If Port_Entered Then

    On Error GoTo Clean_Exit

    If Port_Ready(Port_Number) Then

        Api_Result = (Synchronous_Read(COM_PORT(Port_Number).Handle, Read_Buffer, LONG_1, Read_Byte_Count, NULL_POINTER) <> LONG_0)
        COM_PORT(Port_Number).DLL_Error = IIf(Api_Result, LONG_0, Err.LastDllError)

        If Api_Result And Read_Byte_Count = LONG_1 Then Get_Character = Read_Buffer

    End If

  End If

End If

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then Get_Character = vbNullString

GET_COM_PORT = Get_Character

End Function

Private Function PURGE_COM_PORT(Port_Number As Long) As Boolean

Dim Temp_Result As Boolean

Const HEX_0F As Byte = &HF
Const PURGE_ALL As Long = HEX_0F

Temp_Result = (Com_Port_Purge(COM_PORT(Port_Number).Handle, PURGE_ALL) <> LONG_0)
COM_PORT(Port_Number).DLL_Error = IIf(Temp_Result, LONG_0, Err.LastDllError)

DoEvents

PURGE_COM_PORT = Temp_Result

End Function

Private Sub PURGE_BUFFERS(Port_Number As Long)

With COM_PORT(Port_Number).Buffers

    .Read_Result = vbNullString
    .Read_Buffer = vbNullString
    .Write_Result = vbNullString
    .Write_Buffer = vbNullString
    .Receive_Result = vbNullString
    .Receive_Buffer = vbNullString
    .Receive_Length = LONG_0
    .Transmit_Length = LONG_0
    .Transmit_Result = vbNullString
    .Transmit_Buffer = vbNullString
    .Read_Buffer_Empty = True
    .Synchronous_Bytes_Read = LONG_0
    .Synchronous_Bytes_Sent = LONG_0
   
End With

End Sub

Private Function SET_PORT_TIMERS(Port_Number As Long) As Boolean

Dim Temp_Result As Boolean

Const NO_TIMEOUT As Long = MAXDWORD
Const WRITE_CONSTANT As Long = LONG_3000

With COM_PORT(Port_Number)

    .Timeouts.Read_Interval_Timeout = NO_TIMEOUT              ' Timeouts not used for file reads.
    .Timeouts.Read_Total_Timeout_Constant = LONG_0            '
    .Timeouts.Read_Total_Timeout_Multiplier = LONG_0          '

    .Timeouts.Write_Total_Timeout_Constant = WRITE_CONSTANT
    .Timeouts.Write_Total_Timeout_Multiplier = LONG_0

Temp_Result = (Set_Com_Timers(.Handle, .Timeouts) <> LONG_0)
.DLL_Error = IIf(Temp_Result, LONG_0, Err.LastDllError)

End With

SET_PORT_TIMERS = Temp_Result

End Function

Public Function CHECK_COM_PORT(ByVal Port_Number As Long) As Long

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.
' https://docs.microsoft.com/en-us/office/vba/api/Excel.Application.Volatile

Dim Temp_Queue As Long
Dim Port_Entered As Boolean

Temp_Queue = LONG_NEG_1                    ' -1 = queue length not available (invalid, stopped, busy or failed)

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                       ' non-blocking per-port operation guard

  If Port_Entered Then

    On Error GoTo Clean_Exit

    If Port_Ready(Port_Number) Then

        With COM_PORT(Port_Number)

            If Com_Port_Clear(.Handle, .Errors, .Status) <> LONG_0 Then

                .DLL_Error = LONG_0
                .Error_History = .Error_History Or .Errors ' preserve error evidence consumed by this poll

                Temp_Queue = .Status.QUEUE_IN

            Else

                .DLL_Error = Err.LastDllError

            End If

        End With

    End If

  End If

End If

DoEvents                                                   ' yields while the guard is held, so a status query cannot reset the last error before the caller reads it

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then Temp_Queue = LONG_NEG_1

CHECK_COM_PORT = Temp_Queue

End Function

Public Function CLEAR_TO_SEND(Port_Number As Long) As Boolean

' returns True if port valid, started and COM Port CTS signal is asserted.
' CTS = Clear To Send, from attached serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Signal_State As Boolean
Dim Modem_Signals As Long

Const CTS_ON As Long = HEX_10

On Error GoTo Clean_Exit

If Port_Ready(Port_Number) Then

    If READ_MODEM_SIGNALS(Port_Number, Modem_Signals) Then Signal_State = Modem_Signals And CTS_ON

End If

Clean_Exit:

If Err.Number <> LONG_0 Then Signal_State = False

CLEAR_TO_SEND = Signal_State

End Function

Private Function READ_MODEM_SIGNALS(Port_Number As Long, ByRef Modem_Signals As Long) As Boolean

' Shared modem-status read for DEVICE_READY / DEVICE_CALLING / CLEAR_TO_SEND / CARRIER_DETECT.
' Reads into a caller-supplied local so a status query from a worksheet cell, timer or callback can
' run while a port operation is in progress WITHOUT overwriting the operation's shared diagnostics.
' Shared state is only updated when the port is idle.

Dim Temp_Result As Boolean

Temp_Result = (Get_Port_Modem(COM_PORT(Port_Number).Handle, Modem_Signals) <> LONG_0)

If Not COM_PORT(Port_Number).Port_Busy Then                ' persist only when no operation is in flight

    COM_PORT(Port_Number).DLL_Error = IIf(Temp_Result, LONG_0, Err.LastDllError)

    If Temp_Result Then COM_PORT(Port_Number).Signals = Modem_Signals

End If

READ_MODEM_SIGNALS = Temp_Result

End Function

Public Function DEVICE_READY(Port_Number As Long) As Boolean

' returns True if port valid, started and COM Port DSR signal is asserted.
' DSR = Data Set Ready,from attached serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Signal_State As Boolean
Dim Modem_Signals As Long

Const DSR_ON As Long = HEX_20

On Error GoTo Clean_Exit

If Port_Ready(Port_Number) Then

    If READ_MODEM_SIGNALS(Port_Number, Modem_Signals) Then Signal_State = Modem_Signals And DSR_ON

End If

Clean_Exit:

If Err.Number <> LONG_0 Then Signal_State = False

DEVICE_READY = Signal_State

End Function

Public Function DEVICE_CALLING(Port_Number As Long) As Boolean

' returns True if port valid, started and COM Port RI signal is asserted.
' RI = Ring Indicator from attached modem, serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Signal_State As Boolean
Dim Modem_Signals As Long

Const RING_ON As Long = HEX_40

On Error GoTo Clean_Exit

If Port_Ready(Port_Number) Then

    If READ_MODEM_SIGNALS(Port_Number, Modem_Signals) Then Signal_State = Modem_Signals And RING_ON

End If

Clean_Exit:

If Err.Number <> LONG_0 Then Signal_State = False

DEVICE_CALLING = Signal_State

End Function

Public Function CARRIER_DETECT(Port_Number As Long) As Boolean

' returns True if port valid, started and COM Port RLSD/CD signal is asserted.
' RLSD/CD = Carrier Detect from attached serial device or cable configuration.

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Signal_State As Boolean
Dim Modem_Signals As Long

Const DCD_ON As Long = HEX_80

On Error GoTo Clean_Exit

If Port_Ready(Port_Number) Then

    If READ_MODEM_SIGNALS(Port_Number, Modem_Signals) Then Signal_State = Modem_Signals And DCD_ON

End If

Clean_Exit:

If Err.Number <> LONG_0 Then Signal_State = False

CARRIER_DETECT = Signal_State

End Function

Public Function SIGNAL_COM_PORT(ByVal Port_Number As Long, Signal_Function As Long) As Boolean

' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-escapecommfunction
' set/clear BREAK, DTR, RTS port signals from list above.
'
' Accepted function codes 1 to 9, as defined in the Windows SDK (winbase.h):
' 1 = SETXOFF, 2 = SETXON, 3 = SETRTS, 4 = CLRRTS, 5 = SETDTR, 6 = CLRDTR,
' 7 = RESETDEV (reset device if possible - defined in winbase.h, omitted from the online table),
' 8 = SETBREAK, 9 = CLRBREAK.

Dim Signal_Valid As Boolean
Dim Signal_Result As Boolean
Dim Port_Entered As Boolean

Signal_Valid = Signal_Function > LONG_0 And Signal_Function < LONG_10

If Port_Valid(Port_Number) And Signal_Valid Then

  Port_Entered = PORT_ENTER(Port_Number)                   ' signals mutate port control state - guard like any operation

  If Port_Entered Then

    On Error GoTo Clean_Exit

    If Port_Ready(Port_Number) Then

        Signal_Result = (Set_Com_Signal(COM_PORT(Port_Number).Handle, Signal_Function) <> LONG_0)
        COM_PORT(Port_Number).DLL_Error = IIf(Signal_Result, LONG_0, Err.LastDllError)

    End If

  End If

End If

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then Signal_Result = False

SIGNAL_COM_PORT = Signal_Result

End Function

Public Function REQUEST_TO_SEND(ByVal Port_Number As Long, RTS_State As Boolean) As Boolean

Dim RTS_Signal As Long
Dim RTS_Result As Boolean
Dim Port_Entered As Boolean

Const SIGNAL_RTS_1 As Long = LONG_3
Const SIGNAL_RTS_0 As Long = LONG_4

RTS_Signal = IIf(RTS_State, SIGNAL_RTS_1, SIGNAL_RTS_0)

If Port_Valid(Port_Number) Then

  Port_Entered = PORT_ENTER(Port_Number)                   ' RTS mutates port control state - guard like any operation

  If Port_Entered Then

    On Error GoTo Clean_Exit

    If Port_Ready(Port_Number) Then

        RTS_Result = (Set_Com_Signal(COM_PORT(Port_Number).Handle, RTS_Signal) <> LONG_0)
        COM_PORT(Port_Number).DLL_Error = IIf(RTS_Result, LONG_0, Err.LastDllError)

        If RTS_Result Then Kernel_Sleep_MilliSeconds LONG_50
        ' optional - allow local and remote hardware devices to settle.

    End If

  End If

End If

Clean_Exit:

If Port_Entered Then PORT_LEAVE Port_Number

If Err.Number <> LONG_0 Then RTS_Result = False

REQUEST_TO_SEND = RTS_Result

End Function

Private Function CLEAN_PORT_SETTINGS(Port_Settings As String) As String

Dim New_Settings As String

Const TEXT_COMMA As String = ","
Const TEXT_SPACE As String = " "
Const TEXT_EQUALS As String = "="
Const TEXT_DOUBLE_SPACE As String = "  "
Const TEXT_EQUALS_SPACE As String = "= "
Const TEXT_SPACE_EQUALS As String = " ="

New_Settings = Trim(Port_Settings)
New_Settings = UCase(New_Settings)

If InStr(New_Settings, TEXT_EQUALS) > LONG_0 Then New_Settings = Replace(New_Settings, TEXT_COMMA, TEXT_SPACE, , , vbTextCompare)   ' key=value form only - a comma list such as 9600,N,8,1 is a valid BuildCommDCB string
New_Settings = Replace(New_Settings, TEXT_DOUBLE_SPACE, TEXT_SPACE, , , vbTextCompare)   ' collapse runs first, or "A  =  B" keeps a space around "="
New_Settings = Replace(New_Settings, TEXT_DOUBLE_SPACE, TEXT_SPACE, , , vbTextCompare)
New_Settings = Replace(New_Settings, TEXT_DOUBLE_SPACE, TEXT_SPACE, , , vbTextCompare)
New_Settings = Replace(New_Settings, TEXT_SPACE_EQUALS, TEXT_EQUALS, , , vbTextCompare)
New_Settings = Replace(New_Settings, TEXT_EQUALS_SPACE, TEXT_EQUALS, , , vbTextCompare)

CLEAN_PORT_SETTINGS = New_Settings

End Function

Public Function GET_HOST_MICROSECONDS() As Currency

' Application.Volatile  ' - remove comment mark to allow function to recalculate in Excel Worksheet cell.

Dim Temp_QPC As Currency

QPC Temp_QPC

GET_HOST_MICROSECONDS = TICKS_TO_MICROSECONDS(Temp_QPC)

End Function

Private Function HOST_TICK_FREQUENCY() As Currency

' https://docs.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancefrequency
'
' The performance counter frequency is fixed at system boot, so it is queried once and cached.
' QPC and QPF values are both read into Currency, which scales the raw 64-bit integer by 1/10000.
' Both values carry the same scaling, so the ticks/frequency ratio - and therefore elapsed time - is correct.
' The counter frequency is NOT assumed to be 10 MHz: hosts and virtual machines do use other frequencies.

Dim Temp_QPF As Currency

Const ASSUMED_FREQUENCY As Currency = 1000    ' Currency-scaled 10 MHz, used only if QPF is unavailable

If Host_Ticks_Per_Second = 0 Then

    If QPF(Temp_QPF) <> LONG_0 Then Host_Ticks_Per_Second = Temp_QPF

    If Host_Ticks_Per_Second <= 0 Then Host_Ticks_Per_Second = ASSUMED_FREQUENCY

End If

HOST_TICK_FREQUENCY = Host_Ticks_Per_Second

End Function

Private Function TICKS_TO_MICROSECONDS(Elapsed_Ticks As Currency) As Currency

TICKS_TO_MICROSECONDS = Int(CDbl(Elapsed_Ticks) / CDbl(HOST_TICK_FREQUENCY()) * LONG_1E6)

End Function

Public Function Port_Started(Port_Number As Long) As Boolean

' Alias of Port_Ready - the main module uses this name, so code using it compiles with either module.

Port_Started = Port_Ready(Port_Number)

End Function

Public Function Port_Number_Valid(Port_Number As Long) As Boolean

' Alias of Port_Valid - the main module uses this name, so code using it compiles with either module.

Port_Number_Valid = Port_Valid(Port_Number)

End Function

Public Function GET_HOST_MILLISECONDS() As Currency

' Same as the main module: Currency, not Long - an absolute millisecond count overflows a Long after about 24.9 days.

Dim Temp_QPC As Currency

QPC Temp_QPC

GET_HOST_MILLISECONDS = Int(TICKS_TO_MICROSECONDS(Temp_QPC) / LONG_1000)

End Function

Public Function Port_Ready(Port_Number As Long) As Boolean

Dim Temp_Result As Boolean

If Port_Valid(Port_Number) Then

    Temp_Result = COM_PORT(Port_Number).Handle > LONG_0

End If

Port_Ready = Temp_Result

End Function

Public Function Port_Valid(Port_Number As Long) As Boolean

Port_Valid = Not Port_Number < COM_PORT_MIN And Not Port_Number > COM_PORT_MAX

End Function

Private Function COM_PORT_CLOSED(Port_Number As Long) As Boolean

Dim Temp_Result As Boolean

If Port_Valid(Port_Number) Then

    Temp_Result = COM_PORT(Port_Number).Handle < LONG_1
    
End If

COM_PORT_CLOSED = Temp_Result

End Function

Public Function CONVERT_PARITY(DCB_PARITY As Byte) As String

Dim Parity_Text As String

Select Case DCB_PARITY

Case LONG_0:    Parity_Text = "N"
Case LONG_1:    Parity_Text = "O"
Case LONG_2:    Parity_Text = "E"
Case LONG_3:    Parity_Text = "M"
Case LONG_4:    Parity_Text = "S"

Case Else:      Parity_Text = "?"

End Select

CONVERT_PARITY = Parity_Text

End Function

Public Function CONVERT_STOPBITS(DCB_STOPBITS As Byte) As String

Dim Stop_Text As String

Select Case DCB_STOPBITS

Case LONG_0:    Stop_Text = "1"
Case LONG_1:    Stop_Text = "1.5"
Case LONG_2:    Stop_Text = "2"

Case Else:      Stop_Text = "?"

End Select

CONVERT_STOPBITS = Stop_Text

End Function

Public Function DEBUG_COM_PORT(Optional Port_Number As Long, Optional Debug_State As Variant) As Boolean

DEBUG_COM_PORT = False

End Function


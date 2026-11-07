Private Function Query_Port_DCB(ByVal Port_Handle As LongPtr, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Long
    If T_FAIL_QUERY Then
        T_SetLastError 87
        Query_Port_DCB = 0
    Else
        Port_DCB = T_SeedDCB
        Query_Port_DCB = 1
    End If
End Function

Private Function Apply_Port_DCB(ByVal Port_Handle As LongPtr, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Long
    T_ApplyCount = T_ApplyCount + 1
    T_Applied = Port_DCB
    If T_FAIL_APPLY Then
        T_SetLastError 87
        Apply_Port_DCB = 0
    Else
        Apply_Port_DCB = 1
    End If
End Function

Private Function Com_Port_Purge(ByVal Port_Handle As LongPtr, ByVal Purge_Flags As Long) As Long
    T_PurgeMask = Purge_Flags
    Com_Port_Purge = 1
End Function

Private Function Set_Com_Timers(ByVal Port_Handle As LongPtr, ByRef TIMEOUT As COM_PORT_TIMEOUTS) As Long
    Set_Com_Timers = 1
End Function

Private Function Com_Port_Clear(ByVal Port_Handle As LongPtr, ByRef Port_Error_Mask As Long, ByRef Port_Status As COM_PORT_STATUS) As Long
    T_ClearCalls = T_ClearCalls + 1
    If T_CANCEL_ON_CLEAR Then COM_PORT(1).Cancel_Requested = True
    If T_RAISE_CLEAR Then Err.Raise 5, , "T_RAISE_CLEAR"
    If T_FAIL_CLEAR Then
        T_SetLastError 87
        Com_Port_Clear = 0
        Exit Function
    End If
    Port_Error_Mask = 0
    Port_Status.QUEUE_IN = 0
    If T_ClearCalls <= T_QOUT_BUSY_CALLS Then Port_Status.QUEUE_OUT = 1 Else Port_Status.QUEUE_OUT = 0
    Com_Port_Clear = 1
End Function

Private Function Set_Com_Signal(ByVal Port_Handle As LongPtr, ByVal Signal_Function As Long) As Long
    If T_RAISE_SIGNAL Then Err.Raise 5, , "T_RAISE_SIGNAL"
    Set_Com_Signal = 1
End Function

Private Sub T_HOOK(Port_Number As Long)
    Dim i As Long
    T_HOOKS = T_HOOKS + 1
    If T_HOOK_MODE = 1 And T_HOOKS = 1 Then T_Max = 0
    If T_HOOK_MODE = 3 And T_HOOKS = 1 Then T_WRITE_LATE
    If T_HOOK_MODE = 2 And T_HOOKS = 1 Then
        ReDim T_Arr(0 To 3999)
        For i = 0 To 3999: T_Arr(i) = &H55: Next i
    End If
End Sub

Private Sub T_WRITE_LATE()
    Dim h As LongPtr, b(0 To 2) As Byte, n As Long, r As Long
    b(0) = 65: b(1) = 66: b(2) = 67
    h = Com_Port_Create(T_LATE_PATH, &HC0000000, 3, NULL_POINTER, 3, 0, NULL_POINTER)
    r = Synchronous_Write_Bytes(h, b(0), 3, n, NULL_POINTER)
    r = Com_Port_Close(h)
End Sub

Public Function T_RESET() As String
    T_PATH = vbNullString
    T_RAISE_CONFIG = False: T_RAISE_PRINT = False: T_RAISE_CLEAR = False: T_RAISE_SIGNAL = False
    T_FAIL_QUERY = False: T_FAIL_APPLY = False: T_FAIL_CLEAR = False
    T_ApplyCount = 0: T_ClearCalls = 0: T_QOUT_BUSY_CALLS = 0
    T_HOOKS = 0: T_HOOK_MODE = 0: T_LOG = vbNullString: T_CANCEL_ON_CLEAR = False
    COM_PORT(1).Port_Busy = False
    COM_PORT(1).Cancel_Requested = False
    COM_PORT(1).Handle = 0
'@M COM_PORT(1).Name = vbNullString
'@M COM_PORT(1).DLL_Error = 0
'@M COM_PORT(1).VBA_Error = 0
    COM_PORT(1).Error_History = 0
    T_RESET = "ok"
End Function

Private Sub T_SEEDDCB_SET(ByVal bits As Long, ByVal xon As Long, ByVal xoff As Long, ByVal baud As Long, ByVal par As Long)
    Dim d As DEVICE_CONTROL_BLOCK
    d.LENGTH_DCB = 28: d.BAUD_RATE = baud: d.BIT_FIELD = bits: d.BYTE_SIZE = 8
    d.PARITY = par: d.CHAR_XON = xon: d.CHAR_XOFF = xoff
    T_SeedDCB = d
End Sub

Public Function T_OPEN(ByVal path As String, ByVal baud As Long, ByVal create As Boolean) As String
    Dim h As LongPtr, disp As Long
    disp = 3: If create Then disp = 4
    h = Com_Port_Create(path, &HC0000000, 3, NULL_POINTER, disp, 0, NULL_POINTER)
    If h = HANDLE_INVALID Then T_OPEN = "FAIL": Exit Function
    COM_PORT(1).Handle = h
'@M COM_PORT(1).Name = "COM Port 1,"
    With COM_PORT(1).DCB
        .LENGTH_DCB = 28: .BAUD_RATE = baud: .BYTE_SIZE = 8: .PARITY = 0: .STOP_BITS = 0
    End With
    PURGE_BUFFERS 1
    T_OPEN = IIf(SET_PORT_VALUES(1), "ok", "VALUES-FAIL")
End Function

Public Function T_CLOSE() As String
    Dim r As Long
    r = Com_Port_Close(COM_PORT(1).Handle)
    COM_PORT(1).Handle = 0
    T_CLOSE = "closed=" & r
End Function

Public Function T_CLEAN(ByVal s As String) As String
    T_CLEAN = "[" & CLEAN_PORT_SETTINGS(s) & "]"
End Function

Public Function T_SPC(ByVal settings As String, ByVal bits As Long, ByVal xon As Long, ByVal xoff As Long, ByVal baud As Long) As String
    Dim r As Boolean, c As String
    T_RESET
    T_SEEDDCB_SET bits, xon, xoff, baud, 0
    COM_PORT(1).Handle = 1
    c = CLEAN_PORT_SETTINGS(settings)
    r = SET_PORT_CONFIG(1, c)
    T_SPC = "res=" & r & ";applies=" & T_ApplyCount & ";bits=" & Hex$(COM_PORT(1).DCB.BIT_FIELD) & ";xon=" & COM_PORT(1).DCB.CHAR_XON & ";xoff=" & COM_PORT(1).DCB.CHAR_XOFF & ";baud=" & COM_PORT(1).DCB.BAUD_RATE & ";par=" & COM_PORT(1).DCB.PARITY
    COM_PORT(1).Handle = 0
End Function

Private Function T_DLLERR() As String
'@M T_DLLERR = CStr(COM_PORT(1).DLL_Error)
'@ND T_DLLERR = "n/a"
End Function

Private Function T_VBAERR() As String
'@M T_VBAERR = CStr(COM_PORT(1).VBA_Error)
'@ND T_VBAERR = "n/a"
End Function

Public Function T_START_NOSET(ByVal path As String, ByVal bits As Long, ByVal failApply As Boolean) As String
    Dim r As Boolean
    T_RESET
    T_PATH = path
    T_SEEDDCB_SET bits, 17, 19, 9600, 0
    T_FAIL_APPLY = failApply
    r = START_COM_PORT(1)
    T_START_NOSET = "res=" & r & ";applies=" & T_ApplyCount & ";bits=" & Hex$(COM_PORT(1).DCB.BIT_FIELD) & ";started=" & (COM_PORT(1).Handle <> 0) & ";dll=" & T_DLLERR()
    If COM_PORT(1).Handle <> 0 Then T_CLOSE
End Function

Public Function T_START_RAISE(ByVal path As String) As String
    Dim r As Boolean
    T_RESET
    T_PATH = path
    T_SEEDDCB_SET 0, 17, 19, 9600, 0
    T_RAISE_CONFIG = True
    r = START_COM_PORT(1)
    T_START_RAISE = "res=" & r & ";handle_nonzero=" & (COM_PORT(1).Handle <> 0) & ";busy=" & Port_Busy(1)
    If COM_PORT(1).Handle <> 0 Then T_CLOSE
End Function

Public Function T_START_CFGFAIL(ByVal path As String) As String
    Dim r As Boolean
    T_RESET
    T_PATH = path
    T_FAIL_QUERY = True
    r = START_COM_PORT(1)
    T_START_CFGFAIL = "res=" & r & ";handle=" & COM_PORT(1).Handle & ";dll=" & T_DLLERR()
    If COM_PORT(1).Handle <> 0 Then T_CLOSE
End Function

Public Function T_START_STALE_ERR(ByVal path As String) As String
    Dim r As Boolean
    T_RESET
    T_PATH = path
    T_SEEDDCB_SET 0, 17, 19, 9600, 0
    On Error Resume Next
    Err.Raise 5
    r = START_COM_PORT(1)
    On Error GoTo 0
    T_START_STALE_ERR = "res=" & r
    If COM_PORT(1).Handle <> 0 Then T_CLOSE
End Function

Public Function T_TRAP_CHECK(ByVal path As String) As String
    Dim r As Long
    T_RESET
    T_OPEN path, 9600, False
    T_RAISE_CLEAR = True
    r = CHECK_COM_PORT(1)
    T_TRAP_CHECK = "r=" & r & ";busy=" & Port_Busy(1) & ";vba=" & T_VBAERR()
    T_RAISE_CLEAR = False
    T_CLOSE
End Function

Public Function T_TRAP_PRINT(ByVal path As String) As String
    Dim r As Long, e As Long
    T_RESET
    T_OPEN path, 9600, False
    T_RAISE_CLEAR = True
    T_RAISE_PRINT = True
    On Error Resume Next
    r = CHECK_COM_PORT(1)
    e = Err.Number
    On Error GoTo 0
    T_RAISE_PRINT = False: T_RAISE_CLEAR = False
    T_TRAP_PRINT = "busy_after=" & Port_Busy(1) & ";err_seen_by_caller=" & e
    COM_PORT(1).Port_Busy = False
    T_CLOSE
End Function

Public Function T_STOP_BOGUS(ByVal path As String) As String
    Dim r As Boolean, h As LongPtr, r2 As Long
    T_RESET
    h = Com_Port_Create(path, &HC0000000, 3, NULL_POINTER, 3, 0, NULL_POINTER)
    r2 = Com_Port_Close(h)
    COM_PORT(1).Handle = h
'@M COM_PORT(1).Name = "COM Port 1,"
    r = STOP_COM_PORT(1)
    T_STOP_BOGUS = "res=" & r & ";handle_nonzero=" & (COM_PORT(1).Handle <> 0) & ";busy=" & Port_Busy(1)
    COM_PORT(1).Handle = 0
End Function

Public Function T_DRAIN(ByVal path As String) As String
    Dim r As Boolean
    T_RESET
    T_OPEN path, 9600, False
    T_QOUT_BUSY_CALLS = 1
    r = DRAIN_PORT_OUTPUT(1, 10)
    T_DRAIN = "drained=" & r & ";polls=" & T_ClearCalls
    T_CLOSE
End Function

Public Function T_DEBUG_STALE() As String
    Dim r As Boolean, e As Long
    T_RESET
    On Error Resume Next
    Err.Raise 5
    r = DEBUG_COM_PORT(99, True)
    e = Err.Number
    On Error GoTo 0
    T_DEBUG_STALE = "r=" & r & ";err_after=" & e
End Function

Public Function T_SEND(ByVal path As String, ByVal kind As String) As String
    Dim r As Boolean, a() As Byte
    T_RESET
    T_OPEN path, 9600, True
    Select Case kind
    Case "error": r = SEND_COM_PORT(1, CVErr(2007))
    Case "bytes": ReDim a(0 To 3): r = SEND_COM_PORT(1, a)
    Case "range1": ActiveSheet.Range("A1").Value = "hi": r = SEND_COM_PORT(1, ActiveSheet.Range("A1"))
    Case "range2": ActiveSheet.Range("A1:A2").Value = "hi": r = SEND_COM_PORT(1, ActiveSheet.Range("A1:A2"))
    Case "text": r = SEND_COM_PORT(1, "Hello")
    End Select
    T_SEND = "r=" & r & ";busy=" & Port_Busy(1)
    T_CLOSE
End Function

Public Function T_TX(ByVal path As String, ByVal kind As String) As String
    Dim r As Boolean, a() As Byte, i As Long
    T_RESET
    T_OPEN path, 9600, True
    Select Case kind
    Case "transmit": r = TRANSMIT_COM_PORT(1, "Hello" & vbCrLf)
    Case "put": r = PUT_COM_PORT(1, "Zq")
    Case "bytes": ReDim a(0 To 255): For i = 0 To 255: a(i) = i: Next i: r = TRANSMIT_BYTES(1, a)
    Case "bytes_big": ReDim a(0 To 1999): For i = 0 To 1999: a(i) = &HAA: Next i: r = TRANSMIT_BYTES(1, a)
    Case "bytes_changed"
        ReDim T_Arr(0 To 1999): For i = 0 To 1999: T_Arr(i) = &HAA: Next i
        T_HOOK_MODE = 2
        r = TRANSMIT_BYTES(1, T_Arr)
    End Select
    T_TX = "r=" & r
    T_CLOSE
End Function

Public Function T_READ256(ByVal path As String) As String
    Dim s As String, i As Long, o As String
    T_RESET
    T_OPEN path, 9600, False
    s = READ_COM_PORT(1, 256)
    o = "len=" & Len(s) & ";"
    For i = 1 To Len(s): o = o & AscW(Mid$(s, i, 1)) & ",": Next i
    T_READ256 = o
    T_CLOSE
End Function

Public Function T_RECV(ByVal path As String, ByVal which As String, ByVal maxb As Long, ByVal totalms As Long, ByVal charWait As Long) As String
    Dim s As String, n As Long, a() As Byte, t0 As Currency, t1 As Currency
    T_RESET
    T_OPEN path, 9600, False
    If charWait > 0 Then COM_PORT(1).Timers.Char_Loop_Wait = charWait
    t0 = GET_HOST_MICROSECONDS
    If which = "str" Then
        s = RECEIVE_COM_PORT(1, maxb, totalms): n = Len(s)
    Else
        n = RECEIVE_BYTES(1, a, maxb, totalms)
    End If
    t1 = GET_HOST_MICROSECONDS
    T_RECV = "n=" & n & ";ms=" & CLng((t1 - t0) / 1000)
    T_CLOSE
End Function

Public Function T_RB_FIXED(ByVal path As String) As String
    Dim fixedArr(0 To 9) As Byte, dyn() As Byte, r1 As Long, r2 As Long
    T_RESET
    T_OPEN path, 9600, False
    fixedArr(0) = 7
    r1 = RECEIVE_BYTES(1, fixedArr)
    r2 = RECEIVE_BYTES(1, dyn)
    T_RB_FIXED = "fixed_r=" & r1 & ";then_dynamic_r=" & r2 & ";busy=" & Port_Busy(1) & ";fixed0=" & fixedArr(0)
    T_CLOSE
End Function

Public Function T_STOP_DRAIN(ByVal path As String, ByVal busyCalls As Long, ByVal cancel As Boolean, ByVal failClear As Boolean) As String
    Dim r As Boolean, o As Boolean, r2 As Boolean
    T_RESET
    T_OPEN path, 9600, False
    T_QOUT_BUSY_CALLS = busyCalls
    T_CANCEL_ON_CLEAR = cancel
    T_FAIL_CLEAR = failClear
    T_PurgeMask = -1
    r = STOP_COM_PORT(1, 30)
    o = (COM_PORT(1).Handle <> 0)
    T_FAIL_CLEAR = False
    T_CANCEL_ON_CLEAR = False
    T_STOP_DRAIN = "res=" & r & ";open=" & o & ";busy=" & Port_Busy(1) & ";purge=" & T_PurgeMask
    If o Then
        r2 = STOP_COM_PORT(1)
        T_STOP_DRAIN = T_STOP_DRAIN & ";abort=" & r2 & ";open_after=" & (COM_PORT(1).Handle <> 0) & ";abort_purge=" & T_PurgeMask
    End If
    If COM_PORT(1).Handle <> 0 Then T_CLOSE
End Function

Public Function T_DRAIN_TIME(ByVal path As String, ByVal drainMs As Long) As String
    Dim r As Boolean, t0 As Currency
    T_RESET
    T_OPEN path, 9600, False
    T_QOUT_BUSY_CALLS = 1000000
    t0 = GET_HOST_MICROSECONDS
    r = DRAIN_PORT_OUTPUT(1, drainMs)
    T_DRAIN_TIME = "r=" & r & ";ms=" & CLng((GET_HOST_MICROSECONDS - t0) / 1000)
    T_CLOSE
End Function

Private Sub T_MAKE_EMPTY(ByVal path As String)
    Dim h As LongPtr, r As Long
    h = Com_Port_Create(path, &HC0000000, 3, NULL_POINTER, 2, 0, NULL_POINTER)
    r = Com_Port_Close(h)
End Sub

Public Function T_RECV_LATE(ByVal path As String, ByVal late As Boolean) As String
    Dim s As String, t0 As Currency
    T_RESET
    T_MAKE_EMPTY path
    T_OPEN path, 9600, False
    T_LATE_PATH = path
    If late Then T_HOOK_MODE = 3
    t0 = GET_HOST_MICROSECONDS
    s = RECEIVE_COM_PORT(1)
    T_RECV_LATE = "s=[" & s & "];ms_at_least_50=" & ((GET_HOST_MICROSECONDS - t0) / 1000 >= 50)
    T_CLOSE
End Function

'@NEW Public Function T_LAST_ERROR(ByVal path As String, ByVal missing As String) As String
'@NEW     Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, h As LongPtr, s As String, r As Boolean
'@NEW     T_RESET
'@NEW     a = PORT_LAST_ERROR(99)
'@NEW     T_MAKE_EMPTY path
'@NEW     T_OPEN path, 9600, False
'@NEW     s = GET_COM_PORT(1)
'@NEW     b = PORT_LAST_ERROR(1)
'@NEW     T_FAIL_CLEAR = True
'@NEW     c = CHECK_COM_PORT(1)
'@NEW     d = PORT_LAST_ERROR(1)
'@NEW     T_FAIL_CLEAR = False
'@NEW     s = GET_COM_PORT(1)
'@NEW     e = PORT_LAST_ERROR(1)
'@NEW     h = COM_PORT(1).Handle
'@NEW     g = Com_Port_Close(h)
'@NEW     s = GET_COM_PORT(1)
'@NEW     f = PORT_LAST_ERROR(1)
'@NEW     COM_PORT(1).Handle = 0
'@NEW     T_LAST_ERROR = "invalid=" & a & ";no_data=" & b & ";check=" & c & ";clear_fail=" & d & ";recovered=" & e & ";get_closed=[" & s & "]" & f
'@NEW     T_RESET
'@NEW     T_PATH = missing
'@NEW     T_SEEDDCB_SET 0, 17, 19, 9600, 0
'@NEW     r = START_COM_PORT(1)
'@NEW     T_LAST_ERROR = T_LAST_ERROR & ";start_missing=" & r & ":" & PORT_LAST_ERROR(1)
'@NEW     T_RESET
'@NEW     T_PATH = path
'@NEW     T_SEEDDCB_SET 0, 17, 19, 9600, 0
'@NEW     T_FAIL_APPLY = True
'@NEW     r = START_COM_PORT(1, "baud=9600")
'@NEW     T_LAST_ERROR = T_LAST_ERROR & ";start_cfgfail=" & r & ":" & PORT_LAST_ERROR(1)
'@NEW End Function

Public Function T_RB_LOCKED(ByVal path As String) As String
    Dim a() As Byte, v As Variant, r As Long, e As Long
    T_RESET
    T_OPEN path, 9600, False
    ReDim a(0 To 2)
    r = 99
    On Error Resume Next
    For Each v In a
        r = RECEIVE_BYTES(1, a)
        e = Err.Number
        Exit For
    Next v
    On Error GoTo 0
    T_RB_LOCKED = "r=" & r & ";err_seen_by_caller=" & e & ";busy=" & Port_Busy(1)
    COM_PORT(1).Port_Busy = False
    T_CLOSE
End Function

Public Function T_RB_BUSY(ByVal path As String) As String
    Dim a() As Byte, r As Long, ub As String
    T_RESET
    T_OPEN path, 9600, False
    ReDim a(0 To 2): a(0) = 7: a(1) = 8: a(2) = 9
    COM_PORT(1).Port_Busy = True
    r = RECEIVE_BYTES(1, a)
    COM_PORT(1).Port_Busy = False
    On Error Resume Next
    ub = CStr(UBound(a)) & ":" & a(0) & a(1) & a(2)
    If Err.Number <> 0 Then ub = "array-erased"
    On Error GoTo 0
    T_RB_BUSY = "r=" & r & ";array=" & ub
    T_CLOSE
End Function

Public Function T_BYVAL(ByVal path As String) As String
    Dim a() As Byte, n As Long
    T_RESET
    T_OPEN path, 9600, False
    T_Max = 1500
    T_HOOK_MODE = 1
    n = RECEIVE_BYTES(1, a, T_Max)
    T_BYVAL = "n=" & n
    T_CLOSE
End Function

Public Function T_WAITS(ByVal path As String, ByVal baud As Long) As String
    T_RESET
    T_OPEN path, baud, False
    With COM_PORT(1).Timers
        T_WAITS = "bps=" & .Bytes_Per_Second & ";exit=" & .Exit_Loop_Wait & ";char=" & .Char_Loop_Wait & ";data=" & .Data_Loop_Wait & ";line=" & .Line_Loop_Wait & ";slice=" & .Timeslice_Bytes
    End With
    T_CLOSE
End Function

Public Function T_DLLTEXT() As String
'@M T_DLLTEXT = DLL_ERROR_TEXT(21) & " | " & DLL_ERROR_TEXT(15)
'@ND T_DLLTEXT = "n/a"
End Function

Public Function T_SHOW_DCB(ByVal bits As Long) As String
'@M Dim p As Long, q As Long
'@M T_RESET
'@M COM_PORT(1).Handle = 1
'@M COM_PORT(1).DCB.BIT_FIELD = bits
'@M SHOW_PORT_DCB 1
'@M p = InStr(T_LOG, "Bits 1")
'@M q = InStr(p, T_LOG, vbLf)
'@M T_SHOW_DCB = Mid$(T_LOG, p, q - p)
'@M COM_PORT(1).Handle = 0
'@ND T_SHOW_DCB = "n/a"
End Function

Public Function T_SHOW_ERR() As String
'@M Dim r As Boolean, L As Variant, o As String
'@M T_RESET
'@M COM_PORT(1).Handle = 1
'@M COM_PORT(1).Error_History = 4
'@M COM_PORT(1).Port_Errors = 8
'@M T_FAIL_CLEAR = True
'@M r = SHOW_PORT_ERRORS(1)
'@M For Each L In Split(T_LOG, vbLf)
'@M     If InStr(L, "PARITY") > 0 Or InStr(L, "FRAMING") > 0 Or InStr(L, "Failed") > 0 Or InStr(L, "FAILED") > 0 Or InStr(L, "INVALID") > 0 Then o = o & Trim$(L) & " // "
'@M Next L
'@M T_SHOW_ERR = "r=" & r & " :: " & o
'@M COM_PORT(1).Handle = 0
'@ND T_SHOW_ERR = "n/a"
End Function

'@NEW Public Function T_ALIASES() As String
'@NEW@M     Dim n As Long: n = 1: Dim m As Long: m = 99
'@NEW@M     T_ALIASES = "Port_Ready(1)=" & Port_Ready(n) & ";Port_Valid(1)=" & Port_Valid(n) & ";Port_Valid(99)=" & Port_Valid(m)
'@NEW@ND     Dim n As Long: n = 1: Dim m As Long: m = 99: Dim b As Byte: b = 2: Dim sb As Byte: sb = 1
'@NEW@ND     T_ALIASES = "Port_Started(1)=" & Port_Started(n) & ";Port_Number_Valid(1)=" & Port_Number_Valid(n) & ";Port_Number_Valid(99)=" & Port_Number_Valid(m) & ";MS>0=" & (GET_HOST_MILLISECONDS() > 0) & ";PAR(2)=" & CONVERT_PARITY(b) & ";STOP(1)=" & CONVERT_STOPBITS(sb)
'@NEW End Function
'@NEW Public Function T_HELPER(ByVal bits As Long, ByVal xon As Long, ByVal xoff As Long) As String
'@NEW     T_RESET
'@NEW     COM_PORT(1).DCB.BIT_FIELD = bits: COM_PORT(1).DCB.CHAR_XON = xon: COM_PORT(1).DCB.CHAR_XOFF = xoff
'@NEW     SET_DCB_SAFE_FLAGS 1
'@NEW     T_HELPER = "bits=" & Hex$(COM_PORT(1).DCB.BIT_FIELD) & ";xon=" & COM_PORT(1).DCB.CHAR_XON & ";xoff=" & COM_PORT(1).DCB.CHAR_XOFF
'@NEW End Function

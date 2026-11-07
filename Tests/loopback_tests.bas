Attribute VB_Name = "LOOPBACK_TESTS"
Option Explicit
'---------------------------------------------------------------------------
' Loopback tests for the serial modules, run against a REAL or VIRTUAL COM port by Tests\run_loopback.ps1.
' Do not import this into a workbook you use. It sends test data on the ports you give it.
'
' Every public test takes (PortA, PortB, ...) and returns ONE line:
'     PASS|name|detail   FAIL|name|detail   INFO|name|detail   SKIP|name|reason
' A VBA error is returned as FAIL with the error text. Ports are stopped on every path.
' PortA = PortB means plug mode: the port is opened once and used as both ends (TX wired to RX).
' A port that will not start (for example port 0 from -DryRun) gives SKIP.
'
' Uses only functions that are public in both SERIAL_PORT_VBA.bas and SERIAL_PORT_VBA_NO_DEBUG.bas.
'---------------------------------------------------------------------------

Private Const SET_9600 As String = "baud=9600 parity=N data=8 stop=1"
Private Const SET_1200 As String = "baud=1200 parity=N data=8 stop=1"
Private Const SET_FLOW_A As String = "baud=9600 parity=N data=8 stop=1 octs=on rts=on"
Private Const SET_FLOW_B As String = "baud=9600 parity=N data=8 stop=1 rts=on"
Private Const SET_EVEN As String = "baud=9600 parity=E data=8 stop=1"
Private Const SET_ODD As String = "baud=9600 parity=O data=8 stop=1"

Private Saved_Have As Boolean
Private Saved_A As String
Private Saved_B As String

'------------------------------ helpers -------------------------------------

Private Function Fmt(ByVal Kind As String, ByVal Name As String, ByVal Detail As String) As String
Detail = Replace(Replace(Detail, vbCr, " "), vbLf, " ")
Fmt = Kind & "|" & Name & "|" & Detail
End Function

Private Function Pair_Mode(ByVal PortA As Long, ByVal PortB As Long) As Boolean
Pair_Mode = (PortA <> PortB)
End Function

Private Function Open_Ends(ByVal PortA As Long, ByVal PortB As Long, ByVal SetA As String, ByVal SetB As String) As Boolean
Dim Dummy As Boolean
Open_Ends = False
If Not START_COM_PORT(PortA, SetA) Then Exit Function
If PortB <> PortA Then
    If Not START_COM_PORT(PortB, SetB) Then
        Dummy = STOP_COM_PORT(PortA)
        Exit Function
    End If
End If
Open_Ends = True
End Function

Private Sub Close_Ends(ByVal PortA As Long, ByVal PortB As Long)
Dim Dummy As Boolean
On Error Resume Next
If Port_Started(PortA) Then Dummy = STOP_COM_PORT(PortA)
If PortB <> PortA Then
    If Port_Started(PortB) Then Dummy = STOP_COM_PORT(PortB)
End If
Err.Clear
End Sub

Private Sub Pause(ByVal Milliseconds As Long)
Dim T0 As Currency
T0 = GET_HOST_MILLISECONDS()
Do While (GET_HOST_MILLISECONDS() - T0) < Milliseconds
    DoEvents
Loop
End Sub

Private Sub Drain_Port(ByVal Port As Long)
Dim Junk() As Byte, N As Long
N = RECEIVE_BYTES(Port, Junk, 0, 300)
End Sub

Private Sub Fill_All_Bytes(Data() As Byte)          ' 00-FF twice = 512 bytes
Dim I As Long
ReDim Data(0 To 511)
For I = 0 To 511
    Data(I) = CByte(I Mod 256)
Next I
End Sub

Private Sub Fill_Seq_Bytes(Data() As Byte, ByVal Count As Long)
Dim I As Long
ReDim Data(0 To Count - 1)
For I = 0 To Count - 1
    Data(I) = CByte(I Mod 251)
Next I
End Sub

Private Function Pattern_Text(ByVal Count As Long) As String       ' printable ASCII 33-126, no CR/LF
Dim I As Long, S As String
S = Space$(Count)
For I = 1 To Count
    Mid$(S, I, 1) = Chr$(33 + ((I - 1) Mod 94))
Next I
Pattern_Text = S
End Function

Private Function First_Diff(Expected() As Byte, Got() As Byte, ByVal GotCount As Long, ByVal ExpectedOffset As Long) As Long
' Index in Got of the first byte that differs from Expected(ExpectedOffset + index), -1 if the GotCount bytes all match.
Dim I As Long
First_Diff = -1
For I = 0 To GotCount - 1
    If Got(I) <> Expected(ExpectedOffset + I) Then
        First_Diff = I
        Exit Function
    End If
Next I
End Function

Private Function Err_Text() As String
Err_Text = "VBA error " & Err.Number & ": " & Err.Description
End Function

Private Function Elapsed(ByVal T0 As Currency) As Long
Elapsed = CLng(GET_HOST_MILLISECONDS() - T0)
End Function

'--------------------- 1. save and restore settings --------------------------

Private Function Parse_Saved(ByVal Text As String) As String
' GET_PORT_SETTINGS gives baud-data-parity-stop, for example 9600-8-N-1. Returns the key=value form, or "" if it cannot be used.
' Handshaking is set explicitly to off, with DTR and RTS on. Without that, the key=value form keeps whatever the
' tests left, and the flow-control test turns on CTS handshaking. The original handshaking cannot be read
' through the module, so it is not restored.
Const NO_HANDSHAKE As String = " to=off xon=off odsr=off octs=off dtr=on rts=on idsr=off"
Dim P() As String
P = Split(Text, "-")
If UBound(P) <> 3 Then Exit Function
If InStr(P(2), "?") > 0 Or InStr(P(3), "?") > 0 Then Exit Function
Parse_Saved = "baud=" & P(0) & " parity=" & P(2) & " data=" & P(1) & " stop=" & P(3) & NO_HANDSHAKE
End Function

Private Function Read_Existing(ByVal Port As Long, ByRef Text As String) As Boolean
Dim Dummy As Boolean
Text = vbNullString
If Not START_COM_PORT(Port) Then Exit Function
Text = GET_PORT_SETTINGS(Port)
Dummy = STOP_COM_PORT(Port)
Read_Existing = True
End Function

Public Function LB_SAVE_SETTINGS(ByVal PortA As Long, ByVal PortB As Long) As String
Const NM As String = "save_settings"
On Error GoTo Fail
Saved_Have = False
Saved_A = vbNullString
Saved_B = vbNullString
If Not Read_Existing(PortA, Saved_A) Then
    LB_SAVE_SETTINGS = Fmt("SKIP", NM, "port " & PortA & " did not start")
    Exit Function
End If
If PortB <> PortA Then
    If Not Read_Existing(PortB, Saved_B) Then
        Saved_A = vbNullString
        LB_SAVE_SETTINGS = Fmt("SKIP", NM, "port " & PortB & " did not start")
        Exit Function
    End If
End If
Saved_Have = True
LB_SAVE_SETTINGS = Fmt("INFO", NM, "settings found on COM" & PortA & " = " & Saved_A & IIf(PortB <> PortA, ", COM" & PortB & " = " & Saved_B, "") & " (baud-data-parity-stop)")
Exit Function
Fail:
LB_SAVE_SETTINGS = Fmt("FAIL", NM, Err_Text())
Close_Ends PortA, PortB
End Function

Private Function Restore_One(ByVal Port As Long, ByVal Saved As String) As String
' "" = restored and checked, otherwise the problem
Dim Want As String, Got As String, Dummy As Boolean
Want = Parse_Saved(Saved)
If Len(Want) = 0 Then Restore_One = "COM" & Port & " saved text '" & Saved & "' cannot be restored": Exit Function
If Not START_COM_PORT(Port, Want) Then Restore_One = "COM" & Port & " did not start with '" & Want & "'": Exit Function
Got = GET_PORT_SETTINGS(Port)
Dummy = STOP_COM_PORT(Port)
If Got <> Saved Then Restore_One = "COM" & Port & " reads back " & Got & ", was " & Saved
End Function

Public Function LB_RESTORE_SETTINGS(ByVal PortA As Long, ByVal PortB As Long) As String
Const NM As String = "restore_settings"
Dim Problem As String, P2 As String
On Error GoTo Fail
If Not Saved_Have Then
    LB_RESTORE_SETTINGS = Fmt("SKIP", NM, "no settings were saved")
    Exit Function
End If
Problem = Restore_One(PortA, Saved_A)
If PortB <> PortA Then
    P2 = Restore_One(PortB, Saved_B)
    If Len(P2) > 0 Then Problem = Problem & IIf(Len(Problem) > 0, "; ", "") & P2
End If
If Len(Problem) = 0 Then
    LB_RESTORE_SETTINGS = Fmt("PASS", NM, "restored COM" & PortA & " = " & Saved_A & IIf(PortB <> PortA, ", COM" & PortB & " = " & Saved_B, "") & " (baud, parity, data, stop; handshaking set to off with DTR and RTS on, because the original cannot be read)")
Else
    LB_RESTORE_SETTINGS = Fmt("FAIL", NM, Problem)
End If
Close_Ends PortA, PortB
Exit Function
Fail:
LB_RESTORE_SETTINGS = Fmt("FAIL", NM, Err_Text())
Close_Ends PortA, PortB
End Function

Public Function LB_CLEANUP(ByVal PortA As Long, ByVal PortB As Long) As String
' Run by the runner before the restore. A port still started here means a test left it open.
Const NM As String = "cleanup"
Dim Leaked As String
On Error Resume Next
If Port_Started(PortA) Then Leaked = "COM" & PortA
If PortB <> PortA Then
    If Port_Started(PortB) Then Leaked = Leaked & IIf(Len(Leaked) > 0, ", ", "") & "COM" & PortB
End If
Close_Ends PortA, PortB
If Len(Leaked) > 0 Then
    LB_CLEANUP = Fmt("FAIL", NM, "a test left " & Leaked & " started, now stopped")
Else
    LB_CLEANUP = Fmt("PASS", NM, "no port left started")
End If
End Function

'------------------------------ 2. settings ----------------------------------

Private Function Settings_Test(ByVal PortA As Long, ByVal PortB As Long, ByVal NM As String, ByVal Settings As String, ByVal ExpectStart As Boolean, ByVal ExpectGet As String) As String
Dim Ports(0 To 1) As Long, Count As Long, I As Long
Dim Started As Boolean, Got As String, Dummy As Boolean, Probe As String
Dim Problem As String, Detail As String
On Error GoTo Fail
Ports(0) = PortA: Count = 1
If PortB <> PortA Then Ports(1) = PortB: Count = 2
For I = 0 To Count - 1
    ' the port must start with no settings first, or this is SKIP (this is what port 0 gives)
    If Not START_COM_PORT(Ports(I)) Then
        Settings_Test = Fmt("SKIP", NM, "port " & Ports(I) & " did not start")
        Exit Function
    End If
    Dummy = STOP_COM_PORT(Ports(I))
Next I
For I = 0 To Count - 1
    Started = START_COM_PORT(Ports(I), Settings)
    Got = GET_PORT_SETTINGS(Ports(I))
    If Started <> ExpectStart Then
        Problem = Problem & "COM" & Ports(I) & " start returned " & Started & ", wanted " & ExpectStart & "; "
    ElseIf ExpectStart = False Then
        If Port_Started(Ports(I)) Then Problem = Problem & "COM" & Ports(I) & " left started after the failed start; "
    ElseIf Len(ExpectGet) > 0 Then
        If Got <> ExpectGet Then Problem = Problem & "COM" & Ports(I) & " GET_PORT_SETTINGS = " & Got & ", wanted " & ExpectGet & "; "
    End If
    Detail = Detail & "COM" & Ports(I) & " start=" & Started & " get=" & Got & "; "
    If Port_Started(Ports(I)) Then Dummy = STOP_COM_PORT(Ports(I))
Next I
If Len(Problem) = 0 Then
    Settings_Test = Fmt("PASS", NM, "'" & Settings & "' " & Detail)
Else
    Settings_Test = Fmt("FAIL", NM, "'" & Settings & "' " & Problem & Detail)
End If
Exit Function
Fail:
Settings_Test = Fmt("FAIL", NM, Err_Text())
Close_Ends PortA, PortB
End Function

Public Function LB_SETTINGS_KEYVALUE(ByVal PortA As Long, ByVal PortB As Long) As String
LB_SETTINGS_KEYVALUE = Settings_Test(PortA, PortB, "settings_keyvalue", SET_9600, True, "9600-8-N-1")
End Function

Public Function LB_SETTINGS_COMMA(ByVal PortA As Long, ByVal PortB As Long) As String
LB_SETTINGS_COMMA = Settings_Test(PortA, PortB, "settings_comma_list", "9600,N,8,1", True, vbNullString)
End Function

Public Function LB_SETTINGS_BAUD_ONLY(ByVal PortA As Long, ByVal PortB As Long) As String
LB_SETTINGS_BAUD_ONLY = Settings_Test(PortA, PortB, "settings_baud_only", "9600", True, vbNullString)
End Function

Public Function LB_SETTINGS_BAD(ByVal PortA As Long, ByVal PortB As Long) As String
LB_SETTINGS_BAD = Settings_Test(PortA, PortB, "settings_bad_text", "abc", False, vbNullString)
End Function

'------------------------------ 3. bytes 00-FF --------------------------------

Private Function Bytes_Test(ByVal NM As String, ByVal PortA As Long, ByVal PortB As Long, ByVal TxPort As Long, ByVal RxPort As Long) As String
Dim Tx() As Byte, Rx() As Byte, N As Long, D As Long, Sent As Boolean
On Error GoTo Fail
If Not Open_Ends(PortA, PortB, SET_9600, SET_9600) Then
    Bytes_Test = Fmt("SKIP", NM, "port did not start")
    Exit Function
End If
Drain_Port RxPort
Fill_All_Bytes Tx
Sent = TRANSMIT_BYTES(TxPort, Tx)
N = RECEIVE_BYTES(RxPort, Rx, 0, 10000)
If Not Sent Then
    Bytes_Test = Fmt("FAIL", NM, "TRANSMIT_BYTES returned False; received " & N & " bytes; last error " & PORT_LAST_ERROR(TxPort))
ElseIf N <> 512 Then
    Bytes_Test = Fmt("FAIL", NM, "received " & N & " of 512 bytes")
Else
    D = First_Diff(Tx, Rx, N, 0)
    If D < 0 Then
        Bytes_Test = Fmt("PASS", NM, "512 bytes (00-FF twice) COM" & TxPort & " to COM" & RxPort & " identical")
    Else
        Bytes_Test = Fmt("FAIL", NM, "byte " & D & " sent " & Hex$(Tx(D)) & " received " & Hex$(Rx(D)))
    End If
End If
Close_Ends PortA, PortB
Exit Function
Fail:
Bytes_Test = Fmt("FAIL", NM, Err_Text())
Close_Ends PortA, PortB
End Function

Public Function LB_BYTES_A_TO_B(ByVal PortA As Long, ByVal PortB As Long) As String
LB_BYTES_A_TO_B = Bytes_Test("bytes_00_FF_A_to_B", PortA, PortB, PortA, PortB)
End Function

Public Function LB_BYTES_B_TO_A(ByVal PortA As Long, ByVal PortB As Long) As String
If Not Pair_Mode(PortA, PortB) Then
    LB_BYTES_B_TO_A = Fmt("SKIP", "bytes_00_FF_B_to_A", "plug mode, there is only one port")
Else
    LB_BYTES_B_TO_A = Bytes_Test("bytes_00_FF_B_to_A", PortA, PortB, PortB, PortA)
End If
End Function

'------------------------------ 4. text ---------------------------------------

Public Function LB_TEXT(ByVal PortA As Long, ByVal PortB As Long) As String
Const NM As String = "text_round_trip"
Dim S As String, R As String, I As Long, C As Long, Sent As Boolean
On Error GoTo Fail
If Not Open_Ends(PortA, PortB, SET_9600, SET_9600) Then
    LB_TEXT = Fmt("SKIP", NM, "port did not start")
    Exit Function
End If
Drain_Port PortB
For I = 1 To 3
    For C = 32 To 126
        S = S & Chr$(C)
    Next C
    S = S & vbCrLf
Next I
Sent = TRANSMIT_COM_PORT(PortA, S)
R = RECEIVE_COM_PORT(PortB, 0, 10000)
If Not Sent Then
    LB_TEXT = Fmt("FAIL", NM, "TRANSMIT_COM_PORT returned False; received " & Len(R) & " characters")
ElseIf R = S Then
    LB_TEXT = Fmt("PASS", NM, Len(S) & " characters, printable ASCII with CR/LF, identical")
Else
    LB_TEXT = Fmt("FAIL", NM, "sent " & Len(S) & " characters, received " & Len(R))
End If
Close_Ends PortA, PortB
Exit Function
Fail:
LB_TEXT = Fmt("FAIL", NM, Err_Text())
Close_Ends PortA, PortB
End Function

'------------------------------ 5. exact Max_Bytes ----------------------------

Public Function LB_MAXBYTES_BYTES(ByVal PortA As Long, ByVal PortB As Long) As String
Const NM As String = "max_bytes_RECEIVE_BYTES"
Dim Tx() As Byte, Rx1() As Byte, Rx2() As Byte, N1 As Long, N2 As Long, D1 As Long, D2 As Long
On Error GoTo Fail
If Not Open_Ends(PortA, PortB, SET_9600, SET_9600) Then
    LB_MAXBYTES_BYTES = Fmt("SKIP", NM, "port did not start")
    Exit Function
End If
Drain_Port PortB
Fill_Seq_Bytes Tx, 300
If Not TRANSMIT_BYTES(PortA, Tx) Then
    LB_MAXBYTES_BYTES = Fmt("FAIL", NM, "TRANSMIT_BYTES returned False")
    Close_Ends PortA, PortB
    Exit Function
End If
N1 = RECEIVE_BYTES(PortB, Rx1, 100, 10000)
N2 = RECEIVE_BYTES(PortB, Rx2, 0, 10000)
If N1 <> 100 Then
    LB_MAXBYTES_BYTES = Fmt("FAIL", NM, "first receive with Max_Bytes=100 gave " & N1 & "; second gave " & N2)
ElseIf N2 <> 200 Then
    LB_MAXBYTES_BYTES = Fmt("FAIL", NM, "first receive gave 100; second gave " & N2 & ", wanted 200")
Else
    D1 = First_Diff(Tx, Rx1, N1, 0)
    D2 = First_Diff(Tx, Rx2, N2, 100)
    If D1 >= 0 Then
        LB_MAXBYTES_BYTES = Fmt("FAIL", NM, "first 100 bytes differ at " & D1)
    ElseIf D2 >= 0 Then
        LB_MAXBYTES_BYTES = Fmt("FAIL", NM, "remaining 200 bytes differ at " & D2 & " (not in order or lost)")
    Else
        LB_MAXBYTES_BYTES = Fmt("PASS", NM, "300 sent, 100 then 200 received in order")
    End If
End If
Close_Ends PortA, PortB
Exit Function
Fail:
LB_MAXBYTES_BYTES = Fmt("FAIL", NM, Err_Text())
Close_Ends PortA, PortB
End Function

Public Function LB_MAXBYTES_TEXT(ByVal PortA As Long, ByVal PortB As Long) As String
Const NM As String = "max_bytes_RECEIVE_COM_PORT"
Dim S As String, R1 As String, R2 As String
On Error GoTo Fail
If Not Open_Ends(PortA, PortB, SET_9600, SET_9600) Then
    LB_MAXBYTES_TEXT = Fmt("SKIP", NM, "port did not start")
    Exit Function
End If
Drain_Port PortB
S = Pattern_Text(300)
If Not TRANSMIT_COM_PORT(PortA, S) Then
    LB_MAXBYTES_TEXT = Fmt("FAIL", NM, "TRANSMIT_COM_PORT returned False")
    Close_Ends PortA, PortB
    Exit Function
End If
R1 = RECEIVE_COM_PORT(PortB, 100, 10000)
R2 = RECEIVE_COM_PORT(PortB, 0, 10000)
If Len(R1) <> 100 Then
    LB_MAXBYTES_TEXT = Fmt("FAIL", NM, "first receive with Max_Bytes=100 gave " & Len(R1) & "; second gave " & Len(R2))
ElseIf Len(R2) <> 200 Then
    LB_MAXBYTES_TEXT = Fmt("FAIL", NM, "first receive gave 100; second gave " & Len(R2) & ", wanted 200")
ElseIf R1 <> Left$(S, 100) Or R2 <> Mid$(S, 101) Then
    LB_MAXBYTES_TEXT = Fmt("FAIL", NM, "lengths right, content differs or out of order")
Else
    LB_MAXBYTES_TEXT = Fmt("PASS", NM, "300 sent, 100 then 200 received in order")
End If
Close_Ends PortA, PortB
Exit Function
Fail:
LB_MAXBYTES_TEXT = Fmt("FAIL", NM, Err_Text())
Close_Ends PortA, PortB
End Function

'------------------------------ 6. late reply ---------------------------------

Private Function Late_Scenario(ByVal PortA As Long, ByVal PortB As Long, ByRef Started As Boolean, ByRef Problem As String, ByRef TxMs As Long, ByRef TotalMs As Long, ByRef Sent As Long, ByRef Received As Long, ByRef Same As Boolean) As Boolean
Dim S As String, R As String, T0 As Currency, TxOk As Boolean
On Error GoTo Fail
Started = Open_Ends(PortA, PortB, SET_1200, SET_1200)
If Not Started Then Exit Function
Drain_Port PortB
S = Pattern_Text(120)
Sent = Len(S)
T0 = GET_HOST_MILLISECONDS()
TxOk = TRANSMIT_COM_PORT(PortA, S)
TxMs = Elapsed(T0)
R = RECEIVE_COM_PORT(PortB, 0, 10000)
TotalMs = Elapsed(T0)
Received = Len(R)
Same = (R = S)
If Not TxOk Then Problem = "TRANSMIT_COM_PORT returned False"
Close_Ends PortA, PortB
Late_Scenario = True
Exit Function
Fail:
Problem = Err_Text()
Close_Ends PortA, PortB
Late_Scenario = True
End Function

Public Function LB_LATE_REPLY(ByVal PortA As Long, ByVal PortB As Long) As String
Const NM As String = "late_reply_1200"
Dim Started As Boolean, Problem As String, TxMs As Long, TotalMs As Long, Sent As Long, Received As Long, Same As Boolean, Dummy As Boolean
On Error GoTo Fail
Dummy = Late_Scenario(PortA, PortB, Started, Problem, TxMs, TotalMs, Sent, Received, Same)
If Not Started Then
    LB_LATE_REPLY = Fmt("SKIP", NM, "port did not start")
ElseIf Len(Problem) > 0 Then
    LB_LATE_REPLY = Fmt("FAIL", NM, Problem)
ElseIf Same Then
    LB_LATE_REPLY = Fmt("PASS", NM, "all " & Sent & " bytes came back at 1200 baud when receive was called straight after transmit")
Else
    LB_LATE_REPLY = Fmt("FAIL", NM, "sent " & Sent & ", received " & Received & " (or content differs)")
End If
Exit Function
Fail:
LB_LATE_REPLY = Fmt("FAIL", NM, Err_Text())
Close_Ends PortA, PortB
End Function

Public Function LB_LATE_TIMING(ByVal PortA As Long, ByVal PortB As Long) As String
Const NM As String = "late_reply_timing"
Dim Started As Boolean, Problem As String, TxMs As Long, TotalMs As Long, Sent As Long, Received As Long, Same As Boolean, Dummy As Boolean
Dim Wire As Long
On Error GoTo Fail
Dummy = Late_Scenario(PortA, PortB, Started, Problem, TxMs, TotalMs, Sent, Received, Same)
If Not Started Then
    LB_LATE_TIMING = Fmt("SKIP", NM, "port did not start")
ElseIf Len(Problem) > 0 Then
    LB_LATE_TIMING = Fmt("FAIL", NM, Problem)
Else
    Wire = CLng(Sent * 10# * 1000# / 1200#)
    LB_LATE_TIMING = Fmt("INFO", NM, "transmit call took " & TxMs & " mS, wire time about " & Wire & " mS (" & Sent & " bytes at 1200 baud); transmit plus receive took " & TotalMs & " mS. " & _
        IIf(TxMs < Wire * 0.7, "The write returned before the data was sent (the driver buffers it).", "The write waited for most of the wire time (the driver sends before returning).") & _
        " Virtual ports such as com0com do not pace at the baud rate, so a short time there is expected.")
End If
Exit Function
Fail:
LB_LATE_TIMING = Fmt("FAIL", NM, Err_Text())
Close_Ends PortA, PortB
End Function

'------------------------------ 7. CHECK_COM_PORT ----------------------------

Public Function LB_CHECK_COM_PORT(ByVal PortA As Long, ByVal PortB As Long) As String
Const NM As String = "check_com_port"
Const N_BYTES As Long = 50
Dim Tx() As Byte, Rx() As Byte, Waiting As Long, Drained As Long
On Error GoTo Fail
If Not Open_Ends(PortA, PortB, SET_9600, SET_9600) Then
    LB_CHECK_COM_PORT = Fmt("SKIP", NM, "port did not start")
    Exit Function
End If
Drain_Port PortB
Fill_Seq_Bytes Tx, N_BYTES
If Not TRANSMIT_BYTES(PortA, Tx) Then
    LB_CHECK_COM_PORT = Fmt("FAIL", NM, "TRANSMIT_BYTES returned False")
    Close_Ends PortA, PortB
    Exit Function
End If
Pause 600
Waiting = CHECK_COM_PORT(PortB)
Drained = RECEIVE_BYTES(PortB, Rx, 0, 5000)
If Waiting <> N_BYTES Then
    LB_CHECK_COM_PORT = Fmt("FAIL", NM, "sent " & N_BYTES & ", CHECK_COM_PORT reported " & Waiting & "; draining gave " & Drained)
ElseIf Drained <> N_BYTES Then
    LB_CHECK_COM_PORT = Fmt("FAIL", NM, "CHECK_COM_PORT reported " & Waiting & " but draining gave " & Drained)
Else
    LB_CHECK_COM_PORT = Fmt("PASS", NM, "reported " & Waiting & " waiting, drained " & Drained)
End If
Close_Ends PortA, PortB
Exit Function
Fail:
LB_CHECK_COM_PORT = Fmt("FAIL", NM, Err_Text())
Close_Ends PortA, PortB
End Function

'------------------------------ 8. PORT_LAST_ERROR ----------------------------

Public Function LB_LAST_ERROR(ByVal PortA As Long, ByVal PortB As Long) As String
Const NM As String = "port_last_error"
Dim Tx() As Byte, Rx() As Byte, WriteErr As Long, ReadErr As Long, N As Long, Sent As Boolean
On Error GoTo Fail
If Not Open_Ends(PortA, PortB, SET_9600, SET_9600) Then
    LB_LAST_ERROR = Fmt("SKIP", NM, "port did not start")
    Exit Function
End If
Drain_Port PortB
Fill_Seq_Bytes Tx, 20
Sent = TRANSMIT_BYTES(PortA, Tx)
WriteErr = PORT_LAST_ERROR(PortA)
N = RECEIVE_BYTES(PortB, Rx, 0, 5000)
ReadErr = PORT_LAST_ERROR(PortB)
If Not Sent Or N <> 20 Then
    LB_LAST_ERROR = Fmt("FAIL", NM, "the transfer did not succeed (sent=" & Sent & ", received " & N & "), so the check is not valid; last errors write=" & WriteErr & " read=" & ReadErr)
ElseIf WriteErr = 0 And ReadErr = 0 Then
    LB_LAST_ERROR = Fmt("PASS", NM, "0 after a successful write and a successful read")
Else
    LB_LAST_ERROR = Fmt("FAIL", NM, "after write " & WriteErr & ", after read " & ReadErr & ", both should be 0")
End If
Close_Ends PortA, PortB
Exit Function
Fail:
LB_LAST_ERROR = Fmt("FAIL", NM, Err_Text())
Close_Ends PortA, PortB
End Function

'------------------------------ 9. graceful stop under flow control ----------

Public Function LB_FLOW_STOP(ByVal PortA As Long, ByVal PortB As Long, ByVal RtsCtsLoop As Boolean) As String
Const NM As String = "flow_control_graceful_stop"
Const NOTE As String = " Microsoft documents the driver's transmit queue count as zero for non-overlapped writes, which these modules use. A drain that returns at once means the graceful stop gives no delivery guarantee on this driver."
Dim Tx() As Byte, Rx() As Byte, TxOk As Boolean, TxMs As Long, StopOk As Boolean, StopMs As Long
Dim T0 As Currency, StartedAfter As Boolean, Broken As String, Rcvd As Long, RcvdKnown As Boolean
Dim Detail As String, PlainOk As Boolean, Pair As Boolean
On Error GoTo Fail
Pair = Pair_Mode(PortA, PortB)
If Not Pair And Not RtsCtsLoop Then
    LB_FLOW_STOP = Fmt("SKIP", NM, "plug mode needs -RtsCtsLoop (RTS wired to CTS)")
    Exit Function
End If
If Not Open_Ends(PortA, PortB, SET_FLOW_A, SET_FLOW_B) Then
    LB_FLOW_STOP = Fmt("SKIP", NM, "port did not start")
    Exit Function
End If
Drain_Port PortB
If Not REQUEST_TO_SEND(PortB, False) Then
    LB_FLOW_STOP = Fmt("FAIL", NM, "REQUEST_TO_SEND(COM" & PortB & ", False) returned False, last error " & PORT_LAST_ERROR(PortB))
    Close_Ends PortA, PortB
    Exit Function
End If
Pause 300
If CLEAR_TO_SEND(PortA) Then
    TxOk = REQUEST_TO_SEND(PortB, True)
    Close_Ends PortA, PortB
    LB_FLOW_STOP = Fmt("SKIP", NM, "RTS is not wired to CTS (CTS stayed high with RTS low)")
    Exit Function
End If
Fill_Seq_Bytes Tx, 20
T0 = GET_HOST_MILLISECONDS()
TxOk = TRANSMIT_BYTES(PortA, Tx)
TxMs = Elapsed(T0)
T0 = GET_HOST_MILLISECONDS()
StopOk = STOP_COM_PORT(PortA, 1500)
StopMs = Elapsed(T0)
StartedAfter = Port_Started(PortA)
Detail = "INFO: write with CTS low returned " & TxOk & " after " & TxMs & " mS. STOP_COM_PORT(A, 1500) returned " & StopOk & " after " & StopMs & " mS. "
If StopOk Then
    If StartedAfter Then Broken = "STOP_COM_PORT returned True but the port is still started"
    If Pair Then
        If REQUEST_TO_SEND(PortB, True) Then
            Pause 300
            Rcvd = RECEIVE_BYTES(PortB, Rx, 0, 2000)
            RcvdKnown = True
        End If
    End If
Else
    If Not StartedAfter Then
        Broken = "STOP_COM_PORT returned False but the port was closed (it must stay open)"
    Else
        TxOk = REQUEST_TO_SEND(PortB, True)
        Pause 300
        Rcvd = RECEIVE_BYTES(PortB, Rx, 0, 2000)
        RcvdKnown = True
        PlainOk = STOP_COM_PORT(PortA)
        If Not PlainOk Then
            Broken = "plain STOP_COM_PORT returned False after RTS was raised"
        ElseIf Port_Started(PortA) Then
            Broken = "plain STOP_COM_PORT returned True but the port is still started"
        End If
    End If
End If
Detail = Detail & "INFO: the drain " & IIf(StopMs > 1000, "waited", "did not wait") & " (" & StopMs & " mS). " & _
    IIf(RcvdKnown, "INFO: B received " & Rcvd & " bytes after RTS was raised. ", "INFO: bytes received after RTS was raised not measured (port closed in plug mode). ") & NOTE
Close_Ends PortA, PortB
If Len(Broken) > 0 Then
    LB_FLOW_STOP = Fmt("FAIL", NM, "contract broken: " & Broken & ". " & Detail)
Else
    LB_FLOW_STOP = Fmt("PASS", NM, "stop contract held. " & Detail)
End If
Exit Function
Fail:
Detail = Err_Text()
On Error Resume Next
TxOk = REQUEST_TO_SEND(PortB, True)
Close_Ends PortA, PortB
LB_FLOW_STOP = Fmt("FAIL", NM, Detail)
End Function

'------------------------------ 10. parity detection --------------------------

Public Function LB_PARITY(ByVal PortA As Long, ByVal PortB As Long) As String
Const NM As String = "parity_detection"
Dim Tx() As Byte, Rx() As Byte, Hist As Long, N As Long, Waiting As Long
On Error GoTo Fail
If Not Pair_Mode(PortA, PortB) Then
    LB_PARITY = Fmt("SKIP", NM, "plug mode, needs two ports")
    Exit Function
End If
If Not Open_Ends(PortA, PortB, SET_EVEN, SET_ODD) Then
    LB_PARITY = Fmt("SKIP", NM, "port did not start")
    Exit Function
End If
Drain_Port PortB
Hist = PORT_ERROR_HISTORY(PortB, True)
Fill_All_Bytes Tx
If TRANSMIT_BYTES(PortA, Tx) Then
    Pause 800
    Waiting = CHECK_COM_PORT(PortB)          ' a status poll records the driver's error flags in the history
    N = RECEIVE_BYTES(PortB, Rx, 0, 5000)
End If
Hist = PORT_ERROR_HISTORY(PortB, True)
LB_PARITY = Fmt("INFO", NM, "A even, B odd, 512 bytes sent; B had " & Waiting & " waiting, read " & N & "; PORT_ERROR_HISTORY(B) = " & Hist & _
    IIf((Hist And 4) <> 0, " (parity bit 4 set). ", " (parity bit 4 not set). ") & _
    "A real UART pair should flag parity (bit 4). A virtual pair such as com0com never will, because it does not model parity.")
Close_Ends PortA, PortB
Exit Function
Fail:
LB_PARITY = Fmt("FAIL", NM, Err_Text())
Close_Ends PortA, PortB
End Function

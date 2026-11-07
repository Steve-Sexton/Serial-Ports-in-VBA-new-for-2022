Attribute VB_Name = "SERIAL_PORT_RIBBON"
Option Explicit

'----------------------------------------
' Change Com port number here if required
'
  Const Port_Number As Long = 1
'----------------------------------------

' All message and result variables are local to each callback: every port call can yield (DoEvents), so a second
' callback may run before the first has finished, and module-level variables would be overwritten under it.

Const Message_Box_Buttons As Long = vbInformation + vbOKOnly
'
Sub COM_PORT_CONTROL_1(Optional control As Variant)   'Callback for COM_PORT_START onAction

Dim Start_Result As Boolean, Port_Settings As String
Dim Message_Box_Text As String, Message_Box_Title As String, Message_Box_Result As Long

Message_Box_Title = "Start COM Port " & Port_Number

Start_Result = START_COM_PORT(Port_Number)

Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Port_Number) & vbCrLf & vbCrLf

Message_Box_Text = Port_Settings & "Start Result = " & Start_Result

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_CONTROL_2(Optional control As Variant)   'Callback for COM_PORT_STOP onAction

Dim Stop_Result As Boolean, Was_Busy As Boolean
Dim Message_Box_Text As String, Message_Box_Title As String, Message_Box_Result As Long

Message_Box_Title = "Stop COM Port " & Port_Number

Was_Busy = Port_Busy(Port_Number)                    ' captured before the call - the call itself requests the cancel

Stop_Result = STOP_COM_PORT(Port_Number)

If Was_Busy And Not Stop_Result Then
    Message_Box_Text = "Port busy: cancel requested, click Stop again"
Else
    Message_Box_Text = "Stop Result = " & Stop_Result
End If

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_DATA_1(Optional control As Variant)  'Callback for COM_PORT_CHECK onAction

Dim Characters_Waiting As Long, Was_Busy As Boolean, Port_Settings As String
Dim Message_Box_Text As String, Message_Box_Title As String, Message_Box_Result As Long

Message_Box_Title = "Check for data waiting"

Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Port_Number) & vbCrLf & vbCrLf

Was_Busy = Port_Busy(Port_Number)

Characters_Waiting = CHECK_COM_PORT(Port_Number)

If Was_Busy Then
    Message_Box_Text = Port_Settings & "Port busy: another operation is active, try again"
Else
    Message_Box_Text = Port_Settings & "Characters Waiting = " & Characters_Waiting
End If

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_DATA_2(Optional control As Variant)  'Callback for COM_PORT_READ onAction

Dim Characters_Read As String, Was_Busy As Boolean, Port_Settings As String
Dim Message_Box_Text As String, Message_Box_Title As String, Message_Box_Result As Long

Message_Box_Title = "COM Port Data Read Test"

Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Port_Number) & vbCrLf & vbCrLf

Was_Busy = Port_Busy(Port_Number)

Characters_Read = READ_COM_PORT(Port_Number, 20)

If Was_Busy Then
    Message_Box_Text = Port_Settings & "Port busy: another operation is active, try again"
Else
    Message_Box_Text = Port_Settings & "Characters Read = " & Characters_Read
End If

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_DATA_3(Optional control As Variant)  'Callback for COM_PORT_WRITE onAction

Dim Send_Result As Boolean, Send_Message As String, Port_Settings As String
Dim Message_Box_Text As String, Message_Box_Title As String, Message_Box_Result As Long

Message_Box_Title = "COM Port Data Send Test"

Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Port_Number) & vbCrLf & vbCrLf

Send_Message = Application.Name & " " & Application.Version & " @ " & Time & vbCrLf

Send_Result = SEND_COM_PORT(Port_Number, Send_Message)

Message_Box_Text = Port_Settings & IIf(Send_Result, "Characters Sent = " & Len(Send_Message), "Send Failed")

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_SIGNAL_1(Optional control As Variant)    'Callback for COM_PORT_RTS_ON onAction

Dim Signal_Result As Boolean, Port_Settings As String
Dim Message_Box_Text As String, Message_Box_Title As String, Message_Box_Result As Long

Message_Box_Title = "COM Port " & Port_Number & " - Request To Send ON"

Signal_Result = REQUEST_TO_SEND(Port_Number, 1)

Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Port_Number) & vbCrLf & vbCrLf

Message_Box_Text = Port_Settings & "Set RTS Result = " & Signal_Result

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

Sub COM_PORT_SIGNAL_2(Optional control As Variant)    'Callback for COM_PORT_RTS_OFF onAction

Dim Signal_Result As Boolean, Port_Settings As String
Dim Message_Box_Text As String, Message_Box_Title As String, Message_Box_Result As Long

Message_Box_Title = "COM Port " & Port_Number & " - Request To Send OFF"

Signal_Result = REQUEST_TO_SEND(Port_Number, 0)

Port_Settings = "Port Settings = " & GET_PORT_SETTINGS(Port_Number) & vbCrLf & vbCrLf

Message_Box_Text = Port_Settings & "Set RTS Result = " & Signal_Result

Message_Box_Result = MsgBox(Message_Box_Text, Message_Box_Buttons, Message_Box_Title)

End Sub

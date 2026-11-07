
'==================== TEST HARNESS (test builds only) ====================
Private Declare PtrSafe Sub T_SetLastError Lib "Kernel32.dll" Alias "SetLastError" (ByVal e As Long)

Private T_PATH As String
Private T_RAISE_CONFIG As Boolean
Private T_RAISE_PRINT As Boolean
Private T_RAISE_CLEAR As Boolean
Private T_RAISE_SIGNAL As Boolean
Private T_FAIL_QUERY As Boolean
Private T_FAIL_APPLY As Boolean
Private T_FAIL_CLEAR As Boolean
Private T_SeedDCB As DEVICE_CONTROL_BLOCK
Private T_Applied As DEVICE_CONTROL_BLOCK
Private T_ApplyCount As Long
Private T_ClearCalls As Long
Private T_QOUT_BUSY_CALLS As Long
Private T_HOOKS As Long
Private T_HOOK_MODE As Long
Private T_Max As Long
Private T_Arr() As Byte
Private T_LOG As String
Private T_PurgeMask As Long
Private T_CANCEL_ON_CLEAR As Boolean
Private T_LATE_PATH As String


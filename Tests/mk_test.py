"""Builds a test copy of a serial module. Never writes to the module itself.

The copy renames the Comm API declarations so tests.bas can replace them with stubs, opens a file
instead of a COM port, adds fault-injection hooks, and appends the tests.

Usage: mk_test.py <module.bas> <output.bas> <M|ND> <new|head> <tests.bas>
  new  = current code, with every test
  head = code from before commit cefd14c, without tests of functions it does not have
"""
import sys, re
src, dst, kind, mode, testsf = sys.argv[1:6]
d = open(src, 'rb').read().decode('latin-1').replace('\r\n', '\n')
tests = open(testsf, 'rb').read().decode('latin-1').replace('\r\n', '\n')

def rep(old, new, cnt=1):
    global d
    assert d.count(old) == cnt, (old, d.count(old))
    d = d.replace(old, new)

for fn in ['Query_Port_DCB', 'Apply_Port_DCB', 'Set_Com_Timers', 'Com_Port_Clear', 'Set_Com_Signal', 'Com_Port_Purge']:
    rep('Declare PtrSafe Function %s Lib' % fn, 'Declare PtrSafe Function %s_Real Lib' % fn)
rep('Device_Path = DEVICE_PREFIX & CStr(Port_Number)', 'Device_Path = T_PATH')
rep('Clean_Settings = CLEAN_PORT_SETTINGS(Port_Settings)\n', 'Clean_Settings = CLEAN_PORT_SETTINGS(Port_Settings)\nIf T_RAISE_CONFIG Then Err.Raise 5\n')
rep('Private Function OPERATION_INTERRUPTED(Port_Number As Long, Expected_Handle As LongPtr) As Boolean\n',
    'Private Function OPERATION_INTERRUPTED(Port_Number As Long, Expected_Handle As LongPtr) As Boolean\nT_HOOK Port_Number\n')
if kind == 'M':
    if mode == 'new':
        rep('On Error Resume Next                                       \' a diagnostic print must never itself raise\n',
            'On Error Resume Next\nIf T_RAISE_PRINT Then Err.Raise 6, , "T_RAISE_PRINT"\n')
    rep('Public Sub PRINT_DEBUG_TEXT(Module_Text As String, Result_Text As String, Message_Text As String)\n',
        'Public Sub PRINT_DEBUG_TEXT(Module_Text As String, Result_Text As String, Message_Text As String)\nT_LOG = T_LOG & Module_Text & "|" & Result_Text & "|" & Message_Text & vbLf\n' + ('' if mode == 'new' else 'If T_RAISE_PRINT Then Err.Raise 6, , "T_RAISE_PRINT"\n'))
    rep('Public Sub PRINT_SHOW_TEXT(Device_Text As String, Prefix_Text As String, Detail_Text As String, Result_Text As Variant)\n',
        'Public Sub PRINT_SHOW_TEXT(Device_Text As String, Prefix_Text As String, Detail_Text As String, Result_Text As Variant)\nT_LOG = T_LOG & Device_Text & "|" & Prefix_Text & "|" & Detail_Text & "|" & CStr(Result_Text) & vbLf\n')
decl = open(testsf.replace('tests.bas', 'tests_decl.bas'), 'rb').read().decode('latin-1').replace('\r\n', '\n')
rep('Public Function START_COM_PORT(', decl + '\nPublic Function START_COM_PORT(')
out = []
for line in tests.split('\n'):
    if line.startswith("'@"):
        mm = re.match(r"^'@(NEW)?(?:@)?(M|ND)? ?(.*)$", line)
        isnew, k, rest = mm.group(1), mm.group(2), mm.group(3)
        if isnew and mode != 'new':
            continue
        if k and k != kind:
            continue
        out.append(rest)
    else:
        out.append(line)
d = d.rstrip('\n') + '\n' + '\n'.join(out) + '\n'
open(dst, 'wb').write(d.replace('\n', '\r\n').encode('latin-1'))
print('built', dst)

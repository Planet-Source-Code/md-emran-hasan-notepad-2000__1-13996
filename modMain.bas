Attribute VB_Name = "modMain"
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const WM_GETTEXTLENGTH = &HE
' Win32 Declarations for DisableX
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE And KEY_ENUMERATE_SUB_KEYS And KEY_NOTIFY And KEY_CREATE_SUB_KEY And KEY_CREATE_LINK And KEY_SET_VALUE
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_OPTION_VOLATILE = 1
Private Const REG_SZ = 1

Private Const ERROR_SUCCESS = 0&

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

Public cEnumValues As Collection
Public cEnumKeys As Collection

Public Function RGCreateKey(hKey As Long, SubKey As String)
    Dim lngRet As Long
    Dim lngResult As Long
    Dim lngDis As Long
    

    lngRet = RegCreateKeyEx(hKey, SubKey, 0&, 0&, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lngResult, lngDis)
    lngRet = RegCloseKey(lngResult) 'Close key
End Function

Public Function RGDeleteKey(hKey As Long, SubKey As String)
    RegDeleteKey hKey, SubKey
End Function

Public Function RGSetKeyValue(hKey As Long, SubKey As String, ValueName As String, sValue As String)
    
    Dim lngRet As Long
    Dim lngResult As Long
    
    
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then
    
        RegSetValueEx lngResult, ValueName, 0, REG_SZ, ByVal sValue, Len(sValue)
        RegFlushKey lngResult
        RegCloseKey lngResult
    End If
End Function
Public Function Associate(Program As String, Extension As String, Description As String, Optional Icon As String)
    '** Description:
    '** Associate file with Cool Pad 2000
    RGCreateKey HKEY_CLASSES_ROOT, "." & Extension
    RGSetKeyValue HKEY_CLASSES_ROOT, "." & Extension, "", Extension & "file"
    
    RGCreateKey HKEY_CLASSES_ROOT, Extension & "file"
    RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell"
    If LCase(Extension) = "bat" Then
        RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\edit"
        RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\edit\command"
        RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file\shell\edit\command", "", Program & " " & "%1" 'Set file path
    Else
        RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\open"
        RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\open\command"
        RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file\shell\open\command", "", Program & " " & "%1" 'Set file path
    End If
    RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\DefaultIcon"
    
    RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file", "", Description 'Set file description
    RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file\DefaultIcon", "", Icon 'Set file icon
End Function


Public Sub DisableX(TheForm As Form)
    '** Description:
    '** Disable X in upper right corner of the form
    Dim lngMenu As Long
    lngMenu = GetSystemMenu(TheForm.hwnd, False)
    DeleteMenu lngMenu, 6, MF_BYPOSITION
End Sub


Public Sub MakeNormal(Handle As Long)
    SetWindowPos Handle, HWND_NOTTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MakeTopMost(Handle As Long)
    SetWindowPos Handle, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub


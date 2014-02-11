VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Rem ======================================= For registry stuff ===================================
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const REG_SZ = 1 ' Unicode nul terminated string

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal samDesired As Long, phkResult As Long) As Long

Public Function GetRegistry(ByVal key As Long, ByRef KeyPath As String, ByRef ValueName As String) As String
Dim lngRet As Long
Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long

On Error Resume Next

    Rem - Retrieve information about the key
    Call RegOpenKey(key, KeyPath, lngRet)
    lResult = RegQueryValueEx(lngRet, ValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        Rem - Create a buffer & extract the data
        strBuf = String(lDataBufSize, Chr$(0))
        lResult = RegQueryValueEx(lngRet, ValueName, 0, 0, ByVal strBuf, lDataBufSize)
        If lResult = 0 Then If Len(strBuf) > 0 Then GetRegistry = left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
    End If
    Call RegCloseKey(lngRet)

End Function

Public Sub SetRegistry(ByVal key As Long, ByVal Path As String, ByVal Keyname As String, ByVal KeyVal As String)
Dim lngHKey As Long

    Dim lngH As Long, lngRet As Long
    Call RegCreateKey(key, Path, lngH)
    lngRet = RegSetValueEx(lngH, Keyname, 0, REG_SZ, ByVal KeyVal, Len(KeyVal))
    Call RegCloseKey(lngH)
    If lngRet <> 0 Then Call ZenMB("Sorry, you need to be an administrator to write to the registry and change this setting.")
    
End Sub

Public Sub DelRegistry(ByVal key As Long, ByRef KeyPath As String, Optional ByRef ValueName As String = "DELETEKEY")
On Error Resume Next
Dim lngHKey As Long
'Const HKEY_CURRENT_USER = &H80000001
'Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
'Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
'Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
'Const STANDARD_RIGHTS_ALL = &H1F0000
'Const SYNCHRONIZE = &H100000
'Const READ_CONTROL = &H20000
'Const STANDARD_RIGHTS_READ = (READ_CONTROL)
'Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
'Const KEY_CREATE_LINK = &H20
'Const KEY_CREATE_SUB_KEY = &H4
'Const KEY_ENUMERATE_SUB_KEYS = &H8
'Const KEY_NOTIFY = &H10
'Const KEY_QUERY_VALUE = &H1
'Const KEY_SET_VALUE = &H2
'Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
'Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
'Const KEY_EXECUTE = (KEY_READ)
'Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Const KEY_ALL_ACCESS = 983103

    Dim Result As Long
    RegOpenKeyEx key, KeyPath, 0, KEY_ALL_ACCESS, Result
    If ValueName = "DELETEKEY" Then
        Rem - Delete the whole key
        RegDeleteKey Result, ""
    Else
        Rem - Delete the Value
        RegCreateKey lngHKey, KeyPath, Result
        RegDeleteValue Result, ValueName
    End If
    Rem - close the handle
    RegCloseKey Result

End Sub


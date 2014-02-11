Attribute VB_Name = "modShellExe"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_NORMAL = 1


Public Function ShellExe(ByVal FName As String, Optional Params As String = "", Optional ShowCommand As Long = SW_NORMAL, Optional ByVal StartDir As String = vbNullString) As Long
    
    ShellExe = ShellExecute(0, "open", FName, Params, StartDir, ShowCommand)
    If ShellExe <= 32 Then Call ZenMB("It appears that '" & FName & "' could not be found. Please ensure that the path to this file or folder is correct and accessible.")

End Function

Attribute VB_Name = "modLog"
Option Explicit
#If LOGMODE > 0 Then
Public LOG_FNum As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Public Sub LOG_Open()
    
    LOG_FNum = FreeFile
    Open App.Path & "\ZenLog.txt" For Append As #LOG_FNum
    
End Sub

Public Sub LOG_Close()
    Close #LOG_FNum
End Sub

Public Sub LOG_Write(ByVal Text As String)

    Print #LOG_FNum, Text

End Sub

Public Sub LOG_LastDllError()

    Call LOG_Write("    - LastDLLError = " & CStr(GetLastError))

End Sub
#End If

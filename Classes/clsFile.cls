VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text



Public Sub DoAction(ByRef prop As clsZenDictionary)
Dim strFile As String, strParam As String

    Rem - Process and parse the Action string
    strParam = prop("Param")
    strFile = prop("Action")
    
    If InStr(strFile, "%") > 0 Then strFile = InsertSpecialFolder(strFile)
    If InStr(strParam, "%") > 0 Then strParam = InsertSpecialFolder(strParam)
    
    If strFile <> "rundll32.exe" Then
        If prop("NewInstance") <> "True" Then
            If ShowExeWindow(strFile) <> 0 Then Exit Sub
        End If
    End If
    
    Call ShellExe(strFile, strParam, , GetStartDir(strFile, prop("ChangeDir")))
        
End Sub





Private Function GetStartDir(Optional ByVal strFile As String = vbNullString, Optional ByVal ChangeDir As String = vbNullString) As String
Dim strNewDir As String

On Error Resume Next

    Rem - Values for ChangeDir
    Rem - If ChangeDir = "No" - Stay in current dir
    Rem - If InStr(ChangeDir, "\") > 0 - Changes to the specified dir
    Rem - Else changes to App dir

    Rem - Only change dirs if strFile contains a path
    If ChangeDir <> "No" Then ' defaults to app's path
        If InStr(ChangeDir, "\") > 0 Then
            Rem - ChangeDir specifies the folder that we should switch to ...
            GetStartDir = ChangeDir
        ElseIf CBool(InStr(strFile, "\") > 0) Then
            Rem - Use the app path as the folder that we should switch to ...
            Dim k As Long
            For k = Len(strFile) To 0 Step -1
                If Mid(strFile, k, 1) = "\" Then
                    GetStartDir = left(strFile, k - 1)
                    Exit For
                End If
            Next k
        End If
    End If
    
    

End Function

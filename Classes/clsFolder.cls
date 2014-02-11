VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub DoAction(ByRef prop As clsZenDictionary)
Dim strFolder As String


    Rem - Process and parse the Action string
    strFolder = prop("Action")
    If InStr(strFolder, "%") > 0 Then strFolder = InsertSpecialFolder(strFolder)
    Call ShellExe("explorer.exe", strFolder)
    
End Sub





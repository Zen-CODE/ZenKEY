VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZenDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private prop As String
Public Function FromINI(ByVal FileName As String) As Boolean
Rem - Loads a ini file into a RichClient dictionary
Rem - A ini file comprises lines of two string separated by an equals sign. The left side is the key, the right side the value.
    
    If Len(Dir(FileName)) > 0 Then
        Dim fileNum As Long, line As String
        fileNum = FreeFile
        Open FileName For Input As #fileNum
            If LenB(prop) = 0 Then prop = "|"
            While Not EOF(fileNum)
                Line Input #fileNum, line
                prop = "|" & line & prop
            Wend
        Close fileNum
        FromINI = True
    End If
    
End Function

Public Property Get Item(ByRef key As String) As String
Attribute Item.VB_UserMemId = 0
    
    Item = Prop_Get(key, prop)
    
End Property

Public Property Let Item(ByRef key As String, ByVal Value As String)

    Call Prop_Set(key, Value, prop)
    
End Property

Public Sub ToINI(ByRef FileName As String)

    Dim fileNum As Long, k As Long
    Dim items() As String
    
    items = Split(prop, "|")
    
    fileNum = FreeFile
    Open FileName For Output As #fileNum
        For k = UBound(items) To 0 Step -1
            If Len(items(k)) > 0 Then Print #fileNum, items(k)
        Next k
    Close #fileNum

End Sub


Public Sub FromProp(ByRef newProp As String)
    
    If Len(prop) > 0 Then
        prop = newProp & Mid(prop, 2)
    Else
        prop = newProp
    End If
    

End Sub

Public Function Copy() As clsZenDictionary
    
    Dim zCopy As New clsZenDictionary
    Call zCopy.FromProp(ToProp)
    Set Copy = zCopy

End Function


Public Function ToProp() As String
    ToProp = prop
End Function

Public Function IsEmpty() As Boolean

    IsEmpty = CBool(Len(prop) < 3)
    
End Function

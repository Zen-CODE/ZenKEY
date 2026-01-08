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
Private dic As Dictionary
Public Function FromINI(ByVal FileName As String) As Boolean
Rem - Loads a ini file into a RichClient dictionary
Rem - A ini file comprises lines of two string separated by an equals sign. The left side is the key, the right side the value.
    
    If Len(Dir(FileName)) > 0 Then
        Dim fileNum As Long, line As String
        fileNum = FreeFile
        Open FileName For Input As #fileNum
            While Not EOF(fileNum)
                Line Input #fileNum, line
                Call FromProp(line)
            Wend
        Close fileNum
        FromINI = True
    End If
    
End Function

Public Property Get Item(ByRef key As String) As String
Attribute Item.VB_UserMemId = 0
    Item = dic(key) ' Will return null string by default
End Property

Public Property Let Item(ByRef key As String, ByVal Value As String)
' Use the second param ByVal to assist implicit string conversion

    If Len(Value) = 0 Then
        If dic.Exists(key) Then Call dic.Remove(key)
    Else
        If dic.Exists(key) Then dic(key) = Value Else Call dic.Add(key, Value)
    End If
    
End Property

Public Sub ToINI(ByRef FileName As String)

    Dim fileNum As Long, k As Long
    fileNum = FreeFile
    Open FileName For Output As #fileNum
        For k = dic.Count - 1 To 0 Step -1
            Print #fileNum, dic.KeyByIndex(k) & "=" & dic.ItemByIndex(k)
        Next k
    Close #fileNum

End Sub


Public Sub FromProp(ByRef Prop As String)
Dim k As Long
Dim items() As String, intPos As Integer

    items = Split(Prop, "|")
    'Move backards, so we preserve the ZenKEY behaviour of using the first occurence
    For k = UBound(items) To 0 Step -1 ' Split using pipes
        intPos = InStr(items(k), "=")
        If intPos > 0 Then 'Has a key/value pair
            Me(left(items(k), intPos - 1)) = Mid(items(k), intPos + 1) ' Use the zen Item prop to handle removal
        End If
    Next k

End Sub

Public Function ToProp() As String
Dim max As Long, k As Long
Dim arr() As String

    max = dic.Count - 1
    If max > -1 Then
        ReDim arr(0 To max)
        For k = 0 To max
            arr(k) = dic.KeyByIndex(k) & "=" & dic.ItemByIndex(k)
        Next k
        ToProp = "|" & Join(arr, "|") & "|"
    Else
        ToProp = vbNullString
    End If
    

End Function

Public Function Copy() As clsZenDictionary
Dim max As Long, k As Long, zCopy As clsZenDictionary

    Set zCopy = New clsZenDictionary
    max = dic.Count - 1
    For k = 0 To max
        zCopy(dic.KeyByIndex(k)) = dic.ItemByIndex(k)
    Next k
    Set Copy = zCopy

End Function

Public Function Count() As Long
    Count = dic.Count
End Function

Private Sub Class_Initialize()
    Set dic = New clsZenDictionary
    dic.StringCompareMode = TextCompare
End Sub



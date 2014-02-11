Attribute VB_Name = "modPropExtract"
Option Explicit



Public Function Prop_Get(ByRef PropName As String, ByRef PropString As String) As String
Const strSep As String = "|"
Dim parts() As String, lngPos As Long

    parts = Split(PropString, strSep & PropName & "=", , vbTextCompare)
    If UBound(parts) > 0 Then
        lngPos = InStr(parts(1), strSep)
        If lngPos > 0 Then
            Prop_Get = Left$(parts(1), lngPos - 1)
        Else
            Prop_Get = parts(1)
        End If
    End If
    
End Function
 
Public Sub Prop_Set(ByRef PropName As String, ByRef Value As String, ByRef PropString As String)
Dim lngPos As Long
Dim lngEnd As Long
Const strSep As String = "|"
Dim key As String

    key = strSep & PropName & "="
    lngPos = InStr(1, PropString, key, vbTextCompare)
    If lngPos > 0 Then
        Rem - Replace the current property
        lngEnd = InStr(Mid$(PropString, lngPos + 1), strSep)
        If Len(Value) > 0 Then
            PropString = Left$(PropString, lngPos - 1) & key & Value & Mid$(PropString, lngEnd + lngPos)
        Else
            PropString = Left$(PropString, lngPos - 1) & Mid$(PropString, lngEnd + lngPos)
        End If
    ElseIf Len(Value) > 0 Then
        Rem - Add a new property
        PropString = key & Value & IIf(LenB(PropString) > 0, PropString, strSep)
    End If

End Sub


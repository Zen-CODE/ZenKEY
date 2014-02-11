Attribute VB_Name = "modActDescrip"
Option Explicit
Option Compare Text
Public Type StringArray
    Action() As String
End Type
Public Actions() As StringArray
Public Function Actions_GetActIndex(ByVal Act As String, ByVal ClassIndex As Long) As Long
Dim max As Long
Dim k As Long

    max = UBound(Actions(ClassIndex).Action())
    For k = 1 To max
        If Act = Prop_Get("Action", Actions(ClassIndex).Action(k)) Then
            Actions_GetActIndex = k - 1
            Exit Function
        End If
    Next k


End Function

Public Sub Actions_Load()
On Error GoTo ErrorTrap

    Dim FNum As Long, lngGroup As Long, lngItem As Long
    Dim strLine As String
    
    FNum = FreeFile
    lngGroup = -1
    Open App.Path & "\Actions.ini" For Input As #FNum
    While Not EOF(FNum)
        Line Input #FNum, strLine
        Select Case Prop_Get("Class", strLine)
            Case "EndGroup"
                'lngGroup = lngGroup + 1
            Case ""
                lngItem = lngItem + 1
                ReDim Preserve Actions(lngGroup).Action(0 To lngItem)
                Actions(lngGroup).Action(lngItem) = strLine
            Case Else
                Rem - New group
                lngGroup = lngGroup + 1
                lngItem = 0
                ReDim Preserve Actions(0 To lngGroup)
                ReDim Actions(lngGroup).Action(0 To 0)
                Actions(lngGroup).Action(0) = strLine
        End Select
    Wend
    Close FNum

    Exit Sub
    
ErrorTrap:
    Call MsgBox(CStr(Err.Number) & ", " & Err.Description & " in Actions_Init", vbInformation)
    Err.Clear

End Sub
Public Function Actions_GetClassIndex(ByVal ActClass As String) As String
Dim k As Integer

    If Len(ActClass) = 0 Then ActClass = "File"
    For k = UBound(Actions()) To 0 Step -1
        If ActClass = Prop_Get("Class", Actions(k).Action(0)) Then
            Actions_GetClassIndex = k
            Exit Function
        End If
    Next k

End Function
Public Function Actions_GetDescrip(ByRef prop As clsZenDictionary) As String
Dim strClass As String, strAct As String

    strClass = prop("Class")
    strAct = prop("Action")

Dim lngGroup As Long, lngItem As Long
Dim GMax As Long, IMax As Long
Dim k As Long, i As Long

    'TODO: Use ZenDic lookup?
    GMax = UBound(Actions())
    For k = 0 To GMax
        If Prop_Get("Class", Actions(k).Action(0)) = strClass Then
            IMax = UBound(Actions(k).Action())
            For i = 0 To IMax
                If Prop_Get("Action", Actions(k).Action(i)) = strAct Then
                    Actions_GetDescrip = Prop_Get("Caption", Actions(k).Action(i))
                    Exit Function
                End If
            Next i
        End If
    Next k
    Actions_GetDescrip = strAct

End Function

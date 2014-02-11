VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub DoAction(ByVal Prop As String)
Dim strCriteria As String

    'strCriteria = InputBox("Search for what?", "ZenKEY Internet search", vbNullString)
    MainForm.Show
    MainForm.SetFocus
    strCriteria = InputBox("Search for what?", Prop_Get("Caption", Prop))
    If Prop_Get("HideForm", ZenProperties) = "True" Then MainForm.Visible = False
    
    If Len(strCriteria) > 0 Then
        Dim strSearch As String, intPos As Integer
        
        strSearch = Prop_Get("Action", Prop)
        intPos = InStr(strSearch, "<Criteria>")
        strSearch = Left(strSearch, intPos - 1) & PrepCriteria(strCriteria) & Mid(strSearch, intPos + 10)
        Call ShellExe(strSearch)
    End If

End Sub

Private Function PrepCriteria(ByVal Criteria As String) As String
Dim k As Long

    PrepCriteria = Criteria
    k = InStr(PrepCriteria, " ")
    While k > 0
        Mid(PrepCriteria, k, 1) = "+"
        k = InStr(PrepCriteria, " ")
    Wend

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub DoAction(ByVal Action As String)
    'Call MsgBox("Stub DoAction fired! - " & Action, vbInformation)
    Forms(0).txtTimer.Text = Forms(0).txtTimer.Text & Action & vbCrLf
End Sub

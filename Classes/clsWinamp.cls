VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private lngWinamp As Long
Private Const WM_COMMAND = &H111


Public Sub DoAction(ByRef prop As clsZenDictionary)
Dim lngVal As Long

    lngVal = CLng(Val(prop("Action")))
    Rem - Send the message to winamp
    If Winamp_Link Then Call PostMessage(lngWinamp, WM_COMMAND, lngVal, 0)
    Rem - Prevent the holding down of FFwd Rew from overloading & crashing winamp
    If lngVal = 40060 Or lngVal = 40061 Then Call Sleep(100)
        
End Sub
Public Function Winamp_Link() As Boolean

    Rem - Check if the Handle is valid
    Winamp_Link = IsWindow(lngWinamp)
    If Not Winamp_Link Then
        Rem - Okay, see if we can find the handle....
        lngWinamp = FindWindow("Winamp v1.x", vbNullString)
        Winamp_Link = IsWindow(lngWinamp)
        If Not Winamp_Link Then Call ZenMB("Winamp does not appear to be running. " & _
            "Please start Winamp (or compatible) and try again.")
    End If
    
End Function

Private Sub Class_Initialize()
    lngWinamp = FindWindow("Winamp v1.x", vbNullString)
End Sub



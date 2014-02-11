VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Sub DoAction(ByRef prop As clsZenDictionary)
Dim lngCommand As Long
Dim strWClass As String
Dim lngHWnd As Long
Const WM_APPCOMMAND = &H319

    Rem - Get the media command
    lngCommand = CLng(Val(prop("Action")))
    
    Select Case strWClass
        Case vbNullString, "Active": lngHWnd = GetForegroundWindow
        Case Else
            strWClass = prop("Window Class")
            lngHWnd = FindWindow(strWClass, vbNullString)
    End Select
    
    Rem - Sends to active app & broadcasts. Cool Cascades nicely
    Const HWND_BROADCAST = &HFFFF
    lngHWnd = PostMessage(lngHWnd, WM_APPCOMMAND, HWND_BROADCAST, ByVal (lngCommand * &H10000))
    If lngHWnd <> 1 Then Call ZenMB("It appears that there are no active applications that that can handle this command. Please launch a media-enable application and try again.")

End Sub


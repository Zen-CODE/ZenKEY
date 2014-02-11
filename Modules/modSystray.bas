Attribute VB_Name = "modSystray"
Option Explicit

Rem =========================================================================
Rem - For restoring the system tray icon if explorer is restarted
Public lngTaskBarMsg As Long
Rem ============  For system tray icon
'Private Type NOTIFYICONDATA
'    cbSize As Long
'    hwnd As Long
'    uId As Long
'    uFlags As Long
'    ucallbackMessage As Long
'    HIcon As Long
'    szTip As String * 64
'End Type
Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  HIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  'guidItem As Guid
End Type
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Const APP_SYSTRAY_ID = 999

Public Sub Balloon_Show(ByVal Caption As String)
Const NIF_INFO = &H10
Const NIM_MODIFY = &H1
'Const NIM_DELETE = &H2
Dim nid As NOTIFYICONDATA
Const Limit = 1000
Static strMessage As String
   
    With nid
        .cbSize = Len(nid)
        .hwnd = MainForm.hwnd
        .uID = APP_SYSTRAY_ID
        .uFlags = NIF_INFO
        If Len(Caption) = 0 Then
            Rem - Fired from the main form timer. Countdown to clearing
            strMessage = vbNullString
        Else
            Rem - A message should be added to the lsit
            Select Case Len(strMessage)
                Case 0
                    strMessage = Caption
                Case Is > Limit
                    strMessage = Caption & vbCr & vbTab & left(strMessage, Limit)
                Case Else
                    strMessage = Caption & vbCr & vbTab & strMessage
            End Select
        End If
        
        Rem - Hide if it is already shown
        MainForm.tmrBalloon.Enabled = False ' Disable/reset the timer interval
        Call Shell_NotifyIcon(NIM_MODIFY, nid) 'Remove the previous
        If Len(strMessage) > 0 Then
            Rem - Okay, there are messages. Show em!
            .szInfoTitle = "--- ZenKEY ---" & vbNullChar
            .szInfo = strMessage & vbNullChar
            .dwInfoFlags = 1 ' 0 = None, 1 = Warning, 2 = Information, 3 = Error
            Call Shell_NotifyIcon(NIM_MODIFY, nid)
            MainForm.tmrBalloon.Interval = 3000
            MainForm.tmrBalloon.Enabled = True
        End If
    End With
    
End Sub

Public Function Systray_Add(ByRef TheForm As Object, ByVal HIcon As Long, ByVal ToolTip As String) As Boolean
'Public Function Systray_Add(ByRef TheForm As Form, ByVal HIcon As Long, ByVal ToolTip As String) As Boolean
Const NIM_ADD = &H0
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4
Const WM_RBUTTONDOWN = &H204
        
    Dim TrayI As NOTIFYICONDATA
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = TheForm.hwnd
    'TrayI.uID = 1& 'Application-defined identifier of the taskbar icon
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.uID = APP_SYSTRAY_ID
    
    Rem - Set the callback message
    TrayI.uCallbackMessage = WM_RBUTTONDOWN 'WM_LBUTTONDOWN
    TrayI.HIcon = HIcon 'TheForm.Icon ''Set the picture (must be an icon!)
    TrayI.szTip = ToolTip & Chr$(0) 'Set the tooltiptext
    Systray_Add = CBool(0 <> Shell_NotifyIcon(NIM_ADD, TrayI))

End Function




Public Sub Systray_Del(ByRef TheForm As Form)
Const NIM_DELETE = &H2
Dim TrayI As NOTIFYICONDATA

    Rem - Remove the icon
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = TheForm.hwnd
    TrayI.uID = APP_SYSTRAY_ID '1&
    Call Shell_NotifyIcon(NIM_DELETE, TrayI)
    
End Sub



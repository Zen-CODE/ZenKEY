Attribute VB_Name = "modMenu"
Option Explicit
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public MENU_OldProc As Long  ' pointer to Form1's previous window procedure

Public Function WindowMenuProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Rem - The following function is called by clsDynamicMenu when a menu item is clicked
Const WM_COMMAND = &H111

Const WM_USER As Long = &H400
Const NIN_BALLOONSHOW = (WM_USER + 2)
Const NIN_BALLOONHIDE = (WM_USER + 3)
Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Const NIN_BALLOONUSERCLICK = (WM_USER + 5)
Const WM_RBUTTONDOWN = &H204

    Select Case uMsg
        Case WM_COMMAND
            Rem - A local event has been fired
            Select Case wParam
                Case Is > 0 ' 0 To 10000
                    Call MainForm.DoAction(zenDic("Action", "FIREACTION", "INDEX", wParam))
                    WindowMenuProc = 0
                Case Else
                    WindowMenuProc = CallWindowProc(MENU_OldProc, hwnd, uMsg, wParam, lParam)
            End Select
        Case lngTaskBarMsg
            Rem - Hook for explorer being restarted
            Rem - Refresh System tray icons
            If MainForm.Icon.Handle <> 0 Then Call Systray_Add(MainForm, MainForm.Icon, ZenKEYCap)
            Dim k As Long
            For k = 0 To ST_Count - 1
                Call ST_TrayForms(k).SendToTray
            Next k
        Case WIN_ZKAction
            Rem - A notification that an action should be fired.
            Dim zAction As New clsZenDictionary
            Call zAction.FromProp(Registry.GetRegistry(HKCU, "SOFTWARE\ZenCODE\ZenKEY", "Action"))
            Call ZK_GetObject(zAction("Class")).DoAction(zAction)
        Case WM_RBUTTONDOWN
            Select Case lParam
                Case NIN_BALLOONHIDE, NIN_BALLOONTIMEOUT, NIN_BALLOONSHOW, NIN_BALLOONUSERCLICK
                    Rem - Let the message die.
                Case Else
                    Rem - If this is some other message, let the previous procedure handle it.
                    WindowMenuProc = CallWindowProc(MENU_OldProc, hwnd, uMsg, wParam, lParam)
            End Select

        Case Else
            Rem - If this is some other message, let the previous procedure handle it.
            WindowMenuProc = CallWindowProc(MENU_OldProc, hwnd, uMsg, wParam, lParam)
    End Select
    
End Function



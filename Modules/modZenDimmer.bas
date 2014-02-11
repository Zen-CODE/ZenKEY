Attribute VB_Name = "modZenDimmer"
Option Explicit
Public Const MONITORINFOF_PRIMARY = &H1
Public Const MONITOR_DEFAULTTONEAREST = &H2
Public Const MONITOR_DEFAULTTONULL = &H0
Public Const MONITOR_DEFAULTTOPRIMARY = &H1
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type
Public Type POINT
    X As Long
    Y As Long
End Type
Public Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Public Declare Function MonitorFromPoint Lib "user32.dll" (ByVal X As Long, ByVal Y As Long, ByVal dwFlags As Long) As Long
Public Declare Function MonitorFromRect Lib "user32.dll" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
Public Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
Public Declare Function EnumDisplayMonitors Lib "user32.dll" (ByVal hdc As Long, ByRef lprcClip As Any, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' For transparency
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public lngTrans As Long
Public Sub CloseApp()
Dim k As Long

    Call ShowCursor(1)
    For k = Forms.Count To 1 Step -1
        Unload Forms(k - 1)
    Next k

End Sub


Public Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, ByVal dwData As Long) As Long
Const HWND_TOPMOST = -1
'Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Dim MI As MONITORINFO, R As RECT
Dim Form As New frmDimmer
        

    MI.cbSize = Len(MI)
    GetMonitorInfo hMonitor, MI
    With Form
        Call .Move(.ScaleX(MI.rcMonitor.Left, vbPixels, vbTwips), _
            .ScaleY(MI.rcMonitor.Top, vbPixels, vbTwips), _
            .ScaleX(MI.rcMonitor.Right - MI.rcMonitor.Left, vbPixels, vbTwips), _
            .ScaleY(MI.rcMonitor.Bottom - MI.rcMonitor.Top, vbPixels, vbTwips))
        Call SetTrans(.hwnd, lngTrans)
        Call .Show
        Call SetWindowPos(.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
    End With
    'Continue enumeration
    MonitorEnumProc = 1
End Function


Private Function SetTrans(ByVal lngHWnd As Long, ByVal TransLevel As Long) As Boolean
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Dim Ret As Long, lngTrans As Long
    
    On Error Resume Next
    Rem - Set the window style to 'Layered'
    Ret = GetWindowLong(lngHWnd, GWL_EXSTYLE)
    Rem - 255 = Totally opague
    lngTrans = 255 * (TransLevel / 100)
    If lngTrans < 1 Then
        Call SetLayeredWindowAttributes(lngHWnd, 0, 255, LWA_ALPHA)
        Call SetWindowLong(lngHWnd, GWL_EXSTYLE, Ret And (Not WS_EX_LAYERED))
    Else
        Call SetWindowLong(lngHWnd, GWL_EXSTYLE, Ret Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(lngHWnd, 0, lngTrans, LWA_ALPHA)
    End If
            
    SetTrans = CBool(Err.Number = 0)
        
        
End Function
Public Sub Main()
        
    ' Get the transarency level
    lngTrans = Val(Prop_Get("Trans", Command$)) Mod 100
    If lngTrans = 0 Then lngTrans = 60 Else lngTrans = 100 - lngTrans

    Call ShowCursor(0)
    EnumDisplayMonitors ByVal 0&, ByVal 0&, AddressOf MonitorEnumProc, ByVal 0&
   

End Sub

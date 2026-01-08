VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Rem =====================  For minimizing all windows
Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Rem ===================== For Shrinking/Growing =====================
'Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Type WindowRecord
    LastHWnd As Long
    LastWinPlc As WINDOWPLACEMENT
    LastAction As String
    Properties As String
End Type
'Private Const SW_NORMAL = 1
Private History() As WindowRecord ' An array 1 to HistDepth big
Private HistCurrent As Long
'Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private HistDepth As Long
'Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public colSelected As Collection
Private booMulti As Boolean
Rem - For 'Pin to desktop' functionality
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Sub MakeTransparent(ByVal hwnd As Long, Optional ByVal Undo As Boolean = False, Optional ByVal TransColor As Variant)
Rem - If TransColor is not specified, the default window brackground colour is used
Dim Ret As Long
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const LWA_COLORKEY = &H1
Const COLOR_WINDOW = 5 'Windows background

    'Set the window style to 'Layered'
    Ret = GetWindowLong(hwnd, GWL_EXSTYLE)
    If Not Undo Then
        Ret = Ret Or WS_EX_LAYERED
        Call SetWindowLong(hwnd, GWL_EXSTYLE, Ret)
        If IsMissing(TransColor) Then TransColor = GetSysColor(COLOR_WINDOW)
        Call SetLayeredWindowAttributes(hwnd, TransColor, 255, LWA_ALPHA Or LWA_COLORKEY)
    Else
        Ret = Ret And Not WS_EX_LAYERED
        Call SetWindowLong(hwnd, GWL_EXSTYLE, Ret)
        Call RefreshWindow(hwnd)
    End If

End Sub
Public Sub DoAction(ByRef Prop As clsZenDictionary)
Dim lngHWnd As Long
Const SW_MAXIMIZE = 3
Const SW_MINIMIZE = 6
Const SW_RESTORE = 9
Dim booUndo As Boolean
Dim Action As String

    Action = Prop("Action")
    If p_UnLoggedAct(Action) Then Exit Sub ' Exit if the action was an 'Undo'
    If Not booMulti Then
        If colSelected.Count > 0 Then
            If p_IsMultiCompat(Action) Then
                Rem -------------------------------------------------------------------------------------------------------
                Rem - There are multiple windows selected. Perform on the action each
                Rem -------------------------------------------------------------------------------------------------------
                Dim k As Long
                booMulti = True
                For k = colSelected.Count To 1 Step -1
                    Prop("HWnd") = colSelected.Item(k)
                    Call DoAction(Prop)
                    Rem - In case the windows have been moved, update their new location
                Next k
                Prop("Hwnd") = vbNullString ' Prevent later usage
                booMulti = False
                Exit Sub
            End If
        End If
    End If
        
    Rem -------------------------------------------------------------------------------------------------------
    Rem - Perform the action
    Rem -------------------------------------------------------------------------------------------------------

    Rem - Okay. Lets do the action then.....
    Call History_Add(lngHWnd, Action)
    Select Case Action
        Case "DESKTOP"
            lngHWnd = 1
            Call ShellExe(App.Path & "\Showde~1.scf")
        Case "RESTOREALL"
            lngHWnd = 1
            Call EnumWindows(AddressOf RestoreAll, ByVal 0&)
        Case "MOVEALLUP", "MOVEALLDOWN", "MOVEALLRIGHT", "MOVEALLLEFT"
            lngHWnd = 1
            Select Case Action
                Case "MOVEALLUP"
                    Call MoveAllWindows(0, -WIN_Shift, False)
                Case "MOVEALLDOWN"
                    Call MoveAllWindows(0, WIN_Shift, False)
                Case "MOVEALLLEFT"
                    Call MoveAllWindows(-WIN_Shift, 0, False)
                Case "MOVEALLRIGHT"
                    Call MoveAllWindows(WIN_Shift, 0, False)
            End Select
        Case "MINALL"
            lngHWnd = 1
            Call EnumWindows(AddressOf MinimizeAll, ByVal 0&)
        Case "RestoreFromTray", "RestoreMe"
            Rem - ResoreMe is send by the click on the system tray icon. It sets the 'FormIndex' property in the string
            If ST_Count > 0 Then
                Dim lngIndex As Long
                If InStr(Prop, "FormIndex") > 0 Then lngIndex = CLng(Prop("FormIndex")) Else lngIndex = ST_Count - 1
                lngHWnd = p_FromTray(lngIndex)
            Else
                Call ZenMB("There are no more ZenKEY items in the system tray.")
            End If
        Case "UNDOLASTAWT"
            If settings("AutoTrans") = "True" Then
                If AWT_LastTrans <> 0 Then
                    Call SetTrans(AWT_LastTrans, -1)
                Else
                    Call ZenMB("Sorry, but it appears that no window has of yet been made transparent.")
                End If
            Else
                Call ZenMB("Sorry, but the 'Auto-window transparency' function is not enabled.")
            End If
        Case "WUMTOP" ' Window under mouse to top
            Dim pt As POINTAPI
            Call GetCursorPos(pt)
            lngHWnd = WindowFromPoint(pt.X, pt.Y)
            If WindowIsUsable(lngHWnd) Then
                Call SetWinPos(lngHWnd, HWND_TOP, True)
            Else
                lngHWnd = 0
            End If
        Case Else
            lngHWnd = Val(Prop("Hwnd"))
            If lngHWnd = 0 Then lngHWnd = GetWindowToUse
            If lngHWnd <> 0 Then
                Call GetWindowPlacement(lngHWnd, History(HistCurrent).LastWinPlc)
                History(HistCurrent).LastHWnd = lngHWnd
                Select Case Action
                    Case "TILEL2VERT", "TILEL2HOR", "TILEL2VERTNR", "TILEL2HORNR"
                        Rem - Tile last 2 windows vertically, horizontally....
                        If ActiveWindow(1) = 0 Then
                            Call ZenMB("Sorry, but you must have a window history depth of at least 2 to use this function.")
                        Else
                            Rem - Finalise this windows details for the history stack
                            History(HistCurrent).LastHWnd = ActiveWindow(0)
                            Call GetWindowPlacement(History(HistCurrent).LastHWnd, History(HistCurrent).LastWinPlc)
                            Call History_Add(ActiveWindow(1), Action) ' Add the next window to the stack.
                            History(HistCurrent).LastHWnd = ActiveWindow(1)
                            Call GetWindowPlacement(History(HistCurrent).LastHWnd, History(HistCurrent).LastWinPlc)
                            Select Case Action
                                Case "TILEL2VERT"
                                    Call ResizeWin(zenDic("Action", "LEFTHALF"), ActiveWindow(0))
                                    Call ResizeWin(zenDic("Action", "RIGHTHALF"), ActiveWindow(1))
                                Case "TILEL2HOR"
                                    Call ResizeWin(zenDic("Action", "TOPHALF"), ActiveWindow(0))
                                    Call ResizeWin(zenDic("Action", "BOTHALF"), ActiveWindow(1))
                                Case "TILEL2VERTNR"
                                    Call ResizeWin(zenDic("Action", "CENTERRIGHT"), ActiveWindow(0))
                                    Call ResizeWin(zenDic("Action", "CENTERLEFT"), ActiveWindow(1))
                                Case "TILEL2HORNR"
                                    Call ResizeWin(zenDic("Action", "CENTERTOP"), ActiveWindow(0))
                                    Call ResizeWin(zenDic("Action", "CENTERBOT"), ActiveWindow(1))
                            End Select
                        End If
                    Case "TILEL4", "TILEL4NR"
                        Rem - Tile last 4 windows
                        If ActiveWindow(3) = 0 Then
                            Call ZenMB("Sorry, but you must have a window history depth of at least 4 to use this function.")
                        Else
                            Rem - Finalise this windows details for the history stack
                            History(HistCurrent).LastHWnd = ActiveWindow(0)
                            Call GetWindowPlacement(History(HistCurrent).LastHWnd, History(HistCurrent).LastWinPlc)
                            For lngHWnd = 1 To 3
                                Call History_Add(ActiveWindow(lngHWnd), Action) ' Add the next window to the stack.
                                History(HistCurrent).LastHWnd = ActiveWindow(lngHWnd)
                                Call GetWindowPlacement(History(HistCurrent).LastHWnd, History(HistCurrent).LastWinPlc)
                            Next lngHWnd
                            Select Case Action
                                Case "TILEL4"
                                    For lngHWnd = 0 To 3
                                        Call ResizeWin(zenDic("Action", Choose(lngHWnd + 1, "TOPLEFT", "TOPRIGHT", "BOTTOMLEFT", "BOTTOMRIGHT")), ActiveWindow(lngHWnd))
                                    Next lngHWnd
                                Case "TILEL4NR"
                                    For lngHWnd = 0 To 3
                                        Call ResizeWin(zenDic("Action", Choose(lngHWnd + 1, "CENTERTOPLEFT", "CENTERTOPRIGHT", "CENTERBOTLEFT", "CENTERBOTRIGHT")), ActiveWindow(lngHWnd))
                                    Next lngHWnd
                            End Select
                        End If
                    Case "AOT"
                        Call SetWinPos(lngHWnd, HWND_TOPMOST, False)
                    Case "NAOT"
                        Call SetWinPos(lngHWnd, HWND_NOTOPMOST, True)
                    Case "MINALLBUTTOP"
                        Call EnumWindows(AddressOf MinimizeAll, ByVal 0&)
                        Call SetWindowPlacement(lngHWnd, History(HistCurrent).LastWinPlc)
                    Case "MINIMIZE"
                        Call ShowWindow(lngHWnd, SW_MINIMIZE)
                    Case "MAXIMIZE"
                        Call ShowWindow(lngHWnd, SW_MAXIMIZE)
                    Case "TOP"
                        Call SetWinPos(lngHWnd, HWND_TOP, True)
                    Case "BOTTOM"
                        Call SetWinPos(lngHWnd, HWND_BOTTOM, False)
                    Case "NORMAL"
                        Call ShowWindow(lngHWnd, SW_NORMAL)
                    Case "KILLACTIVE"
                        Const WM_CLOSE = &H10
                        Call PostMessage(lngHWnd, WM_CLOSE, 0&, ByVal 0&)
                    Case "ENABLE"
                        Call EnableWindow(lngHWnd, 1)
                    Case "DISABLE"
                        Call EnableWindow(lngHWnd, 0)
                        'Call ZenMB("Due to the dangerous nature of disabling a window, this window will be automatically enabled when you perform another ZenKEY window operation.")
                    Case "SENDTOTRAY"
                        Call p_ToTray(lngHWnd)
                        booUndo = True
                    Case "ICONIFY"
                        Call Icon_Make(lngHWnd, True)
                    Case "ICONIFYCREATE"
                        Call Icon_Make(lngHWnd, False)
                    Case "ICONIFYCLEAR"
                        Call Icon_Remove(lngHWnd)
                    Case "ICONFLUSHDEAD"
                        Call Icon_FlushDead
                    Case "ICONREACQUIRE"
                        Call Icon_Reacquire
                    Case "ICONFLUSHALL"
                        Call Icon_RemoveAll
                    Case "PINTODESKTOP"
                        Call PinToDesktop(True, lngHWnd)
                    Case "UNPINFROMDESKTOP"
                        Call PinToDesktop(False, lngHWnd)
                    Case "MAKETRANS"
                        Call MakeTransparent(lngHWnd)
                    Case "MAKETRANSUNDO"
                        Call MakeTransparent(lngHWnd, True)
                    Case "SHOWINFO"
                        Call ShowInfo(lngHWnd)
                    Case Else
                        If InStr(Action, "SETTRANSPARENCY") > 0 Then
                            Call SetTrans(lngHWnd, Val(Mid(Action, 17)))
                        Else
                            lngHWnd = ResizeWin(Prop, lngHWnd)
                        End If
                End Select
                Rem - For when the Exe is launched/brought to front, remove any icons from tray
                If Not booUndo Then
                    Dim rtn As Long
                    For rtn = ST_Count - 1 To 0 Step -1 ' Sysem tray icons
                        If ST_TrayForms(rtn).OwnerForm = lngHWnd Then Call p_FromTray(rtn)
                    Next rtn
                End If
            End If
    End Select
    
End Sub




Public Function ResizeWin(ByRef Prop As clsZenDictionary, ByVal lngWin As Long) As Long
Dim strMessage As String
Dim RecTan As RECT
Const MIN = 50
Dim strAction As String
        
    strAction = Prop("Action")

    Rem - Ensure the window is displayed normally
    If IsIconic(lngWin) Or IsZoomed(lngWin) Then Call ShowWindow(lngWin, SW_NORMAL)
    
    If strAction = "ROLLUP" Or strAction = "ROLLDOWN" Then
        Rem - Names preserved for legacy
        Rem - Toggle the window Rollup status
        If IsWindowVisible(lngWin) Then
            strMessage = ROL_Toggle(lngWin, CBool(strAction = "ROLLUP"))
            Call ROL_RemoveNorm(lngWin) ' Remove if normalised
            ROL_Check = False ' Prevent immediate unrolling
        Else
            strMessage = "Sorry, the selected window is not visible."
        End If ' If IsWindowVisible(lngWin) Then
    ElseIf GetWindowRect(lngWin, RecTan) <> 1 Then
        strMessage = "Unable to determine this window's placement for some reason. How strange?"
    Else
        Rem - Normal windows resizing
        
        Dim lngWidth As Long, lngHeight As Long
        Dim lngAnchor As Long, rctDesktop As RECT
        
                                                
        Rem - Use the desktop
        With RecTan
            
            lngWidth = .Right - .left
            lngHeight = .Bottom - .Top
            Select Case settings("WinAnchor")
                Case "", "Center": lngAnchor = 0
                Case "TopLeft": lngAnchor = 1
            End Select
        
            Select Case strAction
                Case "PLACERECT"
                    .left = Val(Prop("Left"))
                    .Right = Val(Prop("Right"))
                    .Top = Val(Prop("Top"))
                    .Bottom = Val(Prop("Bottom"))
                Case "SHRINK"
                    If ((.Bottom - .Top) < 2 * WIN_Shift + 1) Or ((4 / 3) * (.Right - .left) < 2 * WIN_Shift + 1) Then
                        strMessage = "Sorry, but this window is too small to shrink any further"
                    Else
                        Select Case lngAnchor
                            Case 0 ' Center
                                .Bottom = .Bottom - 0.5 * WIN_Shift
                                .Top = .Top + 0.5 * WIN_Shift
                                .Right = .Right - 0.5 * WIN_Shift * (4 / 3)
                                .left = .left + 0.5 * WIN_Shift * (4 / 3)
                            Case 1 ' Top Left
                                .Right = .Right - WIN_Shift * (4 / 3)
                                .Bottom = .Bottom - WIN_Shift
                        End Select
                    End If
                Case "GROW"
                    Select Case lngAnchor
                        Case 0 ' Center
                            .Bottom = .Bottom + 0.5 * WIN_Shift
                            .Top = .Top - 0.5 * WIN_Shift
                            .Right = .Right + 0.5 * WIN_Shift / (3 / 4)
                            .left = .left - 0.5 * WIN_Shift / (3 / 4)
                        Case 1 ' Top left
                            .Right = .Right + WIN_Shift * (4 / 3)
                            .Bottom = .Bottom + WIN_Shift
                    End Select
                Case "FATTEN"
                    Select Case lngAnchor
                        Case 0 ' Center
                            .left = .left - 0.5 * WIN_Shift
                            .Right = .Right + 0.5 * WIN_Shift
                        Case 1 ' Top Left
                            .Right = .Right + WIN_Shift
                    End Select
                Case "THIN"
                    If .Right - .left < 2 * WIN_Shift + 1 Then
                        strMessage = "Sorry, but this window is to thin to thin."
                    Else
                        Select Case lngAnchor
                            Case 0 ' Center
                                .left = .left + 0.5 * WIN_Shift
                                .Right = .Right - 0.5 * WIN_Shift
                            Case 1 ' Top left
                                .Right = .Right - WIN_Shift
                        End Select
                    End If
                Case "HEIGHTEN"
                    Select Case lngAnchor
                        Case 0 ' Center
                            .Top = .Top - 0.5 * WIN_Shift
                            .Bottom = .Bottom + 0.5 * WIN_Shift
                        Case 1 ' Top left
                            .Bottom = .Bottom + WIN_Shift
                    End Select
                Case "SHORTEN"
                    If .Bottom - .Top < 2 * WIN_Shift + 1 Then
                        strMessage = "Sorry, but this window is too short to shorten."
                    Else
                        Select Case lngAnchor
                            Case 0 ' Center
                                .Top = .Top + 0.5 * WIN_Shift
                                .Bottom = .Bottom - 0.5 * WIN_Shift
                            Case 1 ' Teop left
                                .Bottom = .Bottom - WIN_Shift
                        End Select
                    End If
                Case "UP"
                    .Top = .Top - WIN_Shift
                    .Bottom = .Bottom - WIN_Shift
                Case "DOWN"
                    .Top = .Top + WIN_Shift
                    .Bottom = .Bottom + WIN_Shift
                Case "LEFT"
                    .left = .left - WIN_Shift
                    .Right = .Right - WIN_Shift
                Case "RIGHT"
                    .left = .left + WIN_Shift
                    .Right = .Right + WIN_Shift
                    
                Case Else
                    'Call GetWindowRect(GetDesktopWindow(), rctDesktop) ' Does not take taskbar into account
                    Const SPI_GETWORKAREA = 48
                    Call SystemParametersInfo(SPI_GETWORKAREA, 0&, rctDesktop, 0&)

                    Select Case strAction
                        Case "MAXHOR"
                            .left = rctDesktop.left
                            .Right = rctDesktop.Right
                        Case "MAXVERT"
                            .Top = rctDesktop.Top
                            .Bottom = rctDesktop.Bottom
                    Case "CENTER"
                        .left = (rctDesktop.left + rctDesktop.Right - lngWidth) / 2
                        .Top = (rctDesktop.Top + rctDesktop.Bottom - lngHeight) / 2
                        .Right = .left + lngWidth
                        .Bottom = .Top + lngHeight
                    Case "CENTERTOP" '
                        .left = (rctDesktop.left + rctDesktop.Right - lngWidth) / 2
                        .Right = .left + lngWidth
                        .Top = rctDesktop.Top + 0.25 * (rctDesktop.Bottom - rctDesktop.Top) - 0.5 * lngHeight
                        .Bottom = .Top + lngHeight
                    Case "CENTERBOT"
                        .left = (rctDesktop.left + rctDesktop.Right - lngWidth) / 2
                        .Right = .left + lngWidth
                        .Top = rctDesktop.Top + 0.75 * (rctDesktop.Bottom - rctDesktop.Top) - 0.5 * lngHeight
                        .Bottom = .Top + lngHeight
                    Rem -----------------------------------------------------------------------------------------------------------------
                    Rem - Placing functions (1/2)
                    Case "CENTERLEFT"
                        .left = rctDesktop.left + 0.25 * (rctDesktop.Right - rctDesktop.left) - 0.5 * lngWidth
                        .Right = .left + lngWidth
                        .Top = (rctDesktop.Top + rctDesktop.Bottom - lngHeight) / 2
                        .Bottom = .Top + lngHeight
                    Case "CENTERRIGHT"
                        .left = rctDesktop.left + 0.75 * (rctDesktop.Right - rctDesktop.left) - 0.5 * lngWidth
                        .Right = .left + lngWidth
                        .Top = (rctDesktop.Top + rctDesktop.Bottom - lngHeight) / 2
                        .Bottom = .Top + lngHeight
                    Rem -----------------------------------------------------------------------------------------------------------------
                    Rem - Placing functions (1/4)
                    Case "CENTERTOPLEFT"
                        .left = rctDesktop.left + 0.25 * (rctDesktop.Right - rctDesktop.left) - 0.5 * lngWidth
                        .Right = .left + lngWidth
                        .Top = rctDesktop.Top + 0.25 * (rctDesktop.Bottom - rctDesktop.Top) - 0.5 * lngHeight
                        .Bottom = .Top + lngHeight
                    Case "CENTERTOPRIGHT"
                        .left = rctDesktop.left + 0.75 * (rctDesktop.Right - rctDesktop.left) - 0.5 * lngWidth
                        .Right = .left + lngWidth
                        .Top = rctDesktop.Top + 0.25 * (rctDesktop.Bottom - rctDesktop.Top) - 0.5 * lngHeight
                        .Bottom = .Top + lngHeight
                    Case "CENTERBOTLEFT"
                        .left = rctDesktop.left + 0.25 * (rctDesktop.Right - rctDesktop.left) - 0.5 * lngWidth
                        .Right = .left + lngWidth
                        .Top = rctDesktop.Top + 0.75 * (rctDesktop.Bottom - rctDesktop.Top) - 0.5 * lngHeight
                        .Bottom = .Top + lngHeight
                    Case "CENTERBOTRIGHT"
                        .left = rctDesktop.left + 0.75 * (rctDesktop.Right - rctDesktop.left) - 0.5 * lngWidth
                        .Right = .left + lngWidth
                        .Top = rctDesktop.Top + 0.75 * (rctDesktop.Bottom - rctDesktop.Top) - 0.5 * lngHeight
                        .Bottom = .Top + lngHeight
                    Case "LEFTHALF"
                        .left = rctDesktop.left
                        .Right = 0.5 * (rctDesktop.Right + rctDesktop.left)
                        .Top = rctDesktop.Top
                        .Bottom = rctDesktop.Bottom
                    Case "RIGHTHALF"
                        .left = 0.5 * (rctDesktop.Right + rctDesktop.left)
                        .Right = rctDesktop.Right
                        .Top = rctDesktop.Top
                        .Bottom = rctDesktop.Bottom
                    Case "TOPHALF"
                        .left = rctDesktop.left
                        .Right = rctDesktop.Right
                        .Top = rctDesktop.Top
                        .Bottom = 0.5 * (rctDesktop.Bottom + rctDesktop.Top)
                    Case "BOTHALF"
                        .left = rctDesktop.left
                        .Right = rctDesktop.Right
                        .Top = 0.5 * (rctDesktop.Bottom + rctDesktop.Top)
                        .Bottom = rctDesktop.Bottom
                    Case "TOPRIGHT"
                        .left = 0.5 * (rctDesktop.Right + rctDesktop.left)
                        .Right = rctDesktop.Right
                        .Top = rctDesktop.Top
                        .Bottom = 0.5 * (rctDesktop.Bottom + rctDesktop.Top)
                    Case "TOPLEFT"
                        .left = rctDesktop.left
                        .Right = 0.5 * (rctDesktop.Right + rctDesktop.left)
                        .Top = rctDesktop.Top
                        .Bottom = 0.5 * (rctDesktop.Bottom + rctDesktop.Top)
                    Case "BOTTOMLEFT"
                        .left = rctDesktop.left
                        .Right = 0.5 * (rctDesktop.Right + rctDesktop.left)
                        .Top = 0.5 * (rctDesktop.Bottom + rctDesktop.Top)
                        .Bottom = rctDesktop.Bottom
                    Case "BOTTOMRIGHT"
                        .left = 0.5 * (rctDesktop.Right + rctDesktop.left)
                        .Right = rctDesktop.Right
                        .Top = 0.5 * (rctDesktop.Bottom + rctDesktop.Top)
                        .Bottom = rctDesktop.Bottom
                    Case Else
                        strMessage = "Not yet implemented! - " & Prop("Action")
                End Select
            End Select
            #If ZKCONFIG <> 1 Then
                If Len(strMessage) = 0 Then
                    Call PlaceWindow(lngWin, RecTan)
                    ResizeWin = lngWin
                End If
            #End If
        End With
    End If
    If Prop("SILENT") <> "Y" Then
        If Len(strMessage) > 0 Then Call ZenMB(strMessage)
    End If

End Function



Private Sub ShowInfo(ByVal hwnd As Long)
Rem - Display information about the selected window

    'Call ZenMB("Info for " & CStr(hwnd))
    Dim k As Long, Msg As String
    Dim strExe As String, strClass As String
    
    For k = 0 To 15
        strExe = GetExeFromHandle(ActiveWindow(k))
        strClass = ClassName(ActiveWindow(k))
        Msg = Msg & CStr(k) & ". " & strExe & " (" & CStr(ActiveWindow(k)) & "), " & strClass & vbCrLf
    Next k
    Call ZenMB(Msg)

End Sub

Private Sub Class_Initialize()
    
    Set colSelected = New Collection
    HistDepth = Val(settings("HistDepth"))
    If HistDepth < 1 Then HistDepth = 10
    
End Sub

Private Sub p_ToTray(ByVal hwnd As Long)
#If ZKCONFIG <> 1 Then
    Dim it As frmSystray
    
        Set it = New frmSystray
        With it
            .OwnerForm = 0 ' Initilaise
            .OwnerForm = hwnd
            .SendToTray
        End With
#End If
End Sub

Private Function p_FromTray(ByVal lngIndex As Long) As Long
Rem - Returns the handle to the window being restored
#If ZKCONFIG <> 1 Then
Dim it As frmSystray

    If lngIndex > -1 Then
        If lngIndex < ST_Count Then
            Set it = ST_TrayForms(lngIndex)
            p_FromTray = ST_TrayForms(lngIndex).OwnerForm
            Unload it
            Rem - Only return the number if the window is valid
            If 0 = IsWindow(p_FromTray) Then p_FromTray = 0
        End If
    End If
#End If

End Function


Private Sub p_Focus(ByVal hwnd As Long)
    
    If IsWindowVisible(hwnd) Then Call SetForegroundWindow(hwnd)

End Sub

Public Sub SetTrans(ByVal lngHWnd As Long, ByVal lngTrans As Long)
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Dim Ret As Long

    Rem - Set the window style to 'Layered'
    Ret = GetWindowLong(lngHWnd, GWL_EXSTYLE)
    Rem - 255 = Totally opague
    If lngTrans < 1 Then
        Call SetLayeredWindowAttributes(lngHWnd, 0, 255, LWA_ALPHA)
        Call SetWindowLong(lngHWnd, GWL_EXSTYLE, Ret And (Not WS_EX_LAYERED))
    Else
        lngTrans = 255 * (lngTrans / 100)
        Call SetWindowLong(lngHWnd, GWL_EXSTYLE, Ret Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(lngHWnd, 0, lngTrans, LWA_ALPHA)
    End If
    Call ZK_Win.RefreshWindow(lngHWnd)

End Sub

Public Function Tray_FlushExe(ByVal ExeName As String) As Long
Rem - For when the Exe is launched/brought to front, remove any icons from tray
Dim k As Long, strTApp As String

    For k = ST_Count - 1 To 0 Step -1
        strTApp = GetExeFromHandle(ST_TrayForms(k).OwnerForm)
        If strTApp = ExeName Then Tray_FlushExe = p_FromTray(k)
    Next k
    
End Function
Private Sub History_Add(ByVal hwnd As Long, ByVal Action As String, Optional UndoLast As Boolean = False)
    
    If UndoLast Then
        Rem - Pop off the last window that was added to the histroy
        HistCurrent = (HistCurrent - 1 + HistDepth) Mod HistDepth
        'If HistCurrent > 0 Then ReDim Preserve History(1 To HistCurrent)
        
    Else
        Rem - Add the referenced window to the history list
        If HistCurrent < HistDepth Then
            Rem - Add to the array
            HistCurrent = HistCurrent + 1
            ReDim Preserve History(1 To HistCurrent)
        Else
            Rem - Shift the array down and put at the end
            Dim k As Long
            For k = 2 To HistCurrent
                History(k - 1) = History(k)
            Next k
        End If
        
        Rem - Add the details of the referenced window
        With History(HistCurrent)
            .LastWinPlc.Length = Len(.LastWinPlc)
            'Call GetWindowPlacement(HWnd, .LastWinPlc)
            .LastHWnd = hwnd
            .LastAction = Action
        End With
    End If

End Sub

Public Sub Selected(ByVal Action As String, ByVal lngHWnd As Long)
Rem - Adds window to, toggles or clears the current windows selection
Dim k As Long
    
    Select Case Action
        Case "Toggle", "Add"
            Dim strHWnd As String
            Rem - If control is down, either select or uselect the thing
            If WindowIsUsable(lngHWnd) Then
                Dim booFound As Boolean
                For k = colSelected.Count To 1 Step -1
                    If lngHWnd = colSelected.Item(k) Then
                        Rem - The window is already in the selected list.
                        If Action = "Toggle" Then Call colSelected.Remove(k)
                        booFound = True
                        Exit For
                    End If
                Next k
                If Not booFound Then Call colSelected.Add(lngHWnd)
            End If
            
        #If ZKCONFIG <> 1 Then
            Case "Clear"
                Call dhc_RemoveAll(colSelected)
            Case "Remove"
                Call dhc_Remove(colSelected, lngHWnd)
            Case "Purge"
                For k = colSelected.Count To 1 Step -1
                    If IsWindow(colSelected.Item(k)) = 0 Then Call colSelected.Remove(k)
                Next k
        #End If
    End Select
    
End Sub

Private Function p_UnLoggedAct(ByVal Action As String) As Boolean

    p_UnLoggedAct = True
    Select Case Action
        Case "SELADD": Call Selected("Add", ActiveWindow(0))
        Case "SELTOGGLE": Call Selected("Toggle", ActiveWindow(0))
        Case "SELCLEAR": Call Selected("Clear", 0)
        Case "SELREMOVE": Call Selected("Remove", ActiveWindow(0))
        Case "WINALIGNLEFT", "WINALIGNRIGHT", "WINALIGNTOP", "WINALIGNBOT", "WINALIGNCENTERV", "WINALIGNCENTERH"
            Call WinAlign(Action)
        Case "Undo"
            If HistCurrent > 0 Then
                With History(HistCurrent)
                    If (.LastAction = "RestoreFromTray") Or (.LastAction = "RestoreMe") Then
                        Call p_ToTray(.LastHWnd)
                    Else
                        Dim k As Long
                        Call SetWindowPlacement(.LastHWnd, .LastWinPlc)
                        Select Case .LastAction
                            Case "PLACERECT"
                                Dim lngWins As Long
                                
                                Rem - Check if the action is part of an alignment. If it is, undo all the window movements.
                                lngWins = Val(Prop_Get("NumWins", History(HistCurrent).Properties))
                                If lngWins > 1 Then
                                    Call History_Add(0, "", True)
                                    For k = 2 To lngWins
                                        Call DoAction(zenDic("Action", "Undo"))
                                    Next k
                                    Exit Function
                                End If
                            Case "TOP", "WUMTOP"
                                Call SetWinPos(.LastHWnd, HWND_BOTTOM, False)
                            Case "BOTTOM"
                                Call SetWinPos(.LastHWnd, HWND_TOP, False)
                            Case "MINIMIZE", "NORMAL", "MAXIMIZE"
                                Call SetWindowPlacement(.LastHWnd, .LastWinPlc)
                           Case "AOT"
                                Call SetWinPos(.LastHWnd, HWND_NOTOPMOST, False)
                            Case "NOAT"
                                Call SetWinPos(.LastHWnd, HWND_TOPMOST, True)
                            Case "DESKTOP"
                                Call ShellExe(App.Path & "\Showde~1.scf")
                            Case "SENDTOTRAY"
                                Call p_FromTray(ST_Count - 1)
                            Case "MINALL"
                                Call EnumWindows(AddressOf RestoreAll, ByVal 0&)
                            Case "MINALLBUTTOP"
                                Call EnumWindows(AddressOf RestoreAll, ByVal 0&)
                            Case "DISABLE"
                                Call EnableWindow(.LastHWnd, 1)
                            Case "PINTODESKTOP"
                                Call PinToDesktop(False, .LastHWnd)
                            Case "UNPINFROMDESKTOP"
                                Call PinToDesktop(True, .LastHWnd)
                            Case "ENABLE"
                                Call ZenMB("Enabling a window cannot be undone. If it was disabled, please just disable it again.")
                            Case "RESTOREALL"
                                Call ZenMB("Restoring all windows cannot be undone. If you wish to minimize all windows, please just use this action.")
                            Case "KILLACTIVE"
                                Call ZenMB("Sorry, killing windows cannot be undone...")
                            Case "MOVEALLUP", "MOVEALLDOWN", "MOVEALLRIGHT", "MOVEALLLEFT"
                                Select Case .LastAction
                                    Case "MOVEALLUP"
                                        Call MoveAllWindows(0, WIN_Shift, False)
                                    Case "MOVEALLDOWN"
                                        Call MoveAllWindows(0, -WIN_Shift, False)
                                    Case "MOVEALLLEFT"
                                        Call MoveAllWindows(WIN_Shift, 0, False)
                                    Case "MOVEALLRIGHT"
                                        Call MoveAllWindows(-WIN_Shift, 0, False)
                                End Select
                            Case "ROLLUP", "ROLLDOWN"
                                Rem - If the Rolluing down/up has been undone, remove it from the collection
                                k = ROL_GetIndex(.LastHWnd)
                                If k > 0 Then Call ROL_Remove(k)
                            Case "MAKETRANS"
                                Call MakeTransparent(.LastHWnd, True)
                            Case "MAKETRANSUNDO"
                                Call MakeTransparent(.LastHWnd)
                            Case "UNDOLASTAWT"
                                Call ZenMB("Sorry, undoing the last 'Auto-window transparency' cannot be undone...")
                            Case Else
                                If InStr(.LastAction, "SETTRANSPARENCY") > 0 Then Call SetTrans(.LastHWnd, -1)
                        End Select
                    End If
                End With
                Call History_Add(0, "", True)
            ElseIf p_UnLoggedAct Then
                Call ZenMB("There are no more actions that can be undone!")
            End If
        Case Else: p_UnLoggedAct = False
    End Select

End Function
Public Sub WinAlign(ByVal strAction As String)
Dim rctRect As RECT, lngVal As Long
Dim lngHWnd As Long, lngMax As Long
Dim k As Long, lngTemp As Long

    If colSelected.Count < 2 Then
        Call ZenMB("More than one window must be selected to perform alignment. Please create a window selection by using the Desktop map or by using window selection actions.")
    Else
        lngHWnd = colSelected.Item(1)
        If GetWindowRect(lngHWnd, rctRect) <> 0 Then
            lngMax = colSelected.Count
            booMulti = True
            Select Case strAction
                Case "WINALIGNLEFT": lngVal = rctRect.left
                Case "WINALIGNRIGHT": lngVal = rctRect.Right
                Case "WINALIGNTOP": lngVal = rctRect.Top
                Case "WINALIGNBOT": lngVal = rctRect.Bottom
                Case "WINALIGNCENTERV": lngVal = CLng(0.5 * (rctRect.Top + rctRect.Bottom))
                Case "WINALIGNCENTERH": lngVal = CLng(0.5 * (rctRect.Right + rctRect.left))
            End Select
            
            Rem - Now alignm the other windows
            For k = 2 To lngMax
                lngHWnd = colSelected.Item(k)
                If GetWindowRect(lngHWnd, rctRect) <> 0 Then
                    With rctRect
                        Select Case strAction
                            Case "WINALIGNLEFT"
                                lngTemp = .Right - .left
                                .left = lngVal
                                .Right = lngVal + lngTemp
                            Case "WINALIGNRIGHT"
                                lngTemp = .Right - .left
                                .Right = lngVal
                                .left = lngVal - lngTemp
                            Case "WINALIGNTOP"
                                lngTemp = .Bottom - .Top
                                .Top = lngVal
                                .Bottom = .Top + lngTemp
                            Case "WINALIGNBOT"
                                lngTemp = .Bottom - .Top
                                .Bottom = lngVal
                                .Top = .Bottom - lngTemp
                            Case "WINALIGNCENTERV"
                                lngTemp = .Bottom - .Top
                                .Top = lngVal - 0.5 * lngTemp
                                .Bottom = .Top + lngTemp
                            Case "WINALIGNCENTERH"
                                lngTemp = .Right - .left
                                .left = lngVal - 0.5 * lngTemp
                                .Right = .left + lngTemp
                        End Select
                        Call DoAction(zenDic("Action", "PlaceRect", "Left", .left, "Right", .Right, "Top", .Top, _
                            "Bottom", .Bottom, "HWnd", lngHWnd))
                    End With
                End If
            Next k
            booMulti = False
                        
            If lngMax > 1 Then
                Rem - If this is the last window, add a tag for undoing all the window placements at once
                Call Prop_Set("NumWins", CStr(lngMax - 1), History(HistCurrent).Properties)
            End If
            
            

        End If
    End If
    
End Sub

Private Function p_IsMultiCompat(ByVal Action As String) As Boolean

    Select Case Action
        Case "AOT", "NAOT", "MINIMIZE", "MAXIMIZE", "TOP", "BOTTOM", "NORMAL", "KILLACTIVE", "ENABLE", "DISABLE", "SENDTOTRAY"
            p_IsMultiCompat = True
        Case "SHRINK", "GROW", "MAXHOR", "MAXVERT", "CENTER", "CENTERTOP", "CENTERBOT", "CENTERLEFT", "CENTERRIGHT", "CENTERTOPLEFT"
            p_IsMultiCompat = True
        Case "CENTERTOPRIGHT", "CENTERBOTLEFT", "CENTERBOTRIGHT", "FATTEN", "THIN", "HEIGHTEN", "SHORTEN", "UP", "DOWN", "LEFT"
            p_IsMultiCompat = True
        Case "RIGHT", "LEFTHALF", "RIGHTHALF", "TOPHALF", "BOTHALF", "TOPRIGHT", "TOPLEFT", "BOTTOMLEFT", "BOTTOMRIGHT" ', "MOVEWIN"
            p_IsMultiCompat = True
        Case "ICONIFY", "ICONIFYCLEAR", "ROLLUP", "ROLLDOWN", "PINTODESKTOP", "UNPINFROMDESKTOP", "MAKETRANS", "MAKETRANSUNDO"
            p_IsMultiCompat = True
        Case Else
            If left(Action, 15) = "SETTRANSPARENCY" Then
                p_IsMultiCompat = True
            Else
                p_IsMultiCompat = False
            End If
    End Select
    
    If p_IsMultiCompat Then
        Rem - Clear the collection of dead windows before allowing
        Call Selected("Purge", 0)
    End If

End Function
Public Sub PinToDesktop(ByVal bPin As Boolean, ByVal hwnd As Long)
    
    If bPin Then
        Call SetParent(hwnd, FindWindow("progman", vbNullString))
    Else
        Call SetParent(hwnd, FindWindow("ExploreWClass", vbNullString))
    End If

End Sub

Public Sub RefreshWindow(ByVal hwnd As Long)
Const RDW_FRAME = &H400
Const RDW_ALLCHILDREN = &H80
Const RDW_ERASE = &H4
Const RDW_INVALIDATE = &H1

    Call RedrawWindow(hwnd, ByVal 0&, ByVal 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_FRAME Or RDW_ALLCHILDREN)

End Sub


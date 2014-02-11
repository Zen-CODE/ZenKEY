Attribute VB_Name = "modWindows"
 Option Explicit
Option Compare Text
Rem=========================================
Rem - This module is neccesary as AddressOf operator cannot use  a class function
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetAncestor Lib "user32.dll" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
Private Const GA_ROOTOWNER = 3
Private Const GA_ROOT = 2

Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_NORMAL = 1
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Rem - For getting the desktop area
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Rem - Common
Public Type RECT
        left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
'Private Declare Function Get\ Lib "user32" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal Hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Rem ======================= Declarations for Gettings ExeFromHandle ===================
Private SE_ExeName As String
Private SE_HWnd As Long

'Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Private Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Rem ======================= Declarations for Gettings ExeFromHandle ===================
Rem ======================= Declarations for WidowHooks - WH_XX ===================
'Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
''Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
'Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
'Public WH_HookID As Long
Rem ======================= Declarations for WidowHooks - WH_XX ===================
Rem --------------------------------- For Systray forms ---------------------------------
'Public ST_TrayForms() As frmSystray
Public ST_TrayForms() As Object
Public ST_Count As Long

Rem - See modSystray
Private WMV_X As Long ' The direction in which tomove the windows for the MoveAllWindows command
Private WMV_Y As Long
Public WMV_TotX As Long
Public WMV_TotY As Long
Public ActiveWindow(0 To 15) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Rem - The followng two entries are stored as public variables and not in 'ZenProperties'
Rem - for reasons of optimzation for the callback functions
Public IDT_Enabled As Boolean
Public DTM_Enabled As Boolean
Public IDT_AutoFocus As Boolean
Public IDT_ActiveApp As String

Rem - For the new SetWindowPosition functions
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Private booExeSearch As Boolean
Private DtArea As RECT ' For moving all windows whever the taskbar is, we need to know the desktop area
Rem - Testing the window to see if it belongs to the desktop
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public DTP_Area As RECT
Public DTP_Handle As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal HIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal HIcon As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public ZK_WinUMouse As Boolean
Rem - For efficient Infinite Desktop updates
Public Enum WindowsStatus
    Normal = 0
    Desktop = 1
    OffScreen = 2
End Enum
Public Type RecInfo
    TheRect As RECT
    Status As WindowsStatus
    Caption As String
    hwnd As Long
End Type
Public WIN_RecList() As RecInfo
Public WIN_RecMax As Long
Public WIN_RecCurrent As Long
Public WIN_Changed As Boolean
Private lngPrevRecMax As Long
Private booWinampWarn As Boolean
Rem -------------------------------- For registry stuff
Public Const HKCR = &H80000000 'HKEY_CLASSES_ROOT
Public Const HKCU = &H80000001 'HKEY_CURRENT_USER
Public Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Rem ------------------------------- For windows version
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetProcessImageFileName Lib "PSAPI.DLL" Alias "GetProcessImageFileNameA" (ByVal hProcess As Long, ByVal lpImageFileName As String, ByVal nSize As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long


Public Function IsWindows7OrBetter() As String
    
    Dim WinVer As Long
    WinVer = GetVersion() And &HFFFF&
    ' For major version, WinVer Mod 256. For minor, WinVer \ 256
    
    IsWindows7OrBetter = CBool((WinVer \ 256 > 0) And (WinVer Mod 256 > 5))
    'GetWinVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed") XP = 5.01, 7 = 6.1
    
End Function
Public Function IsWindows8() As String
    
    Dim WinVer As Long
    WinVer = GetVersion() And &HFFFF&
    ' For major version, WinVer Mod 256. For minor, WinVer \ 256
    
    IsWindows8 = CBool((WinVer Mod 256 = 6) And (WinVer \ 256 > 1)) ' 6  for win 7
    'GetWinVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed") XP = 5.01, 7 = 6.1
    
End Function

Private Sub p_AddToDiplay(ByVal hwnd As Long, ByRef rctRect As RECT)
Dim strTemp As String
    
    If rctRect.Top = -30000 Then
        If Not booWinampWarn Then
            strTemp = "It appears that you are using Winamp with a modern skin." & vbCr & vbCr & _
                "Please be advised that Winamp employs non-standard methods to render 'Modern skins'. " & _
                "As Nullsoft do not document or respond to any queries in this regard, we are unable to ascertain why 'Winamp' windows behave as they do." & _
                vbCr & vbCr & "As such, please be warned that Winamp windows may not respond to standard ZenKEY/Windows commands."
            Call ZenMB(strTemp)
            booWinampWarn = True
        End If
        Exit Sub ' Winamp off-screen window messy coverup for Nullsoft &@^#!!!
    End If
    
    If WIN_RecMax >= lngPrevRecMax Then
        ReDim Preserve WIN_RecList(0 To WIN_RecMax)
        lngPrevRecMax = WIN_RecMax
    End If
    
    Dim booSame As Boolean
    With WIN_RecList(WIN_RecCurrent)
        If (hwnd = .hwnd) Then
            If (rctRect.left = .TheRect.left) And (rctRect.Right = .TheRect.Right) Then
                If (rctRect.Top = .TheRect.Top) And (rctRect.Bottom = .TheRect.Bottom) Then booSame = True
            End If
        End If
        If Not booSame Then
            WIN_Changed = True
            .TheRect = rctRect
            
            Rem - Now decide the caption for the item.
            Dim booDesktop As Boolean
            Dim lngRet As Long
            .hwnd = hwnd
            booDesktop = CBool(DTP_Handle = hwnd)

            Rem - Decide the colour/status
            Select Case True
                Case booDesktop
                    .Status = Desktop
                Case Else
                    If (rctRect.Right <= DTP_Area.left) Or (rctRect.left >= DTP_Area.Right) Then
                        .Status = OffScreen
                    ElseIf (rctRect.Bottom <= DTP_Area.Top) Or (rctRect.Top >= DTP_Area.Bottom) Then
                        .Status = OffScreen
                    Else
                        .Status = Normal
                    End If
            End Select
        End If
    End With
    WIN_RecCurrent = WIN_RecCurrent + 1
    WIN_RecMax = WIN_RecMax + 1

End Sub

Public Function RecordAll(ByVal hwnd As Long, ByVal lParam As Long) As Boolean

    If IsWindowVisible(hwnd) Then
        If IsIconic(hwnd) = 0 Then
            Dim rctRect As RECT
            Call GetWindowRect(hwnd, rctRect)
            With rctRect
                If .Right - .left > 0 Then
                    If .Bottom - .Top > 0 Then Call p_AddToDiplay(hwnd, rctRect)
                End If
            End With
        End If
    End If
    RecordAll = True

End Function
Public Function SendToZenKEY(ByVal prop As String) As Boolean
Dim lngHWnd As Long

    lngHWnd = Val(Registry.GetRegistry(HKCU, "SOFTWARE\ZenCODE\ZenKEY", "WindowHandle"))
    If lngHWnd > 0 Then
        If IsWindow(lngHWnd) Then
            Call Registry.SetRegistry(HKCU, "SOFTWARE\ZenCODE\ZenKEY", "Action", prop)
            Rem - Now notify the running ZenKEY of the command
            Call PostMessage(lngHWnd, WIN_ZKAction, 0, 0)
            SendToZenKEY = True
        End If
    End If
    


End Function

Public Function GetFileName(ByVal FileName As String) As String
Dim parts() As String, max As Long
    
    parts = Split(FileName, "\")
    max = UBound(parts())
    If max > -1 Then GetFileName = parts(max)
        
End Function
Public Function GetExeFromHandle(ByVal hwnd As Long) As String
Const PROCESS_QUERY_INFORMATION = (&H400)
Const PROCESS_VM_READ = (&H10)
    
    Rem - Get ID for window thread
    Dim lngProcID As Long, hProcess As Long
    Call GetWindowThreadProcessId(hwnd, lngProcID)
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lngProcID)
    If hProcess Then
        Dim sChar As Long, sBuf As String
        sBuf = String(256, Chr$(0))
        sChar = GetProcessImageFileName(hProcess, sBuf, 256)
        GetExeFromHandle = left$(sBuf, InStr(sBuf, Chr$(0)) - 1)
        GetExeFromHandle = ReplaceDevName(GetExeFromHandle)
    End If
    
End Function
Public Function ShowExeWindow(ByVal FileName As String) As Long
Rem - This function either activates the window belonging to the exe named 'FileName" (SetWinList = False)
Rem - or returns a window list belonging to the 'FileName' module (SetWInList = True)

    #If ZKCONFIG <> 1 And ZenWiz <> 1 Then
    SE_ExeName = GetFileName(FileName)

    Rem ---------------------------------------------------
    Rem - Show any Exe Window
    Rem ---------------------------------------------------
    Rem - Just clear form the system tray if it has been sent there
    SE_HWnd = Icon_FlushExe(FileName)
    If SE_HWnd = 0 Then
        SE_HWnd = ZK_Win.Tray_FlushExe(FileName)
        If SE_HWnd = 0 Then
            Rem - The program is not in the tray or iconized
            Select Case Right(FileName, 4)
                Case ".exe", ".bin": booExeSearch = True
                Case Else: booExeSearch = False
            End Select
            Call EnumWindows(AddressOf EnumWinExeProc, ByVal 0&)
        End If ' Not Icon_FlushExe
    End If ' Not Tray_FlushExe
    ShowExeWindow = SE_HWnd
    #End If

End Function
Private Function EnumWinExeProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
Dim strApp As String
Dim booShow As Boolean
    
    If IsWindowVisible(hwnd) Then
    
        Rem - Use on error as Win NT does not support this call
        On Error Resume Next
        
        If booExeSearch Then
            Rem - Match according to owner handle
            strApp = GetExeFromHandle(hwnd)
            If GetFileName(strApp) = SE_ExeName Then booShow = True
        Else
            Rem - Match according to title bar
            Dim strCap As String
            strCap = String(255, Chr$(0))
            Call GetWindowText(hwnd, strCap, 255)
            strCap = left$(strCap, InStr(strCap, Chr$(0)) - 1)
            If Val(InStr(strCap, SE_ExeName)) > 0 Then booShow = True
        End If
    
        If booShow Then
            If SE_HWnd = 0 Then SE_HWnd = hwnd ' Use the first window as this seems to be the top level window
            Call SetWinPos(hwnd, HWND_TOP, True)
        End If

    End If
    Rem - continue enumeration
    EnumWinExeProc = True
    
End Function

Public Function GetWindowToUse() As Long
Dim strMessage As String

    If ZK_WinUMouse Then
        Rem - Use Window under the mouse
        If MainForm.HWndUnderMenu = 0 Then
            GetWindowToUse = WindowFromCursor
        Else
            GetWindowToUse = MainForm.HWndUnderMenu
            MainForm.HWndUnderMenu = 0
        End If
    Else
        Rem - Get Window from last active or window under mouse
        GetWindowToUse = ActiveWindow(0)
    End If

    If GetWindowToUse = 0 Then
        Call ZenMB("Unable to find any window to act on? Strange....", "OK")
    ElseIf Not WindowIsUsable(GetWindowToUse) Then
        GetWindowToUse = 0
    End If
    
End Function

Public Function MinimizeAll(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
Dim lngParent As Long
Dim lngStyle As Long

    If IsWindowVisible(hwnd) Then
        Rem - First determine if it has a parent
        lngParent = GetAncestor(hwnd, GA_ROOTOWNER)
        If lngParent > 0 Then hwnd = lngParent
        lngStyle = GetWindowLong(hwnd, GWL_STYLE)
        Rem - Check if it has a minimis box
        If lngStyle And WS_MINIMIZEBOX Then
            If Not (lngStyle And WS_MINIMIZE) Then Call CloseWindow(hwnd)
        End If
    End If
    MinimizeAll = True

End Function
#If ZKCONFIG <> 1 Then
Private Function MoveAll(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
Const SWP_FRAMECHANGED = &H20
    
    If IsWindowVisible(hwnd) Then
        Dim DestRect As RECT
        
        If Not (IsIconic(hwnd) Or IsZoomed(hwnd)) Then
            Dim lngRet
            lngRet = GetWindowRect(hwnd, DestRect)
            If lngRet <> 0 Then
                With DestRect
                    .left = .left + WMV_X
                    .Right = .Right + WMV_X
                    .Top = .Top + WMV_Y
                    .Bottom = .Bottom + WMV_Y
                End With
                If IDT_Enabled Then
                    If ZK_IDT.IsUseable(hwnd) Then Call PlaceWindow(hwnd, DestRect)
                Else
                    Call PlaceWindow(hwnd, DestRect)
                End If
            End If
        End If
    End If
    MoveAll = True
    
End Function
#End If
Public Function RestoreAll(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
Dim lngParent As Long
Dim lngStyle As Long
Const SW_RESTORE = 9

    If IsWindowVisible(hwnd) Then
        Rem - First determine if it has a parent
        lngParent = GetAncestor(hwnd, GA_ROOTOWNER)
        If lngParent > 0 Then hwnd = lngParent
        lngStyle = GetWindowLong(hwnd, GWL_STYLE)
        Rem - Check if it has a minimis boz
        If lngStyle And WS_MINIMIZEBOX Then Call ShowWindow(hwnd, SW_RESTORE)
    End If
    RestoreAll = True

End Function




Public Function WindowFromCursor() As Long
 Dim CurPos As POINTAPI
 Dim lngOwner As Long
 
    Call GetCursorPos(CurPos)
    WindowFromCursor = WindowFromPoint(CurPos.X, CurPos.Y)
    Rem - Get the root window so that we don't move textboxes and labels within a window!
    lngOwner = GetAncestor(WindowFromCursor, GA_ROOTOWNER)
    If lngOwner <> 0 Then WindowFromCursor = lngOwner
    
End Function

Public Function WindowIsUsable(ByVal lngHWnd As Long) As Boolean
    
    ' TODO: Consolidate with IsLegalWindow
    If lngHWnd <> 0 Then
        Rem - Retrieve the class name to check that it is not the desktop ....
        Select Case ClassName(lngHWnd)
            Case "PROGMAN", "SHELL_TRAYWND", "SHELLDLL_DefView", "WorkerW", "SystemTray_Main"
                Rem - Progman = Parent Desktop window classname = SHELLDLL_DefView
                Rem - SHELL_TRAYWND = Taskbar
                Rem - WorkerW seems to replace the desktop on Win7 when changing backgrounds or other settings...
                'strMessage = "Sorry, but Window actions cannot be perfomed on the Desktop or Taskbar."
            'Case "RunDLL"
                Rem - Once a RunDLL window is activated, it does not fire again when it is re-activated?
                Rem - This causes it to remain transparent, so take it out of transparency
            Case "Thunderform", "ThunderRT6FormDC", "ThunderFormDC"
                #If IDE = 1 Then
                    WindowIsUsable = CBool(Right(GetExeFromHandle(lngHWnd), 7) <> "Vb6.exe")
                #Else
                    WindowIsUsable = CBool(Right(GetExeFromHandle(lngHWnd), 10) <> "ZenKEY.exe")
                #End If
                
            Case Else
                WindowIsUsable = True
        End Select
    End If
    
End Function




Public Sub MoveAllWindows(ByVal XPixShift As Long, ByVal YPixShift As Long, ByVal DTMRec As Boolean)
#If ZKCONFIG <> 1 Then
Rem - Use the desktop

    WMV_X = XPixShift
    WMV_Y = YPixShift
    If DTMRec Then
        WMV_TotX = WMV_TotX + XPixShift
        WMV_TotY = WMV_TotY + YPixShift
    End If
    Call EnumWindows(AddressOf MoveAll, ByVal 0&)
#End If
End Sub

Public Sub SetWinPos(ByVal hwnd As Long, ByVal Layer As Long, ByVal Activate As Boolean)
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

    On Error Resume Next ' Get can't show modal form error swhen message box dispalyed
    If IsIconic(hwnd) Then Call OpenIcon(hwnd)
    Call SetWindowPos(hwnd, Layer, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
    If Activate Then Call SetForegroundWindow(hwnd)

End Sub
Public Function ClassName(ByVal lngHWnd As Long) As String
Dim RetVal As Long, lpClassName As String
    
    Rem - Retrieve the class name to check that it is not the desktop ....
    lpClassName = Space(256)
    RetVal = GetClassName(GetAncestor(lngHWnd, GA_ROOTOWNER), lpClassName, 256)
    ClassName = left$(lpClassName, RetVal)
 
End Function

#If ZKCONFIG <> 1 Then
Public Sub PlaceWindow(ByVal lngWin As Long, ByRef DestRect As RECT)

    With DestRect
        If Not IDT_Enabled Then Call p_KeepOnScreen(DestRect)
        Call MoveWindow(lngWin, .left, .Top, .Right - .left, .Bottom - .Top, True) ' Dont repaint if off screen
    End With
    
End Sub
#End If
Private Sub p_KeepOnScreen(ByRef AppRect As RECT)

    With AppRect
        Rem - Check Top and Bottom
        Select Case True
            Case .Top <= DTP_Area.Top
                .Bottom = .Bottom - .Top
                .Top = 0
            Case .Bottom >= DTP_Area.Bottom
                .Top = .Top - (.Bottom - DTP_Area.Bottom)
                .Bottom = DTP_Area.Bottom
        End Select
        
        Rem - Check Left and Right
        Select Case True
            Case .left <= DTP_Area.left
                .Right = .Right - .left
                .left = 0
            Case .Right > DTP_Area.Right
                .left = .left - (.Right - DTP_Area.Right)
                .Right = DTP_Area.Right
        End Select
    End With

End Sub

Private Function ReplaceDevName(ByRef filePath As String) As String
'\Device\HarddiskVolume5\My Documents\ZenCODE\ZenKEY\ZenKEY\ZKConfig.exe
Dim intPos As Long, devName As String
Dim sLetter As String

    intPos = InStr(Mid$(filePath, 9), "\")
    devName = left$(filePath, 9 + intPos - 2)
    sLetter = GetDriveForNtDeviceName(devName)
    If Len(sLetter) > 0 Then
        ReplaceDevName = Replace(filePath, devName & "\", sLetter)
    Else
        ReplaceDevName = filePath
    End If

End Function

Public Function IsOnDesktop(ByRef hwnd As Long) As Boolean
Dim rct As RECT

    If GetWindowRect(hwnd, rct) Then
        If DTP_Area.left <= rct.left Then
            If DTP_Area.Top <= rct.Top Then
                If DTP_Area.Right >= rct.Right Then
                    If DTP_Area.Bottom >= rct.Bottom Then
                        IsOnDesktop = True
                    End If
                End If
            End If
        End If
    End If
    
End Function

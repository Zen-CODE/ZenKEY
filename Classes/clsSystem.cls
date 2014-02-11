VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type LUID
    LowPart As Long
    HighPart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type
Private Const ANYSIZE_ARRAY = 1
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Private Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal wReserved As Long)
Private Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long
Rem - TO eanble shutdown/Suspend privisledges
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Rem - Screensaver stuff
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND As Long = &H112&
Private Const SC_SCREENSAVE As Long = &HF140&
Private Const SW_SHOWNORMAL = 1
Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Rem - For disabling the screen saver .....
'Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private booSCRDisabled As Boolean
Private Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Rem - For the CD Door opening and closing
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Rem - For the focus rectangle
Private Declare Function GetFocus Lib "user32" () As Long
'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Rem - For moouse input

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Function GetDriveNo(ByVal Number As Long) As String
Dim strDrive As String
Dim lngCount As Long
Dim strAllDrives As String
Dim k As Long
  
  strAllDrives = Space(64)
  
  Rem - Get all drives on your PC as one long string
  k = GetLogicalDriveStrings(Len(strAllDrives), strAllDrives)
  Rem - Trim off any trailing spaces. strAllDrives 'now contains all the drive letters.
  strAllDrives = left(strAllDrives, k)

  Rem - There is a character "chr$(0)" between each drive letter that we must filter out.
  While Len(strAllDrives) > 0
        k = InStr(strAllDrives, Chr(0))
        strDrive = left(strAllDrives, k - 1)
        If GetDriveType(strDrive) = 5 Then ' CD-ROM
          lngCount = lngCount + 1
          If Number = lngCount Then
              GetDriveNo = strDrive
              Exit Function
          End If
        End If
        strAllDrives = Mid(strAllDrives, k + 1)
  Wend
 
End Function

Public Sub OpenCloseCD(ByVal CDNum As Long)
Dim strAlias As String
Dim RetVal As Long
Dim strDrive As String

    strDrive = GetDriveNo(CDNum)
    strAlias = "Drive" & left(strDrive, 1)
    
    Rem - To close
    RetVal = mciSendString("set " & strAlias & " door closed", vbNullString, 0&, 0&)
    RetVal = mciSendString("close " & strAlias, vbNullString, 0&, 0&)
    If RetVal <> 0 Then
        Rem - To open
        RetVal = mciSendString("open " & strDrive & ": type cdaudio alias " & strAlias & " wait", vbNullString, 0&, 0&)
        RetVal = mciSendString("set " & strAlias & " door open", vbNullString, 0&, 0&)
    End If
    

End Sub
Private Function StartScreenSaver() As Boolean
Dim lRet As Long

    lRet = SendMessage(GetDesktopWindow(), WM_SYSCOMMAND, SC_SCREENSAVE, 0)
    StartScreenSaver = CBool(lRet = 0)

End Function
Private Sub EnableShutDown()
    Dim hProc As Long
    Dim hToken As Long
    Dim mLUID As LUID
    Dim mPriv As TOKEN_PRIVILEGES
    Dim mNewPriv As TOKEN_PRIVILEGES
    
    hProc = GetCurrentProcess()
    OpenProcessToken hProc, TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, hToken
    LookupPrivilegeValue "", "SeShutdownPrivilege", mLUID
    mPriv.PrivilegeCount = 1
    mPriv.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    mPriv.Privileges(0).pLuid = mLUID
    ' enable shutdown privilege for the current application
    AdjustTokenPrivileges hToken, False, mPriv, 4 + (12 * mPriv.PrivilegeCount), mNewPriv, 4 + (12 * mNewPriv.PrivilegeCount)
End Sub



Public Sub DoAction(ByRef prop As clsZenDictionary)
Const EWX_REBOOT = 2
Const EWX_POWEROFF = 8
Const SPI_SETMOUSESPEED = 113
Dim strAction As String
Dim strTemp As String
Dim sFac As Long ' Scaling factor for mouse move, button monitor for mouse clicking

    strAction = prop("Action")
    Select Case strAction
        Case "EMAIL"
            Call ShellExe("mailto:")
        Case "SCREENSAVER"
            Call StartScreenSaver
        Case "STANDBY"
            Rem - Put the system on standby, For Win2000, NOT, xP
            Call EnableShutDown
            Call SetSystemPowerState(True, False) ' Standby
        Case "PREVENTSAVER"
            Rem - Prevent the screen saver from firing....
            Const SPI_SETSCREENSAVEACTIVE = 17
            Const SPIF_UPDATEINIFILE = &H1
            If booSCRDisabled Then
                Call SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, True, ByVal 0, SPIF_UPDATEINIFILE)  'Enable
                strTemp = "Disable screen saver"
            Else
                Call SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, False, ByVal 0, SPIF_UPDATEINIFILE) 'Disable
                strTemp = "Re-enable screen saver"
            End If
            Call cMenu.SetItemProp(CLng(Val(prop("MenuHandle"))), "Caption", strTemp, CLng(Val(prop("MenuID"))))
            booSCRDisabled = Not booSCRDisabled
        Case "RESTART"
            Rem - Get permission.
            Call MainForm.CloseApp
            Call EnableShutDown
            Call ExitWindowsEx(EWX_REBOOT, 0&)  'shut down the computer
        Case "SHUTDOWN"
            Call MainForm.CloseApp
            Call EnableShutDown
            Call ExitWindowsEx(EWX_POWEROFF, 0&)  'shut down the computer
            Call MainForm.CloseApp
        Case "RUNDIALOG"
            Dim oShellApp As Object
            Set oShellApp = CreateObject("Shell.Application")
            oShellApp.FileRun
            Set oShellApp = Nothing
        Case "SHUTDOWNDIALOG"
            Call SHShutDownDialog(0)
        Case "MOUSESLOWEST"
            Call SystemParametersInfo(SPI_SETMOUSESPEED, ByVal 0, ByVal 1, ByVal 0)
        Case "MOUSESLOW"
            Call SystemParametersInfo(SPI_SETMOUSESPEED, ByVal 0, ByVal 5, ByVal 0)
        Case "MOUSEMEDIUM"
            Call SystemParametersInfo(SPI_SETMOUSESPEED, ByVal 0, ByVal 10, ByVal 0)
        Case "MOUSEFAST"
            Call SystemParametersInfo(SPI_SETMOUSESPEED, ByVal 0, ByVal 15, ByVal 0)
        Case "MOUSEFASTEST"
            Call SystemParametersInfo(SPI_SETMOUSESPEED, ByVal 0, ByVal 20, ByVal 0)
        Case "NEXTTASK"
            Call ZK_GetObject("KEYSTROKES").DoAction(zenDic("Action", "4>89>2>81>4>82>5>92>5>"))
        Case "PREVTASK"
            Call ZK_GetObject("KEYSTROKES").DoAction(zenDic("Action", "6>89>0>88>0>82>6>83>7>91>7>94>3>"))
        Case "CTRMOUSE", "CTRMOUSEACTIVE", "CTRMOUSEAC", "MOVEMOUSERGT"
            Call PlaceMouse(strAction)
        Case "DIM50", "DIM25", "DIM75", "DIM50ALL", "DIM25ALL", "DIM75ALL"
            Call Prop_Set("TRANS", Mid(strAction, 4, 2), strTemp)
            If InStr(strAction, "ALL") > 0 Then Call Prop_Set("SCREENS", "ALL", strTemp)
            Call ShellExe(App.Path & "\ZenDim.exe", strTemp)
        Case "MOUSERIGHT", "MOUSELEFT", "MOUSEUP", "MOUSEDOWN", "MOUSERIGHT1PX", "MOUSELEFT1PX", "MOUSEUP1PX", "MOUSEDOWN1PX"
            Dim pt As POINTAPI, xFac As Long, yFac As Long
            Const VK_UP = &H26, VK_RIGHT = &H27
            Const VK_LEFT = &H25, VK_DOWN = &H28
            
            If InStr(strAction, "1PX") > 0 Then sFac = 1 Else sFac = WIN_Shift
            Select Case strAction
                Case "MOUSERIGHT", "MOUSERIGHT1PX": xFac = 1
                Case "MOUSELEFT", "MOUSELEFT1PX": xFac = -1
                Case "MOUSEUP", "MOUSEUP1PX": yFac = -1
                Case Else: yFac = 1 ' "MOUSEDOWN", "MOUSEDOWN1PX"
            End Select
            
            Call GetCursorPos(pt)
            Do
                DoEvents
                If xFac = 0 Then ' They are moving vertically, so allow horizontal adjustments
                    pt.Y = pt.Y + sFac * yFac
                    If GetAsyncKeyState(VK_RIGHT) <> 0 Then
                        pt.X = pt.X + sFac
                    ElseIf GetAsyncKeyState(VK_LEFT) <> 0 Then
                        pt.X = pt.X - sFac
                    End If
                Else ' They are moving horizontally, so allow vertical adjustments
                    pt.X = pt.X + sFac * xFac
                    If GetAsyncKeyState(VK_UP) <> 0 Then
                        pt.Y = pt.Y - sFac
                    ElseIf GetAsyncKeyState(VK_DOWN) <> 0 Then
                        pt.Y = pt.Y + sFac
                    End If
                End If
                
                Call SetCursorPos(pt.X, pt.Y)
                Call Sleep(20)
            Loop While p_KeysAreDown(prop)
        
        Case "MOUSECLICK", "MOUSERCLICK", "MOUSEMCLICK", "MOUSEDCLICK"
            Rem - Wait for key release
            Do
                DoEvents
            Loop While p_KeysAreDown(prop)
            
            Select Case strAction
                Case "MOUSECLICK", "MOUSEDCLICK": sFac = 0
                Case "MOUSERCLICK": sFac = 1
                Case Else: sFac = 2
            End Select
            Call p_Mouse(sFac)
            If strAction = "MOUSEDCLICK" Then Call p_Mouse(sFac)
        
        Case Else
            If left(strAction, 6) = "CDDOOR" Then
                Call OpenCloseCD(Val(Mid(strAction, 7)))
            Else
                Call ZenMB("Unidentified system action in clsSystem. Action = " & prop("Action"), "OK")
            End If
    End Select

End Sub


Private Sub Class_Terminate()
    Rem - If enabled, remove the timer preventing the screen saver from firing...
    If booSCRDisabled Then Call DoAction(zenDic("Action", "PREVENTSAVER"))
End Sub

 

Private Sub PlaceMouse(ByVal Action As String)
Dim lngX As Long, lngY As Long
            
    Select Case Action
        Case "CTRMOUSE"
            Rem - Center mouse in screen
            lngX = MainForm.ScaleX(Screen.Width, vbTwips, vbPixels) \ 2
            lngY = MainForm.ScaleY(Screen.Height, vbTwips, vbPixels) \ 2
        Case "MOVEMOUSERGT"
            Rem - Move of the right of the screen
            Dim pt As POINTAPI
            Call GetCursorPos(pt)
            lngX = MainForm.ScaleX(Screen.Width, vbTwips, vbPixels)
            lngY = pt.Y
        Case "CTRMOUSEACTIVE", "CTRMOUSEAC"
            Rem - Center mouse on the active window / active control
            Dim lngHWnd As Long
            Dim rctRect As RECT
            Dim lngThreadActive As Long ', lngThreadMe As Long
            
            lngHWnd = GetForegroundWindow
            If Action = "CTRMOUSEAC" Then
                Rem - Okay. Now we need to get the active control. Before we can, we need to allow ourselves
                Rem - access to the active applications messages. Ha. Gotta love it...
                lngThreadActive = GetWindowThreadProcessId(lngHWnd, ByVal 0&)
                Call AttachThreadInput(App.ThreadID, lngThreadActive, True) ' Attach to thread
                lngHWnd = GetFocus
                Call AttachThreadInput(App.ThreadID, lngThreadActive, False) 'Dettach
            End If
            
            Call GetWindowRect(lngHWnd, rctRect)
            With rctRect
                lngX = (.left + .Right) \ 2
                lngY = (.Top + .Bottom) \ 2
            End With
            
    End Select
    Call SetCursorPos(lngX, lngY)
    
End Sub

Private Function p_KeysAreDown(ByRef prop As clsZenDictionary) As Boolean ', ByVal All As Boolean) As Boolean
Rem - If 'All' is True, returns True if all keys in the Prop string have been released, else it returns False
Rem - If 'All' is False, returns True if any key in the Prop string has been released, else it returns False
Dim colKeys As Collection
Dim strShift As String
Dim k As Integer
Const VK_SHIFT = &H10
Const VK_CONTROL = &H11
Const VK_MENU = &H12

    Set colKeys = New Collection
    strShift = prop("ShiftKey")
    If InStr(strShift, "Alt") > 0 Then colKeys.Add VK_MENU 'ShiftValue = 1
    If InStr(strShift, "Ctrl") > 0 Then colKeys.Add VK_CONTROL 'ShiftValue = ShiftValue + 2
    If InStr(strShift, "Shift") > 0 Then colKeys.Add VK_SHIFT 'ShiftValue = ShiftValue + 4
    k = CLng(Val(prop("Hotkey")))
    If k <> 0 Then colKeys.Add k
    
    Rem - Wait until all keys are released....
    For k = colKeys.Count To 1 Step -1 ' The key has been released.
        If GetAsyncKeyState(colKeys.item(k)) = 0 Then Exit Function
    Next k
    p_KeysAreDown = True

End Function

Private Sub p_Mouse(ByVal Button As Long)
Rem - Button - 0 = Left, 1 = right, 2 = middle
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_MIDDLEDOWN = &H20
Const MOUSEEVENTF_MIDDLEUP = &H40
Const MOUSEEVENTF_MOVE = &H1
Const MOUSEEVENTF_ABSOLUTE = &H8000
Const MOUSEEVENTF_RIGHTDOWN = &H8
Const MOUSEEVENTF_RIGHTUP = &H10

    Select Case Button
        Case 0
            mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&
        Case 1
            mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0&, 0&, 0&, 0&
        Case Else '"MOUSEMCLICK"
            mouse_event MOUSEEVENTF_MIDDLEDOWN Or MOUSEEVENTF_MIDDLEUP, 0&, 0&, 0&, 0&
    End Select
    DoEvents
            
End Sub

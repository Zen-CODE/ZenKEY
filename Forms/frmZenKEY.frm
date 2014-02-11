VERSION 5.00
Begin VB.Form frmZenKEY 
   BorderStyle     =   0  'None
   ClientHeight    =   1395
   ClientLeft      =   3645
   ClientTop       =   9480
   ClientWidth     =   1995
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmZenKEY.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   93
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   133
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrBalloon 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrHook 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   480
   End
   Begin VB.PictureBox picSubmenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   1560
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picZenKEY 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1200
      Picture         =   "frmZenKEY.frx":058A
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "frmZenKEY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Rem =================================  Printing text, each letter a cutout
'Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Rem ======================================= Round rect
'Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private booEndIt As Boolean
Private booDragOn As Boolean

Rem - For getting the picturebox area for setting the form region ........
'Private Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Public MenuShowing As Boolean

Public HWndUnderMenu As Long
Private CurPos As POINTAPI
Rem - For identifying menu items for changing caption
Public AWT As clsAWT
Private lngRCMenu As Long
Private lngMainMenu As Long
Private lngDTMMenu As Long
Rem ============= For getting posisiton of window
'
'Private Type POINTAPI
'        x As Long
'        y As Long
'End Type
'Private Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type
'Private Type
'        Length As Long
'        flags As Long
'        showCmd As Long
'        ptMinPosition As POINTAPI
'        ptMaxPosition As POINTAPI
'        rcNormalPosition As RECT
'End Type
'Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Rem = For trying to detect movement. Put here for speed - don't want to have to create on each firing
'Private lngLeft As Single
'Private lngTop As Single
Private lngPrevLeft As Long
Private lngPrevTop As Long
Private lngPrevRight As Long
Private lngPrevBottom As Long
Public booMouseDown As Boolean
Rem ============================= For extracting icons =======================================
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long

Rem ============================= For extracting icons =======================================
Private Type MenuItem
    ParentMenu As Long
    ItemCount As Long
    ParentCaption As String
    SubMenuStarted As Boolean
End Type
Dim strSkinPath As String
Dim rctCurrent As RECT
Dim booInTray As Boolean
'Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Dim booActiveMaxed As Boolean
Dim sngWinPos As Single
Dim lngCount As Long
Dim booShowBalloon As Boolean

Private Sub Form_Initialize()

    Dim X As Long
    X = InitCommonControls

End Sub

Private Sub Form_Load()
    
    sngWinPos = Val(settings("WindowPos"))
    If sngWinPos = 0 Then sngWinPos = 70
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const WM_RBUTTONUP = &H205
Const WM_LBUTTONUP = &H202
Dim Msg As Single
    
    'If booModalShown Then Exit Sub
    Rem - Show click menu immediately

    If booInTray Then
        Msg = Me.ScaleX(X, vbPixels, vbTwips) / Screen.TwipsPerPixelX
        If (Msg = WM_LBUTTONUP) Then
            Call ShowMenu(lngMainMenu)
        ElseIf (Msg = WM_RBUTTONUP) Then
            If lngRCMenu <> 0 Then Call ShowMenu(lngRCMenu)
        Else
            Call GetCursorPos(CurPos)
            booMouseDown = True
        End If
    Else
        Call GetCursorPos(CurPos)
        booMouseDown = True
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If booMouseDown Then
        Dim CurPos2 As POINTAPI
        Dim SWidth As Single, SHeight As Single
        Dim FLeft As Single
        Dim FTop As Single
        Dim dx As Single, dy As Single
        
            Call GetCursorPos(CurPos2)
            With Me
                dx = CurPos2.X - CurPos.X
                dy = CurPos2.Y - CurPos.Y
                If Abs(dx) < 3 And Abs(dy) < 3 Then
                    Rem - They have clicked
                    If Button = 2 Then
                        If lngRCMenu <> 0 Then
                            Call DoAction(zenDic("Action", "SHOWMENU", "MenuHandle", lngRCMenu))
                        End If
                    Else
                        Call ShowMenu(lngMainMenu)
                    End If
                Else
                    .Move Me.left + Me.ScaleX(dx, vbPixels, vbTwips), Me.Top + Me.ScaleY(dy, vbPixels, vbTwips)
                End If
            
            End With
        
        booMouseDown = False
    End If

End Sub


Public Sub Menu_Build()
Dim lngMaxItem As Long, lngZKMax As Long
Dim k As Integer
Dim GroupSep As Long ' Number of items per grouping
Dim GroupLim As Long
Dim bNotRestarting As Boolean

    On Error GoTo ErrorTrap
    
    bNotRestarting = Not CBool(Command$ = "RESTART")
    GroupSep = Val(settings("GroupSep"))
    If GroupSep < 1 Then GroupSep = 4
    GroupLim = Val(settings("GroupLim"))
    If GroupLim < 1 Then GroupLim = 50 '100
        
    Rem - DTMMenu
    lngZKMax = UBound(ZKMenu())
    If DTM_Enabled Then
        Call Menu_AddDTM
        lngMaxItem = UBound(ZKMenu())
    Else
        lngMaxItem = lngZKMax
    End If
    
    Rem - Loop though ZKMenu array, creating a popup menu for each item...
    With picZenKEY
        Dim lngLevel As Long, strClass As String
        Dim pMenu() As MenuItem, strCaption As String
        
        Rem - NOTE:*** lngMainMenu is the global base menu handle ***
        ReDim pMenu(0)
        
        pMenu(0).ParentMenu = cMenu.GetFormMenu(Me)
        lngMainMenu = cMenu.Add(pMenu(0).ParentMenu, "Zenkey", True, 0)
        pMenu(0).ParentMenu = lngMainMenu
        lngLevel = 0
        
        For k = 0 To lngMaxItem
            Rem - Add the hotkey to the caption if it has one
            strClass = ZKMenu(k)("Class")
            strCaption = Hotkeys.GetCaption(ZKMenu(k)) 'If either key has used, add the key names to the caption
            If Len(strCaption) > 0 Then
                strCaption = ZKMenu(k)("Caption") & " (" & strCaption & ")"
            Else
                strCaption = ZKMenu(k)("Caption")
            End If
            ZKMenu(k)("MenuID") = CStr(k + MNU_Start)  ' Set ID
            
            Select Case strClass
                Case "Group"
                    Rem -----------------------------
                    Rem - Create a submenu
                    Rem -----------------------------
                    pMenu(lngLevel).ItemCount = pMenu(lngLevel).ItemCount + 1
                    lngLevel = lngLevel + 1
                    ReDim Preserve pMenu(lngLevel)
                    pMenu(lngLevel).ParentMenu = cMenu.Add(pMenu(lngLevel - 1).ParentMenu, strCaption, True, k + MNU_Start)
                    ZKMenu(k)("MenuHandle") = pMenu(lngLevel).ParentMenu
                    Rem - Set the right click menu
                    Rem - DTMMenu
                    Select Case ZKMenu(k)("RIGHTCLICKMENU")
                        Case "True": lngRCMenu = pMenu(lngLevel).ParentMenu
                        Case "DTM"
                            Rem - Set this item as a Desktop map menu, then remove it from the main menu.
                            Dim strName As String
                            strName = ZKMenu(k)("MENU")
                            If DTM_Enabled Then ZK_DTM.MenuHandles.item(strName) = pMenu(lngLevel).ParentMenu
                            If lngLevel = 1 Then Call cMenu.SetItemProp(lngMainMenu, "REMOVE", "", k + MNU_Start)
                    End Select
                    pMenu(lngLevel).ParentCaption = strCaption
                    
                Case ""
                    If ZKMenu(k)("ENDGROUP") = "True" Then
                        Rem -------------------------------------
                        Rem - End the current submenu
                        pMenu(lngLevel).ItemCount = 0
                        lngLevel = lngLevel - 1
                        If pMenu(lngLevel).SubMenuStarted Then
                            pMenu(lngLevel).SubMenuStarted = False
                            pMenu(lngLevel).ItemCount = 0
                            lngLevel = lngLevel - 1
                        End If
                    End If
                Case Else
                    Rem ---------------------------------------------
                    Select Case True
                        Case (pMenu(lngLevel).ItemCount < GroupLim) Or (pMenu(lngLevel).ItemCount Mod GroupLim) <> 0
                            Rem - Do Nothing
                        Case Else ' 'Case pMenu(lngLevel).ItemCount Mod GroupLim = 0
                            Rem - End the previous submenu if we were busy
                            If lngLevel > 0 Then
                                If pMenu(lngLevel - 1).SubMenuStarted Then
                                    pMenu(lngLevel).ItemCount = 0
                                    lngLevel = lngLevel - 1
                                End If
                            End If
                        
                            Rem - Create a submenu
                            pMenu(lngLevel).SubMenuStarted = True
                            pMenu(lngLevel).ItemCount = pMenu(lngLevel).ItemCount + 1
                            lngLevel = lngLevel + 1
                            ReDim Preserve pMenu(lngLevel)
                            pMenu(lngLevel).ParentMenu = cMenu.Add(pMenu(lngLevel - 1).ParentMenu, pMenu(lngLevel - 1).ParentCaption & " cont...", True, k + MNU_Start)
                    End Select
                    Rem - Now add the item itself
                    pMenu(lngLevel).ItemCount = pMenu(lngLevel).ItemCount + 1
                    ZKMenu(k)("MenuHandle") = pMenu(lngLevel).ParentMenu  ' Set parent menu
                    Call cMenu.Add(pMenu(lngLevel).ParentMenu, strCaption, False, k + MNU_Start)
            End Select
            Rem - Add the item to the current menu, only launch if not restarting
            If bNotRestarting Then If Len(ZKMenu(k)("Startup")) > 0 Then Call ShellExe(App.Path & "\ZenKP.exe", ZKMenu(k).ToProp)


            Rem - Add a seperator if appropriate
            Select Case True
                Case lngLevel > 0
                    If ((pMenu(lngLevel).ItemCount - 1) Mod GroupSep = GroupSep - 1) Then
                        If UBound(ZKMenu) >= k + 1 Then
                            If ZKMenu(k + 1)("ENDGROUP") <> "True" Then
                                Call cMenu.Add(pMenu(lngLevel).ParentMenu, "-", False, k + MNU_Start + 2 * lngMaxItem)
                            End If
                        End If
                    End If
                Case (k > 3) And (k < lngZKMax - 2)
                    If ((pMenu(lngLevel).ItemCount + 1) Mod GroupSep = GroupSep - 1) Then Call cMenu.Add(pMenu(lngLevel).ParentMenu, "-", False, k + MNU_Start + 2 * lngMaxItem)
            End Select
        Next k
    End With
    Exit Sub

ErrorTrap:
    Err.Source = "Menu_Build"
    Call ZenErr("k = " & CStr(k) & ", ZKMenu = " & ZKMenu(k).ToProp)
    Resume Next
    
End Sub









Public Sub SetFormRegion()
Dim X1 As Single, Y1 As Single
Dim sngWidth As Single, sngHeight As Single
    
    With Me
        sngWidth = .ScaleX(.Picture.Width, vbHimetric, vbTwips) '+ Me.ScaleX(1, vbPixels, vbTwips)
        sngHeight = .ScaleX(.Picture.Height, vbHimetric, vbTwips) ' + Me.ScaleY(1, vbPixels, vbTwips)
        If Len(progInfo("X1")) > 0 Then
            X1 = Val(progInfo("X1"))
            Y1 = Val(progInfo("Y1"))
        End If
        If X1 >= 1 Then
            Me.Move X1, Y1, sngWidth, sngHeight
        Else
            .Move 0.995 * Screen.Width - Me.Width, 0.01 * Screen.Height, sngWidth, sngHeight
        End If
        Rem - Load the icons if appropriate
        If settings("ICN_Load") = "Y" Then Call Icon_Load(progInfo)
    End With

End Sub

Private Sub Load_Hotkeys()
Dim k As Long
Dim MaxMenu As Long
Dim lngAns As Long
Dim strMessage As String
Dim strHK As String
Dim strAction As String

    MaxMenu = UBound(ZKMenu())
    For k = 0 To MaxMenu
        strHK = ZKMenu(k)("Hotkey")
        strMessage = ZKMenu(k)("ShiftKey")
        If Len(strMessage) > 0 Or Val(strHK) > 0 Then
                Rem - Add a Hotkey to show the menu
                strAction = ZKMenu(k)("Action")
                If Not Hotkeys.AddHotkey(k + MNU_Start, strMessage, Val(strHK), CBool(strAction = "TOGGLEHKEX")) Then
                    If settings("HKWARN") <> "N" Then
                        Rem - Tell them the key can't register
                        If lngAns <> 1 Then
                            Select Case True
                                Case Len(strMessage) = 0: strMessage = Hotkeys.Keyname(Val(strHK)) ' Only a hotkey
                                Case Len(strHK) > 0: strMessage = strMessage & " + " & Hotkeys.Keyname(Val(strHK)) ' Both a hotkey & a shiftkey
                            End Select
                            '"' + '" & Hotkeys.Keyname(Val(strAns)) &
                            strMessage = "Unable to create Hotkey for the '" & ZKMenu(k)("Caption") & ", ZenKEY action, using '" & strMessage & _
                                "'. This normally means it is being used by another application. Either disable the hotkey, or disable the Hotkey in the application using it." & _
                                vbCr & vbCr & "Do you want to be told each time a Hotkey cannot be loaded?"
                            lngAns = ZenMB(strMessage, "Yes", "No")
                        End If
                    End If
                End If
            End If

    Next k
    Hotkeys.booLoaded = True

End Sub



Public Sub DoAction(ByVal prop As clsZenDictionary)
Dim lngIndex As Long, booTemp As Boolean

    On Error GoTo ErrorTrap:

    lngIndex = CLng(Val((prop("Index")))) - MNU_Start
    If lngIndex <> -MNU_Start Then
        Rem - Fired by WaitForMessages ie. they have pressed a hotkey
        Set prop = ZKMenu(lngIndex)
        If booShowBalloon Then Call Balloon_Show("Action fired - " & prop("Caption"))
        If prop("Class") <> "ZenKEY" Then
            Dim it As Object
            Set it = ZK_GetObject(prop("Class"))
            Call it.DoAction(prop)
            Exit Sub
        End If
    End If
    
    Select Case prop("Action") ' UCase(Action)
        Case "TOGGLEAWT"
            booTemp = CBool(AWT Is Nothing)
            If booTemp Then
                Set AWT = New clsAWT
            Else
                Call AWT.AWT_Flush
                Set AWT = Nothing
            End If
            Call p_ToggleMenu(lngIndex, booTemp)
        Case "FOLLOWACTIVE"
            SET_FollowActive = Not SET_FollowActive
            Call p_ToggleMenu(lngIndex, SET_FollowActive)
        Case "WindowUnderMouse"
            ZK_WinUMouse = Not ZK_WinUMouse
            Call p_ToggleMenu(lngIndex, ZK_WinUMouse)
        Case "GROUP"
            Call ShowMenu(Val(prop("MenuHandle")))
        Case "SHOWMENU", "SHOWMENUMAIN"
            Call ShowMenu(lngMainMenu)
        Case "SHOWMENURCLICK"
            If lngRCMenu <> 0 Then
                Call ShowMenu(lngRCMenu)
            Else
                Call ShowMenu(lngMainMenu)
            End If
        Case "HIDEFORM"
            booTemp = Not Me.Visible
            Me.Visible = booTemp
            settings("HideForm") = IIf(booTemp, "False", "True")
            Call p_ToggleMenu(lngIndex, booTemp)
        
        Case "SETAOT"
            booTemp = Not CBool(settings("AOT") = "True")
            Call SetWinPos(Me.hwnd, IIf(booTemp, HWND_TOPMOST, HWND_NOTOPMOST), True)
            settings("AOT") = IIf(booTemp, "True", "False")
            Call p_ToggleMenu(lngIndex, booTemp)
        Case "REFRESHVIS"
            If Not Me.Visible Then Call SetWinPos(Me.hwnd, SET_Layer, False)
        Case "WIZARD"
            Call ShellExe(App.Path & "\ZenWiz.exe")
        Case "EMAIL"
            Call ShellExe("mailto:zenkey.zencode@gmail.com")
        Case "HELP"
            Call ShellExe(App.Path & "\Help\Index.htm")
        Case "TOGGLEHOTKEYS", "TOGGLEHKEX"
            Rem - Do not change the caption as if there are multiple entries in the menu, they become inconsist
            Dim strIcon As String
            booTemp = Hotkeys.booLoaded
            If booTemp Then
                Call Hotkeys.Unload
                strIcon = strSkinPath & "\TrayD.ico"
            Else
                Call Load_Hotkeys
                strIcon = strSkinPath & "\Tray.ico"
            End If
            Call p_ToggleMenu(lngIndex, booTemp)
            If booInTray Then
                If Len(Dir(strIcon)) > 0 Then Set Me.Icon = LoadPicture(strIcon) Else Set Me.Icon = LoadPicture()
                Call DoAction(zenDic("Action", "RefreshTray"))
            End If
        Case "ABOUT"
            Call ShellExe(App.Path & "\ZKConfig.exe", "ABOUT")
        Case "ShowQuote"
            Dim booMore As Boolean, strPrev As String
            strPrev = settings("HideQuotes")
            settings("HideQuotes") = "False"
            Do
                booMore = CBool(1 = ZenMB(ZenKEYCap, "OK", "More"))
            Loop While booMore
            settings("HideQuotes") = strPrev
        Case "EXIT"
            #If LOGMODE > 0 Then
                Call LOG_Write("Exiting.....................")
            #End If
            Call CloseApp
        Case "PositionForm"
            Call PositionForm(prop)
            Exit Sub
        Case "RefreshTray"
            If booInTray Then
                Call Systray_Del(Me)
                Call Systray_Add(Me, Me.Icon, ZenKEYCap)
            Else
                Call ZenMB("Sorry, but this skin does not use a system tray icon.")
            End If
        Case "CONFIG"
            Call ShellExe(App.Path & "\ZKConfig.exe")
        Case "CONFIG_SETTINGS"
            Call ShellExe(App.Path & "\ZKConfig.exe", "SETTINGS")
        Case "CONFIG_ITEMS"
            Call ShellExe(App.Path & "\ZKConfig.exe", "ITEMS")
        Case Else
            If prop("Class") = "GROUP" Then
                ' If it has no action but it's a group, show it
                Call ShowMenu(Val(prop("MenuHandle")))
            Else
                Call ZenMB("Unknown action in ZenKEY - '" & prop("Action") & "'")
            End If
    End Select
    Exit Sub

ErrorTrap:
    Call ZenErr(prop.ToProp)
    Resume Next
        
End Sub

Public Sub ShowMenu(ByVal Handle As Long)
Rem - DTMMenu - Sub ShowMenu(ByVal Handle As Long)
Const TPM_CENTERALIGN = &H4&

    If ZK_WinUMouse Then HWndUnderMenu = WindowFromCursor
        
    'On Error Resume Next ' Get can't show modal form error swhen message box dispalyed
    Call SetWinPos(MainForm.hwnd, SET_Layer, True)
    
    Rem = Shoe the normal menu, just the form or the Windows pop-up menu
    Dim CurPos As POINTAPI
    Call GetCursorPos(CurPos)
    
    Rem - Added to prevent changing of application z-order changing
    tmrHook.Enabled = False

    Call TrackPopupMenu(Handle, TPM_CENTERALIGN, CurPos.X, CurPos.Y, 0, MainForm.hwnd, ByVal 0&)
    If settings("HideForm") = "True" Then MainForm.Hide

    Rem - Added to prevent changing of application z-order changing
    tmrHook.Enabled = True


End Sub
Private Function Menu_SetItemPic(ByRef prop As clsZenDictionary, ByVal IconMode As Long, ByVal MenuID As Long) As PictureBox
Dim lngIndex As Long
Const GROUP = 4 ' Number of items per grouping
Dim lngMenuHandle As Long

    lngMenuHandle = Val(prop("MenuHandle"))
    Select Case IconMode
        Case 2
            Rem - Zen Icons
            Call cMenu.SetItemProp(lngMenuHandle, "PICTURE", picSubmenu(MenuID Mod GROUP), MenuID)
        Case 3
            Rem - Icons using the starting letter of the caption
            Dim strCaption As String
            Dim lngVal As Long

            
            strCaption = prop("Caption")
            If Len(strCaption) > 0 Then lngVal = Asc(LCase(left$(strCaption, 1)))
            Rem- Picboxes 2-27 hold ask letter, 38-47 numbers
            Select Case lngVal
                Case Asc("a") To Asc("z"): lngIndex = lngVal - Asc("a") + 2
                Case Asc("0") To Asc("9"): lngIndex = lngVal - Asc("0") + 38
                Case Else: lngIndex = 1
            End Select
            Call cMenu.SetItemProp(lngMenuHandle, "PICTURE", picSubmenu(lngIndex), MenuID)
        Case 4, 5
            Rem - 4. Program icons
            Rem - 5. Large program icons
            lngIndex = 0
            Select Case prop("Class")
                Case "File", "Folder"
                    Rem - Now draw the associated icon
                    Dim mIcon As Long, RetVal As Long

                    If Mid(prop("Action"), 2, 1) = ":" Then
                        'ExtractIconEx Prop_Get("Action", strProp), 0, ByVal 0, mIcon, 1
                        If IconMode = 5 Then ' Large icons
                            RetVal = ExtractIconEx(prop("Action"), 0, mIcon, ByVal 0, 1)
                        Else ' Small icons
                            RetVal = ExtractIconEx(prop("Action"), 0, ByVal 0, mIcon, 1)
                        End If
                        
                        Const DI_MASK = &H1
                        Const DI_IMAGE = &H2
                        Const DI_NORMAL = DI_MASK Or DI_IMAGE
                        If RetVal = 1 Then
                            
                            Rem - Load the picture box and draw the icons and caption into it.......
                            lngIndex = picSubmenu.UBound + 1
                            Load picSubmenu(lngIndex)
                            With picSubmenu(lngIndex)
                                .Picture = LoadPicture(vbNullString)
                                If IconMode = 5 Then ' Large icons
                                    .Height = 33 '17
                                    .Width = 33 '22
                                    Call DrawIconEx(.hdc, 0, 0, mIcon, 32, 32, 0, 0, DI_NORMAL)
                                Else ' Small icons
                                    .Height = 17
                                    .Width = 22 + .TextWidth(prop("Caption"))
                                    Call DrawIconEx(.hdc, 0, 0, mIcon, 16, 16, 0, 0, DI_NORMAL)
                                    .CurrentY = 1
                                    .CurrentX = 19
                                    picSubmenu(lngIndex).Print prop("Caption")
                                End If
                                
                                Set .Picture = .Image
                            End With
                            Call DestroyIcon(mIcon)
                        End If
                    End If
            End Select
            If lngIndex = 0 Then
                Call cMenu.SetItemProp(lngMenuHandle, "PICTURE", picSubmenu(MenuID Mod GROUP), MenuID)
            Else
                Call cMenu.SetItemProp(lngMenuHandle, "BITMAP", picSubmenu(lngIndex), MenuID)
            End If
    End Select
    
End Function



Public Sub Load_Graphics()
On Error Resume Next

    picZenKEY.Picture = LoadPicture(strSkinPath & "\MainIcon.bmp")
    Me.Picture = LoadPicture(strSkinPath & "\Form.jpg")
    'Me.Width = Me.ScaleX(picForm.Width + 10, Me.ScaleMode, vbTwips)
    If Len(Dir(strSkinPath & "\Tray.ico")) > 0 Then
        Set Me.Icon = LoadPicture(strSkinPath & "\Tray.ico")
    Else
        Set Me.Icon = LoadPicture()
    End If

End Sub

Public Sub Initialise()
    
    Call Menu_LoadINI("ZenKEY.ini")
    
    Rem - Set hooks
    Call cMenu.SetHook(Me) ' Menu hook
    Const GWL_WNDPROC = (-4)
    lngTaskBarMsg = RegisterWindowMessage(ByVal "TaskbarCreated") ' Explorer restart form tray icon
    
    Call Menu_Build
    Call Menu_SetIcons

    Call Load_Hotkeys
    If Me.Icon.Handle <> 0 Then booInTray = Systray_Add(Me, Me.Icon, ZenKEYCap)
    If settings("HideForm") <> "True" Then Call DoAction(zenDic("Action", "REFRESHVIS"))
    
    Call Registry.SetRegistry(HKCU, "SOFTWARE\ZenCODE\ZenKEY", "WindowHandle", CStr(Me.hwnd))
    If settings("AutoTrans") = "True" Then Set AWT = New clsAWT Else Set AWT = Nothing
    
    If SET_Trans > 0 Then Call ZK_Win.SetTrans(MainForm.hwnd, SET_Trans)
    tmrHook.Enabled = True
    
End Sub




Private Sub PositionForm(ByRef prop As clsZenDictionary)
Dim lngH As Long
Dim lngRet As Long, rctRect As RECT
Dim booRepos As Boolean
    
    lngH = Val(prop("Hwnd"))
    lngRet = GetWindowRect(lngH, rctRect)
    
    If lngRet <> 0 Then
        If Not IsIconic(lngH) Then
            With rctRect
                Select Case True
                    Case IsZoomed(lngH)
                        Rem - Do nothing - place in default
                        If Not booActiveMaxed Then
                            lngPrevLeft = (sngWinPos / 100) * MainForm.ScaleX(Screen.Width - Me.Width, vbTwips, vbPixels)
                            lngPrevTop = 0
                            booRepos = True
                            booActiveMaxed = True
                        End If
                    Case (.Right - .left) > 0 And (.Bottom - .Top) > 0
                        Rem - If the window is too small, ignore it
                        Rem - Check that the window has moved
                        booRepos = True
                        If Not booActiveMaxed Then
                            If .left = lngPrevLeft Then
                                If .Top = lngPrevTop Then
                                    If .Right = lngPrevRight Then
                                        If .Bottom = lngPrevBottom Then
                                            booRepos = False
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        booActiveMaxed = False
                        If booRepos Then
                            Rem - Attach to window
                            Dim sngFormWidth As Single, sngFormHeight As Single
                            sngFormWidth = Me.ScaleX(Me.Width, vbTwips, vbPixels)
                            sngFormHeight = Me.ScaleY(Me.Height, vbTwips, vbPixels)
                            lngPrevLeft = .left + (sngWinPos / 100) * (.Right - .left)
                            lngPrevTop = .Top
                            Select Case True
                                Case lngPrevTop > sngFormHeight
                                    Rem - Put above if possible
                                    lngPrevTop = lngPrevTop - sngFormHeight
                                Case (lngPrevTop < sngFormHeight) And ((sngFormHeight + .Bottom) < Me.ScaleY(Screen.Height, vbTwips, vbPixels) - sngFormHeight) '
                                    Rem - If Window at top, or too near to top to fit form above
                                    Rem - Put Below
                                    lngPrevTop = .Bottom
                            End Select
                            
                            Rem - Ensure the form is on screen
                            If lngPrevLeft < 0 Then
                                lngPrevLeft = 0
                            ElseIf lngPrevLeft > Me.ScaleX(Screen.Width, vbTwips, vbPixels) - sngFormWidth Then
                                lngPrevLeft = Me.ScaleX(Screen.Width, vbTwips, vbPixels) - sngFormWidth
                            End If
                            If lngPrevTop < 0 Then
                                lngPrevTop = 0
                            ElseIf lngPrevTop > Me.ScaleY(Screen.Height, vbTwips, vbPixels) - sngFormHeight Then
                                lngPrevTop = Me.ScaleY(Screen.Height, vbTwips, vbPixels) - sngFormHeight
                            End If
                        End If
                End Select
                
                If booRepos Then
                    Rem - Well, place the damn window then
                    Me.Move Me.ScaleX(lngPrevLeft, vbPixels, vbTwips), Me.ScaleY(lngPrevTop, vbPixels, vbTwips)
                    If settings("HideForm") <> "True" Then Call DoAction(zenDic("Action", "REFRESHVIS"))
                    lngPrevLeft = .left
                    lngPrevTop = .Top
                    lngPrevRight = .Right
                    lngPrevBottom = .Bottom
                End If
                
            End With
        End If
    End If ' If lngRet <> 0
    
End Sub

Public Sub Vars_Init()
Dim strTemp As String, k As Long
    
    Set ZK_Win = New clsWindows
    
    Rem - Icon mode. 1 = No icons, 2 = ZenKey icons, 3 = Load from files
    k = Val(settings("IconMode"))
    If k = 0 Then settings("IconMode") = "2"  ' Secs default
    
    Rem - SkinPath
    strTemp = settings("Skin")
    If Len(strTemp) = 0 Then
        strSkinPath = App.Path & "\Skins\Default"
    Else
        strSkinPath = App.Path & "\Skins\" & strTemp
    End If
    
    SET_FollowActive = CBool(settings("FollowActive") = "True")
    SET_Trans = Val(settings("SET_Trans"))
    
    strTemp = settings("SET_Layer")
    If Len(strTemp) > 0 Then
        SET_Layer = Val(settings("SET_Layer"))
    Else
        SET_Layer = 0
    End If
    
    ZK_WinUMouse = CBool(settings("WindowUnderMouse") = "True")
    
    Rem - History depth
    k = Val(settings("HistDepth"))
    If k < 1 Then settings("HistDepth") = "10"
    
    Rem - Notify actions
    strTemp = settings("Notify")
    booShowBalloon = CBool(strTemp <> "None")
    
    Rem - Windows shift values
    WIN_Shift = Val(settings("WinShift"))
    If WIN_Shift = 0 Then WIN_Shift = 10
    
    Rem - Auto-window transparency
    AWT_Depth = Val(settings("AWTDepth"))
    If AWT_Depth < 1 Then AWT_Depth = 1
    
    Rem - Infinite Desktop
    IDT_Enabled = CBool(settings("IDT_Enable") = "Y")
    If IDT_Enabled Then
        Set ZK_IDT = New clsIDT
        IDT_AutoFocus = CBool(settings("IDTAutoFocus") = "True")
    End If
    
    Rem - The desktop map
    strTemp = settings("DTM_Position")
    DTM_Enabled = CBool(strTemp <> "6") ' Do now show desktop map for time
    If DTM_Enabled Then Set ZK_DTM = New frmDesktopMap
    
    Rem - ZenBar
    SET_ZenBar = CBool(settings("ICN_Zenbar") = "Y")
    Set ICN_Forms = New Collection
    
End Sub


Private Sub Menu_SetIcons()
Dim lngIconMode As Long
Dim intMaxItem As Long
Const GROUP = 4 ' Number of items per grouping
Dim k As Long
Dim lngMenuHandle As Long
    
    On Error Resume Next
    
    Rem - DTMMenu

    Rem - Setup the pictures for the various modes
    lngIconMode = Val(settings("IconMode"))

    Select Case lngIconMode
        Case 1
            Rem - No icons
            Rem - DTMMenu
            'Exit Sub
        Case 2, 4, 5
            Rem - 2 - The Zen icons i.e. The Z in 4 different orientations
            Rem - 4 - Program icons if possible
            Rem - 5 - Large program icons if possible
            Rem - Only load and draw these icons if they are required
            For k = 0 To GROUP - 1
                If k > 0 Then Load picSubmenu(k)
                With picSubmenu(k)
                    
                    Rem - Redraw the ZenCode icon uis different orientations
                    Const lngH As Long = 12  'picZenKEY.Height
                    Const lngW As Long = 12 'picZenKEY.Width
                    Select Case k
                        Case 0: picSubmenu(k).PaintPicture picZenKEY.Picture, 0, 0, lngW, lngH
                        Case 1: picSubmenu(k).PaintPicture picZenKEY.Picture, 0, lngH - 1, lngW, -lngH
                        Case 2: picSubmenu(k).PaintPicture picZenKEY.Picture, lngW - 1, 0, -lngW, lngH
                        Case 3: picSubmenu(k).PaintPicture picZenKEY.Picture, lngW - 1, lngH - 1, -lngW, -lngH
                    End Select
                    Set .Picture = .Image
                End With
            Next k
        Case 3
            Rem - 3 - A letter, depending on the caption of the icon. Picture loaded later
            picSubmenu(0).Picture = LoadPicture(strSkinPath & "\MainIcon.bmp")
            Load picSubmenu(1)
            picSubmenu(1).Picture = LoadPicture(strSkinPath & "\Blank.bmp")
            'For k = 0 To 25 'Asc("a") To Asc("z")
            For k = 0 To 35 'Asc("a") To Asc("z")
                Load picSubmenu(k + 2)
                If k < 26 Then
                    Rem- Picboxes 2-27 hold ask letter, 38-47 numbers
                    picSubmenu(k + 2).Picture = LoadPicture(strSkinPath & "\" & Chr$(Asc("a") + k) & ".bmp")
                Else
                    picSubmenu(k + 2).Picture = LoadPicture(strSkinPath & "\" & CStr(k - 26) & ".bmp")
                End If
            Next k
    End Select
    
    intMaxItem = UBound(ZKMenu())
    For k = 0 To intMaxItem
        Rem - Add a picture to main menu item
        Select Case ZKMenu(k)("Menu")
            Case "TRANS"
                If DTM_Enabled Then
                    Rem - If it is the transparency menu in the DTM, check to see if it shoul be checked
                    Select Case SET_Trans
                        Case -1, 0, 100
                            If ZKMenu(k)("Level") = "-1" Then Call ZK_DTM.SetTransIcon(ZKMenu(k))
                        Case Else
                            If SET_Trans = Val(ZKMenu(k)("Level")) Then Call ZK_DTM.SetTransIcon(ZKMenu(k))
                    End Select
                End If
            Case "T"
                lngMenuHandle = Val(ZKMenu(k)("MenuHandle"))
                Call cMenu.SetItemProp(lngMenuHandle, "CHECK", ZKMenu(k)("Caption"), k + MNU_Start)
            Case "C"
                Rem - Do nothing. No picture.
            Case Else
                If lngIconMode <> 1 Then Call Menu_SetItemPic(ZKMenu(k), lngIconMode, k + MNU_Start)
        End Select
    Next k
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call CloseApp
End Sub

Private Sub tmrBalloon_Timer()

    Call Balloon_Show(vbNullString)

End Sub

Private Sub tmrHook_Timer()
Dim lngH As Long
                                                              
    lngH = GetForegroundWindow
    
    If WIN_Active <> lngH Then
        Rem - Check if the Active window had changed
        If (lngH <> 0) Then
            If lngH <> ActiveWindow(0) Then
                Rem - The active window has changed. Maintain the active window list
                If WindowIsUsable(lngH) Then
                    Call AddToActive(lngH)
                    IDT_ActiveApp = GetExeFromHandle(lngH)
                End If
            End If
            WIN_Active = lngH
            
            If DTM_Enabled Then lngCount = 0 ' Trigger an update
            If Not (AWT Is Nothing) Then Call AWT.SetAutoTrans(lngH)
            ROL_Check = CBool(ROL_Count > 0)
        End If
    End If
         
    If ROL_Check Then Call ROL_Focus(lngH)

    Rem - Instead of timers on each form, we try a loop
    If lngCount = 0 Then
        Rem - If we are not displaying a message, the follow the active window (if chosen)
        If SET_FollowActive Then
            If MainForm.Visible Then
                If MainForm.hwnd <> WIN_Active Then Call DoAction(zenDic("Action", "PositionForm", "HWnd", WIN_Active))
            End If
        End If
    
        Rem - Update the icons
        Call Icon_UpdateAll
        Rem - Show the desktop map?
        If DTM_Enabled Then
            If ZK_DTM.Visible Then
                Call ZK_DTM.DrawWindows  ' Show all the time
                Call ZK_DTM.DrawActiveIcon
            End If
        End If
        
        If SET_ZenBar Then Call Icon_Layout
    End If
    lngCount = (lngCount + 1) Mod 5
    
End Sub
Private Sub AddToActive(ByVal hwnd As Long)
Dim k As Long, lngPrev As Long

    'Prevent the buildup of windows that we cannt see or use.
    'Added to help desktop movement get a better history of usable apps
    If Len(ClassName(hwnd)) = 0 Then Exit Sub

    lngPrev = -1
    For k = 0 To 15
        If ActiveWindow(k) = hwnd Then
            lngPrev = k
            Exit For
        End If
    Next k
    
    Select Case lngPrev
        Case 0
            Exit Sub
        Case Is > 0
            Rem - Window already exists in our stack. Move it to the top
            'ActiveWindow(lngPrev) = ActiveWindow(0)
            For k = lngPrev To 1 Step -1
                ActiveWindow(k) = ActiveWindow(k - 1)
            Next k
        Case Is < 0
            Rem - A totally new window. Pop it on top.
            For k = 15 To 1 Step -1
                ActiveWindow(k) = ActiveWindow(k - 1)
            Next k
    End Select
    ActiveWindow(0) = hwnd

End Sub




Private Sub Menu_AddDTM()
Dim lngMax As Long, lngCount As Long
Dim FNum As Long, strLine As String

    FNum = FreeFile
    lngMax = UBound(ZKMenu())
    Open App.Path & "\DTMMenu.ini" For Input As #FNum
        While Not EOF(FNum)
            lngCount = lngCount + 1
            ReDim Preserve ZKMenu(0 To lngMax + lngCount)
            Line Input #FNum, strLine 'ZKMenu(lngMax + lngCount)
            Set ZKMenu(lngMax + lngCount) = New clsZenDictionary
            Call ZKMenu(lngMax + lngCount).FromProp(strLine)
        Wend
    Close #FNum

End Sub

Public Sub CloseApp()
Dim k As Long

On Error Resume Next

    Rem - Save the X and x2 of the form in "ProgInfo.ini"
    If Not booKill Then ' Make sure it is only fired once...
        booKill = True

        Dim pInfo As New clsZenDictionary
        Call pInfo.FromINI(settings("SavePath") & "\ProgInfo.ini")
        pInfo("X1") = CStr(CLng(MainForm.left))
        pInfo("Y1") = CStr(CLng(MainForm.Top))
    
        Rem - OKay, time to die
        Call cMenu.UnHook
        If booInTray Then Call Systray_Del(MainForm)
        If Hotkeys.booLoaded Then Call Hotkeys.Unload
        Set Registry = Nothing
        
        Rem - Also save the active icons
        Call Icon_RemoveAll(pInfo)
        Call pInfo.ToINI(settings("SavePath") & "\ProgInfo.ini")

        Rem - Make the forms that are offscreen visible if they are not...
        If settings("IDT_VisOnExit") = "Y" Then Call ZK_GetObject("IDT").DoAction(zenDic("Action", "MakeAllVisible"))
                
        Call Registry.SetRegistry(HKCU, "SOFTWARE\ZenCODE\ZenKEY", "WindowHandle", vbNullString)
        Call ROL_RestoreAll
        If Not (AWT Is Nothing) Then Call MainForm.AWT.AWT_Flush
        
        #If LOGMODE > 0 Then
            Call LOG_Write("ZenKEY closed - " & Format(Now, "Long date") & ", " & Format(Now, "Long Time") & vbCr)
            Call LOG_Write("=================================================")
            Call LOG_Close
        #End If
        For k = Forms.Count - 1 To 0 Step -1
            Unload Forms(k)
        Next k
        
    End If

End Sub


Private Sub p_ToggleMenu(ByVal Index As Long, ByVal booTick As Boolean)
Dim strCap As String
Dim strState As String

    strCap = ZKMenu(Index)("Caption")
    Call cMenu.SetItemProp(ZKMenu(Index)("MenuHandle"), IIf(booTick, "CHECK", "UNCHECK"), strCap, ZKMenu(Index)("MenuID"))

End Sub


VERSION 5.00
Begin VB.Form frmDesktopMap 
   Appearance      =   0  'Flat
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "ZenKEY Desktop Map"
   ClientHeight    =   3015
   ClientLeft      =   4830
   ClientTop       =   4215
   ClientWidth     =   4305
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmDesktopMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   287
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picDTM 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.Image imiZK 
         Height          =   180
         Index           =   0
         Left            =   960
         Picture         =   "frmDesktopMap.frx":058A
         ToolTipText     =   "Hello"
         Top             =   600
         Width           =   180
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   0
         TabIndex        =   4
         ToolTipText     =   "Move the Desktop left"
         Top             =   0
         Width           =   210
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ù"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Move the Desktop up"
         Top             =   0
         Width           =   210
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ú"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   7
         Left            =   480
         TabIndex        =   2
         ToolTipText     =   "Move the Desktop down"
         Top             =   0
         Width           =   210
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ø"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   8
         Left            =   720
         TabIndex        =   1
         ToolTipText     =   "Move the Desktop right"
         Top             =   0
         Width           =   210
      End
      Begin VB.Shape shpSel 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   2  'Dash
         Height          =   1035
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Shape shpApp 
         BorderColor     =   &H000080FF&
         Height          =   1035
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   2400
   End
End
Attribute VB_Name = "frmDesktopMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim sngStartX As Single, sngStartY As Single
Private Const XBorder = 0.025 ' Border for placing the form as a fraction of screen width
Private Const YBorder = 0.035 '0.085 ' Border for placing the form as a fraction of screen height
Dim DTM_Position As Long
Dim COL_OnScreen As OLE_COLOR
Dim COL_OffScreen As OLE_COLOR
Dim COL_Active As OLE_COLOR
Dim COL_Desktop As OLE_COLOR
Dim COL_Selected As OLE_COLOR
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Dim DTM_X As Single, DTM_Y As Single
Dim booLoaded As Boolean
Dim lngDragIndex As Long
Dim lngPrevDragIndex As Long
Dim booDragOn As Boolean
Dim booControl As Boolean
Dim booMouseDown As Boolean
Rem - DTMMenu
Public MenuHandles As New clsZenDictionary
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Private zPrevTrans As clsZenDictionary
Rem - For manual form draging
Dim lngXGap As Long, lngYGap As Long
Dim ptCursor As POINTAPI
Rem - For delayed DrawWindows after closing
'Dim booWinampWarn As Boolean
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Const booShowDrag = True
Private HIcon As Long
Private cProcInfo As clsProcInfo
Public Sub DoAction(ByRef prop As clsZenDictionary)

    If Not DTM_Enabled Then
        Call ZenMB("Sorry, but the 'Desktop Map' must be enabled in order to perform this function.")
    Else
        Dim strAction As String
        strAction = prop("Action")
        Select Case strAction
            Case "TRANS"
                Rem - Set ZenKEY transparency - ZenIcons cannot be set with Parent = DesktopWindow
                Call SetTransIcon(prop)
                SET_Trans = Val(prop("Level"))
                Call ZK_Win.SetTrans(MainForm.hwnd, SET_Trans)
                Call ZK_Win.SetTrans(Me.hwnd, SET_Trans)
            Case "SHOWDTM"
                Rem - Show the desktop map. It will be dredrawn at the end of the sub anyway
                Me.Visible = True
            Case "CONFIGMAP"
                Rem - DTMMenu
                Call ShellExe(App.Path & "\ZKConfig.exe", "MAP")
            Case "HELP"
                Call ShellExe(App.Path & "\Help\DTM.htm")
            Case "WINACTION"
                Call WinAction(zenDic("Action", prop("WINACT")))
            Case "WINFOCUS"
                If lngDragIndex < 0 Then
                    Call ZenMB("Sorry, but you have not clicked on an application window.")
                Else
                    Call SetWinPos(WIN_RecList(lngDragIndex).hwnd, HWND_TOP, True)
                    tmrRefresh.Enabled = True
                End If
            Case "WINBOTTOM"
                Rem - We need to to extra precautions here to make sure the window doe not pop-up to the top again.
                Call WinAction(zenDic("Action", "BOTTOM"))
            Case "WINTRANS"
                Rem - Prevent the activation from undoing the transparency
                Call WinAction(zenDic("Action", "SETTRANSPARENCY=" & prop("Level")))
            Case "ATW_EXCLADD", "ATW_EXCLREM", "IDT_EXCLADD", "IDT_EXCLREM"
                '|Caption=Exclude from Auto-window transparency|
                '|Caption=Allow Auto-window transparency'|"
                '|Caption=Exclude from Desktop movements|
                '|Caption=Allow Desktop movements|
                Dim booAdd As Boolean
                Dim booAWT As Boolean
                booAdd = CBool(strAction = "ATW_EXCLADD" Or strAction = "IDT_EXCLADD")
                booAWT = CBool(strAction = "ATW_EXCLADD" Or strAction = "ATW_EXCLREM")
                Select Case True
                    Case (lngDragIndex < 0) And (ZK_Win.colSelected.Count < 1)
                        Rem - No windows selected
                        Call ZenMB("Please click within a window's borders to perform actions upon it.")
                    Case (MainForm.AWT Is Nothing) And booAWT
                        Call ZenMB("Sorry, but 'Auto-window transparency' has not been enabled.")
                    Case (ZK_Win.colSelected.Count > 0)
                        Rem - Add all the windows
                        Dim lngHWnd As Long, k As Long
                        For k = ZK_Win.colSelected.Count To 1 Step -1
                            lngHWnd = ZK_Win.colSelected.item(k)
                            If booAWT Then
                                Call MainForm.AWT.AWT_OmitList(lngHWnd, booAdd)
                            Else
                                If IDT_Enabled Then
                                    Call ZK_GetObject("IDT").OmitList(lngHWnd, booAdd)
                                Else
                                    Call ZenMB("Sorry, but the 'Infinite desktop' needs to be enabled to use this feature.")
                                End If
                            End If
                        Next k
                    Case Else
                        Rem - Add only the selected windows
                        If booAWT Then
                            Call MainForm.AWT.AWT_OmitList(WIN_RecList(lngDragIndex).hwnd, booAdd)
                        ElseIf IDT_Enabled Then
                            Call ZK_IDT.OmitList(WIN_RecList(lngDragIndex).hwnd, booAdd)
                        Else
                            Call ZenMB("The 'Infinite Desktop' must be enabled to use this feature.")
                        End If
                End Select
            Case "CENTERONCLICK"
                Call CenterOnClick
            Case "CENTERONCLICKFOCUS"
                If IDT_Enabled Then
                    Call CenterOnClick
                    Call DoAction(zenDic("Action", "WINFOCUS"))
                End If
            Case Else
                Call ZenMB("Function not implemented - " & prop("Action"))
        End Select
    End If
End Sub
Private Sub CenterOnClick()
Dim sngX As Single
Dim sngY As Single
Dim lngPrev As Long
    
    sngX = 0.5 * ScaleX(Screen.Width, vbTwips, vbPixels) - DTM_X
    sngY = 0.5 * ScaleY(Screen.Height, vbTwips, vbPixels) - DTM_Y
    Call MoveAllWindows(sngX, sngY, True)
    lngPrev = lngDragIndex
    If DTM_Enabled Then Call ZK_DTM.DrawWindows
    lngDragIndex = lngPrev

End Sub


Public Sub DrawWindows()
Dim k As Long
Dim sngAspect As Single
    
    If booMouseDown Then Exit Sub ' Prevent cascading events

    Rem - Establish the maximum and minimum boundaries, and the dimensions
    Rem - Reset the history and arrays
    WIN_RecCurrent = 0
    WIN_RecMax = 0

    Call EnumWindows(AddressOf RecordAll, ByVal 0&)
      
    If WIN_Changed Then
        Rem - Now draw  the windows
        WIN_Changed = False
        Call LockWindowUpdate(Me.hwnd)
        
        Rem - Firstly, calculate the maximum bounds of the area
        Dim WinMaxX As Long, WinMaxY As Long
        Dim WinMinX As Long, WinMinY As Long
        For k = 0 To WIN_RecMax - 1
            With WIN_RecList(k).TheRect
                If WinMinX > .left Then WinMinX = .left
                If WinMaxX < .Right Then WinMaxX = .Right
                If WinMinY > .Top Then WinMinY = .Top
                If WinMaxY < .Bottom Then WinMaxY = .Bottom
            End With
        Next k
        
        Rem -----------------------------------------------------------------------------------------------------------------------------------------------------
        Rem - Placement and scaling of map items
        Rem -----------------------------------------------------------------------------------------------------------------------------------------------------
        Dim sngRatio As Single
        Dim sngWidth As Single, sngHeight As Single
        
        sngWidth = WinMaxX - WinMinX
        sngHeight = WinMaxY - WinMinY
        sngAspect = Me.Width / Me.Height ' Aspect ratio on the map / SAspect is aspect ratio of screen
        sngRatio = sngWidth / sngHeight
        If sngAspect > sngRatio Then ' Ajust the aspect ratio so that it maintains the screnen aspect ratio
            Rem - Wider than it is higher
            sngRatio = sngWidth * sngAspect / sngRatio ' New width
            WinMinX = WinMinX + 0.5 * (sngWidth - sngRatio)
            sngWidth = sngRatio
            sngRatio = lblHeader(5).Width / picDTM.ScaleWidth
        Else
            Rem - Higher than it is wider
            sngRatio = sngHeight * sngRatio / sngAspect ' New height
            WinMinY = WinMinY + 0.5 * (sngHeight - sngRatio)
            sngHeight = sngRatio
            sngRatio = lblHeader(5).Height / picDTM.ScaleHeight
        End If
        With picDTM
            .ScaleTop = WinMinY - sngRatio * sngHeight
            .ScaleHeight = sngHeight * (1 + 2 * sngRatio)
            .ScaleLeft = WinMinX - sngRatio * sngWidth
            .ScaleWidth = sngWidth * (1 + 2 * sngRatio)
        End With
        Dim lngShapeMax As Long
        lngShapeMax = shpApp().UBound
        For k = 0 To WIN_RecMax - 1
            If k > lngShapeMax Then Load shpApp(k)
                    
            Rem - Place the window rectangle
            With WIN_RecList(k).TheRect
                shpApp(k).Move .left, .Top, .Right - .left, .Bottom - .Top
                shpApp(k).Visible = True
            End With
        Next k
                
        Call ColourRects
        Rem - Now hide the rects that we dont need
        For k = shpApp().UBound To WIN_RecMax Step -1
            shpApp(k).Visible = False
        Next k
        Call LockWindowUpdate(0)
    End If ' WIN_Changed
            
End Sub



Private Sub Form_Load()
Dim strTemp As String

    booLoaded = False
    COL_OnScreen = Val(settings("COL_OnScreen"))
    COL_OffScreen = Val(settings("COL_OffScreen"))
    COL_Active = Val(settings("COL_Active"))
    COL_Desktop = Val(settings("COL_Desktop"))
    strTemp = settings("COL_Selected")
    If Len(strTemp) > 0 Then
        COL_Selected = Val(strTemp)
    Else
        COL_Selected = 33023 ' Orange'  '16777088 ' Light cyan ''16744576 ' Light blue
    End If
    
    Dim k As Long, clrCol As OLE_COLOR
    'clrCol = RGB(45, 149, 244)
    clrCol = RGB(85, 189, 255)
    For k = 5 To 8
        lblHeader(k).BackColor = clrCol
    Next k
    For k = 1 To 3
        Load imiZK(k)
        imiZK(k).Visible = True
    Next k
    
    Me.BackColor = Val(settings("COL_Back"))
    picDTM.BackColor = Me.BackColor
    DTM_Position = Val(settings("DTM_Position"))
    Me.DrawStyle = vbDot

    Rem - Determine the position
    Dim sngWidth As Single, sngHeight As Single
    sngWidth = 0.2 * Screen.Width
    sngHeight = 0.2 * Screen.Height
    With Me
        Select Case DTM_Position
            Case 0 ' "BottomLeft"
                .Move Screen.Width * XBorder, Screen.Height * (1 - YBorder) - .Height, sngWidth, sngHeight
            Case 3 ' '"TopRight"
                .Move Screen.Width * (1 - XBorder) - .Width, Screen.Height * YBorder, sngWidth, sngHeight
            Case 2 ' '"TopLeft"
                .Move Screen.Width * XBorder, Screen.Height * YBorder, sngWidth, sngHeight
            Case 1 ' ' BottomRight
                .Move Screen.Width * (1 - XBorder) - Width, Screen.Height * (1 - YBorder) - .Height, sngWidth, sngHeight
            Case 4 ' '"Center"
                .Move 0.5 * (Screen.Width - .Width), 0.5 * (Screen.Height - .Height), sngWidth, sngHeight
            Case 5 '- Wherever I place it
                Dim dtmPos As New clsZenDictionary
                If dtmPos.FromINI(settings("SavePath") & "\DTMPos.ini") Then
                    With dtmPos
                        Me.Move Val(dtmPos("DTM_Left")), Val(dtmPos("DTM_Top")), Val(dtmPos("DTM_Width")), Val(dtmPos("DTM_Height"))
                    End With
                Else
                    Me.Move Screen.Width * XBorder, Screen.Height * (1 - YBorder) - Me.Height, sngWidth, sngHeight
                End If
            'Case 6 - disabled
        End Select
    End With
    Me.ForeColor = COL_Active
    If SET_Trans > 0 Then Call ZK_Win.SetTrans(Me.hwnd, SET_Trans)
        
    Rem - Set the captions
    imiZK(0).ToolTipText = "Desktop actions"
    imiZK(1).ToolTipText = "Close map"
    imiZK(2).ToolTipText = "ZenKEY Menu"
    imiZK(3).ToolTipText = "Right-click menu"
    
    lngDragIndex = -1
    booLoaded = True
    If Not Me.Visible Then Call SetWinPos(Me.hwnd, SET_Layer, False)
    Set cProcInfo = New clsProcInfo
    If settings("DTM_Pin") = "Y" Then Call ZK_Win.PinToDesktop(True, Me.hwnd)

End Sub

Private Sub imiZK_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lblHeader_MouseDown(Index, Button, Shift, X, Y)
    
End Sub


Private Sub imiZK_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lblHeader_MouseMove(Index, Button, Shift, X, Y)
End Sub


Private Sub imiZK_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Call lblHeader_MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub picDTM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DTM_X = X
    DTM_Y = Y
        
    Rem - Set the lngDragIndex to the window that they have click on
    Call SetSelToRect(X, Y, CBool(Button = 1))
    booControl = CBool(Shift <> 0)
    booMouseDown = True
        
    Rem - If control is not down, de-select any previous selection and enable the window clicked upon
    If Not booControl Then
        If ZK_Win.colSelected.Count > 0 Then
            If lngDragIndex > -1 Then
                Rem - They have clicked in a rectangle
                If Not Selected_Contains(lngDragIndex) Then
                    Rem - Set the rectangle they have just clicked in as the active
                    WIN_Active = WIN_RecList(lngDragIndex).hwnd
                    Call ZK_Win.Selected("Clear", 0)
                    Call ColourRects
                End If
            Else
                Call ZK_Win.Selected("Clear", 0)
                Call ColourRects
            End If
        End If
    End If
    If Button = 2 Then
        Call ColourRects
    Else
        If Not booControl Then
            If lngDragIndex > -1 Then
                If lngPrevDragIndex > -1 Then Call ColourRect(lngPrevDragIndex, False)
                Call ColourRect(lngDragIndex, True)
            End If
        End If
    End If
    MainForm.tmrHook.Enabled = False

End Sub

Private Sub picDTM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Long

    If Button = 2 Then Exit Sub
    If Not booDragOn Then
        
        Rem - If they are not dragging
        If booMouseDown Then
            Rem - If the mouse is down
            If (Abs(DTM_X - X) > 3) Or (Abs(DTM_Y - Y) > 3) Then
                Rem - If they have moved over 3 pixels
                If (lngDragIndex > -1) Then
                    Rem - They have clicked in a window, so let them drag
                    booDragOn = True
                    Screen.MousePointer = vbSizePointer
                    If booControl Then
                        Call ZK_Win.Selected("Add", WIN_RecList(lngDragIndex).hwnd)
                        Call ColourRects
                    End If
                Else
                    Call SelRect_Pos(X, Y, Shift)
                    shpSel.Visible = True
                End If
            End If
        Else
            If Me.MousePointer <> vbDefault Then Me.MousePointer = vbDefault
        End If
    End If
    
    If booDragOn Then
        If lngDragIndex > -1 Then
            Dim sngX As Single, sngY As Single
            sngX = X - sngStartX
            sngY = Y - sngStartY
            sngStartX = X
            sngStartY = Y
            
            If ZK_Win.colSelected.Count > 0 Then
                Rem - For multiple windows, placing of the shape is done in mouseup....
                Dim rctRect As RECT
                Dim lngIndex As Long
                
                For k = 0 To WIN_RecMax - 1
                    lngIndex = GetSelIndex(WIN_RecList(k).hwnd)
                    If lngIndex > -1 Then shpApp(k).Move shpApp(k).left + sngX, shpApp(k).Top + sngY
                Next k
            Else
                shpApp(lngDragIndex).Move shpApp(lngDragIndex).left + sngX, shpApp(lngDragIndex).Top + sngY
                Call PlaceToShape(lngDragIndex)
            End If
        End If
    ElseIf shpSel.Visible Then
        Call SelRect_Pos(X, Y, Shift)
    End If
    
End Sub


Private Sub picDTM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim strAct As String
Dim k As Long

    Select Case Button
        Case 2
            Rem =====================================================================================
            Rem - Right mouse button
            Rem =====================================================================================
            If ZK_Win.colSelected.Count < 1 Then
                Rem - Re-colour the rectangles to indicate activity.
                If lngDragIndex > -1 Then Call ColourRect(lngDragIndex, True)
            End If
            
            Select Case Shift
                Case 0 ' Nothing pressed
                    strAct = settings("DTM_RClick")
                    If (strAct = "") Or (strAct = "Full") Then strAct = "RClick"
                Case 1 ' Shift
                    strAct = settings("DTM_SRClick")
                    If (strAct = "Full") Or (Len(strAct) = 0) Then strAct = "WinMove"
                Case 2 ' Control
                    strAct = settings("DTM_CRClick")
                    If Len(strAct) = 0 Then strAct = "WinAlter"
            End Select
            Rem - Show the appropriate menu
            If Len(strAct) > 0 Then Call MainForm.ShowMenu(MenuHandles(strAct))
            Screen.MousePointer = vbDefault
        Case 1
            Rem =====================================================================================
            Rem - Left mouse button
            Rem =====================================================================================
            If booDragOn Then
                Rem - They have dragged.  Place the window(s)
                Screen.MousePointer = vbDefault
                If ZK_Win.colSelected.Count > 0 Then
                    For k = 0 To WIN_RecMax - 1
                        If GetSelIndex(WIN_RecList(k).hwnd) > -1 Then Call PlaceToShape(k)
                    Next k
                End If
                
            ElseIf shpSel.Visible Then
                Call SelRect_Set(Shift)
                shpSel.Visible = False
            ElseIf booControl Then
                Rem - They have not dragged.
                If lngDragIndex > -1 Then
                    If lngPrevDragIndex <> lngDragIndex Then
                        If lngPrevDragIndex > -1 Then
                            If ZK_Win.colSelected.Count < 1 Then
                                Rem - Only add the previosly selected window if they are starting a selection.
                                If Not IsZenWindow(lngPrevDragIndex) Then Call ZK_Win.Selected("Add", WIN_RecList(lngPrevDragIndex).hwnd)
                            End If
                        End If
                    End If
                    Call ZK_Win.Selected("Toggle", WIN_RecList(lngDragIndex).hwnd)
                    Call ColourRects
                End If
            Else
                Rem - No dragging. They have just clicked. Activate the window.
                Call ZK_Win.Selected("Clear", 0)
                If lngDragIndex > -1 Then Call SetForegroundWindow(WIN_RecList(lngDragIndex).hwnd)
                Call ColourRects
            End If
            
    End Select
    booDragOn = False
    booMouseDown = False
    MainForm.tmrHook.Enabled = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        Rem - Stop it unloading and just hide it...
        Cancel = 1
        Me.Hide
    Else
        Dim dtmPos As New clsZenDictionary
        dtmPos("DTM_Width") = CStr(CLng(Me.Width))
        dtmPos("DTM_Height") = CStr(CLng(Me.Height))
        dtmPos("DTM_Left") = CStr(CLng(Me.left))
        dtmPos("DTM_Top") = CStr(CLng(Me.Top))
        Call dtmPos.ToINI(settings("SavePath") & "\DTMPos.ini")
        If HIcon <> 0 Then Call DestroyIcon(HIcon)
    End If
    
End Sub

Private Sub Form_Resize()
Const MinWidth = 120
Const MinHeight = 140
Const SM_CYCAPTION = 4 'Height of windows caption

    If Me.ScaleWidth < MinWidth Then Me.Width = Me.ScaleX(MinWidth, vbPixels, vbTwips)
    If Me.ScaleHeight + GetSystemMetrics(SM_CYCAPTION) < MinHeight Then Me.Height = Me.ScaleY(MinHeight, vbPixels, vbTwips)
    picDTM.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - ICN_Height - 2
    
    With picDTM
        
        ' The following are now contained by picDTM
        imiZK(0).Move .ScaleLeft, .ScaleTop ' Top left
        imiZK(1).Move .ScaleLeft + .ScaleWidth - imiZK(2).Width, .ScaleTop ' Top right
        imiZK(2).Move .ScaleLeft, .ScaleTop + .ScaleHeight - imiZK(2).Height ' Bottom left
        imiZK(3).Move imiZK(1).left, imiZK(2).Top ' Right bottom
           
        Rem - Now place the header + the desktop map movement buttons
        'lblHeader(4).Move .ScaleLeft + 0.5 * (.ScaleWidth - lblHeader(4).Width), .ScaleTop ' Header
        lblHeader(5).Move .ScaleLeft + 0.5 * (.ScaleWidth - 4 * lblHeader(5).Width), imiZK(2).Top
        lblHeader(6).Move lblHeader(5).left + lblHeader(5).Width, lblHeader(5).Top
        lblHeader(7).Move lblHeader(6).left + lblHeader(5).Width, lblHeader(5).Top
        lblHeader(8).Move lblHeader(7).left + lblHeader(5).Width, lblHeader(5).Top
        
    End With
    If booLoaded Then
        Call DrawWindows
        Call DrawActiveIcon(True)
    End If
    If SET_ZenBar Then Call Icon_Layout

End Sub












Private Sub lblHeader_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Index
        Case 1
            Rem - Top left
            Call MainForm.ShowMenu(MenuHandles("Main"))
        Case 0, 2
            Call MainForm.DoAction(zenDic("Action", "SHOWMENUMAIN"))
        Case 3
            Call MainForm.DoAction(zenDic("Action", "SHOWMENURCLICK"))
        Case 5 ' Desktop left
            Call ZK_GetObject("IDT").DoAction(zenDic("Action", "DTHALFRIGHT"))
        Case 6 ' Desktop up
            Call ZK_GetObject("IDT").DoAction(zenDic("Action", "DTHALFDOWN"))
        Case 7 ' Desktop down
            Call ZK_GetObject("IDT").DoAction(zenDic("Action", "DTHALFUP"))
        Case 8 ' Desktop right
            Call ZK_GetObject("IDT").DoAction(zenDic("Action", "DTHALFLEFT"))
    End Select
    
End Sub


Private Sub lblHeader_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case Index
        Case 0
            Me.MousePointer = vbArrowQuestion  ' Desktop actions
        
    End Select
    
End Sub


Private Sub tmrRefresh_Timer()
    
    tmrRefresh.Enabled = False
    Call DrawWindows
    
    
End Sub





Private Sub SetSelToRect(ByVal X As Single, ByVal Y As Single, ByVal LeftClick As Boolean)

    lngPrevDragIndex = lngDragIndex
    lngDragIndex = IsInRect(X, Y)
    If lngDragIndex > -1 Then
        sngStartX = X
        sngStartY = Y
    End If
    
End Sub


Private Sub WinAction(ByRef prop As clsZenDictionary)

    Select Case True
        Case prop("Action") = "UNDO"
            Call ZK_Win.DoAction(prop)
        Case (lngDragIndex < 0) And (ZK_Win.colSelected.Count < 1)
            Call ZenMB("Please click within a window's borders to perform actions upon it.")
        Case Else
            If ZK_Win.colSelected.Count < 1 Then prop("HWnd") = CStr(WIN_RecList(lngDragIndex).hwnd)
            Call ZK_Win.DoAction(prop)
            tmrRefresh.Enabled = True
    End Select
    
End Sub

Private Sub ShowMenu(ByRef TheMenu As Menu)

    Rem - A mechanism to prevent the desktop map timer changing the z-order while is is visible
    MainForm.tmrHook.Enabled = False
    tmrRefresh.Enabled = False
    
    Call PopupMenu(TheMenu)
    
    tmrRefresh.Enabled = True
    MainForm.tmrHook.Enabled = True

End Sub

Private Sub ColourRects()
Dim k As Long, i As Long
Dim booColoured As Boolean
Dim lngMax As Long

    lngMax = ZK_Win.colSelected.Count
    For k = 0 To WIN_RecMax - 1
        Rem - Check if they are selected
        booColoured = False
        For i = lngMax To 1 Step -1
            If ZK_Win.colSelected.item(i) = WIN_RecList(k).hwnd Then
                Call ColourRect(k, True)
                booColoured = True
            End If
        Next i
        If Not booColoured Then Call ColourRect(k, False)
    Next k
    
End Sub

Private Function IsInRect(ByVal X As Single, ByVal Y As Single) As Long
Rem - Returns the rectangle index (WIN_RecList()) the X and Y fall within
Rem - Returns - 1 if the X and Y fall are outside of any rect
Dim k As Long

    IsInRect = -1
    For k = WIN_RecMax - 1 To 0 Step -1
        Rem - We step backwards as this way, we seem to get the topmost windows first
        With WIN_RecList(k).TheRect
            If X > .left Then
                If X < .Right Then
                    If Y > .Top Then
                        If Y < .Bottom Then
                            If WindowIsUsable(WIN_RecList(k).hwnd) Then
                                IsInRect = k
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next k

End Function


Private Sub ColourRect(ByVal k As Long, ByVal booSelected As Boolean)

    If k < 0 Then Exit Sub
    If booSelected Then
        shpApp(k).BorderColor = COL_Selected
    ElseIf WIN_RecList(k).hwnd = WIN_Active Then
        shpApp(k).BorderColor = COL_Active
    Else
        Select Case WIN_RecList(k).Status
            Case Desktop
                shpApp(k).BorderColor = COL_Desktop '&HFF8080
            Case OffScreen
                shpApp(k).BorderColor = COL_OffScreen
            Case Else ' Normal
                shpApp(k).BorderColor = COL_OnScreen
        End Select
    End If
    shpApp(k).Visible = True

End Sub

Private Function GetSelIndex(ByVal hwnd As Long) As Long
Dim k As Long

    GetSelIndex = -1
    For k = ZK_Win.colSelected.Count To 1 Step -1
        If hwnd = ZK_Win.colSelected.item(k) Then
            GetSelIndex = k
            Exit Function
        End If
    Next k
    
End Function

Private Sub SelRect_Pos(ByVal X As Single, ByVal Y As Single, ByVal Shift As Single)
Dim sngLeft As Single, sngTop As Single
Dim k As Long
        
    If X > DTM_X Then sngLeft = DTM_X Else sngLeft = X
    If Y > DTM_Y Then sngTop = DTM_Y Else sngTop = Y
    shpSel.Move sngLeft, sngTop, Abs(X - DTM_X), Abs(Y - DTM_Y)

    Rem - Shift = 0 -  No key down - Normal select (erase previous)
    Rem - Shift = vbShiftMask - Shift down - Add to selection (no removing)
    Rem - Shift = vbCtrlMask - Control Down - Reverse selection (De-select selected, select unselected)
    Dim booIsInRect As Boolean
    If Shift = 0 Then Call ZK_Win.Selected("Clear", 0)
    
    For k = 0 To WIN_RecMax - 1
        booIsInRect = RectIsInRect(k, sngLeft, sngTop, Abs(X - DTM_X), Abs(Y - DTM_Y))
        Select Case Shift
            Case 0
                Rem - Normal select (erase previous)
                Call ColourRect(k, booIsInRect)
            Case vbShiftMask
                Rem - Add to selection (no removing)
                If Selected_Contains(k) Then
                    Call ColourRect(k, True)
                Else
                    Call ColourRect(k, booIsInRect)
                End If
            Case vbCtrlMask
                Rem - Reverse selection (De-select selected, select unselected)
                If Selected_Contains(k) Then
                    Call ColourRect(k, Not booIsInRect)
                Else
                    Call ColourRect(k, booIsInRect)
                End If
        End Select
    Next k


End Sub

Private Sub SelRect_Set(ByVal Shift As Long)
Dim k As Long, booIsInRect As Boolean
Dim sngLeft As Single, sngRight As Single
Dim sngTop As Single, sngBottom As Single

    With shpSel
        sngLeft = .left
        sngRight = .left + .Width
        sngTop = .Top
        sngBottom = .Top + .Height
    End With
    
    Rem - Shift = 0 -  No key down - Normal select (erase previous)
    Rem - Shift = vbShiftMask - Shift down - Add to selection (no removing)
    Rem - Shift = vbCtrlMask - Control Down - Reverse selection (De-select selected, select unselected)
    For k = 0 To WIN_RecMax - 1
        booIsInRect = RectIsInRect(k, sngLeft, sngTop, sngRight - sngLeft, sngBottom - sngTop)
        If booIsInRect Then
            Select Case Shift
                Case 0
                    Rem - Normal select (erase previous)
                    Call ZK_Win.Selected("Add", WIN_RecList(k).hwnd)
                Case vbShiftMask
                    Rem - Add to selection (no removing)
                    If Not Selected_Contains(k) Then Call ZK_Win.Selected("Add", WIN_RecList(k).hwnd)
                Case vbCtrlMask
                    Rem - Reverse selection (De-select selected, select unselected)
                    If Selected_Contains(k) Then
                        Call ZK_Win.Selected("Toggle", WIN_RecList(k).hwnd)
                    Else
                        Call ZK_Win.Selected("Add", WIN_RecList(k).hwnd)
                    End If
            End Select
        End If
    Next k
    Call ColourRects

End Sub

Private Function RectIsInRect(ByVal RecIndex As Long, ByVal left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single) As Boolean
Rem - Return TRUE if the RecList(RecIndex) rectanlge lies within the specified rectangular area
    
    With WIN_RecList(RecIndex).TheRect
        If .left > left Then
            If .Right < left + Width Then
                If .Top > Top Then
                    If .Bottom < Top + Height Then
                        RectIsInRect = WindowIsUsable(WIN_RecList(RecIndex).hwnd)
                    End If
                End If
            End If
        End If
    End With

End Function

Private Function Selected_Contains(ByVal Index As Long) As Boolean
Dim k As Long

    For k = ZK_Win.colSelected.Count To 1 Step -1
        If ZK_Win.colSelected.item(k) = WIN_RecList(Index).hwnd Then
            Selected_Contains = True
            Exit Function
        End If
    Next k

End Function





Private Function IsZenWindow(ByVal Index As Long) As Boolean
'    #If IDE = 1 Then
'        IsZenWindow = CBool(Right(GetExeFromHandle(RecList(Index).hwnd), 7) = "Vb6.exe")
'    #Else
'        IsZenWindow = CBool(Right(GetExeFromHandle(RecList(Index).hwnd), 10) = "ZenKEY.exe")
'    #End If

    Select Case WIN_RecList(Index).hwnd
        Case MainForm.hwnd, Me.hwnd: IsZenWindow = True
        Case Else: IsZenWindow = False
    End Select
    
    
End Function

Private Sub PlaceToShape(ByVal RecIndex As Long)
Dim zdAct As clsZenDictionary

    With shpApp(RecIndex)
        Set zdAct = zenDic("Left", Int(.left), "Top", Int(.Top), "Right", Int(.left + .Width), "Bottom", Int(.Top + .Height), _
                "Action", "PlaceRect", "HWnd", WIN_RecList(RecIndex).hwnd)
        Call ZK_Win.DoAction(zdAct)
    End With

End Sub
        

Public Sub SetTransIcon(ByRef prop As clsZenDictionary)
    
    If Not zPrevTrans Is Nothing Then
        Rem - Remove tick from last menu
        Call cMenu.SetItemProp(zPrevTrans("MenuHandle"), "UNCHECK", zPrevTrans("Caption"), zPrevTrans("MenuID"))
    End If
    Call cMenu.SetItemProp(prop("MenuHandle"), "CHECK", prop("Caption"), prop("MenuID"))
    Set zPrevTrans = prop.Copy

End Sub

Public Sub DrawActiveIcon(Optional ByVal Refresh As Boolean = False)
Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = DI_MASK Or DI_IMAGE
Static lngLast As Long
Static strExe As String
Dim sngTop As Single
    
    If Refresh Then lngLast = 0
    If WIN_Active <> 0 Then
        sngTop = picDTM.Height + 4
        If WIN_Active <> lngLast Then
            ' Redraw the icon if it's changed
            Me.Cls
            Me.Line (1, sngTop)-Step(Me.ScaleWidth - 2, ICN_Height - 4), , B
            If HIcon <> 0 Then Call DestroyIcon(HIcon)
            strExe = GetExeFromHandle(WIN_Active)
            HIcon = ExtractIcon(Me.hwnd, strExe, 0)
            Call DrawIconEx(Me.hdc, 1 + ICN_HGap, sngTop + ICN_VGap, HIcon, 32, 32, 0, 0, DI_NORMAL)
            lngLast = WIN_Active
        End If
            
        'Erase previous text
        Me.Line (ICN_TLeft + 2 * ICN_HGap - 2, ICN_VGap + sngTop + 2)-Step(Me.ScaleWidth - (ICN_TLeft - 2 * ICN_HGap) - 12, ICN_Height - 12), Me.BackColor, BF
        Call cProcInfo.Init(lngLast)
        Me.CurrentY = ICN_VGap + sngTop
        Me.CurrentX = ICN_TLeft + 2 * ICN_HGap
        Me.Print GetFileName(strExe)
        Me.CurrentY = ICN_VGap + ICN_LineGap + sngTop
        Me.CurrentX = ICN_TLeft + 2 * ICN_HGap
        Me.Print cProcInfo.CPUUsage & "% cpu, " & cProcInfo.RAM
    End If ' WIN_Active
    
End Sub

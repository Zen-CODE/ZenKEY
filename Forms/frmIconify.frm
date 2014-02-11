VERSION 5.00
Begin VB.Form frmIconify 
   BorderStyle     =   0  'None
   ClientHeight    =   1170
   ClientLeft      =   3420
   ClientTop       =   1845
   ClientWidth     =   2880
   ControlBox      =   0   'False
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
   Icon            =   "frmIconify.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   78
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   192
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuIconize 
         Caption         =   "Iconize window"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
         Begin VB.Menu mnuPreserve 
            Caption         =   "Preserve icon"
         End
         Begin VB.Menu mnuModeMinimize 
            Caption         =   "Minimize window (vs. Hide)"
         End
      End
      Begin VB.Menu mnuAll 
         Caption         =   "All icons"
         Begin VB.Menu mnuColor 
            Caption         =   "Forecolor"
         End
         Begin VB.Menu mnuZenBar 
            Caption         =   "Place in ZenBar positions"
         End
         Begin VB.Menu mnuFlushDead 
            Caption         =   "Flush inactive icons"
         End
         Begin VB.Menu mnuReAcquire 
            Caption         =   "Reacquire process information"
         End
      End
      Begin VB.Menu mnuCloseMain 
         Caption         =   "Close"
         Begin VB.Menu mnuClose 
            Caption         =   "Close"
         End
         Begin VB.Menu mnuCloseClear 
            Caption         =   "Close and clear settings"
         End
         Begin VB.Menu mnuCloseKill 
            Caption         =   "Close icon and application(s)"
         End
      End
   End
End
Attribute VB_Name = "frmIconify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'Public prop As String
Public WinList As Collection
Public FileName As String
Public ExeName As String
Public booIsIconized As Boolean
Rem ----------------------------------------------------------
Dim booDrag As Boolean
Dim booLoading As Boolean
Dim sngStartX As Single, sngStartY As Single ' For dragging purposes
Dim ptCursor As POINTAPI, ptStart As POINTAPI
Const ZenMidGap = 20
Dim cProcInfo As New clsProcInfo
Const SW_MAXIMIZE = 3
Const SW_HIDE = 0
Const SW_SHOW = 1
Private prop As New clsZenDictionary
Private Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long
Dim strLine1 As String, strLine2 As String
Private Declare Function DestroyIcon Lib "user32" (ByVal HIcon As Long) As Long
Private HIcon As Long  'Handle to the extracted icon
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Type CHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Function UIntToDbl(ByVal Value As Long) As Double
    
    If ((Value And &H80000000) <> 0) Then
      UIntToDbl = Value And &H7FFFFFFF
      UIntToDbl = UIntToDbl + 2147483648#
    Else
        UIntToDbl = Value
    End If
      
End Function

Public Sub Iconify()
        
    Rem - Get the windows current state
    #If IDE = 1 Then
        If Right(FileName, 7) <> "Vb6.exe" Then
    #Else
        If Right(FileName, 10) <> "ZenKEY.exe" Then
    #End If
            Call p_AppWindows(False)
            booIsIconized = True
            Call SetFocusToLastActive(1)
        End If
        
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    booDrag = False
    sngStartX = X
    sngStartY = Y
    If Button = 2 Then Call PopupMenu(mnuMain) ' Right click
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const Tolerance = 5
    
    If Button = 1 Then
        Rem - Left mouse button.
        If booDrag Then
            Call GetCursorPos(ptCursor)
            Me.Move Me.ScaleX(ptCursor.X, vbPixels, vbTwips) - sngStartX, Me.ScaleY(ptCursor.Y, vbPixels, vbTwips) - sngStartY
        ElseIf sngStartX <> 0 Then
            booDrag = (Abs(sngStartX - X) > Tolerance) Or (Abs(sngStartY - Y) > Tolerance)
            If booDrag Then
                Call GetCursorPos(ptCursor)
                sngStartX = Me.ScaleX(ptCursor.X, vbPixels, vbTwips) - Me.left
                sngStartY = Me.ScaleY(ptCursor.Y, vbPixels, vbTwips) - Me.Top
                ptStart.X = ptCursor.X
                ptStart.Y = ptCursor.Y
            End If
        End If
    End If
    
End Sub




Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Not booDrag Then
        Rem - Normal click = 1
        If Button = 1 Then
            Call Icon_Reacquire
            If WinList.Count = 0 Then
                Call ShellExe(FileName)
            Else
                Call ShowWin
            End If
            If Not mnuPreserve.Checked Then Call IconClose(SET_ZenBar)
        End If
    Else
        Dim k As Long
        For k = ICN_Forms.Count To 1 Step -1
            Call ICN_Forms(k).DrawIcon
        Next k
        booDrag = False
    End If
    sngStartX = 0
    sngStartY = 0
    
End Sub





Public Sub UnIconify()
       
    Call p_AppWindows(True)
    booIsIconized = False
    If Not mnuPreserve.Checked Then Call IconClose(SET_ZenBar)
 
End Sub

Private Sub UpdateStats()
Dim lngWin As Long
            
    If WinList.Count > 0 Then
        Rem - Check the Window is valid, if not clear it
        lngWin = WinList.item(1)
        If IsWindow(lngWin) = 0 Then
            Call dhc_RemoveAll(WinList)
            lngWin = 0
        End If
    End If
    
    If lngWin = 0 Then
        strLine1 = ExeName
        strLine2 = " - - - -"
        Call dhc_RemoveAll(WinList)
    Else
        Rem - Treat as a normal icon.
        Dim strCap As String
        With cProcInfo
            Call .Init(lngWin)
            strLine2 = .CPUUsage & "% cpu, " & .RAM
        End With
        If booIsIconized Then
            strLine1 = ExeName & "**" & strCap
        Else
            strLine1 = ExeName & strCap
        End If
    End If ' lngWin = 0
        
End Sub

Private Sub Form_Paint()
        
    Call UpdateStats    ' Fill out proc info
    Call PaintDesktop(Me.hdc)
    Call DrawIcon
    Me.CurrentY = ICN_VGap
    Me.CurrentX = ICN_TLeft + 2 * ICN_HGap
    Me.Print strLine1
    Me.CurrentY = ICN_VGap + ICN_LineGap
    Me.CurrentX = ICN_TLeft + 2 * ICN_HGap
    Me.Print strLine2
    Me.Line (0, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1), , B

End Sub

Private Sub mnuClose_Click()

    Call IconClose(SET_ZenBar)
    
End Sub

Private Sub mnuCloseClear_Click()
    
    Rem - Clear the icon settings and then close the icon
    mnuPreserve.Checked = False
    Call IconClose(SET_ZenBar)

End Sub



Private Sub mnuCloseKill_Click()
Dim k As Long
        
    Call Icon_Reacquire
    If WinList.Count > 0 Then
        Const WM_CLOSE = &H10
        For k = 1 To WinList.Count
            Call PostMessage(WinList.item(k), WM_CLOSE, 0&, ByVal 0&)
        Next k
    End If
    Call mnuClose_Click

End Sub

Private Sub mnuFlushDead_Click()
    Call ZK_Win.DoAction(zenDic("Action", "ICONFLUSHDEAD"))
End Sub




Private Sub mnuColor_Click()
Dim cc As CHOOSECOLOR
Dim Custcolor(16) As Long
Dim lReturn As Long
Dim lngColor As Long

    'set the structure size
    cc.lStructSize = Len(cc)
    cc.hWndOwner = Me.hwnd
    cc.hInstance = App.hInstance
    cc.lpCustColors = StrConv("Choose the ZenIcon font color.", vbUnicode)
    'no extra flags
    cc.flags = 0

    'Show the 'Select Color'-dialog
    If CHOOSECOLOR(cc) <> 0 Then
        Dim k As Long
        lngColor = cc.rgbResult
        For k = ICN_Forms.Count To 1 Step -1
            ICN_Forms(k).ForeColor = lngColor
            ICN_Forms(k).Refresh
        Next k
        
    End If
    
End Sub

Private Sub mnuIconize_Click()

    If booIsIconized Then
        Call UnIconify
    ElseIf WinList.Count > 0 Then
        Rem - Check that the window handle is valid. If not, do nothing
        If IsWindow(WinList.item(1)) Then Call Iconify
    End If
    
End Sub

Private Sub mnuPreserve_Click()
    
    If Not booLoading Then mnuPreserve.Checked = Not mnuPreserve.Checked
    
End Sub


Private Sub mnuReAcquire_Click()
    Call Icon_Reacquire
End Sub


Private Sub mnuModeMinimize_Click()

    If Not booLoading Then mnuModeMinimize.Checked = Not mnuModeMinimize.Checked
    
End Sub

Private Sub mnuZenBar_Click()
    Call Icon_Layout
End Sub

Private Sub p_InitSize()
Dim sngHeight As Single

    sngHeight = Me.ScaleY(ICN_Height, vbPixels, vbTwips)
    If Not SET_ZenBar Then
        Dim sngLeft As Single, sngTop As Single
        Dim sngWidth As Single
        
        If Len(prop("Left")) > 0 Then
            Rem - If 'Preserve' is selected and we have a recorded position, please it there.
            sngLeft = Val(prop("Left"))
            sngTop = Val(prop("Top"))
            sngWidth = Val(prop("Width"))
        Else
            Rem - Otheriwse we place it centered around the mouse pointer.
            Dim ptMouse As POINTAPI
            Call GetCursorPos(ptMouse)
            sngLeft = Me.ScaleX(ptMouse.X, vbPixels, vbTwips)
            sngTop = Me.ScaleY(ptMouse.Y, vbPixels, vbTwips) - 0.5 * sngHeight
        End If
        With Me
            If sngWidth = 0 Then
                If .TextWidth(ExeName) < .TextWidth("100% CPU, 1000mb") Then
                    sngWidth = .ScaleX(34 + 2 * ICN_HGap + .TextWidth("100% CPU, 1000mb") + 10, vbPixels, vbTwips)
                Else
                    sngWidth = .ScaleX(34 + 2 * ICN_HGap + .TextWidth(ExeName) + 10, vbPixels, vbTwips)
                End If
            End If
            .Move sngLeft, sngTop, sngWidth, sngHeight
        End With
        
    End If

End Sub


Private Sub p_settings(ByVal Action As String)
Dim strFName As String

    On Error Resume Next
    strFName = settings("SavePath") & "\Icons\" & GetFileName(FileName) & ".ini"
    Select Case Action
        Case "Load"
            Rem - Load
            booLoading = True
            If prop.FromINI(strFName) Then
                Rem - Preserve state and information
                If prop("Preserve") = "Y" Then mnuPreserve.Checked = True
                If prop("Minimize") = "M" Then mnuModeMinimize.Checked = True
            End If
            booLoading = False
        Case "Save"
            Rem - Save
            If Len(Dir(settings("SavePath") & "\Icons", vbDirectory)) = 0 Then Call MkDir(settings("SavePath") & "\Icons")
            Dim zDic As New clsZenDictionary
            zDic("Left") = CStr(CLng(Me.left))
            zDic("Top") = CStr(CLng(Me.Top))
            zDic("Width") = CStr(CLng(Me.Width))
            If mnuModeMinimize.Checked Then zDic("Minimize") = "M"
            If mnuPreserve.Checked Then zDic("Preserve") = "Y"
            Call zDic.ToINI(strFName)
            
         Case "Clear"
            Rem - Clear the settings
            If Len(Dir(strFName)) > 0 Then Call Kill(strFName)
    End Select


End Sub

Private Sub p_AppWindows(ByVal booShow As Boolean)
Dim k As Long
    
    
    Rem - Prevent acting upon ZenKEY windows
    #If IDE = 1 Then
        If CBool(ExeName = "Vb6.exe") Then Exit Sub
    #End If
    
    Rem - Only set the window list on Iconization. The undo can then use this list.
    For k = WinList.Count To 1 Step -1
        Call p_MinAction(WinList.item(k), booShow)
    Next k

End Sub


Private Sub p_MinAction(ByVal hwnd As Long, ByVal booShow As Boolean)

    If IsWindow(hwnd) Then
        If booShow Then
            Call SetWinPos(hwnd, HWND_TOP, True)
        Else
            If mnuModeMinimize.Checked Then
                Const SW_MINIMIZE = 6
                Call ShowWindow(hwnd, SW_MINIMIZE)
            Else
                Call ShowWindow(hwnd, SW_HIDE)
            End If
        End If
    End If
    
End Sub


Public Sub ShowWin()
Rem - A call has been made to show the application. This could be from right clicking
Rem - the icons, or from 'Icon_FlushExe'.
Dim k As Long
        
    For k = 1 To WinList.Count
        Call SetWinPos(WinList.item(k), HWND_TOP, True)
    Next k
    booIsIconized = False

End Sub





Public Sub Init(ByRef FName As String)
' Accept the window handle or the filename string
                
    If WinList Is Nothing Then
        If settings("ICN_Pin") = "Y" Then Call ZK_Win.PinToDesktop(True, Me.hwnd)
        Set WinList = New Collection
    End If
    FileName = FName
    
    Rem - The form can be initialized via a window handle or a FIleName
    If ICN_Forms.Count > 0 Then Me.ForeColor = ICN_Forms(1).ForeColor
    Call ICN_Forms.Add(Me)
    Call p_settings("Load")
    ExeName = StrConv(GetFileName(FileName), vbProperCase)
    Call p_InitSize
    
    
End Sub

Public Sub DrawIcon()
Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = DI_MASK Or DI_IMAGE
    
    If HIcon = 0 Then HIcon = ExtractIcon(Me.hwnd, FileName, 0)
    Call DrawIconEx(Me.hdc, 1 + ICN_HGap, 1 + ICN_VGap, HIcon, 32, 32, 0, 0, DI_NORMAL)

End Sub

Public Sub IconClose(ByRef booLayout As Boolean)
Rem - Hide the icon, saving it's settings if not otherwise requested

    If booIsIconized Then Call p_AppWindows(True)
    Me.Visible = False
    Set WinList = Nothing
    If mnuPreserve.Checked Then Call p_settings("Save") Else Call p_settings("Clear")
    If HIcon <> 0 Then Call DestroyIcon(HIcon)


    Rem - Handle the array changing here
    Dim k As Long
    For k = ICN_Forms.Count To 1 Step -1
        If ICN_Forms(k) Is Me Then Call ICN_Forms.Remove(k)
    Next k
    Call cProcInfo.Unload
    Unload Me
    If booLayout Then Call Icon_Layout
    
End Sub


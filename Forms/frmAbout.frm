VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "| ZenKEY configuration |"
   ClientHeight    =   3075
   ClientLeft      =   6285
   ClientTop       =   3645
   ClientWidth     =   3090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFC0C0&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":058A
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   206
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   2400
   End
   Begin VB.Label lblBy 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ZenKEY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   0
      Left            =   1140
      TabIndex        =   1
      Top             =   120
      Width           =   675
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V. 1.9.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1200
      TabIndex        =   0
      Top             =   1140
      Width           =   585
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Rem -------- Printing text, each letter a cutout
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Rem --------- Create YinTang region
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function SetWindowRgn Lib "user32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Const RGN_AND = 1        'Intersection des deux régions
Private Const RGN_OR = 2         'Addition des deux régions
Private Const RGN_XOR = 3        'Difficile à décrire ... essayez
Private Const RGN_DIFF = 4       'Soustraction de la région 2 à la région 1
Private Const RGN_COPY = 5       'Copie la région 1
Private YY As Long
Rem - For form transparency
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Rem - New
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Sub Display()
    
    'Me.Width = Me.Height
'    lblHeader.Left = 0.5 * Me.ScaleWidth - 0.5 * lblHeader.Width
    'If Command$ = "ABOUT" Then
    '    lblBy.Caption = "by ZenCODE" ' /" & vbCr & "R. T. Larkin"
    '    lblBy.Move 0.5 * Me.ScaleWidth - 0.5 * lblBy.Width, lblBy.Top - 6
    'End If
    
    lblVersion.Caption = "V. " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    'lblVersion.Left = (0.75 + 0.125) * Me.ScaleWidth - 0.5 * lblVersion.Width
    
    Call CentreForm(Me)
    
    If Command$ = "SPLASH" Then
        On Error Resume Next
        lblBy(0).Visible = False
        Timer1.Enabled = True
        'Call YingYang_Round(Me)
        Call SetTrans(Me.hwnd, "SETTRANSPARENCY=50")
    Else
        Rem - About
        'Call YingYang_Square(Me)
        Call SetTrans(Me.hwnd, "SETTRANSPARENCY=75")
        
    End If
    Call YingYang_Square(Me)
    
    If lblBy.UBound < 1 Then Load lblBy(1)
    lblBy(1).Move lblBy(0).left + 2, lblBy(0).Top + 1
    lblBy(1).ForeColor = RGB(240, 240, 240)
    lblBy(1).Visible = True
    Me.Show
    
    
    'Call SetTrans(Me.hwnd, "SETTRANSPARENCY=75")
'    Exit Sub

'    Call TileMe(Me) ' LoadPicture(App.Path & "\Bluerock.bmp"))
'    strText1 = "ZenKEY"
'    strText2 = "by ZenCODE"
'    CircleSize = 12
'    X1 = Me.ScaleWidth / 2
'    Y1 = Me.ScaleHeight / 2
'    Rem - Colour the bit where the circle will be
'    Me.FillStyle = vbFSSolid
'
'    Me.Line (X1 - 2 * CircleSize, Y1 - 1.5 * CircleSize)-(X1, Y1 + 3 * CircleSize), vbBlack, BF 'Colour halves
'    Me.Line (X1 + Me.DrawWidth, Y1 - 1.5 * CircleSize)-Step(CircleSize, 3 * CircleSize), vbWhite, BF
'    Me.FillColor = vbWhite
'    Me.Circle (X1 - 0.5 * CircleSize, Y1), CircleSize / 5, vbWhite
'    Me.FillColor = vbBlack
'    Me.Circle (X1 + 0.5 * CircleSize, Y1), CircleSize / 5, vbBlack
    
'    Call CentreForm(Me)
'    On Error GoTo ErrorTrap:
'
'    With Me
'        Rem - FontName = "Copperplate Gothic Bold"
'        .FontSize = 80
'        X = .ScaleWidth / 2 - .TextWidth(strText1) / 2
'
'        Rem ------------------------------------  Begin to trace the path
'        lngResult = BeginPath(.hdc)
'        If lngResult <> 1 Then Call Err.Raise(vbObjectError + 1)
'        Rem - Print ZenKEY
'        lngResult = TextOut(.hdc, X, 0, strText1, Len(strText1))
'        If lngResult <> 1 Then Call Err.Raise(vbObjectError + 1)
'        Y = 0.8 * .TextHeight(strText1)
'        Rem - Print ZenCODE
'        .Font.Size = 14
'        X = .ScaleWidth / 2 - .TextWidth(strText2) / 2
'        lngResult = TextOut(.hdc, X, Y, strText2, Len(strText2))
'        If lngResult <> 1 Then Call Err.Raise(vbObjectError + 1)
'        Rem - Print Last bit
'        .Font.Size = 14
'        strText1 = String(60, "-")
'        X = .ScaleWidth / 2 - .TextWidth(strText1) / 2
'        lngResult = TextOut(.hdc, X, Y + 25, strText1, Len(strText1))
'        If lngResult <> 1 Then Call Err.Raise(vbObjectError + 1)
'
''        Me.Circle (X1, Y1), CircleSize
'        Rem - Close the path bracket
'        Rem - Convert the path to a region
'        lngResult = EndPath(.hdc)
'        If lngResult <> 1 Then Call Err.Raise(vbObjectError + 1)
'        hRgn = PathToRegion(.hdc)
'        If hRgn = 0 Then Call Err.Raise(vbObjectError + 1)
'        Rem - Set the Window-region
'        lngResult = SetWindowRgn(.hwnd, hRgn, True)
'        If lngResult <> 1 Then Call Err.Raise(vbObjectError + 1)
'    End With
'    Rem - Destroy our region
'    Call DeleteObject(hRgn)
'
'    Exit Sub
'
'ErrorTrap:
'
'    Rem - Something failed. Fine then, if we can't have a fancy splash, make it simple
'    'On Error Resume Next
'    With Me
'        Call EndPath(.hdc)
'        .CurrentX = (.ScaleWidth - .TextWidth(strText1)) / 2
'        Me.Print strText1;
'        .CurrentY = 0.8 * .TextHeight(strText1)
'        .Font.Size = 12
'        .CurrentX = (.ScaleWidth - .TextWidth(strText2)) / 2
'        Me.Print strText2
'    End With
    
End Sub

Private Sub SetTrans(ByVal lngHWnd As Long, ByVal Action As String)
    
    Const LWA_ALPHA = &H2
    Const GWL_EXSTYLE = (-20)
    Const WS_EX_LAYERED = &H80000
    Dim Ret As Long
    
    Rem - Set the window style to 'Layered'
    Ret = GetWindowLong(lngHWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    Call SetWindowLong(lngHWnd, GWL_EXSTYLE, Ret)
    Rem - 255 = Totally opague
    Ret = Abs(255 * Val(Mid$(Action, 17)) / 100) Mod 256
    If Ret < 16 Then Ret = 16 ' Sets level of transparency
    Call SetLayeredWindowAttributes(lngHWnd, 0, Ret, LWA_ALPHA)

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)

    If StartSection = "SPLASH" Then
    
        Rem - Run the Wizard on first execution
        Dim rc As New clsZenDictionary
        Call rc.FromINI(settings("SavePath") & "\ProgInfo.ini")
        
        Dim lngRunCount As Long
        lngRunCount = Val(rc("RunCount"))
        rc("RunCount") = CStr(lngRunCount + 1)
        Call rc.ToINI(settings("SavePath") & "\ProgInfo.ini")
        
        If lngRunCount = 0 Then
            Rem - Being run for the first time
            Dim strDisplay As String
            strDisplay = "Greetings! " & vbCr & vbCr & "You can access ZenKEY by pressing 'Alt + Space'. It comes with a complete configuration which you can edit at anytime, or you can use the Wizard to set it up. "
            strDisplay = strDisplay & vbCr & vbCr & "Would you like ZenKEY to continue showing tips each time it starts up?"
            If 0 = ZenMB(strDisplay, "Yes", "No") Then
                settings("ShowTips") = "True"
            Else
                settings("ShowTips") = ""
            End If
            Call settings.ToINI(settings("SavePath") & "\Settings.ini")
        End If
        
        If settings("ShowTips") = "True" Then
            Dim strTemp As String
            
            Select Case lngRunCount
                Case 0
                    Exit Sub
                Case 1
                    strTemp = "ZenKEY is not an aeroplane!" & vbCr & vbCr & _
                        "ZenKEY is made up of a series of actions, with each action being listed in a menu. You can:" & vbCr & vbCr & _
                        "a) create actions that open programs or documents, position windows or alter their properties." & vbCr & _
                        "b) assign global keystrokes to fire these actions from any program" & vbCr & _
                        "c) create menus containing these actions" & vbCr & _
                        "d) click or right click on the ZenKEY window to show these menus" & vbCr & _
                        "e) assign a keystroke to show any menu in any program" & vbCr & _
                        "f) open a website or internet resource " & vbCr & _
                        "g) act with compassion and respect for all living creatures " & vbCr & vbCr & _
                        "You can use the ZenKEY Configuration utility or the Wizard to setup ZenKEY for your needs. " & _
                        "You can also ignore ZenKEY and hope it goes away. Who knows what is wisest?"
                Case 2: strTemp = "Why go the menu when the menu can come to you? Use 'Alt + Space' (or assign any keys you want) to show menus at your mouse pointer."
                Case 3: strTemp = "Why hunt for program icons? Assign keystrokes to immediately launch any program on your computer."
                Case 4: strTemp = "Why spend time dragging and positioning windows? Assign keystrokes to instantly position, resize or alter any window."
                Case Else
                    strTemp = ZenKEYCap
            End Select

            Me.Hide
            Call ZenMB(strTemp, "OK")
            
        End If
    End If


End Sub




Private Sub Timer1_Timer()
Static booSecond As Boolean

    
    If booSecond Then
        Rem - Show splash screen for xx milliseconds.
        Timer1.Enabled = False
        Unload Me
    Else
        Rem - Delay showing of screen for xx milliseconds
        Timer1.Interval = 2000
        Me.Show
    End If
    
    
    booSecond = Not booSecond
    
End Sub



Public Sub YingYang_Square(obj As Form)
   'Déclaration des différents "handles" des différentes "régions" de la feuille, qui, réunies, formeront le Ying Yang
   Dim Cercle        As Long
   Dim RECT          As Long
   Dim PCercleH      As Long
   Dim PCercleB      As Long
   Dim HCercle       As Long
   Dim Cadre         As Long
   Dim TrouB         As Long
   Dim TrouH         As Long
   Dim CercleBis     As Long
   Dim HCercleBis    As Long
   Dim CercleBisBis  As Long
   Dim Ying_Yang     As Long
   Dim YYang         As Long

   Dim H             As Long
   Dim L             As Long
   Dim HBord         As Long
   Dim LBord         As Long
   Dim HT            As Long
   Dim LT            As Long

   H = obj.Height / Screen.TwipsPerPixelY
   L = obj.Width / Screen.TwipsPerPixelX

   HBord = 0 ' 1.5 * Int(H / 100) ' Border width
   LBord = 0 '1.5 * Int(L / 100)

   HT = 1.2 * Int(H / 10) ' Cut circle diameters
   LT = 1.2 * Int(L / 10)

   'Création des différentes "régions", et combinaisons entre elles
   'Attention : pour réaliser une combinaison, la variable-région de destination
   'doit déjà avoir été intialisée en lui affectant une région auparavant.

   'HCercle = CreateEllipticRgn(((L - (2 * LBord)) / 4) + LBord, ((H - (2 * HBord)) / 2) + HBord, 3 * (((L - (2 * LBord)) / 4) + LBord), (H - HBord))
   'HCercle = CreateEllipticRgn(((L - (2 * LBord)) / 4) + LBord, ((H - (2 * HBord)) / 2) + HBord, 3 * (((L - (2 * LBord)) / 4) + LBord), (H - HBord))



   Dim rgn1 As Long, rgn2 As Long, rgn3 As Long
   Dim lngFinal As Long
   Dim pt As POINTAPI

    pt.X = LBord: pt.Y = HBord
    rgn1 = BeginPath(Me.hdc)
    Call MoveToEx(Me.hdc, LBord, HBord, pt)  ' Top left
    Call LineTo(Me.hdc, LBord, H - HBord - 2) ' Bot left
    Call LineTo(Me.hdc, 0.241 * L, H - HBord - 2) ' Bot right
    Call LineTo(Me.hdc, 0.241 * L, 0.5 * H) ' Mid right
    Call LineTo(Me.hdc, 0.765 * L, 0.5 * H) ' left center
    Call LineTo(Me.hdc, 0.765 * L, HBord)  ' Top center
    Call LineTo(Me.hdc, LBord, HBord)  ' Bot left
    Call EndPath(Me.hdc)

'    rgn1 = BeginPath(Me.hdc)
'    Call MoveToEx(Me.hdc, LBord, HBord, pt)  ' Top left
'    Call LineTo(Me.hdc, LBord, H - HBord - 2) ' Bot left
'    Call LineTo(Me.hdc, 0.735 * L, H - HBord - 2) ' Bot right
'    Call LineTo(Me.hdc, 0.735 * L, 0.5 * H) ' Mid right
'    Call LineTo(Me.hdc, 0.25 * L, 0.5 * H) ' left center
'    Call LineTo(Me.hdc, 0.25 * L, HBord) ' Top center
'    Call LineTo(Me.hdc, LBord, HBord)  ' Bot right
'    Call EndPath(Me.hdc)
    rgn1 = PathToRegion(Me.hdc)

   Rem -  Cut out top square
    pt.X = LBord: pt.Y = HBord
    rgn2 = BeginPath(Me.hdc)
    LBord = LBord + 0.4 * L: HBord = HBord + 0.14 * H
    Call MoveToEx(Me.hdc, LBord, HBord, pt)  ' Top left
    Call LineTo(Me.hdc, LBord + 0.18 * L, HBord) ' across
    Call LineTo(Me.hdc, LBord + 0.18 * L, HBord + 0.18 * H) ' down
    Call LineTo(Me.hdc, LBord, HBord + 0.18 * H)  'left
    Call LineTo(Me.hdc, LBord, HBord)   ' up
    Call EndPath(Me.hdc)
    rgn2 = PathToRegion(Me.hdc)
   CombineRgn rgn1, rgn1, rgn2, RGN_XOR ' Remove center
   
   Rem -  Put back botton square
    pt.X = LBord: pt.Y = HBord
    rgn3 = BeginPath(Me.hdc)
    LBord = LBord + 0.008 * L: HBord = HBord + 0.52 * H
    Call MoveToEx(Me.hdc, LBord, HBord, pt)  ' Top left
    Call LineTo(Me.hdc, LBord + 0.167 * L, HBord) ' across
    Call LineTo(Me.hdc, LBord + 0.167 * L, HBord + 0.173 * H) ' down
    Call LineTo(Me.hdc, LBord, HBord + 0.173 * H)  'left
    Call LineTo(Me.hdc, LBord, HBord)   ' up
    Call EndPath(Me.hdc)
    rgn3 = PathToRegion(Me.hdc)
   CombineRgn rgn1, rgn1, rgn3, RGN_XOR ' Remove center
   
   
   'CombineRgn rgn1, rgn1, rgn2, RGN_XOR ' Remove center

   'rgn2 = CreateRectRgn(0, 0, L, H) ' Outer border
'   rgn3 = CreateEllipticRgn(((L - (2 * LBord)) / 2) + LBord - (LT / 2), ((3 * (H - (2 * HBord)) / 4)) + HBord - (HT / 2), ((L - (2 * LBord)) / 2) + LBord + (LT / 2), ((3 * (H - (2 * HBord)) / 4)) + HBord + (HT / 2)) ' lower circel
'   CombineRgn rgn2, rgn2, rgn3, RGN_XOR ' Remove center

   'CombineRgn rgn1, rgn1, rgn2, RGN_XOR ' Remove center
   lngFinal = rgn1

   SetWindowRgn obj.hwnd, lngFinal, True 'Applique la région finale à la feuille


   'Suppression des régions
   DeleteObject Cercle
   DeleteObject RECT
   DeleteObject PCercleH
   DeleteObject PCercleB
   DeleteObject HCercle
   DeleteObject Cadre
   DeleteObject TrouB
   DeleteObject TrouH
   DeleteObject CercleBis
   DeleteObject HCercleBis
   DeleteObject CercleBisBis
   DeleteObject Ying_Yang
   DeleteObject YYang


'Rem - Original code
'
''   'Déclaration des différents "handles" des différentes "régions" de la feuille, qui, réunies, formeront le Ying Yang
''   Dim Cercle        As Long
''   Dim Rect          As Long
''   Dim PCercleH      As Long
''   Dim PCercleB      As Long
''   Dim HCercle       As Long
''   Dim Cadre         As Long
''   Dim TrouB         As Long
''   Dim TrouH         As Long
''   Dim CercleBis     As Long
''   Dim HCercleBis    As Long
''   Dim CercleBisBis  As Long
''   Dim Ying_Yang     As Long
''   Dim YYang         As Long
''
''   Dim H             As Long
''   Dim L             As Long
''   Dim HBord         As Long
''   Dim LBord         As Long
''   Dim HT            As Long
''   Dim LT            As Long
''
''   H = obj.Height / Screen.TwipsPerPixelY
''   L = obj.Width / Screen.TwipsPerPixelX
''
''   HBord = Int(H / 100)
''   LBord = Int(L / 100)
''
''   HT = Int(H / 10)
''   LT = Int(L / 10)
''
''   'Création des différentes "régions", et combinaisons entre elles
''   'Attention : pour réaliser une combinaison, la variable-région de destination
''   'doit déjà avoir été intialisée en lui affectant une région auparavant.
''
''   HCercle = CreateEllipticRgn(((L - (2 * LBord)) / 4) + LBord, ((H - (2 * HBord)) / 2) + HBord, 3 * (((L - (2 * LBord)) / 4) + LBord), (H - HBord))
''   Cercle = CreateEllipticRgn(LBord, HBord, L - LBord, H - HBord)
''   Rect = CreateRectRgn(L / 2, 0, L, H)
''   CombineRgn HCercle, Cercle, Rect, RGN_DIFF
''
''   HCercleBis = CreateEllipticRgn(LBord, HBord, L - LBord, H - HBord)
''   PCercleB = CreateEllipticRgn(((L - (2 * LBord)) / 4) + LBord, ((H - (2 * HBord)) / 2) + HBord, 3 * (((L - (2 * LBord)) / 4) + LBord), (H - HBord))
''   CombineRgn HCercleBis, HCercle, PCercleB, RGN_DIFF
''
''   CercleBis = CreateEllipticRgn(LBord, HBord, L - LBord, H - HBord)
''   PCercleH = CreateEllipticRgn(((L - (2 * LBord)) / 4) + LBord, HBord, 3 * (((L - (2 * LBord)) / 4) + LBord), ((H - (2 * HBord)) / 2) + HBord)
''   CombineRgn CercleBis, Cercle, PCercleH, RGN_DIFF
''
''   CercleBisBis = CreateEllipticRgn(LBord, HBord, L - LBord, H - HBord)
''   HCercle = CreateEllipticRgn(0, 0, L, H)
''   CombineRgn CercleBisBis, CercleBis, HCercleBis, RGN_DIFF
''
''   Ying_Yang = CreateEllipticRgn(0, 0, L, H)
''   Cadre = CreateEllipticRgn(0, 0, L, H)
''   CombineRgn Ying_Yang, Cadre, CercleBisBis, RGN_DIFF
''
''   YYang = CreateEllipticRgn(0, 0, L, H)
''   TrouB = CreateEllipticRgn(((L - (2 * LBord)) / 2) + LBord - (LT / 2), ((3 * (H - (2 * HBord)) / 4)) + HBord - (HT / 2), ((L - (2 * LBord)) / 2) + LBord + (LT / 2), ((3 * (H - (2 * HBord)) / 4)) + HBord + (HT / 2))
''   CombineRgn YYang, Ying_Yang, TrouB, RGN_OR
''
''   YY = CreateEllipticRgn(0, 0, L, H)
''   TrouH = CreateEllipticRgn(((L - (2 * LBord)) / 2) + LBord - (LT / 2), ((H - (2 * HBord)) / 4) + HBord - (HT / 2), ((L - (2 * LBord)) / 2) + LBord + (LT / 2), ((H - (2 * HBord)) / 4) + HBord + (HT / 2))
''   CombineRgn YY, YYang, TrouH, RGN_DIFF
''
''   SetWindowRgn obj.Hwnd, YY, True 'Applique la région finale à la feuille
''
''   'Suppression des régions
''   DeleteObject Cercle
''   DeleteObject Rect
''   DeleteObject PCercleH
''   DeleteObject PCercleB
''   DeleteObject HCercle
''   DeleteObject Cadre
''   DeleteObject TrouB
''   DeleteObject TrouH
''   DeleteObject CercleBis
''   DeleteObject HCercleBis
''   DeleteObject CercleBisBis
''   DeleteObject Ying_Yang
''   DeleteObject YYang

End Sub




Public Sub YingYang_Round(obj As Form)
'   'Déclaration des différents "handles" des différentes "régions" de la feuille, qui, réunies, formeront le Ying Yang
'   Dim Cercle        As Long
'   Dim RECT          As Long
'   Dim PCercleH      As Long
'   Dim PCercleB      As Long
'   Dim HCercle       As Long
'   Dim Cadre         As Long
'   Dim TrouB         As Long
'   Dim TrouH         As Long
'   Dim CercleBis     As Long
'   Dim HCercleBis    As Long
'   Dim CercleBisBis  As Long
'   Dim Ying_Yang     As Long
'   Dim YYang         As Long
'
'   Dim H             As Long
'   Dim L             As Long
'   Dim HBord         As Long
'   Dim LBord         As Long
'   Dim HT            As Long
'   Dim LT            As Long
'
'   H = obj.Height / Screen.TwipsPerPixelY
'   L = obj.Width / Screen.TwipsPerPixelX
'
'   HBord = Int(H / 100)
'   LBord = Int(L / 100)
'
'   HT = Int(H / 10)
'   LT = Int(L / 10)
'
'   'Création des différentes "régions", et combinaisons entre elles
'   'Attention : pour réaliser une combinaison, la variable-région de destination
'   'doit déjà avoir été intialisée en lui affectant une région auparavant.
'
'   HCercle = CreateEllipticRgn(((L - (2 * LBord)) / 4) + LBord, ((H - (2 * HBord)) / 2) + HBord, 3 * (((L - (2 * LBord)) / 4) + LBord), (H - HBord))
'   Cercle = CreateEllipticRgn(LBord, HBord, L - LBord, H - HBord)
'   RECT = CreateRectRgn(L / 2, 0, L, H)
'   CombineRgn HCercle, Cercle, RECT, RGN_DIFF
'
'   HCercleBis = CreateEllipticRgn(LBord, HBord, L - LBord, H - HBord)
'   PCercleB = CreateEllipticRgn(((L - (2 * LBord)) / 4) + LBord, ((H - (2 * HBord)) / 2) + HBord, 3 * (((L - (2 * LBord)) / 4) + LBord), (H - HBord))
'   CombineRgn HCercleBis, HCercle, PCercleB, RGN_DIFF
'
'   CercleBis = CreateEllipticRgn(LBord, HBord, L - LBord, H - HBord)
'   PCercleH = CreateEllipticRgn(((L - (2 * LBord)) / 4) + LBord, HBord, 3 * (((L - (2 * LBord)) / 4) + LBord), ((H - (2 * HBord)) / 2) + HBord)
'   CombineRgn CercleBis, Cercle, PCercleH, RGN_DIFF
'
'   CercleBisBis = CreateEllipticRgn(LBord, HBord, L - LBord, H - HBord)
'   HCercle = CreateEllipticRgn(0, 0, L, H)
'   CombineRgn CercleBisBis, CercleBis, HCercleBis, RGN_DIFF
'
'   Ying_Yang = CreateEllipticRgn(0, 0, L, H)
'   Cadre = CreateEllipticRgn(0, 0, L, H)
'   CombineRgn Ying_Yang, Cadre, CercleBisBis, RGN_DIFF
'
'   YYang = CreateEllipticRgn(0, 0, L, H)
'   TrouB = CreateEllipticRgn(((L - (2 * LBord)) / 2) + LBord - (LT / 2), ((3 * (H - (2 * HBord)) / 4)) + HBord - (HT / 2), ((L - (2 * LBord)) / 2) + LBord + (LT / 2), ((3 * (H - (2 * HBord)) / 4)) + HBord + (HT / 2))
'   CombineRgn YYang, Ying_Yang, TrouB, RGN_OR
'
'   YY = CreateEllipticRgn(0, 0, L, H)
'   TrouH = CreateEllipticRgn(((L - (2 * LBord)) / 2) + LBord - (LT / 2), ((H - (2 * HBord)) / 4) + HBord - (HT / 2), ((L - (2 * LBord)) / 2) + LBord + (LT / 2), ((H - (2 * HBord)) / 4) + HBord + (HT / 2))
'   CombineRgn YY, YYang, TrouH, RGN_DIFF
'
'   SetWindowRgn obj.hwnd, YY, True 'Applique la région finale à la feuille
'
'   'Suppression des régions
'   DeleteObject Cercle
'   DeleteObject RECT
'   DeleteObject PCercleH
'   DeleteObject PCercleB
'   DeleteObject HCercle
'   DeleteObject Cadre
'   DeleteObject TrouB
'   DeleteObject TrouH
'   DeleteObject CercleBis
'   DeleteObject HCercleBis
'   DeleteObject CercleBisBis
'   DeleteObject Ying_Yang
'   DeleteObject YYang

End Sub

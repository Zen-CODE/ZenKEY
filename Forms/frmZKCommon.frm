VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmZKConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "|                                                              ZenKEY configuration"
   ClientHeight    =   5565
   ClientLeft      =   3030
   ClientTop       =   3090
   ClientWidth     =   7935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   2
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmZKCommon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmZKCommon.frx":058A
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   Begin VB.TextBox txtScetion 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3060
      TabIndex        =   11
      Top             =   112
      Width           =   1815
   End
   Begin VB.ComboBox cmbObjectType 
      Height          =   315
      Left            =   3060
      TabIndex        =   9
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3060
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid msfGrid 
      Height          =   2655
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      Enabled         =   -1  'True
      TextStyleFixed  =   3
      SelectionMode   =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Menu item"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5340
      TabIndex        =   21
      Top             =   4380
      Width           =   1050
   End
   Begin VB.Label Hoteky 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu items"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2160
      TabIndex        =   20
      Top             =   4380
      Width           =   810
   End
   Begin VB.Label lblDisableSec 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disable section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6120
      TabIndex        =   14
      Top             =   720
      Width           =   1515
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&New section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6120
      TabIndex        =   13
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label lblSecCount 
      AutoSize        =   -1  'True
      Caption         =   "1 of 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5160
      TabIndex        =   12
      Top             =   660
      Width           =   570
   End
   Begin VB.Image imiSecRight 
      Height          =   375
      Left            =   5520
      Picture         =   "frmZKCommon.frx":2059
      Top             =   90
      Width           =   375
   End
   Begin VB.Image imiSecLeft 
      Height          =   375
      Left            =   5100
      Picture         =   "frmZKCommon.frx":24E0
      Top             =   90
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1260
      TabIndex        =   10
      Top             =   180
      Width           =   1650
   End
   Begin VB.Label lblObjectType 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Object type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1515
      TabIndex        =   8
      Top             =   1140
      Width           =   1110
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   1260
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   6
      Top             =   675
      Width           =   780
   End
   Begin VB.Label lblEditAction 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Action"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6480
      TabIndex        =   5
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label lblEditDes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Caption"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4080
      TabIndex        =   4
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label lblRemove 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2640
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblAdd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1380
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblEditHK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Hotkey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5280
      TabIndex        =   1
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Image imiHeader 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   6030
      Left            =   120
      Picture         =   "frmZKCommon.frx":2955
      Top             =   120
      Width           =   930
   End
   Begin VB.Shape shpEditHK 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   5220
      Shape           =   4  'Rounded Rectangle
      Top             =   4620
      Width           =   1155
   End
   Begin VB.Shape shpAdd 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   4620
      Width           =   1215
   End
   Begin VB.Shape shpRemove 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   2580
      Shape           =   4  'Rounded Rectangle
      Top             =   4620
      Width           =   1215
   End
   Begin VB.Shape shpEditDes 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   4020
      Shape           =   4  'Rounded Rectangle
      Top             =   4620
      Width           =   1155
   End
   Begin VB.Shape shpEditAction 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   6420
      Shape           =   4  'Rounded Rectangle
      Top             =   4620
      Width           =   1155
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   1260
      Shape           =   4  'Rounded Rectangle
      Top             =   615
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   1260
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   6060
      Shape           =   4  'Rounded Rectangle
      Top             =   1140
      Width           =   1635
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   6060
      Shape           =   4  'Rounded Rectangle
      Top             =   660
      Width           =   1635
   End
   Begin VB.Label lblDeleteSec 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6120
      TabIndex        =   15
      Top             =   240
      Width           =   1515
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   6060
      Shape           =   4  'Rounded Rectangle
      Top             =   180
      Width           =   1635
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Default"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1260
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblReload 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Restore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3060
      TabIndex        =   18
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6420
      TabIndex        =   17
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4740
      TabIndex        =   16
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Shape shpCancel 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   5100
      Width           =   1335
   End
   Begin VB.Shape shpSave 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   4680
      Shape           =   4  'Rounded Rectangle
      Top             =   5100
      Width           =   1335
   End
   Begin VB.Shape shpReload 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   5100
      Width           =   1335
   End
   Begin VB.Shape shpDefault 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   5100
      Width           =   1335
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   1260
      Top             =   4320
      Width           =   2595
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   3960
      Top             =   4320
      Width           =   3675
   End
End
Attribute VB_Name = "frmZKConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intCurrentSection As Integer
Dim booChanged As Boolean
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Dim objSelected As Control
Private Sub cmbGroup_Click()

'        Call Grid_Load(cmbGroup.Text)
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (objSelected Is Nothing) Then Call Set_Focus
End Sub





Private Sub lblAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If objSelected Is Nothing Then
        Call Set_Focus(lblAdd)
    ElseIf objSelected <> lblAdd Then
        Call Set_Focus(lblAdd)
    End If
End Sub


Private Sub lblDefault_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If objSelected Is Nothing Then
        Call Set_Focus(lblDefault)
    ElseIf objSelected <> lblDefault Then
        Call Set_Focus(lblDefault)
    End If

End Sub


Private Sub lblEditAction_Click()
Dim strFName As String, strHeader As String

    strHeader = "Changing file to open "
    If GetOFName(strFName) Then
        With msfGrid
            Rem - Set to the presently selected item. Use the KeyValue and not it's description in the grid!
            Catagory(intCurrentSection).Hotkeys(.Row - 1).Action = strFName
            .TextMatrix(.Row, 1) = ExtractFileName(strFName)
        End With
        booChanged = True
    End If
End Sub

Private Sub lblEditAction_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If objSelected Is Nothing Then
        Call Set_Focus(lblEditAction)
    ElseIf objSelected <> lblEditAction Then
        Call Set_Focus(lblEditAction)
    End If
End Sub


Private Sub lblEditDes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If objSelected Is Nothing Then
        Call Set_Focus(lblEditDes)
    ElseIf objSelected <> lblEditDes Then
        Call Set_Focus(lblEditDes)
    End If
End Sub


Private Sub lblEditHK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If objSelected Is Nothing Then
        Call Set_Focus(lblEditHK)
    ElseIf objSelected <> lblEditHK Then
        Call Set_Focus(lblEditHK)
    End If
End Sub


Private Sub lblExit_Click()
    
    If UnLoadMe Then Unload Me
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Rem - Set the focus to this control if it doesen't have it already
    If objSelected Is Nothing Then
        Call Set_Focus(lblExit)
    ElseIf objSelected <> lblExit Then
        Call Set_Focus(lblExit)
    End If

End Sub



Private Sub lblReload_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If objSelected Is Nothing Then
        Call Set_Focus(lblReload)
    ElseIf objSelected <> lblReload Then
        Call Set_Focus(lblReload)
    End If
End Sub


Private Sub lblRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If objSelected Is Nothing Then
        Call Set_Focus(lblRemove)
    ElseIf objSelected <> lblRemove Then
        Call Set_Focus(lblRemove)
    End If
End Sub


Private Sub lblSave_Click()
           
    Call SaveKeys
    If booChanged Then Call Restart_PK
    booChanged = False

End Sub

Private Sub Form_Load()
    
    Call Menu_LoadINI
    Call CentreForm(Me)
    Call SetGraphics
    Call Grid_Init
    Call Listbox_Init
    'Call SetRegion
    
    Me.AutoRedraw = False
    
    
End Sub



Public Sub Grid_Init()
Dim k As Integer
Dim sngWidth As Single

    With msfGrid
        Rem - Seze the Picture boc
        sngWidth = .Width * 1.21
        For k = 0 To .Cols - 1
            Rem - Format = Caption, Action,  - Ctrl - Alt - Shift - Key
            
            .ColWidth(k) = Me.ScaleX(sngWidth * Choose(k + 1, 0.22, 0.25, 0.18, 0.12), vbPixels, vbTwips)
            .TextMatrix(0, k) = Choose(k + 1, "Caption", "Action", "Shift keys", "Key")
        Next k
    End With
    
End Sub


Public Sub Listbox_Init()
Dim k As Integer

    With cmbObjectType
        .Clear
        For k = 0 To ObjectTypeCount - 1
'            .AddItem Section(k)
'            If StartSection = UCase(Section(k)) Then .ListIndex = k
'        Next k
'        If .ListIndex < 0 Then .ListIndex = 0
    End With

End Sub

Public Sub Grid_Load(ByVal Section As String)
Dim k As Integer
Dim i As Integer
Dim strAction As String
Dim NumSections As Integer
Dim NumKeys As Integer

    Rem - Find which section we should load
    NumSections = UBound(Catagory())
    For k = 0 To NumSections
        If Catagory(k).Catagory = Section Then
            Rem - Bingo! Now we load the data into the grid
            intCurrentSection = k
            msfGrid.Visible = False
            msfGrid.Rows = 1
            NumKeys = UBound(Catagory(k).Hotkeys())
            For i = 0 To NumKeys
                With Catagory(k).Hotkeys(i)
                    strAction = Left$(.Action, 1) & LCase$(Mid$(.Action, 2))
                    Select Case Section
                        Case "Winamp"
                            msfGrid.TextMatrix(0, 1) = "Action"
                            If Len(.strKey) > 0 Then
                                msfGrid.AddItem .Descrip & vbTab & Winamp_GetDescrip(CStr(.Action)) & vbTab & .strShiftKeys & vbTab & KeyName(Asc(.strKey))
                            Else
                                msfGrid.AddItem .Descrip & vbTab & Winamp_GetDescrip(CStr(.Action)) & vbTab & "" & vbTab & ""
                            End If
                        Case "Open"
                            msfGrid.TextMatrix(0, 1) = "File to execute/open"
                            If Len(.strKey) > 0 Then
                                msfGrid.AddItem .Descrip & vbTab & ExtractFileName(.Action) & vbTab & .strShiftKeys & vbTab & KeyName(Asc(.strKey))
                            Else
                                msfGrid.AddItem .Descrip & vbTab & ExtractFileName(.Action) & vbTab & "" & vbTab & ""
                            End If
                        Case Else
                            msfGrid.TextMatrix(0, 1) = "Action"
                            If Len(.strKey) > 0 Then
                                msfGrid.AddItem .Descrip & vbTab & strAction & vbTab & .strShiftKeys & vbTab & KeyName(Asc(.strKey))
                            Else
                                msfGrid.AddItem .Descrip & vbTab & strAction & vbTab & "" & vbTab & ""
                            End If
                    End Select
                    Rem - Make only the relevant buttons visible!!
                    shpAdd.Visible = CBool(.Catagory = "Open")
                    lblAdd.Visible = shpAdd.Visible
                    'shpEditHK.Visible = True
                    'shpEditDes.Visible = True
                    shpEditAction.Visible = CBool(.Catagory = "Open")
                    lblEditAction.Visible = shpEditAction.Visible
                    shpRemove.Visible = lblEditAction.Visible
                    lblRemove.Visible = lblEditAction.Visible
                    
                End With
            Next i
            msfGrid.Visible = True
            
            
        End If
    Next k
End Sub






Public Sub Extract(ByRef TheLine As String, ParamArray Items())
Rem - Pumps the pipe separated items into Items()
Dim k As Integer, intEnd As Integer

    intEnd = InStr(TheLine, "|")
    For k = 0 To UBound(Items())
        Items(k) = Left$(TheLine, intEnd - 1)
        TheLine = Mid$(TheLine, intEnd + 1)
        intEnd = InStr(TheLine, "|")
    Next k
    
End Sub


Public Function Winamp_GetDescrip(ByRef ComValue As String) As String

    Select Case ComValue
        Case "40058": Winamp_GetDescrip = "Volume +" ' ' WINAMP_VOLUMEUP
        Case "40059": Winamp_GetDescrip = "Volume -"  ''Const WINAMP_VOLUMEDOWN
        Case "40060": Winamp_GetDescrip = "Forward" ''Const WINAMP_FFWD5S = 40060
        Case "40061": Winamp_GetDescrip = "Rewind" ''Const WINAMP_REW5S = 40061
        Case "40044": Winamp_GetDescrip = "Previous" ':  Winamp_GetCommand =  'Const WINAMP_BUTTON1 = 40044
        Case "40045": Winamp_GetDescrip = "Play" 'Const WINAMP_BUTTON2 = 40045
        Case "40046": Winamp_GetDescrip = "Pause" ''Const WINAMP_BUTTON3 = 40046
        Case "40047": Winamp_GetDescrip = "Stop" 'Const WINAMP_BUTTON4 = 40047
        Case "40048": Winamp_GetDescrip = "Next" 'Const WINAMP_BUTTON5 = 40048
        Case "40192": Winamp_GetDescrip = "Visualization" 'Const WINAMP_VISPLUGIN = 40192
        Case "40187": Winamp_GetDescrip = "Load dir" ''Const WINAMP_FILE_DIR = 40187
        Case "40188": Winamp_GetDescrip = "ID3 tag" ''Const WINAMP_EDIT_ID3 = 40188
        Case "40040": Winamp_GetDescrip = "Playlist" '"WINAMP_OPTIONS_PLEDIT" = 40040
        Case "40029": Winamp_GetDescrip = "Play file" '"WINAMP_OPTIONS_PLEDIT" = 40040
        Case "LAUNCH": Winamp_GetDescrip = "Launch\Capture"
        Case "40001": Winamp_GetDescrip = "Close Winmap"
        Case "40185": Winamp_GetDescrip = "Open location"
        
        Case Else:: Winamp_GetDescrip = "Undefined"
    End Select


'Const WINAMP_VOLUMEUP = 40058
'Const WINAMP_VOLUMEDOWN = 40059
'Const WINAMP_FFWD5S = 40060
'Const WINAMP_REW5S = 40061
'Const WINAMP_BUTTON1 = 40044
'Const WINAMP_BUTTON2 = 40045
'Const WINAMP_BUTTON3 = 40046
'Const WINAMP_BUTTON4 = 40047
'Const WINAMP_BUTTON5 = 40048
'Const WINAMP_VISPLUGIN = 40192
'Const WINAMP_FILE_DIR = 40187
'Const WINAMP_EDIT_ID3 = 40188
'Const WINAMP_OPTIONS_PLEDIT = 40040


End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode <> vbFormCode Then If Not UnLoadMe Then Cancel = 1
End Sub

Private Sub lblSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Rem - Set the focus to this control if it doesen't have it already
    If objSelected Is Nothing Then
        Call Set_Focus(lblSave)
    ElseIf objSelected <> lblSave Then
        Call Set_Focus(lblSave)
    End If

End Sub

Private Sub lblAdd_Click()

    Select Case Section(intCurrentSection)
        Case "Open": Call AddItem_Launch
    End Select
End Sub

Private Sub lblEditHK_Click()
    Dim FormKey As frmKeyCapture
    Dim strShift As String
    Dim strKey As String
    Dim strMBHeader As String
    
    Set FormKey = New frmKeyCapture
    Set FormKey.Picture = Me.Picture
    With msfGrid
        Rem - Set to the presently selected item. Use the KeyValue and not it's description in the grid!
        FormKey.strKey = Catagory(intCurrentSection).Hotkeys(.Row - 1).strKey
        FormKey.strShift = .TextMatrix(.Row, 2)
        Call FormKey.Initialise
        strMBHeader = Catagory(intCurrentSection).Hotkeys(msfGrid.Row - 1).Catagory & ", " & Catagory(intCurrentSection).Hotkeys(msfGrid.Row - 1).Descrip
        FormKey.Caption = strMBHeader
        
        
        Rem - Now so the form
        Me.Visible = False
        FormKey.Show vbModal
        Me.Visible = True
        If FormKey.booDone Then
            booChanged = True
            strShift = FormKey.strShift
            strKey = FormKey.strKey
            .TextMatrix(.Row, 2) = strShift
            Catagory(intCurrentSection).Hotkeys(.Row - 1).strShiftKeys = strShift
            If Len(strKey) > 0 Then
                .TextMatrix(.Row, 3) = KeyName(Asc(strKey))
            Else
                .TextMatrix(.Row, 3) = vbNullString
            End If
            Catagory(intCurrentSection).Hotkeys(.Row - 1).strKey = strKey
    End If
    End With
    
    Unload FormKey
    Set FormKey = Nothing
    
End Sub

Private Sub lblDefault_Click()

    Call Grid_Init
    Call LoadKeys(True)
    Call Listbox_Init
    booChanged = True
    
End Sub

Private Sub lblEditDes_Click()
Dim strTemp As String, strHeader As String

    strHeader = "Changing description: " & Catagory(intCurrentSection).Hotkeys(msfGrid.Row - 1).Catagory & ", " & Catagory(intCurrentSection).Hotkeys(msfGrid.Row - 1).Descrip
    strTemp = InputBox("Okay, you know better then! So what should the description be, KnowItAll!", strHeader)
    If Len(strTemp) > 0 Then
        With msfGrid
            Rem - Set to the presently selected item. Use the KeyValue and not it's description in the grid!
            Catagory(intCurrentSection).Hotkeys(.Row - 1).Descrip = strTemp
            .TextMatrix(.Row, 0) = strTemp
        End With
        booChanged = True
    End If
    
End Sub

Private Sub lblReload_Click()
    
    Call Grid_Init
    Call LoadKeys
    Call Listbox_Init
    booChanged = False
End Sub


Public Sub SaveKeys()

    Dim intFNum As Integer
    Dim intSection As Integer
    Dim intKey As Integer
    Dim SecIndexes As Integer
    Dim KeyIndexes As Integer
    Dim strLine As String
    
    Rem ======================     Open file and start loading
    Rem - Initialise variables
    intFNum = FreeFile
    SecIndexes = UBound(Section())
    
    Open App.Path & "\Powerkey.ini" For Output As #intFNum
        For intSection = 0 To SecIndexes
            Rem - Initialise for a new section
            Print #intFNum, "[" & Section(intSection) & "]"
            KeyIndexes = UBound(Catagory(intSection).Hotkeys())
            For intKey = 0 To KeyIndexes
                With Catagory(intSection).Hotkeys(intKey)
                    strLine = .Descrip & "|" & UCase(.Action) & "|" & .strShiftKeys & "|" & .strKey & "|"
                    Print #intFNum, strLine
                End With
            Next intKey
            Rem - Now read in the entries
            Print #intFNum, "[End]"
        Next intSection
    Close #intFNum
    Call MsgBox("Hotkeys saved successfully!", vbInformation, "Powerkey")

End Sub


Public Sub SetGraphics()
Dim X As Single, Y As Single

    
    With Me
        .AutoRedraw = True
        Call TileMe(Me)
        Set .Picture = .Image
    End With
    
End Sub

Public Sub Restart_PK()
    
    Const WM_CLOSE = &H10
    Dim k As Integer
    Dim lngHandle As Long
    
    lngHandle = FindWindow(vbNullString, "Powerkey V1")
    If lngHandle > 0 Then
        Rem - If there is an instance, force it to close and then restart it!
        k = FreeFile
        Open App.Path & "\Restart.pk" For Output As #k
        Close #k
        
        Call PostMessage(lngHandle, WM_CLOSE, 0&, ByVal 0&)
        DoEvents
        Call Shell(App.Path & "\Powerkey.exe RESTART", vbNormalFocus)
        
    End If

End Sub

Private Sub lblRemove_Click()

    Select Case Section(intCurrentSection)
        Case "Open"
            If vbYes = MsgBox("But like, removing an item is, like serious dude! Are you like sure?", vbYesNo, "Powerkey") Then
                Call RemItem
                booChanged = True
            End If
    End Select
    
End Sub

Private Sub msfGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    DoEvents
    If Section(intCurrentSection) <> "Open" Then Exit Sub
    
    With msfGrid
        If X > .ColWidth(0) And X < .ColWidth(1) + .ColWidth(2) Then
            Dim k As Integer, booFOund As Boolean
            .ToolTipText = vbNullString
            For k = 1 To .Rows - 1
                If Y > k * .RowHeight(k) And Y < (k + 1) * .RowHeight(k) Then
                    .ToolTipText = Catagory(intCurrentSection).Hotkeys(k - 1).Action
                End If
            Next k
        End If
    End With
End Sub

Private Sub msfGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'If Button <> 1 Then
    '    Call MsgBox("Please just click normally. What's all this right click nonsense?", vbOKOnly, "Powerkey Configurator")
    '    Exit Sub
    'Else
        'mnuAdd.Enabled = CBool(Section(intCurrentSection) = "Open")
        'mnuRemove.Enabled = CBool(Section(intCurrentSection) = "Open")
        'Call PopupMenu(mnuFIle)
    'End If
End Sub



Public Sub AddLaunchItem(ByRef Descrip As String, ByRef EXEPath As String)
Dim intMax As Integer

    With Catagory(intCurrentSection)
        intMax = UBound(.Hotkeys()) + 1
        ReDim Preserve .Hotkeys(intMax)
        .Hotkeys(intMax).Descrip = Descrip
        .Hotkeys(intMax).Action = EXEPath
        .Hotkeys(intMax).strKey = vbNullString
        .Hotkeys(intMax).strShiftKeys = vbNullString
    End With
    With msfGrid
        .AddItem Descrip & vbTab & ExtractFileName(EXEPath) & vbTab & "" & vbTab & ""
    End With
    booChanged = True
    
End Sub

Public Sub AddItem_Launch()
'    Dim OFName As OPENFILENAME
'
'    OFName.lStructSize = Len(OFName)
'    'Set the parent window
'    OFName.hwndOwner = Me.hwnd
'    'Set the application's instance
'    OFName.hInstance = App.hInstance
'    'Select a filter
'    'OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
'    OFName.lpstrFilter = "All Files (*.*)" + Chr$(0)
'    'create a buffer for the file
'    OFName.lpstrFile = Space$(254)
'    'set the maximum length of a returned file
'    OFName.nMaxFile = 255
'    'Create a buffer for the file title
'    OFName.lpstrFileTitle = Space$(254)
'    'Set the maximum length of a returned file title
'    OFName.nMaxFileTitle = 255
'    'Set the initial directory
'    OFName.lpstrInitialDir = "C:\"
'    'Set the title
'    OFName.lpstrTitle = "Select File - Powerkey"
'    'No flags
'    OFName.flags = 0
'
'    'Show the 'Open File'-dialog
'    If GetOpenFileName(OFName) Then
'        'MsgBox "File to Open: " + Trim$(OFName.lpstrFile)
'        Dim strName As String
'        strName = InputBox("Please enter the title of this application.", "Powerkey Open item")
'        If Len(strName) > 0 Then Call AddLaunchItem(strName, Trim(OFName.lpstrFile))
'    End If
'
    Rem - Show the 'Open File'-dialog
    Dim strFName As String
    If GetOFName(strFName) Then
        'MsgBox "File to Open: " + Trim$(OFName.lpstrFile)
        Dim strName As String
        strName = InputBox("Please enter a name for this action.", "Powerkey Open item")
        If Len(strName) > 0 Then Call AddLaunchItem(strName, strFName)
    End If
    
    
'    If GetOpenFileName(OFName) Then
'        'MsgBox "File to Open: " + Trim$(OFName.lpstrFile)
'        Dim strName As String
'        strName = InputBox("Please enter the title of this application.", "Powerkey Open item")
'        If Len(strName) > 0 Then Call AddLaunchItem(strName, Trim(OFName.lpstrFile))
'    End If
'
End Sub

Public Sub RemItem()
Dim intUBound As Integer
    
    With msfGrid
        intUBound = UBound(Catagory(intCurrentSection).Hotkeys())
        If (.Rows - 1 <> .Row) Then
            Rem - Damn! It's not the last row, so lets move it there...
            .TextMatrix(.Row, 0) = .TextMatrix(.Rows - 1, 0)
            .TextMatrix(.Row, 1) = .TextMatrix(.Rows - 1, 1)
            .TextMatrix(.Row, 2) = .TextMatrix(.Rows - 1, 2)
            .TextMatrix(.Row, 3) = .TextMatrix(.Rows - 1, 3)
            Catagory(intCurrentSection).Hotkeys(.Row - 1).Action = Catagory(intCurrentSection).Hotkeys(.Rows - 2).Action
            Catagory(intCurrentSection).Hotkeys(.Row - 1).Descrip = Catagory(intCurrentSection).Hotkeys(.Rows - 2).Descrip
            Catagory(intCurrentSection).Hotkeys(.Row - 1).strKey = Catagory(intCurrentSection).Hotkeys(.Rows - 2).strKey
            Catagory(intCurrentSection).Hotkeys(.Row - 1).strShiftKeys = Catagory(intCurrentSection).Hotkeys(.Rows - 2).strShiftKeys
        End If
        Rem - Okay. Now just remove the last items
        .Rows = .Rows - 1
        ReDim Preserve Catagory(intCurrentSection).Hotkeys(0 To intUBound - 1)
    End With

End Sub

Public Function ExtractFileName(ByRef FullName As String) As String
Dim Buffer As String
    
    Buffer = String(255, 0)
    Call GetFileTitle(FullName, Buffer, Len(Buffer))
    ExtractFileName = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)
    
End Function

Public Function UnLoadMe() As Boolean
    If booChanged Then
        UnLoadMe = CBool(vbYes = MsgBox("You have made changes. Are you sure you want to exit without saving? Your loss. Nothing to do with me...", vbYesNo, "Powerkey"))
    Else
        UnLoadMe = True
    End If
End Function

Public Sub Set_Focus(Optional ByRef DaLabel As Control = Nothing)
Const colSelected = &HC000&
Const colNormal = vbBlue

    If Not objSelected Is Nothing Then objSelected.ForeColor = colNormal
    If Not DaLabel Is Nothing Then DaLabel.ForeColor = colSelected
    Set objSelected = DaLabel

End Sub


Public Sub SetRegion()
Dim hRgn As Long
    
    Rem - Try rounded rect region with 3D logo
    Call CentreForm(Me)
    Rem - Include this line for the Ice header
    With Me
        Call RoundRect(.hdc, 1, 1, .ScaleWidth, .ScaleHeight, 30, 30)
        hRgn = CreateRoundRectRgn(1, 1, .ScaleWidth + 2, .ScaleHeight + 2, 30, 30)
        Call SetWindowRgn(.hwnd, hRgn, True)
        If hRgn > 0 Then Call DeleteObject(hRgn)
        
        
'        Dim hBrush As Long
'        hBrush = CreateSolidBrush(.ForeColor)
'        hRgn = CreateRoundRectRgn(1, 1, .ScaleWidth, .ScaleHeight, 40, 50)
'        Call FrameRgn(.hdc, hRgn, hBrush, 1, 1)
'        Call SetWindowRgn(.hwnd, hRgn, True)
'        If hRgn > 0 Then Call DeleteObject(hRgn)
'        If hBrush > 0 Then Call DeleteObject(hBrush)
    End With
    

End Sub

Public Function GetOFName(ByRef FName As String) As Boolean
    Dim OFName As OPENFILENAME
    
    OFName.lStructSize = Len(OFName)
    'Set the parent window
    OFName.hwndOwner = Me.hwnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Select a filter
    'OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OFName.lpstrFilter = "All Files (*.*)" + Chr$(0)
    'create a buffer for the file
    OFName.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    OFName.nMaxFile = 255
    'Create a buffer for the file title
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = "C:\"
    'Set the title
    OFName.lpstrTitle = "Select File - Powerkey"
    'No flags
    OFName.flags = 0

    'Show the 'Open File'-dialog
    GetOFName = GetOpenFileName(OFName)
    If GetOFName Then FName = Trim(OFName.lpstrFile)

End Function

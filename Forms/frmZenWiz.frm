VERSION 5.00
Begin VB.Form frmZenWiz 
   Caption         =   "| ZenKEY - Wizard |"
   ClientHeight    =   4500
   ClientLeft      =   7500
   ClientTop       =   4680
   ClientWidth     =   8955
   Icon            =   "frmZenWiz.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmZenWiz.frx":058A
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   Begin VB.CommandButton zbNext 
      Caption         =   "Next >>"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3780
      TabIndex        =   0
      Top             =   3945
      Width           =   1455
   End
   Begin VB.CommandButton zbCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3945
      Width           =   1455
   End
   Begin VB.Image imiBack 
      Height          =   2505
      Left            =   -240
      Picture         =   "frmZenWiz.frx":3125
      Top             =   900
      Visible         =   0   'False
      Width           =   4065
   End
   Begin VB.Image imiBot 
      Height          =   150
      Left            =   0
      Picture         =   "frmZenWiz.frx":4BB4
      Top             =   4380
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   4020
      Picture         =   "frmZenWiz.frx":5B44
      Top             =   2040
      Width           =   960
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome!"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   1020
      TabIndex        =   1
      Top             =   840
      Width           =   6975
   End
End
Attribute VB_Name = "frmZenWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit



Private Sub Form_Initialize()
    Dim X As Long
    X = InitCommonControls
    Call Init_ZK
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then End
End Sub

Private Sub zbCancel_Click()
    End
End Sub

Private Sub Form_Load()
Dim strTemp As String
    
    Call SetGraphics
    Set PicBack = Me.Picture
    
    Call ZW_Next("Start")
    Set HotKeys = New clsHotkey
    
    Rem - Initialise other stuff
    strTemp = "Welcome!" & vbCr & vbCr
    strTemp = strTemp & "This wizard will guide you through the process of setting up ZenKEY to do what you want it to."
    
    lblMessage.Caption = strTemp
    
End Sub


Private Sub SetGraphics()
Dim sngStartY As Single

    With Me
        .AutoRedraw = True
        sngStartY = .ScaleY(.Picture.Height, vbHimetric, .ScaleMode)
        
Dim sngWidth As Single, sngHeight As Single
Dim k As Single, i As Single
    
        sngWidth = .ScaleX(imiBack.Picture.Width, vbHimetric, .ScaleMode)
        sngHeight = .ScaleY(imiBack.Picture.Height, vbHimetric, .ScaleMode)
        For i = 0 To .ScaleHeight Step sngHeight
            For k = 0 To .ScaleWidth Step sngWidth
                .PaintPicture imiBack.Picture, k, i + sngStartY
            Next k
        Next i
            
        Rem - Paint on bottom bar
        sngHeight = .ScaleY(imiBot.Picture.Height, vbHimetric, .ScaleMode)
        .PaintPicture imiBot.Picture, -1, Me.ScaleHeight - sngHeight
    
        Set .Picture = .Image
        .AutoRedraw = False
        
    End With
        
        
        
    
End Sub

Private Sub zbNext_Click()

    Call ZW_Next("Type")

End Sub




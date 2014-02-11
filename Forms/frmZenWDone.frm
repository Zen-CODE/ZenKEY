VERSION 5.00
Begin VB.Form frmZenWDone 
   Caption         =   "| ZenKEY - Wizard |"
   ClientHeight    =   4500
   ClientLeft      =   7710
   ClientTop       =   5265
   ClientWidth     =   8955
   Icon            =   "frmZenWDone.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   Begin VB.CheckBox chkRestart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Restart ZenKEY so the changes take effect now."
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
      Height          =   195
      Left            =   2580
      TabIndex        =   4
      ToolTipText     =   "If ticked, any active instance of the program will be killed and restarted."
      Top             =   3240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton zbFinnish 
      Caption         =   "Finish"
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
      Left            =   7320
      TabIndex        =   1
      Top             =   3945
      Width           =   1455
   End
   Begin VB.CommandButton zbBack 
      Caption         =   "<< Back"
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
      TabIndex        =   3
      Top             =   3945
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   4020
      Picture         =   "frmZenWDone.frx":058A
      Top             =   1920
      Width           =   960
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "What exactly do you wish to do?"
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
      Height          =   615
      Left            =   1020
      TabIndex        =   2
      Top             =   1020
      Width           =   6975
   End
End
Attribute VB_Name = "frmZenWDone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    If ZK_Running Then chkRestart.Visible = True Else chkRestart.Value = 0
    Set zbFinnish.Picture = zbBack.Picture
    Set zbCancel.Picture = zbBack.Picture

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then End
End Sub


Private Sub zbBack_Click()
    Call ZW_Next("Previous")
End Sub


Private Sub zbCancel_Click()
    End
End Sub

Private Sub zbFinnish_Click()
    Call ZW_Next("Finnish")
End Sub



VERSION 5.00
Begin VB.Form frmZenWTemplate 
   Caption         =   "ZenKEY Configuration Wizard"
   ClientHeight    =   4440
   ClientLeft      =   4185
   ClientTop       =   5640
   ClientWidth     =   7470
   Icon            =   "frmWizTemplate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   296
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   Begin VB.CommandButton zbBack 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3930
      Width           =   1455
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "<< Back"
   End
   Begin VB.CommandButton zbNext 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   3930
      Width           =   1455
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Next >>"
   End
   Begin VB.CommandButton zbCancel 
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   3930
      Width           =   1455
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancel"
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "What exactly do you wish to do?"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   6975
   End
End
Attribute VB_Name = "frmZenWTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Option1_Click()

End Sub


Private Sub zbBack_Click()
    Call ZW_Next("Previous")
End Sub


Private Sub zbCancel_Click()
    End
End Sub


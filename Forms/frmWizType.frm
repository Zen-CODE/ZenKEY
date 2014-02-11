VERSION 5.00
Begin VB.Form frmZenWType 
   Caption         =   "| ZenKEY - Wizard |"
   ClientHeight    =   4500
   ClientLeft      =   4335
   ClientTop       =   5685
   ClientWidth     =   8955
   Icon            =   "frmWizType.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   Begin VB.OptionButton optRemove 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remove an item"
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
      Height          =   255
      Left            =   3300
      TabIndex        =   7
      ToolTipText     =   "Remove one of the items from ZenKEY"
      Top             =   2820
      Width           =   2535
   End
   Begin VB.OptionButton optEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit an existing item"
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
      Height          =   255
      Left            =   3300
      TabIndex        =   6
      ToolTipText     =   "Change one of the items in ZenKEY"
      Top             =   2340
      Width           =   2535
   End
   Begin VB.OptionButton optAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add a new item"
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
      Height          =   255
      Left            =   3300
      TabIndex        =   5
      ToolTipText     =   "Add a new item to ZenKEY"
      Top             =   1860
      Width           =   2535
   End
   Begin VB.OptionButton optNew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Set up ZenKEY from scratch"
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
      Height          =   255
      Left            =   3300
      TabIndex        =   4
      ToolTipText     =   "Setup an entirely new configuration"
      Top             =   1380
      Value           =   -1  'True
      Width           =   2535
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
      TabIndex        =   1
      Top             =   3945
      Width           =   1455
   End
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
      Left            =   7320
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "What would you like to do?"
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
      TabIndex        =   3
      Top             =   840
      Width           =   6975
   End
End
Attribute VB_Name = "frmZenWType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Set zbNext.Picture = zbBack.Picture
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

Private Sub zbNext_Click()
    
    Select Case True
        Case optRemove.Value
            Call ZW_Next("Remove")
        Case optEdit.Value
            Call ZW_Next("Edit")
        Case optAdd.Value
            Call ZW_Next("NewItem")
        Case optNew.Value
            Call ZW_Next("NewConfig")
            
    End Select
    
End Sub



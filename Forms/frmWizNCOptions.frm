VERSION 5.00
Begin VB.Form frmWizNCOptions 
   Caption         =   "| ZenKEY - Wizard |"
   ClientHeight    =   4500
   ClientLeft      =   7215
   ClientTop       =   7545
   ClientWidth     =   8955
   Icon            =   "frmWizNCOptions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   Begin VB.CheckBox chkControlPanel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "open Control Panel items"
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
      Left            =   1440
      TabIndex        =   17
      ToolTipText     =   "Open internet sites and searches with the press of a button."
      Top             =   2460
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkDTM 
      BackColor       =   &H00FFFFFF&
      Caption         =   "enable the 'Desktop map'"
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
      Left            =   4560
      TabIndex        =   16
      ToolTipText     =   "This feature allows you to move windows off-screen, expanding your desktop to almost any size."
      Top             =   3540
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkIDT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "enable the 'Infinite desktop'"
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
      Left            =   1440
      TabIndex        =   15
      ToolTipText     =   "This feature allows you to move windows off-screen, expanding your desktop to almost any size."
      Top             =   3540
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkAWT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "enable 'Auto-window transparency' (Graphics card recommended)"
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
      Left            =   1440
      TabIndex        =   14
      ToolTipText     =   "'Auto-window transparency' makes window transparent whenever they enter of lose focus."
      Top             =   3240
      Width           =   5355
   End
   Begin VB.CheckBox chkWinamp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "control Winamp or compatible players"
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
      Height          =   240
      Left            =   4560
      TabIndex        =   12
      ToolTipText     =   "Send these commands via Winamp messages instead of using Windows media commands."
      Top             =   1830
      Width           =   3435
   End
   Begin VB.CheckBox chkInternet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "access Internet resources"
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
      Left            =   1440
      TabIndex        =   10
      ToolTipText     =   "Open internet sites and searches with the press of a button."
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkFolders 
      BackColor       =   &H00FFFFFF&
      Caption         =   "open folders"
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
      Left            =   1440
      TabIndex        =   9
      ToolTipText     =   "Open folders with the press of a button."
      Top             =   1860
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkSystem 
      BackColor       =   &H00FFFFFF&
      Caption         =   " issue Windows' System commands"
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
      Height          =   240
      Left            =   4560
      TabIndex        =   8
      ToolTipText     =   "Shutdown, start the screensaver or do more with the press of a button."
      Top             =   2460
      Value           =   1  'Checked
      Width           =   3435
   End
   Begin VB.CheckBox chkWinPrograms 
      BackColor       =   &H00FFFFFF&
      Caption         =   "access Windows utilities and applets"
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
      Height          =   240
      Left            =   4560
      TabIndex        =   7
      ToolTipText     =   "Access hidden Windows utilities and programs with the press of a button."
      Top             =   2145
      Value           =   1  'Checked
      Width           =   3435
   End
   Begin VB.CheckBox chkWinMove 
      BackColor       =   &H00FFFFFF&
      Caption         =   "move and control program windows"
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
      Left            =   4560
      TabIndex        =   6
      ToolTipText     =   "Control the windows a program appears in with the press of a button"
      Top             =   1260
      Value           =   1  'Checked
      Width           =   3435
   End
   Begin VB.CheckBox chkDocuments 
      BackColor       =   &H00FFFFFF&
      Caption         =   "open documents"
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
      Left            =   1440
      TabIndex        =   5
      ToolTipText     =   "Open documents with the press of a button."
      Top             =   1560
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkPrograms 
      BackColor       =   &H00FFFFFF&
      Caption         =   "launch programs"
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
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "Open programs with the press of a button."
      Top             =   1260
      Value           =   1  'Checked
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
   Begin VB.CheckBox chkMedia 
      BackColor       =   &H00FFFFFF&
      Caption         =   "control Media players"
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
      Height          =   240
      Left            =   4560
      TabIndex        =   11
      ToolTipText     =   "Play, pause, alter volume or do more with the press of a button."
      Top             =   1530
      Value           =   1  'Checked
      Width           =   3435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "and you can use it to change the way your Windows Operating System behaves:"
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
      Left            =   900
      TabIndex        =   13
      Top             =   2880
      Width           =   6975
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "ZenKEY can be used in many different ways. You can use it to perform these functions:"
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
      Left            =   1020
      TabIndex        =   3
      Top             =   780
      Width           =   6975
   End
End
Attribute VB_Name = "frmWizNCOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim booLoading As Boolean

Private Sub chkAWT_Click()
    If chkAWT.Enabled Then Call ZenMB(ZK_TransWarn, "OK")
End Sub

Private Sub chkMedia_Click()

    If Not booLoading Then
        Rem - Warn them if they have enabled media commands and Winamp
        Call CheckWinampClash
    End If
    
End Sub

Private Sub chkWinamp_Click()

    If Not booLoading Then
        Call CheckWinampClash
    End If
    
End Sub


Private Sub Form_Activate()
    chkWinamp.Refresh
End Sub

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
    If Form_Valid Then
        Call ZW_Next("Next")
    Else
        Call ZenMB("You should choose to use ZenKEY for something, otherwise there is no point.")
    End If
End Sub



Private Function Form_Valid() As Boolean

Dim It As Control
    
    Form_Valid = False
    For Each It In Me.Controls
        If TypeOf It Is CheckBox Then
            If It.Value = 1 Then
                Form_Valid = True
                Exit Function
            End If
        End If
    Next It
    
    
End Function



Private Sub CheckWinampClash()

    If (chkWinamp.Value = 1) And (chkMedia.Value = 1) Then
        Dim strTemp As String
        strTemp = "Please note that the default Hotkeys for Winamp and Media player commands are the same. " & _
            "This means that you will need to change the Hotkeys later in the Wizard process to prevent them from clashing and failing." & _
            vbCr & vbCr & "You can also prevent this by enabling only 'control Winamp...' and not 'control Media players' or visa-versa."
        Call ZenMB(strTemp, "OK")
    End If
    
End Sub

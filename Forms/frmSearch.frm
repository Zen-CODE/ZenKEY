VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ZenKEY Search"
   ClientHeight    =   1635
   ClientLeft      =   8580
   ClientTop       =   6915
   ClientWidth     =   4965
   ClipControls    =   0   'False
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4965
   Begin VB.CommandButton zbButton 
      Caption         =   "Go"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1150
      Width           =   1335
   End
   Begin VB.TextBox txtSearch 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   645
      TabIndex        =   1
      Top             =   600
      Width           =   3675
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the text to search for."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   690
      TabIndex        =   0
      Top             =   180
      Width           =   3585
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim prop As clsZenDictionary





Public Sub SetGraphics()
    
    Call TileMe(Me, LoadPicture(App.Path & "\Help\clouds.jpg"))

End Sub

Private Sub Form_Activate()
On Error Resume Next

    If Me.Visible Then
        txtSearch.SetFocus
        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch.Text)
    End If
    
End Sub

Public Sub DoAction(ByRef zProp As clsZenDictionary)

    Call Initialise
    Set prop = zProp
    If Len(Command$) > 0 Then
        ' If fired for testing, stop here
        Me.Show vbModal
    Else
        Call SetWinPos(Me.hwnd, HWND_TOP, True)
    End If


End Sub
Private Function PrepCriteria(ByVal Criteria As String) As String
Dim k As Long

    PrepCriteria = Criteria
    k = InStr(PrepCriteria, " ")
    While k > 0
        Mid(PrepCriteria, k, 1) = "+"
        k = InStr(PrepCriteria, " ")
    Wend

End Function

Public Sub Initialise()
Dim k As Integer
    
    Rem - Do the colours
    
    Me.AutoRedraw = True
    Me.ForeColor = vbBlack
    Call SetGraphics
    Me.Move 0.5 * (Screen.Width - Me.Width), 0.5 * (Screen.Height - Me.Height)

End Sub



Private Sub txtSearch_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 13
            Call ZBButton_Click
            KeyAscii = 0
        Case 27
            Me.Hide
    End Select

End Sub


Private Sub ZBButton_Click()
Dim strCriteria As String
    
    Me.Hide
    strCriteria = txtSearch.Text
    If Len(strCriteria) > 0 Then
        Dim strSearch As String, intPos As Integer
        strSearch = prop("Action")
        intPos = InStr(strSearch, "<Criteria>")
        If intPos > 0 Then
            strSearch = left(strSearch, intPos - 1) & PrepCriteria(strCriteria) & Mid(strSearch, intPos + 10)
            Call ShellExe(strSearch)
        End If
    End If
    
End Sub



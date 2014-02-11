VERSION 5.00
Begin VB.Form frmWindowCapture 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "| ZenKEY - Window capture |"
   ClientHeight    =   3450
   ClientLeft      =   9135
   ClientTop       =   7125
   ClientWidth     =   4800
   Icon            =   "frmWindowCapture.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   230
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.CommandButton zbCancel 
      Caption         =   "Cancel"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton zbOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2100
      Picture         =   "frmWindowCapture.frx":058A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      ToolTipText     =   "Drag this cross over the program you wish to select."
      Top             =   1200
      Width           =   510
   End
   Begin VB.TextBox txtEXE 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "This value tells you which program owns the window. Don't worry if you do now know what it means - just drag, drop and click ""OK""."
      Top             =   2280
      Width           =   2475
   End
   Begin VB.TextBox txtWClass 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1860
      Width           =   2475
   End
   Begin VB.Label lblDrag 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drag the cross below over the window you wish to select and click 'OK'."
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
      Height          =   735
      Left            =   420
      TabIndex        =   5
      Top             =   420
      Width           =   4020
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWExe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Owner file :"
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
      Left            =   885
      TabIndex        =   2
      ToolTipText     =   "This value tells you which program owns the window. Don't worry if you do now know what it means - just drag, drop and click ""OK""."
      Top             =   2340
      Width           =   870
   End
   Begin VB.Label lblWClass 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Window class :"
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
      Left            =   780
      TabIndex        =   0
      ToolTipText     =   "This value tells you the window ""class"" name. Do not worry if it looks strange to you - you are not really supposed to see it."
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2595
      Left            =   180
      Top             =   180
      Width           =   4395
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectDrag 
         Caption         =   "Select by 'Drag and drop'"
      End
      Begin VB.Menu mnuPrograms 
         Caption         =   "Look for Programs in 'Program files' folder"
      End
      Begin VB.Menu mnuMydocuments 
         Caption         =   "Look for documents in 'My documents'"
      End
      Begin VB.Menu mnuBrowse 
         Caption         =   "I'm hardcore. Just let me browse"
      End
   End
End
Attribute VB_Name = "frmWindowCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private booValid As Boolean
Private booDone As Boolean
Private ExeName As String
Private pClassName As String

'Public strFileName As String
Public SelMode As String
Public CallingForm As Form
'Public CallingForm As frmAction

Dim booDragging As Boolean
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Sub Form_Activate()
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2

    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call zbCancel_Click
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    booDone = True
End Sub


Private Sub mnuBrowse_Click()

    SelMode = "Browse"

End Sub



Private Sub mnuMydocuments_Click()

    SelMode = "MyDocuments"
    
End Sub


Private Sub mnuPrograms_Click()

    SelMode = "ProgramFiles"
    
End Sub

Private Sub mnuSelectDrag_Click()

    SelMode = "Drag"
    
End Sub

Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    booDragging = True
    Set Screen.MouseIcon = picDrag.Picture
    Screen.MousePointer = vbCustom
    
End Sub


Private Sub picDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If booDragging Then
    
        Dim ptPoint As POINTAPI, hwnd As Long
        Dim lngParent As Long
        
        Rem - Get the window to work with
        Call GetCursorPos(ptPoint)
        hwnd = WindowFromPoint(ptPoint.X, ptPoint.Y)
        lngParent = GetParent(hwnd)
        If lngParent <> 0 Then hwnd = lngParent
        
        pClassName = ClassName(hwnd)
        
        txtWClass.Text = pClassName
        
        ExeName = GetExeFromHandle(hwnd)
        txtEXE.Text = GetFileName(ExeName) & "  (" & ExeName & ")"
    End If
    
End Sub


Private Sub picDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If booDragging Then
        Screen.MousePointer = vbDefault
        booDragging = False
    End If
    
End Sub



Private Sub zbCancel_Click()
    
    booValid = False
    booDone = True
    Me.Hide
    
End Sub


Private Sub zbOK_Click()
    
    Rem - Check the exe file is valid
    If Len(ExeName) > 0 Then
        booValid = True
        booDone = True
        Me.Hide
    Else
        Rem - Force the message box to be above....
        Me.ZOrder 1
        Me.Enabled = False
        Call ZenMB("Please select the program by dragging the cross over the window of the program you wish to select.", "OK")
        Me.Enabled = True
    End If
    
End Sub



Private Sub SetGraphics()

    Set zbCancel.Picture = zbOK.Picture

    Me.AutoRedraw = True
    Set Me.Picture = CallingForm.Picture
    Call TileMe(Me, LoadPicture(App.Path & "\Help\cloudsdark.jpg"))
    Me.AutoRedraw = False
    
    Me.Move 0.5 * (Screen.Width - Me.Width), 0.5 * (Screen.Height - Me.Height)
    
    lblDrag.ForeColor = COL_Zen
    lblWClass.ForeColor = COL_Zen
    lblWExe.ForeColor = COL_Zen

    Dim strTemp As String
    strTemp = "1. Start the program you wish to select." & vbCr
    strTemp = strTemp & "2. Drag the cross below over its visible window" & vbCr
    strTemp = strTemp & "3. Click 'OK' to accept the selection."
    lblDrag.Caption = strTemp

End Sub

Public Function SelectExe(ByRef strExe As String) As Boolean

    Call SetGraphics
    Me.Visible = True
    Do
        DoEvents
    Loop While Not booDone
    SelectExe = booValid
    If booValid Then strExe = ExeName

End Function
Public Function SelectClass(ByRef strClass As String) As Boolean

    Call SetGraphics
    Me.Visible = True
    Do
        DoEvents
    Loop While Not booDone
    SelectClass = booValid
    If booValid Then strClass = pClassName

End Function

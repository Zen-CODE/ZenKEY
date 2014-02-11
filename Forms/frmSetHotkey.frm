VERSION 5.00
Begin VB.Form frmSetHotKey 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "| ZenKEY - Capture the Hotkey |"
   ClientHeight    =   2205
   ClientLeft      =   11445
   ClientTop       =   4560
   ClientWidth     =   5175
   Icon            =   "frmSetHotkey.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   Begin VB.CommandButton zbDone 
      Caption         =   "Done"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton zbCancel 
      Caption         =   "Cancel"
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
      Left            =   3720
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton zbClear 
      Caption         =   "Clear"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1650
      Width           =   1335
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox cmbKey 
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
      Left            =   3420
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Set key combination that will fire the action. The number is brackets is the Windows code for this key, which you can ignore...."
      Top             =   1020
      Width           =   1455
   End
   Begin VB.ComboBox cmbShift 
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Set key combination that will fire the action"
      Top             =   1020
      Width           =   2115
   End
   Begin VB.Label lblPlus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Left            =   3240
      TabIndex        =   4
      ToolTipText     =   "Set key combination that will fire the action"
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press the Keys you wish to use, or select them below."
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
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Use this to start this program when ZenKEY starts (ignored on ZenKEY Restart)"
      Top             =   360
      Width           =   4005
   End
   Begin VB.Label lblHotkey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hotkeys"
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
      Left            =   210
      TabIndex        =   0
      ToolTipText     =   "Set key combination that will fire the action"
      Top             =   1080
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Index           =   3
      Left            =   120
      Top             =   900
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   120
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmSetHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public booValid As Boolean
Public booDone As Boolean
Public prop As clsZenDictionary
Public CallingForm As Form
Public EditIndex As Long
Private Const HK_MIN = 19
Private Const HK_NumPad5 = 12
Private booWinKey As Boolean

Private Sub SetGraphics()
    
    Me.Move 0.5 * (Screen.Width - Me.Width), 0.5 * (Screen.Height - Me.Height)
    Me.AutoRedraw = True
    Call TileMe(Me, LoadPicture(App.Path & "\Help\cloudsdark.jpg"))
    Me.AutoRedraw = False
    
Dim It As Control

    For Each It In Me.Controls
        If TypeOf It Is Label Then
            It.ForeColor = COL_Zen
        End If
    Next It
    
    Set zbClear.Picture = zbDone.Picture
    Set zbCancel.Picture = zbDone.Picture
    
End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Rem - Keypad 0 = 46 or 110
    Rem - Windows key = KeyCode 91. We have to workaround the fact that is does not appear a "Shift" key, but works like one
    Rem - For this, we use booWinKey
    
    txtText.Text = ""
    txtText.SetFocus
    If KeyCode = 91 Then booWinKey = True
    If Shift <> 0 Or booWinKey Then
        Dim lngShift As Long, lngTemp As Long
        Rem - Here 1 = Shift          , Windows shift = 4
        Rem - Here 4  = Alt         , Windows 1 = ALt
        Rem = Here 2 = Cntrl       , Window 2 =Control
        
        lngShift = 0
        lngTemp = Shift
        If booWinKey Then lngShift = 8
        If lngTemp >= 4 Then lngShift = lngShift + 1: lngTemp = lngTemp - 4
        If lngTemp >= 2 Then lngShift = lngShift + 2: lngTemp = lngTemp - 2
        If lngTemp >= 1 Then lngShift = lngShift + 4
    
        cmbShift.Text = HotKeys.ShiftValToStr(lngShift)
    End If
    
    'If KeyCode >= HK_MIN Then
    If (KeyCode >= HK_MIN) Or (KeyCode = HK_NumPad5) Then
        Call HKCombo_Display(KeyCode, cmbKey)
    ElseIf KeyCode = vbKeyReturn Then
        Call zbDone_Click
    End If
    
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight: KeyCode = 0
        Case vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd: KeyCode = 0
    End Select
    
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 91 Then booWinKey = False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    booDone = True
End Sub

Private Sub zbCancel_Click()
    
    booDone = True
    Me.Hide
    
    
End Sub




Private Sub zbClear_Click()
    
    cmbKey.ListIndex = 0
    cmbShift.ListIndex = 0
    If Set_Action() Then
        booValid = True
        booDone = True
        Me.Hide
    End If
    
End Sub

Private Sub zbDone_Click()
    
    If Set_Action() Then
        booValid = True
        booDone = True
        Me.Hide
    End If
    
End Sub





Public Sub Init()

    txtText.Move -100, -100
    Call SetGraphics
    Call Hotkeys_Init
    
    If Len(prop("ShiftKey")) <> 0 Then cmbShift.Text = HotKeys.ShiftValToStr(HotKeys.ShiftValue(prop("ShiftKey")))         ' COnversion required for compatability
    Dim strKey As String
    strKey = prop("Hotkey")
    If Len(strKey) <> 0 Then Call HKCombo_Display(strKey, cmbKey)

End Sub

Private Function Set_Action() As Boolean
On Error GoTo ErrorTrap

    Dim strHK As String, strShift As String
        
    If cmbKey.ListIndex > 0 Then strHK = HKCombo_GetValue(cmbKey)
    If cmbShift.ListIndex > 0 Then strShift = cmbShift.Text
    
    Rem - Check that the has a valid Hotkey
    If HKIsOkay(strShift, strHK, EditIndex) Then
        prop("Hotkey") = strHK
        prop("ShiftKey") = strShift
        Set_Action = True
    End If
    Exit Function

ErrorTrap:
    Call ZenMB(Err.Description, "OK")
    Err.Clear
    
End Function


Private Sub Hotkeys_Init()
Dim k As Integer
    
    With cmbShift
        Rem - The listindex corresponds directly to the ShiftValue -1! How kewl!
        .Clear
        .AddItem "<None>"
        For k = 1 To 15
            .AddItem HotKeys.ShiftValToStr(k)
        Next k
        .ListIndex = 0
    End With
    
    Call HKCombo_Init(cmbKey)
    
End Sub



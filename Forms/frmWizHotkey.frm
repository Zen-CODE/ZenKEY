VERSION 5.00
Begin VB.Form frmWizHotkey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "| ZenKEY - Configure action |"
   ClientHeight    =   4500
   ClientLeft      =   7710
   ClientTop       =   5115
   ClientWidth     =   8955
   Icon            =   "frmWizHotkey.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   Begin VB.TextBox Text1 
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
      Left            =   840
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2040
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
      Left            =   7500
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Set key combination that will fire the action.  The number is brackets is the Windows code for this key, which you can ignore...."
      Top             =   1350
      Width           =   1215
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
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Set key combination that will fire the action"
      Top             =   1350
      Width           =   2115
   End
   Begin VB.CommandButton zbBack 
      Caption         =   "<< Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
   Begin VB.CommandButton zbNext 
      Caption         =   "Next >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3780
      TabIndex        =   4
      Top             =   3945
      Width           =   1455
   End
   Begin VB.CommandButton zbCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   3945
      Width           =   1455
   End
   Begin VB.Label lblPlus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5910
      TabIndex        =   7
      ToolTipText     =   "Set key combination that will fire the action"
      Top             =   1380
      Width           =   90
   End
   Begin VB.Label lblHotkey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1. This action should happen when I press.............................................................."
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
      TabIndex        =   0
      ToolTipText     =   "Set key combination that will fire the action"
      Top             =   1440
      Width           =   6825
   End
End
Attribute VB_Name = "frmWizHotkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Const HK_MIN = 32
Dim booLoading As Boolean


Private Sub Form_Activate()
    If Text1.Visible Then Text1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        If Shift <> 0 Then
            Dim lngShift As Long, lngTemp As Long
            Rem - Here 1 = Shift          , Windows shift = 4
            Rem - Here 4  = Alt         , Windows 1 = ALt
            Rem = Here 2 = Cntrl       , Window 2 =Control
            
            lngShift = 0
            lngTemp = Shift
            If lngTemp >= 4 Then lngShift = 1: lngTemp = lngTemp - 4
            If lngTemp >= 2 Then lngShift = lngShift + 2: lngTemp = lngTemp - 2
            If lngTemp >= 1 Then lngShift = lngShift + 4
        
            cmbShift.Text = HotKeys.ShiftValToStr(lngShift)
        End If
        
        If KeyCode >= HK_MIN Then
            Call HKCombo_Display(KeyCode, cmbKey)
        End If
        
        Select Case KeyCode
            Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight: KeyCode = 0
            Case vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd: KeyCode = 0
        End Select
        
    Text1.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Rem - Prevent alt from removing focus from active control.....

    On Error Resume Next
    Text1.SetFocus
    
End Sub




Private Sub Form_Load()

    Text1.left = -100
    booLoading = True
    Call Hotkeys_Init
    booLoading = False
        
    Rem - Set objects for Zenkey Config compatibility
    Set MainForm = Me


End Sub

Private Sub zbBack_Click()

    Call ZW_Next("Previous")
    
End Sub

Private Sub zbCancel_Click()
    
    End
    
End Sub









Private Function Set_Hotkeys() As Boolean
Dim i As Long, j As Long

On Error GoTo ErrorTrap

    Rem - Check that the item is valid
    
    Rem - Check Hotkeys
    Dim strHK As String, strShift As String
    
    If cmbKey.ListIndex > 0 Then strHK = HKCombo_GetValue(cmbKey)
    If cmbShift.ListIndex > 0 Then strShift = cmbShift.Text
    If HKIsOkay(strShift, strHK, -1) Then
        ZW_NewItem("Hotkey") = strHK
        ZW_NewItem("ShiftKey") = strShift
        Set_Hotkeys = True
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


Private Sub zbNext_Click()
    If Set_Hotkeys Then Call ZW_Next("Next")
End Sub



Public Sub Init()

    
    If Len(ZW_NewItem("ShiftKey")) <> 0 Then cmbShift.Text = HotKeys.ShiftValToStr(HotKeys.ShiftValue(ZW_NewItem("ShiftKey")))         ' COnversion required for compatability
    Dim strKey As String
    strKey = ZW_NewItem("Hotkey")
    If Len(strKey) > 0 Then Call HKCombo_Display(Val(strKey), cmbKey)
    

End Sub

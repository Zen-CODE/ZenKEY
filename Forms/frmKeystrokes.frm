VERSION 5.00
Begin VB.Form frmKeystrokes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "| ZenKEY - Keystroke recording and playback |"
   ClientHeight    =   3375
   ClientLeft      =   8520
   ClientTop       =   3000
   ClientWidth     =   9855
   ClipControls    =   0   'False
   Icon            =   "frmKeystrokes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   657
   Begin VB.CommandButton zbClearLast 
      Caption         =   "Clear last"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8340
      TabIndex        =   24
      Top             =   2910
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear all"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6690
      TabIndex        =   23
      Top             =   2910
      Width           =   1335
   End
   Begin VB.CommandButton cmdCapture 
      Caption         =   "Start recording"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5040
      TabIndex        =   22
      Top             =   2910
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   330
      Left            =   3480
      TabIndex        =   21
      Top             =   2910
      Width           =   1335
   End
   Begin VB.CommandButton cmdDone 
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
      Height          =   330
      Left            =   1830
      TabIndex        =   20
      Top             =   2910
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpenCapture 
      Caption         =   "Record >>"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   180
      TabIndex        =   19
      Top             =   2910
      Width           =   1335
   End
   Begin VB.CommandButton cmdFire 
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton zbHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4080
      TabIndex        =   17
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Insert pause"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5220
      TabIndex        =   16
      ToolTipText     =   "Clicking here will record a pause inbetween keypresses so that Windows can do some processing if neccesary."
      Top             =   2040
      Width           =   1995
   End
   Begin VB.CheckBox chkKey 
      Caption         =   "Windows"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Clicking here will record a 'Windows key' keypress without pressing the 'Windows' key."
      Top             =   1620
      Width           =   1275
   End
   Begin VB.CheckBox chkKey 
      Caption         =   "Control"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Clicking here will record a 'Control' keypress without pressing the 'Control' key."
      Top             =   1620
      Width           =   1275
   End
   Begin VB.CheckBox chkKey 
      Caption         =   "Alt"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Clicking here will record an 'Alt' keypress without pressing the 'Alt' key."
      Top             =   1620
      Width           =   1275
   End
   Begin VB.CheckBox chkKey 
      Caption         =   "Shift"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Clicking here will record a 'Shift' keypress without pressing the 'Shift' key."
      Top             =   1380
      Width           =   1275
   End
   Begin VB.CheckBox chkKey 
      Caption         =   "Tab"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Clicking here will record a 'Tab' keypress without pressing the 'Tab' key."
      Top             =   1380
      Width           =   1275
   End
   Begin VB.CheckBox chkRelease 
      Caption         =   "Release keys"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8220
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Clicking here will make ZenKEY release any keys it has recorded as being held down."
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CheckBox chkHold 
      Caption         =   "Hold keys"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8220
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Clicking here whilst keys are pressed will prevent ZenKEY from recording their release when you let go of them."
      Top             =   900
      Width           =   1215
   End
   Begin VB.TextBox txtSeconds 
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
      Height          =   285
      Left            =   7500
      TabIndex        =   5
      Text            =   "0.5"
      ToolTipText     =   "The recorded a pause will pauset for this amount of time."
      Top             =   2010
      Width           =   450
   End
   Begin VB.CheckBox chkAllowRepeats 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Depressed keys fire repeated keypresses"
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
      Left            =   5160
      TabIndex        =   3
      ToolTipText     =   "Holding a key down usually results in repeated keystrokes. This option preserves this behaviour whilst recording keystrokes."
      Top             =   2400
      Width           =   3975
   End
   Begin VB.ComboBox cmbSeconds 
      Height          =   315
      ItemData        =   "frmKeystrokes.frx":058A
      Left            =   2670
      List            =   "frmKeystrokes.frx":05A3
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Wait for this amount of time before simulating the keystrokes."
      Top             =   2400
      Width           =   615
   End
   Begin VB.ListBox lstKeys 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   180
      TabIndex        =   0
      ToolTipText     =   "This box contains a list of all the keystrokes that will be simulated."
      Top             =   780
      Width           =   4605
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9960
      Top             =   720
   End
   Begin VB.Label lblKeyState 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current key state : "
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
      Height          =   495
      Left            =   5220
      TabIndex        =   9
      ToolTipText     =   "This box displays the state of the keys as being recorded by ZenKEY."
      Top             =   900
      Width           =   2550
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRecOpt 
      BackStyle       =   0  'Transparent
      Caption         =   "Recording panel"
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
      Left            =   5160
      TabIndex        =   7
      Top             =   300
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   5040
      Top             =   150
      Width           =   4635
   End
   Begin VB.Label lblKeystrokes 
      BackStyle       =   0  'Transparent
      Caption         =   "Keystroke sequence"
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
      Left            =   300
      TabIndex        =   6
      Top             =   300
      Width           =   2655
   End
   Begin VB.Label lblof 
      BackStyle       =   0  'Transparent
      Caption         =   "of              second(s)"
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
      Left            =   7260
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblIn 
      BackStyle       =   0  'Transparent
      Caption         =   "these keystokes in                second(s)"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   2460
      Width           =   2955
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   180
      Top             =   150
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1995
      Index           =   3
      Left            =   5040
      Top             =   780
      Width           =   4635
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   1
      Left            =   180
      Top             =   2340
      Width           =   4620
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuInsert2 
         Caption         =   "Insert keystrokes"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove key"
      End
   End
   Begin VB.Menu mnuRecord 
      Caption         =   "Record"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add keystrokes"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert keystrokes"
      End
   End
End
Attribute VB_Name = "frmKeystrokes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Dim KeyArray() As Integer
Dim KeyDown() As Boolean
Dim KeyCount As Long
Dim lngCurShift As Long
Dim booCapture As Boolean
Dim lngLastKey As Long
Dim lngTick As Long
Rem - Public properies --
Public prop As clsZenDictionary
Public booValid As Boolean
Public CallingForm As Form
Rem - Public properies --

Dim booHold As Boolean
Dim colHoldKeys As Collection
Dim booLoading As Boolean

Rem - For insert
Dim lngInsIndex As Long
Dim InsKeyArray()
Dim InsKeyDown()
Dim InsKeyCount As Long
Private Sub chkKey_Click(Index As Integer)

    If booLoading Or (Not booCapture) Then Exit Sub
    
    booLoading = True
    Dim booKeyDown As Boolean
    booKeyDown = CBool(chkKey(Index).Value = 1)
    Select Case Index
        Case 0 ' 9 - Tab
            Call Capture_AddKey(9, booKeyDown)
        Case 1 '16 ' Shift
            Call Capture_AddKey(16, booKeyDown)
        Case 2 '17 ' Control
            Call Capture_AddKey(17, booKeyDown)
        Case 3 ' 91 ' Windows
            Call Capture_AddKey(91, booKeyDown)
        Case 4 '18 ' Alt
            Call Capture_AddKey(18, booKeyDown)
    End Select
    booLoading = False
    
End Sub


Private Sub chkRelease_Click()
    
    If chkRelease.Value = 1 Then
        If booCapture Then
            booLoading = True
            chkHold.Value = 0
            chkRelease.Value = 0
            Call Capture_ReleaseKeys
            booLoading = False
        End If
    End If
    
End Sub

Private Sub cmdCancel_Click()
    
    Me.Hide
    Unload Me
    
End Sub


Private Sub cmdDone_Click()

    Call Action_Set
    booValid = True
    Me.Hide
    
End Sub



Private Sub cmdFire_Click()

    If booCapture Then Call cmdCapture_Click
    DoEvents
    If KeyCount < 1 Then
        Call ZenMB("Please record a series of Keystrokes by clicking on 'Record >>' and then 'Start recording'.", "OK")
    Else
        Timer1.Interval = 1000 * Val(cmbSeconds.Text)
        Timer1.Enabled = True
        Me.Enabled = False
    End If
    
End Sub

Private Sub cmdCapture_Click()
Dim booOfferInsert As Boolean

    Rem - Records where we should begin the insert
    If (KeyCount > 1) Then
        If Not booCapture Then
            If lstKeys.ListIndex > -1 Then
                booOfferInsert = True
            End If
        End If
    End If
    
    If Not booOfferInsert Then
        Call mnuAdd_Click
    Else
        Call PopupMenu(mnuRecord)
    End If
    
    
End Sub

Private Sub cmdClear_Click()

    If KeyCount > 0 Then
        If 0 = ZenMB("This will erase all your recorded keystrokes. Are you sure you wish to do this?", "Yes", "No") Then
            Call Capture_Clear
        End If
    End If
    
End Sub

Private Sub chkHold_Click()
    
    booHold = CBool(chkHold.Value = 1)
    If Not booLoading Then
        If Not booHold Then Call Capture_ReleaseKeys
    End If
    
End Sub

Private Sub cmdOpenCapture_Click()
    If cmdOpenCapture.Caption = "Record >>" Then
        cmdOpenCapture.Caption = "Record <<"
        Me.Width = 9950 '9690
        Rem - Ensure we don't go off the screen
        If Me.left + Me.Width > Screen.Width Then Me.left = Screen.Width - Me.Width
        
    Else
        cmdOpenCapture.Caption = "Record >>"
        Me.Width = 5100 '5235
    End If
    cmdOpenCapture.Refresh
    
End Sub

Private Sub cmdPause_Click()

    If booCapture Then Call Capture_AddKey(-CLng(Val(txtSeconds.Text) * 1000), False)
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If booCapture Then
        
        'Debug.Print "Key down = " & CStr(KeyCode)

        Rem - Do not fire the keyup event if the keys where down and 'Hold keys' was clicked
        If booHold Then
            Dim lngVal As Long
            On Error Resume Next
            lngVal = colHoldKeys.Item(CStr(KeyCode))
            If lngVal = KeyCode Then Exit Sub
        End If
        
        
        Rem - Prevent held keys from firing multiple keydowns?
        If lngLastKey <> 0 Then
            If lngLastKey = KeyCode Or lngLastKey = Shift Then
                If chkAllowRepeats.Value <> 1 Then Exit Sub
            End If
        End If
        
        Rem - Add the key to the array
        Dim lngKey As Long
        If KeyCode = 0 Then
            If Shift <> lngCurShift Then
                lngKey = Shift
                lngCurShift = Shift
                lngLastKey = Shift
            Else
                Exit Sub
            End If
        Else
            lngKey = KeyCode
            lngLastKey = KeyCode
        End If
        Call Capture_AddKey(lngKey, True)
        
        If KeyCode = 18 Then KeyCode = 0 ' For some funny reason, Alt does not fire a second time without this ?
        
    End If
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If booCapture Then
        Dim lngKey As Long
    
        Rem - Do not fire the keyup event if the keys where down and 'Hold keys' was clicked
        If booHold Then
            Dim lngVal As Long
            On Error Resume Next
            lngVal = colHoldKeys.Item(CStr(KeyCode))
            If lngVal = KeyCode Then Exit Sub
        End If
        
        Rem - Record the Keyup
        If KeyCode = 0 Then lngKey = Shift Else lngKey = KeyCode
        Call Capture_AddKey(lngKey, False)
        
        lngLastKey = 0
        
    End If
    
End Sub


Private Sub Form_Load()
    Set colHoldKeys = New Collection
End Sub

Private Sub lstKeys_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstKeys.ListIndex > -1 Then
        If Button = 2 Then
            mnuInsert2.Enabled = Not booCapture
            Call PopupMenu(mnuFile)
        End If
    End If

End Sub

Private Sub mnuAdd_Click()

    booCapture = Not booCapture
    If booCapture Then
        Call Capture_Open
    Else
        Call Capture_Close
    End If
    Call Focus_Set
    txtSeconds.Enabled = Not booCapture

End Sub

Private Sub mnuInsert_Click()

    Call Insert_Start
    
End Sub

Private Sub mnuInsert2_Click()
    Call mnuInsert_Click
End Sub

Private Sub mnuRemove_Click()
    
    If lstKeys.ListIndex > -1 Then Call Capture_RemoveKey(lstKeys.ListIndex)
    

End Sub


Private Sub Timer1_Timer()

    Timer1.Enabled = False
    Call Action_Set
       
    Call TestAction(zenDic("Class", "Keystrokes", "Action", prop("Action")))
    Me.Enabled = True
    
End Sub




Private Sub Capture_AddKey(ByVal KeyVal As Long, ByVal booKeyDown As Boolean)
Dim strItem As String
    
    Rem - Add the Key to the array
    ReDim Preserve KeyArray(0 To KeyCount)
    ReDim Preserve KeyDown(0 To KeyCount)
    
    Rem - Add to the listbox
    Call Display_Add(KeyVal, booKeyDown, KeyCount + 1)
    
    Rem - Add key to the array
    KeyArray(KeyCount) = KeyVal
    KeyDown(KeyCount) = booKeyDown
    KeyCount = KeyCount + 1
        
    Rem - If keep track of which keys are down or up ...
    If Not booHold Then
        If KeyVal > 0 Then
            On Error Resume Next
            If booKeyDown Then
                colHoldKeys.Add Item:=CStr(KeyVal), key:=CStr(KeyVal)
            Else
                colHoldKeys.Remove (CStr(KeyVal))
            End If
        End If
    End If
    
    Rem - Display the current keystate
    Dim strKeys As String, strKey As Variant
    For Each strKey In colHoldKeys
        strKeys = strKeys & HotKeys.Keyname(Val(strKey)) & " + "
    Next strKey
    If Len(strKeys) > 0 Then
        chkHold.Enabled = True
        'chkHold.Value = 0
    Else
        chkHold.Value = 0
        chkHold.Enabled = False
    End If

    Rem -                   If the kys are held, they KeyVal is not added to the collection
    If booHold Then If booKeyDown Then strKeys = strKeys & HotKeys.Keyname(KeyVal) & " + "
    If Len(strKeys) > 0 Then
        strKeys = left(strKeys, Len(strKeys) - 3)
    Else
        strKeys = "none"
    End If
    Call Display_KeyState(strKeys)
    If Me.Enabled Then Call Focus_Set
    
End Sub


Private Sub Capture_Close()
Dim k As Long

    cmdCapture.Caption = "Start Recording"
    chkHold.Value = 0
    chkHold.Enabled = False
    
    Call Capture_ReleaseKeys
    chkRelease.Enabled = False
    Call Display_KeyState(" - not recording - ")
    For k = 0 To 4
        chkKey(k).Enabled = False
    Next k
    cmdPause.Enabled = False
    
    If lngInsIndex > -1 Then Insert_End

End Sub

Private Sub Action_Set()
Dim k As Long
Dim strAction As String
    
    Rem - Format of action
    Rem - KeyCount>KeyPressed1>KeyDown>KeyPressed2>KeyDown>...[-ms Pause>False].....

    Rem - Encryption scheme -----
    Rem - KeyCount>KeyPressed1+71+1>Odd = KeyDown, Even = KeyUp>KeyPressed2+71+ 2>Odd = KeyDown, Even = KeyUp>...[-ms Pause>False].....
    Rem - KeyCount - No change
    Rem - KeyPressed - Key value = KeyValue + 70 + KeyNumber
    Rem - KeyDown - Even = "Y", Odd = "N"
    Rem - Pause - No Change
    Rem - Encryption scheme -----
    
    strAction = CStr(KeyCount) & ">"
    For k = 1 To KeyCount
        'strAction = strAction & CStr(KeyArray(k - 1)) & ">"
        'strAction = strAction & IIf(CStr(KeyDown(k - 1)), "Y", "N") & ">"
        
        Rem - Encrypt the key values
        Select Case KeyArray(k - 1)
            Case Is < 0 ' Pause - No change
                strAction = strAction & CStr(KeyArray(k - 1)) & ">"
            Case Else
                strAction = strAction & CStr(KeyArray(k - 1) + 70 + k) & ">"
        End Select
        strAction = strAction & CStr(2 * FnRd(0, 4) + IIf(KeyDown(k - 1), 0, 1)) & ">"
        
    Next k
    prop("Action") = strAction
    
End Sub

Public Sub Capture_Open()

    cmdCapture.Caption = "Stop Recording"
    Call Display_KeyState("none")
    chkRelease.Enabled = True
    cmdPause.Enabled = True
    
    Dim k As Long
    For k = 0 To 4
        chkKey(k).Enabled = True
    Next k
    
End Sub





Private Sub Focus_Set()
    lstKeys.SetFocus
    If lstKeys.ListCount > 0 Then lstKeys.ListIndex = lstKeys.ListCount - 1
    
End Sub


Private Sub Capture_ReleaseKeys()
Dim colKeysDown As Collection
Dim k As Long
    Set colKeysDown = New Collection
    
    Rem - This is a dirty trick, but it works & is pretty damn efficiently as collections cannot hold duplciate keys...
    On Error Resume Next
    
    Dim max As Long
    max = KeyCount - 1
    For k = 0 To max
        If KeyArray(k) > 0 Then
            If KeyDown(k) Then
                colKeysDown.Add Item:=CStr(KeyArray(k)), key:=CStr(KeyArray(k))
            Else
                colKeysDown.Remove (CStr(KeyArray(k)))
            End If
        End If
    Next k
    
    'UnLoad colHoldKeys
    'Set colHoldKeys = New Collection
    While colHoldKeys.Count > 0
        colHoldKeys.Remove colHoldKeys.Count
    Wend

    For k = colKeysDown.Count To 1 Step -1
        Call Capture_AddKey(colKeysDown.Item(k), False)
    Next k
    lngLastKey = 0 ' Prevent Form_Keydown from ignoring the next keydown
    Call Display_KeyState("none")
    If Me.Enabled Then Call Focus_Set

End Sub


Private Sub Action_Load()
Dim strAction As String
Dim k As Long
Dim lngPos As Long
        
    Rem - Format of action
    Rem - KeyCount>KeyPressed1>KeyDown>KeyPressed2>KeyDown>...[-ms Pause>False].....

    Rem - Encryption scheme -----
    Rem - KeyCount>KeyPressed1+71+1>Odd = KeyDown, Even = KeyUp>KeyPressed2+71+ 2>Odd = KeyDown, Even = KeyUp>...[-ms Pause>False].....
    Rem - KeyCount - No change
    Rem - KeyPressed - Key value = KeyValue + 70 + KeyNumber
    Rem - KeyDown - Even = "Y", Odd = "N"
    Rem - Pause - No Change
    Rem - Encryption scheme -----
    
    On Error GoTo ErrorTrap:
    
    strAction = prop("Action")
    KeyCount = Val(strAction)
    
    If (KeyCount > 0) And (InStr(strAction, ">") > 0) Then
        lstKeys.Clear
        ReDim KeyArray(0 To KeyCount - 1)
        ReDim KeyDown(0 To KeyCount - 1)
        For k = 0 To KeyCount - 1
            lngPos = InStr(strAction, ">")
            strAction = Mid(strAction, lngPos + 1)
            KeyArray(k) = Val(strAction)
            Rem - Use encryption
            If KeyArray(k) > 0 Then KeyArray(k) = KeyArray(k) - 71 - k
            
            lngPos = InStr(strAction, ">")
            strAction = Mid(strAction, lngPos + 1)
            'KeyDown(k) = (Left(strAction, 1) = "Y")
            Rem - Use encryption
            KeyDown(k) = CBool(Val(strAction) Mod 2 = 0) ' 0 = True, all else False
            
            
            Call Display_Add(KeyArray(k), KeyDown(k), k + 1)
        Next k
    End If
    Exit Sub
    
ErrorTrap:
    Rem - Perhaps we are trying to load an action that is not a series of keystrokes?
    Err.Clear
End Sub

Private Sub Display_Add(ByVal KeyVal As Long, ByVal KeyDown As Boolean, ByVal KeyCount As Long)
Dim strItem As String

    If KeyVal < 0 Then
        strItem = "Pause - " & CStr(Abs(KeyVal)) & " milliseconds"
    ElseIf KeyDown Then
        'strItem = "Keydown - " & CStr(KeyVal)
        'strItem = "Press the key - " & Hotkeys.Keyname(KeyVal)
        strItem = "'" & HotKeys.Keyname(KeyVal) & "' key down"
    Else
        'strItem = "Keyup - " & CStr(KeyVal)
        'strItem = "Release the key - " & Hotkeys.Keyname(KeyVal)
        strItem = "'" & HotKeys.Keyname(KeyVal) & "' key up"
    End If
    strItem = strItem & " - (No. " & CStr(KeyCount) & ")"
    lstKeys.AddItem strItem

End Sub

Public Sub Display_KeyState(ByVal strKeys As String)

    lblKeyState.Caption = "Current key state" & vbCr & strKeys
    
    Rem - Display the current keystate in the buttons

    Rem - Tab = 9
    Rem - Control = 17
    Rem - Windows = 91
    Rem - Alt = 18
    Rem ------------------------------------------
    Rem - chkKey(0) = Tab
    Rem - chkKey(1) = Shift
    Rem - chkKey(2) = Control
    Rem - chkKey(3) = Windows
    Rem - chkKey(4) = Alt
    
    Dim strKey As Variant
    Dim booDown(0 To 4) As Boolean
    Dim lngKey As Long

    Rem - Determine key status
    For Each strKey In colHoldKeys
        lngKey = Val(strKey)
        Select Case lngKey
            Case 9 ' Tab
                booDown(0) = True
            Case 16 ' Shift
                booDown(1) = True
            Case 17 ' Control
                booDown(2) = True
            Case 91 ' Windows
                booDown(3) = True
            Case 18 ' Alt
                booDown(4) = True
        End Select
    Next strKey
    Rem - Show key status
    booLoading = True
    For lngKey = 0 To 4
        chkKey(lngKey).Value = IIf(booDown(lngKey), 1, 0)
    Next lngKey
    booLoading = False
    
End Sub

Public Sub Initialise()
    
    lngInsIndex = -1
    cmbSeconds.ListIndex = 1
    'Set KeyStrokes = New clsKeyStroke
    Call SetGraphics
    Call Display_KeyState(" - not recording - ")
    Call Action_Load
    
End Sub

Private Sub Capture_RemoveKey(ByVal Index As Long)
Dim booKeyDown As Boolean
Dim lngKey As Long
Dim lngStep As Long
Dim lngEnd As Long
Dim k As Long
Dim lngIndex2 As Long

    Rem - Determine the key values, and whether we go up or down?
    lngKey = KeyArray(Index)
    booKeyDown = KeyDown(Index)
    If booKeyDown Then
        lngStep = 1
        lngEnd = KeyCount - 1
    Else
        lngStep = -1
        lngEnd = 0
    End If
    
    Rem - See if we can locate a matching key up/ keydown to remove
    lngIndex2 = -1
    If lngKey > 0 Then
        For k = Index + lngStep To lngEnd Step lngStep
            If KeyArray(k) = lngKey Then
                Rem - Found ?
                If KeyDown(k) = (Not booKeyDown) Then lngIndex2 = k
                Exit For
            End If
        Next k
    End If
    
    Rem - Remove they key, or both
    lngStep = 0
    For k = 0 To KeyCount - 1
        Select Case k
            Case Index, lngIndex2
                Rem - Skip these keys
            Case Else
                KeyArray(lngStep) = KeyArray(k)
                KeyDown(lngStep) = KeyDown(k)
                lngStep = lngStep + 1
        End Select
    Next k
    If lngStep > 0 Then
        ReDim Preserve KeyArray(0 To lngStep - 1)
        ReDim Preserve KeyDown(0 To lngStep - 1)
    End If
    KeyCount = lngStep
    
    Rem - Seems like this key is being held. Remove it
    If booKeyDown And lngIndex2 < 0 Then
        For k = 0 To colHoldKeys.Count - 1
            If lngKey = Val(colHoldKeys.Item(k)) Then colHoldKeys.Remove (CStr(lngKey))
        Next k
    End If
    
    Rem - Re-populate the list
    lstKeys.Clear
    For k = 0 To KeyCount - 1
        Call Display_Add(KeyArray(k), KeyDown(k), k + 1)
    Next k
    If Me.Enabled Then Call Focus_Set

End Sub




Private Sub zbClearLast_Click()
    If lstKeys.ListCount > 0 Then Call Capture_RemoveKey(lstKeys.ListCount - 1)

End Sub


Private Sub zbHelp_Click()
    Call ShellExe(App.Path & "\Help\Keystrokes.htm")
End Sub





Private Function FnRd(ByVal Low As Long, ByVal high As Long) As Long
    FnRd = Int((high * Rnd) + Low)
End Function

Private Sub Capture_Clear()
    KeyCount = 0
    lstKeys.Clear
    Erase KeyArray()
    Erase KeyDown()
    lstKeys.SetFocus
    Call Capture_ReleaseKeys
End Sub

Public Sub SetGraphics()


    Me.AutoRedraw = True
    Me.Width = 9630 '9690
    Call TileMe(Me, LoadPicture(App.Path & "\Help\cloudsdark.jpg"))
    Me.Width = 5100 '5235
    
    Dim It As Object
    For Each It In Me.Controls
        If TypeOf It Is Label Then
            If It.Font.Bold Then It.ForeColor = COL_Zen
        End If
    Next It
    
    Set cmdDone.Picture = cmdOpenCapture.Picture
    Set cmdCancel.Picture = cmdOpenCapture.Picture
    Set cmdCapture.Picture = cmdOpenCapture.Picture
    Set cmdClear.Picture = cmdOpenCapture.Picture
    Set zbClearLast.Picture = cmdOpenCapture.Picture
    Set zbHelp.Picture = cmdFire.Picture
    
    Me.AutoRedraw = False

End Sub

Private Sub Insert_Start()
Dim k As Long
    
    Rem - Copy the current key strokes into the Ins arrays
    lngInsIndex = lstKeys.ListIndex
    InsKeyCount = KeyCount

    ReDim InsKeyArray(0 To KeyCount - 1)
    ReDim InsKeyDown(0 To KeyCount - 1)
    For k = 0 To KeyCount - 1
        InsKeyArray(k) = KeyArray(k)
        InsKeyDown(k) = KeyDown(k)
    Next k
    
    Call Capture_Clear
    Call mnuAdd_Click
    

End Sub

Private Sub Insert_End()
Dim k As Long
Dim NewKeyArray() As Integer
Dim NewKeyDown() As Boolean

    Rem - Create a new array
    ReDim NewKeyArray(0 To InsKeyCount + KeyCount - 1)
    ReDim NewKeyDown(0 To InsKeyCount + KeyCount - 1)

    Rem - Copy the first part in (before the insert)
    For k = 0 To lngInsIndex - 1
        NewKeyArray(k) = InsKeyArray(k)
        NewKeyDown(k) = InsKeyDown(k)
    Next k

    Rem - Now insert the new keystrokes
    For k = lngInsIndex To lngInsIndex + KeyCount - 1
        NewKeyArray(k) = KeyArray(k - lngInsIndex)
        NewKeyDown(k) = KeyDown(k - lngInsIndex)
    Next k

    Rem - Now copy the end part of original sequence
    For k = lngInsIndex + KeyCount To InsKeyCount + KeyCount - 1
        NewKeyArray(k) = InsKeyArray(k - KeyCount)
        NewKeyDown(k) = InsKeyDown(k - KeyCount)
    Next k

    Rem - Now copy eveything back into the original array
    lstKeys.Clear
    KeyCount = InsKeyCount + KeyCount
    ReDim KeyArray(0 To KeyCount - 1)
    ReDim KeyDown(0 To KeyCount - 1)
    For k = 0 To KeyCount - 1
        KeyArray(k) = NewKeyArray(k)
        KeyDown(k) = NewKeyDown(k)
        Call Display_Add(KeyArray(k), KeyDown(k), k + 1)
    Next k
    lngInsIndex = -1



'Dim k As Long
'
'    Rem - Shift the original array up to make space for the insertion
'    ReDim Preserve InsKeyArray(0 To InsKeyCount + KeyCount - 1)
'    ReDim Preserve InsKeyDown(0 To InsKeyCount + KeyCount - 1)
'    For k = lngInsIndex To lngInsIndex + KeyCount - 1 / lngInsIndex + KeyCount
'        InsKeyArray(k + KeyCount) = InsKeyArray(k)
'        InsKeyDown(k + KeyCount) = InsKeyDown(k)
'    Next k
'
'    Rem - Now insert the new keystrokes
'    For k = 0 To KeyCount - 1
'        InsKeyArray(lngInsIndex + k) = KeyArray(k)
'        InsKeyDown(lngInsIndex + k) = KeyDown(k)
'    Next k
'
'    Rem - Now copy eveything back into the original array
'    lstKeys.Clear
'    KeyCount = InsKeyCount + KeyCount
'    ReDim KeyArray(0 To KeyCount - 1)
'    ReDim KeyDown(0 To KeyCount - 1)
'    For k = 0 To KeyCount - 1
'        KeyArray(k) = InsKeyArray(k)
'        KeyDown(k) = InsKeyDown(k)
'        Call Display_Add(KeyArray(k), KeyDown(k), k + 1)
'    Next k
'    lngInsIndex = -1
'
End Sub

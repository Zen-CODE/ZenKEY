VERSION 5.00
Begin VB.Form frmAction 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "| Configure action |"
   ClientHeight    =   5295
   ClientLeft      =   9555
   ClientTop       =   5025
   ClientWidth     =   5250
   Icon            =   "frmAction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   Begin VB.CommandButton zbSetHK 
      Caption         =   "Set"
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
      Left            =   4560
      TabIndex        =   27
      Top             =   2580
      Width           =   495
   End
   Begin VB.CommandButton zbTest 
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
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   3120
      Width           =   1455
   End
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
      TabIndex        =   25
      Top             =   3120
      Width           =   1455
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
      Left            =   3600
      TabIndex        =   24
      Top             =   3120
      Width           =   1455
   End
   Begin VB.ComboBox cmbRunIn 
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
      ItemData        =   "frmAction.frx":0D0F
      Left            =   1080
      List            =   "frmAction.frx":0D1C
      Style           =   2  'Dropdown List
      TabIndex        =   17
      ToolTipText     =   "Describes the type of action to be performed"
      Top             =   4620
      Width           =   3855
   End
   Begin VB.CommandButton cmdKeyStrokes 
      Caption         =   "Edit Keystrokes"
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
      Left            =   3600
      TabIndex        =   21
      ToolTipText     =   "Click here to alter the Keysequence you would like to simulate."
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtKeyStrokes 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Keystokes"
      ToolTipText     =   "Your computer will behave as if you had pressed these keys."
      Top             =   4080
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.ComboBox cmbStartup 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   15
      ToolTipText     =   "Use this to start this program when ZenKEY starts (ignored on ZenKEY Restart)"
      Top             =   3660
      Width           =   2835
   End
   Begin VB.CheckBox chkBringToForeground 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bring to foreground if active"
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
      Left            =   2475
      TabIndex        =   4
      ToolTipText     =   $"frmAction.frx":0D63
      Top             =   1200
      Width           =   2475
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   4620
      TabIndex        =   9
      ToolTipText     =   "Browse for the folder which contains these files or folders."
      Top             =   1500
      Width           =   315
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1080
      TabIndex        =   3
      Text            =   "Select the file/applet to be opened."
      ToolTipText     =   "Browse for or type in the name of a folder. For 'Windows Special Folders', just enter the number."
      Top             =   1470
      Width           =   3555
   End
   Begin VB.TextBox txtCaption 
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
      Left            =   1080
      TabIndex        =   0
      Text            =   "Enter the caption here"
      ToolTipText     =   "Set caption on the item as it appears in the menu"
      Top             =   210
      Width           =   2895
   End
   Begin VB.CheckBox chkEnabled 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enabled"
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
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "Enable or disable this group or item. This will prevent the item/group from showing in 'ZenKEY'."
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox cmbSpecial 
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
      ItemData        =   "frmAction.frx":0DF4
      Left            =   2640
      List            =   "frmAction.frx":0DFB
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Select the special folder to be opened."
      Top             =   1920
      Width           =   2355
   End
   Begin VB.TextBox txtParameter 
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
      TabIndex        =   5
      ToolTipText     =   "Enter the parameter."
      Top             =   1920
      Width           =   1275
   End
   Begin VB.ComboBox cmbAction 
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Describes the type of action to be performed"
      Top             =   825
      Width           =   3555
   End
   Begin VB.ComboBox cmbActItem 
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
      ItemData        =   "frmAction.frx":0E0B
      Left            =   1440
      List            =   "frmAction.frx":0E12
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   "Select the action to be taken"
      Top             =   1470
      Width           =   3555
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
      Left            =   3300
      Style           =   2  'Dropdown List
      TabIndex        =   14
      ToolTipText     =   "Set key combination that will fire the action. The number is brackets is the Windows code for this key, which you can ignore...."
      Top             =   2580
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   13
      ToolTipText     =   "Set key combination that will fire the action"
      Top             =   2580
      Width           =   2040
   End
   Begin VB.Label lblPlus 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3180
      TabIndex        =   23
      ToolTipText     =   "Sets how many window actions can be undone"
      Top             =   2640
      Width           =   90
   End
   Begin VB.Label lblRunIn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Run in"
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
      Left            =   330
      TabIndex        =   22
      ToolTipText     =   "Select the action to be taken"
      Top             =   4680
      Width           =   450
   End
   Begin VB.Label lblStartup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "When ZenKEY starts,"
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
      Left            =   300
      TabIndex        =   20
      ToolTipText     =   "Use this to start this program when ZenKEY starts (ignored on ZenKEY Restart)"
      Top             =   3720
      Width           =   1590
   End
   Begin VB.Label lblSpecial 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special folder"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1500
      TabIndex        =   19
      ToolTipText     =   "Sets how many window actions can be undone"
      Top             =   1980
      Width           =   960
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
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
      Left            =   360
      TabIndex        =   18
      ToolTipText     =   "The caption on the item as it appears in the menu"
      Top             =   240
      Width           =   600
   End
   Begin VB.Label lblAction 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Action type"
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
      Left            =   330
      TabIndex        =   12
      ToolTipText     =   "Describes the type of action to be performed"
      Top             =   855
      Width           =   870
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File"
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
      Left            =   330
      TabIndex        =   11
      ToolTipText     =   "Select the action to be taken"
      Top             =   1545
      Width           =   255
   End
   Begin VB.Label lblParameter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Param"
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
      Left            =   330
      TabIndex        =   10
      Top             =   1980
      Width           =   465
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
      Left            =   330
      TabIndex        =   7
      ToolTipText     =   "Set key combination that will fire the action"
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   180
      Top             =   120
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1035
      Index           =   1
      Left            =   180
      Top             =   1350
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   180
      Top             =   735
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Index           =   3
      Left            =   180
      Top             =   2460
      Width           =   4935
   End
   Begin VB.Menu mnuWeb 
      Caption         =   "Web"
      Visible         =   0   'False
      Begin VB.Menu mnuMyFav 
         Caption         =   "Browse 'My Favourites'"
      End
      Begin VB.Menu mnuAllFav 
         Caption         =   "Browse 'All Users favourties'"
      End
   End
End
Attribute VB_Name = "frmAction"
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
Public EditIndex As Long ' The index of the item being edited

'Dim ValidKeys() As String
Dim booLoading As Boolean
Private lngSpecial As Long
Private CurClass As String
Dim booDupeKeyWarning As Boolean
Dim FormKeyStrokes As frmKeystrokes

Private Sub cmbAction_Click()
Dim lngIndex As Long
Dim k As Long, max As Long

    If Not booLoading Then
        lngIndex = cmbAction.ListIndex
        max = UBound(Actions(lngIndex).Action())
        CurClass = Prop_Get("Class", Actions(lngIndex).Action(0))
        With cmbActItem
            .Clear
            For k = 1 To max
                .AddItem Prop_Get("Caption", Actions(lngIndex).Action(k))
            Next k
            If max > 0 Then
                .ListIndex = 0
                .Visible = True
            Else
                .Visible = False
            End If
        End With
        Call Form_SetMode(prop)
    End If

End Sub


Private Sub cmbActItem_Click()

    If CurClass = "SEARCH" Then
        If cmbActItem.Tag <> "CHANGING" Then
            If cmbActItem.ListIndex > 0 Then txtParameter.Text = Prop_Get("Action", Searches(cmbActItem.ListIndex - 1))
        End If
    End If
    
End Sub

Private Sub cmbRunIn_Click()
    If Not booLoading Then
        If cmbRunIn.ListIndex = 2 Then
            Dim strFName As String
            If FBR_BrowseForFolder("Select the folder in which you wish you application start.", strFName) Then
                Rem - cmbRunIn list indices.
                Rem - 0 - Applications ' own folder
                Rem - 1 - Current folder
                Rem - 2 - Another selected folder
                Rem - 3 - <Any custom folder that they have selected>
                booLoading = True
                If cmbRunIn.ListCount > 3 Then cmbRunIn.RemoveItem 3
                cmbRunIn.AddItem strFName
                cmbRunIn.ListIndex = 3
                booLoading = False
            End If
            
        End If
    End If
End Sub

Private Sub cmbShift_Change()
    booDupeKeyWarning = False
End Sub

Private Sub cmbSpecial_Click()
Dim k As Long

    If cmbSpecial.Tag = "CHANGING" Then Exit Sub

    Select Case Prop_Get("Class", Actions(cmbAction.ListIndex).Action(0))
        Case "File", "Folder"
            If cmbSpecial.ListIndex <> 0 Then
                k = Val(Extract(cmbSpecial.Text, 1, "("))
                txtFile.Tag = "CHANGING"
                txtFile.Text = InsertSpecialFolder("%" & CStr(k) & "%")
                txtFile.Tag = ""
            End If
        Case "MEDIA"
            If cmbSpecial.ListIndex = 3 Then
                Dim It As frmWindowCapture, strClass As String
                
                Set It = New frmWindowCapture
                Set It.CallingForm = Me
                If It.SelectClass(strClass) Then
                    'If cmbSpecial.ListCount > 3 Then cmbSpecial.RemoveItem 4
                    If cmbSpecial.ListCount < 4 Then
                        cmbSpecial.AddItem strClass
                    Else
                        cmbSpecial.List(4) = strClass
                    End If
                    'cmbSpecial.AddItem strClass, 4
                    cmbSpecial.ListIndex = 4
                Else
                    cmbSpecial.ListIndex = 0
                End If
                
            End If
            
        Case "URL"
            Dim strText As String
            Dim booStripped As Boolean, strPrefix As String
            Dim strNewPrefix As String
            
            Select Case cmbSpecial.ListIndex
                Case 0
                    Rem - Do not do any checks or changes if they have selected exact address
                    Exit Sub
                Case 1 ' .AddItem "Web address (http)"
                    strNewPrefix = "http://"
                Case 2 '"Secure Web address (https)"
                    strNewPrefix = "https://"
                Case 3 ' .AddItem "ftp Site (ftp)|
                    strNewPrefix = "ftp://"
            End Select
            
            strText = Trim(txtFile.Text)
            Select Case True ' ---- First detect the  current type of address
                Case left$(strText, 7) = "http://"
                    Rem - A normal www url
                    strPrefix = "http://"
                Case left$(strText, 8) = "https://"
                    Rem - Case 2 '"Secure Web address (https)"
                    strPrefix = "https://"
                Case left$(strText, 6) = "ftp://"
                    Rem -  ftp Site (ftp)
                    strPrefix = "ftp://"
            End Select
            
            If strPrefix <> strNewPrefix Then
                If Len(strPrefix) > 0 Then strText = Mid$(strText, Len(strPrefix) + 1)
                strText = strNewPrefix & strText
                
                If Len(strText) > 0 Then
                    txtFile.Tag = "CHANGING"
                    txtFile.Text = strText
                    txtFile.Tag = ""
                End If
            End If
            
    End Select
    
End Sub


Private Sub cmbStartup_Click()

    If Not booLoading Then
        ' 0 = "do nothing"
        ' 1 - "fire this action"
        ' 2 = "fire this action after a delay"
        ' 3 - fire action of X seconds'
        Select Case cmbStartup.ListIndex
            Case 2
                Dim strNew As String, booValid As Boolean
                strNew = InputBox("Please enter the number of seconds delay before firing.", "ZenKEY - Startup delay")
                If Len(strNew) = 0 Then
                    Rem - Do nothing
                ElseIf Not IsNumeric(strNew) Then
                    Call ZenMB("Sorry, but this value must be a number.", "OK")
                ElseIf Val(strNew < 1) Or Val(strNew) > 30000 Then
                    Call ZenMB("Sorry, but this value must be a number between 1 and 30,000.", "OK")
                Else
                    booValid = True
                End If
                If booValid Then
                    Call Startup_Set(strNew)
                Else
                    Rem - Revert back to the previosu settings
                    Call Startup_Set(cmbStartup.Tag)
                End If
            Case 0
                Call Startup_Set("")
            Case 1
                Call Startup_Set("0")
        End Select
    End If

End Sub

Private Sub cmdBrowse_Click()
Dim strFName As String

    On Error GoTo ErrorTrap

    Select Case Prop_Get("Class", Actions(cmbAction.ListIndex).Action(0))
        Case "FOLDER"
            If FBR_BrowseForFolder("Select a folder that contains the files or folders.", strFName) Then txtFile.Text = strFName
        Case "CPAPPLET"
            Dim strCPName As String, lngMax As Long
            Dim i As Long
            
            strFName = InsertSpecialFolder("%37%") ' "Windows\System32"
            If FBR_GetOFName("Select the file to Applet to open.", strFName, "CP Applet (*.cpl)") Then
                strFName = GetFileName(strFName)
                txtFile.Text = strFName
            End If
        Case "Url"
            Rem _ URL xperiment
            Call PopupMenu(mnuWeb)
        
        Case Else ' FIle
            'Call PopupMenu(mnuMain)
            If SelectFileDlg(Me, strFName) Then txtFile.Text = strFName
    End Select
    Exit Sub
        
ErrorTrap:
    Call ZenMB("Oops. Error " & Err.Number & ", " & Err.Description & " in sub cmdBrowse_Click", "OK")
    Err.Clear
    
End Sub


Private Sub SetGraphics()
Dim It As Control

    Me.Move 0.5 * (Screen.Width - Me.Width), 0.5 * (Screen.Height - Me.Height)
    Me.AutoRedraw = True
    Call TileMe(Me, LoadPicture(App.Path & "\Help\cloudsdark.jpg"))
    Me.AutoRedraw = False
    

    For Each It In Me.Controls
        If TypeOf It Is Label Then
            It.ForeColor = COL_Zen
        End If
    Next It
    
    Set zbDone.Picture = zbTest.Picture
    Set zbCancel.Picture = zbTest.Picture
    
End Sub





Private Sub cmdKeyStrokes_Click()

    With FormKeyStrokes
        Set .prop = prop.Copy
        Set .CallingForm = Me
        Call .Initialise
        Call CentreForm(FormKeyStrokes)
        .Show
        Me.Visible = False
        Do
            DoEvents
        Loop While .Visible
        Me.Visible = True
        If .booValid Then
            Set prop = .prop.Copy
            txtKeyStrokes.Text = KS_GetDescription(prop("Action"))
        End If
    End With

    
End Sub

Private Sub Form_Activate()
On Error Resume Next

    If prop.IsEmpty Then
        txtCaption.SetFocus
        txtCaption.SelStart = 0
        txtCaption.SelLength = Len(txtCaption.Text)
    Else
        zbSetHK.SetFocus
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call zbCancel_Click
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    booDone = True
End Sub

Private Sub lblHotkey_Click()
    booDupeKeyWarning = False
End Sub

Private Sub mnuAllFav_Click()
Dim strFName  As String

    strFName = InsertSpecialFolder("%31%") ' All users favourties
    If Len(strFName) = 0 Then strFName = InsertSpecialFolder("%6%") ' My Favourites
    If FBR_GetOFName("Select the 'URL' open.", strFName, "Web address (*.url)") Then
        txtFile.Text = strFName
        cmbSpecial.ListIndex = 1
    End If

End Sub


Private Sub mnuMyFav_Click()
Dim strFName  As String

    strFName = InsertSpecialFolder("%6%") ' My Favourites
    If Len(strFName) = 0 Then strFName = InsertSpecialFolder("%31%") ' All users favourties
    If FBR_GetOFName("Select the 'URL' open.", strFName, "Web address (*.url)") Then
        txtFile.Text = strFName
        cmbSpecial.ListIndex = 1
    End If

End Sub



Private Sub zbSetHK_Click()
Dim It As frmSetHotKey

    Set It = New frmSetHotKey
    Set It.CallingForm = Me
    Set It.prop = Me.prop.Copy
    It.EditIndex = EditIndex
    Call It.Init
    Me.Visible = False
    It.Visible = True
    While Not It.booDone
        DoEvents
    Wend
    If It.booValid Then
        Set Me.prop = It.prop
        Call HK_SetCombo(prop)
    End If
    Me.Visible = True
    It.Visible = False
    Unload It
    Set It = Nothing
    booDupeKeyWarning = True

End Sub

Private Sub txtParameter_Change()
    
    If CurClass = "SEARCH" Then
        Dim k As Long
        For k = UBound(Searches()) To 0 Step -1
            If txtParameter.Text = Prop_Get("Action", Searches(k)) Then
                cmbActItem.Tag = "CHANGING"
                cmbActItem.ListIndex = k + 1
                cmbActItem.Tag = ""
                Exit Sub
            End If
        Next k
        cmbActItem.ListIndex = 0
    End If

End Sub

Private Sub zbCancel_Click()
    
    booDone = True
    Me.Hide
    
    
End Sub




Private Sub zbDone_Click()
    
    If Set_Action(prop) Then
        booValid = True
        booDone = True
        Me.Hide
    End If
    
End Sub






Private Sub zbTest_Click()
Dim zDic As New clsZenDictionary
    
    If Set_Action(zDic) Then
        booDupeKeyWarning = True
        zDic("HWnd") = Me.hwnd
        Call TestAction(zDic)
    End If
    
End Sub

















Public Sub Init()

    booLoading = True
    Call SetGraphics
    Call Hotkeys_Init
    Call Actions_Init
    booLoading = False
    
    CurClass = prop("Class")
    If CurClass = "File" Then
        If prop("Action") = "rundll32.exe" Then CurClass = "CPAPPLET"
    End If
    cmbAction.ListIndex = Actions_GetClassIndex(CurClass)
    
    If Not prop.IsEmpty Then

        Me.Caption = "| ZenKEY - Configure Item |"
        chkEnabled.Value = IIf(Not prop("Disabled") = "True", 1, 0)
        txtCaption.Text = prop("Caption")
        Select Case CurClass
            Case "Windows", "System", "Winamp", "ZenKEY", "IDT"
                'Call cmbAction_Click
                cmbActItem.ListIndex = Actions_GetActIndex(prop("Action"), cmbAction.ListIndex)
            Case "File"
                Rem - A file
                txtFile.Text = prop("Action")
                txtParameter.Text = prop("Param")
                If prop("NewInstance") = "True" Then
                    chkBringToForeground.Value = 0
                Else
                    chkBringToForeground.Value = 1
                End If
            Case "Media"
                cmbActItem.ListIndex = Actions_GetActIndex(prop("Action"), cmbAction.ListIndex)

                Select Case prop("Window Class")
                    Case "Active", vbNullString: cmbSpecial.ListIndex = 0
                    Case "Sonique2 Window Class": cmbSpecial.ListIndex = 1
                    Case "WMPlayerApp": cmbSpecial.ListIndex = 2
                    Case Else: cmbSpecial.ListIndex = 4 ' User defined window class
                End Select

            Case "Folder"
                Rem - A folder
                txtFile.Text = InsertSpecialFolder(prop("Action"))
            Case "SpecialFolder"
                Rem - A folder
                On Error Resume Next
                Dim strAct As String
                cmbActItem.Text = SpecialFolderCaption(Val(Mid(prop("Action"), 2)))
                
            Case "URL"
                txtFile.Text = prop("Action")
            'Case "Keystrokes"
            '    txtKeyStrokes.Tag = prop("Action")
            Case "Search"
                Rem - List the appropraite search match, or display 'Custom'
                txtParameter.Text = prop("Action")
                Call Search_SetCombo
                
        End Select

        If CurClass = "Group" Then
            cmbAction.Enabled = False
            zbTest.Visible = False
        Else
            cmbAction.Enabled = True
            zbTest.Visible = True
        End If
        Call HK_SetCombo(prop)

    Else
        Me.Caption = "| ZenKEY - New Item |"
        chkEnabled.Value = 1
'        cmbAction.ListIndex = Actions_GetClassIndex("FILE")
        chkBringToForeground.Value = 1
    End If
    Call Startup_Set(prop("StartUp"))

End Sub

Private Sub txtFile_Change()

    If txtFile.Tag <> "CHANGING" Then
        If CurClass = "URL" Then Call URL_SetCombo
    End If
    
End Sub



Private Function Set_Action(ByRef actDict As clsZenDictionary) As Boolean
Dim strTemp As String
Dim i As Long, j As Long
Dim k As Long
Dim strNewAct As String
Dim strClass As String

On Error GoTo ErrorTrap

    Rem - Check that the item is valid
    
    strClass = CurClass
    actDict("Class") = strClass
    Select Case strClass
        Case "GROUP"
            strTemp = prop("Class")
            If Len(strTemp) > 0 And strTemp <> "Group" Then Call Err.Raise(vbObjectError + 1, , "Sorry, but you cannot change an Item into a group. Try creating a new group.")
        Case "FILE"
            Rem - File to run / open
            If Len(Trim(txtFile.Text)) = 0 Or txtFile.Text = "Select the file/applet to be opened." Then Call Err.Raise(vbObjectError + 1, , "Please select a program to run or a file to open.")
            strNewAct = txtFile.Text
            If Len(txtParameter.Text) > 0 Then actDict("Param") = txtParameter.Text
            If chkBringToForeground.Value = 0 Then
                actDict("NewInstance") = "True"
            Else
                If Right(strNewAct, 4) <> ".exe" Then Call ZenMB("Please note that the 'Bring to foreground if active' setting for documents (or any non-executables) will only work if the program which opens it displays the file name in its title bar.", "OK")
                actDict("NewInstance") = vbNullString
            End If
            
            Rem - cmbRunIn list indices.
            Rem - 0 - Applications ' own folder
            Rem - 1 - Current folder
            Rem - 2 - Another selected folder
            Rem - 3 - <Any custoom folder that they have selected>
            
            Rem - Values for ChangeDir
            Rem - If ChangeDir = "No" - Stay in current dir
            Rem - If InStr(ChangeDir, "\") > 0 - Changes to the specified dir
            Rem - Else changes to App dir
    
            Select Case cmbRunIn.ListIndex
                Case 0: actDict("ChangeDir") = ""  ' Default
                Case 1: actDict("ChangeDir") = "No"
                Case 3: actDict("ChangeDir") = cmbRunIn.Text
            End Select
            
        Case "FOLDER"
            Rem - Open folder
            If Len(Trim(txtFile.Text)) = 0 Then Call Err.Raise(vbObjectError + 1, , "Please select a folder to open.")
            strNewAct = txtFile.Text
        
        Case "SPECIALFOLDER"
            Rem - Open a special folder
            i = InStr(cmbActItem.Text, "(")
            j = Val(Mid(cmbActItem.Text, i + 1))
            strNewAct = "%" & CStr(j) & "%"
        Case "KEYSTROKES"
            Rem - Simulate a series of keypresses
            strTemp = prop("Action") ' txtKeyStrokes.Tag
            If Val(strTemp) < 1 Then Call Err.Raise(vbObjectError + 1, , "Please record a series of Keystrokes to simulate by clicking on the 'Edit' button above.")
            strNewAct = strTemp
        Case "CPAPPLET"
            Rem - Control panel applet
            If Len(Dir(InsertSpecialFolder("%SYSTEM32%\") & txtFile.Text)) = 0 Then Call Err.Raise(vbObjectError + 1, , "This Control Panel applet does not appear to lie in the Windows System folder. Please ensure the name and spelling are correct, or browse for the Applet file (*.cpl).")
            strNewAct = "rundll32.exe"
            actDict("Param") = "shell32.dll,Control_RunDLL " & txtFile.Text
            actDict("Class") = "File"
            
        Case "SYSTEM", "Winamp", "Windows", "ZenKEY"
            Rem - Check that they have not changed the caption
            If Prop_Get("Action", Actions(cmbAction.ListIndex).Action(cmbActItem.ListIndex + 1)) = "PREVENTSAVER" Then Call ZenMB("Please note that the caption of this item is determined by ZenKEY of the fly, and not by what you see here..", "OK")
        Case "Media"
            Select Case cmbSpecial.ListIndex
                Case 0 ' Active window
                    actDict("Window Class") = "Active"
                Case 1 ' Sonique 2
                    actDict("Window Class") = "Sonique2 Window Class"
                Case 2 ' Windows media player
                    actDict("Window Class") = "WMPlayerApp"
                Case Else ' User defined window class / ' An already defined User defined window class
                    actDict("Window Class") = cmbSpecial.Text
            End Select
        Case "URL"
            strTemp = Trim(txtFile.Text)
            If Len(strTemp) < 1 Then Call Err.Raise(vbObjectError + 1, , "Please enter the 'URL', or address of the internet resource.")
            strNewAct = strTemp
        Case "SEARCH"
            strNewAct = Trim(txtParameter.Text)
            If InStr(strNewAct, "<Criteria>") = 0 Then Call Err.Raise(vbObjectError + 1, , "In order to pass your keywords to the Search engine, your address must contain a '<Criteria>' field, where your keywords will be inserted. e.g. http....dotcom/search?q=<Criteria>")
        Case "ZenKEY"
            Rem - Check that they have not changed the caption
           Select Case Prop_Get("Action", Actions(cmbAction.ListIndex).Action(cmbActItem.ListIndex + 1))
                Case "SETAOT", "FOLLOWACTIVE", "HIDEFORM", "WindowUnderMouse", "ToggleHotkeys"
                    Call ZenMB("Please note that the caption of this item is determined by ZenKEY on the fly, and not by what you see here..", "OK")
            End Select
    End Select
    actDict("Startup") = cmbStartup.Tag
    
    
    If Len(strNewAct) < 1 Then strNewAct = Prop_Get("Action", Actions(cmbAction.ListIndex).Action(cmbActItem.ListIndex + 1))
    actDict("Action") = strNewAct
    
    Rem - Check caption
    If Len(Trim(txtCaption.Text)) < 1 Then Call Err.Raise(vbObjectError + 1, , "Please enter a caption for this item.")
    actDict("Caption") = Trim(txtCaption.Text)
    
    Rem - Set startup
    actDict("Startup") = cmbStartup.Tag
    
    Rem - Check enabling
    If chkEnabled.Value = 0 Then actDict("Disabled") = "True"
    
    Rem - Check Hotkeys
    Dim strHK As String, strShift As String
        
    If cmbKey.ListIndex > 0 Then strHK = HKCombo_GetValue(cmbKey)
    If cmbShift.ListIndex > 0 Then strShift = cmbShift.Text
    
    If HKIsOkay(strShift, strHK, EditIndex) Then
        actDict("Hotkey") = strHK
        actDict("ShiftKey") = strShift
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





Public Sub Folders_Init()
Dim k As Integer
Dim strTemp As String

    With cmbSpecial
        .Clear
        .AddItem "<Click here>", 0
        For k = 1 To 255
            'strTemp = InsertSpecialFolder("%" & CStr(k) & "%")
            'If Len(strTemp) > 0 Then cmbSpecial.AddItem FBR_GetLastFolder(strTemp) & " (" & CStr(k) & ")"
            strTemp = SpecialFolderCaption(k)
            If Len(strTemp) > 0 Then cmbSpecial.AddItem strTemp
        Next k
        .ListIndex = 0
    End With

End Sub



Private Sub Media_Init()

    With cmbSpecial
        .Clear
        .AddItem "Active window", 0
        .AddItem "Sonique 2", 1
        .AddItem "Windows media player", 2
        .AddItem "User defined window class", 3
        
        Select Case prop("Window Class")
            Case vbNullString: .ListIndex = 0
            Case "Sonique2 Window Class": .ListIndex = 1
            Case "WMPlayerApp": .ListIndex = 2
            Case Else
                .AddItem prop("Window Class"), 4
                .ListIndex = 4
            
        End Select
        
        
    End With
    
End Sub

Public Sub URL_Init()
Dim k As Integer
Dim strTemp As String

    With cmbSpecial
        .Clear
        .AddItem "Exact address", 0
        .AddItem "Web address - http://", 1
        .AddItem "Secure Web address - https://", 2
        .AddItem "ftp Site - ftp://", 3
        .ListIndex = 0
    End With
    
End Sub

Public Sub URL_SetCombo()
Dim strText As String, lngIndex As Long
    
    strText = Trim(txtFile.Text)
    Select Case True ' ---- First detect the  current type of address
        Case left$(strText, 7) = "http://"
            Rem - A normal www url
            lngIndex = 1
        Case left$(strText, 8) = "https://"
            Rem - Case 2 '"Secure Web address (https)"
            lngIndex = 2
        Case left$(strText, 6) = "ftp://"
            Rem -  ftp Site (ftp)
            lngIndex = 3
    End Select
    
    cmbSpecial.Tag = "Changing"
    cmbSpecial.ListIndex = lngIndex
    cmbSpecial.Tag = ""
    
    

End Sub

Private Sub Actions_Init()
Dim k As Long, max As Long

    max = UBound(Actions())
    With cmbAction
        .Clear
        For k = 0 To max
            .AddItem Prop_Get("Caption", Actions(k).Action(0))
        Next k
    End With
    
    cmbStartup.Clear
    cmbStartup.Tag = vbNullString
    Call cmbStartup.AddItem("do nothing")
    Call cmbStartup.AddItem("fire this action")
    Call cmbStartup.AddItem("fire this action after a delay")
    Set FormKeyStrokes = Nothing
    
    cmbRunIn.ListIndex = 0

End Sub


Private Sub Search_Init()
Dim max As Long, k As Long

    cmbActItem.Clear
    cmbActItem.AddItem "Custom"
    max = UBound(Searches())
    For k = 0 To max
        cmbActItem.AddItem Prop_Get("Caption", Searches(k))
    Next k
    cmbActItem.ListIndex = 0
    
End Sub

Private Sub Form_SetSize(ByVal NumLines As Long)

    Rem - Hide all unneccesary controls
    chkBringToForeground.Visible = False
    'chkChangeDir.Visible = False
    cmbStartup.Visible = False
    lblStartup.Visible = False
    txtFile.Visible = False
    cmdBrowse.Visible = False
    lblFile.Visible = False
    lblParameter.Visible = False
    lblSpecial.Visible = False
    cmbSpecial.Visible = False
    txtParameter.Visible = False
    cmbActItem.Visible = False
    txtKeyStrokes.Visible = False
    cmdKeyStrokes.Visible = False
    lblRunIn.Visible = False
    cmbRunIn.Visible = False

    NumLines = NumLines + 1
    'txtFile.Top = 100
    'cmdBrowse.Top = 100
    'lblFile.Top = 104
    'txtParameter.Width = 89
    
    Rem - Size the form
    Dim sngShift As Single
    'sngShift = 30 * (NumLines - 1)
    sngShift = txtFile.Height * 3 / 2.25 * (NumLines - 1)
    Shape1(1).Height = 38 + sngShift
    
    Rem - Add the universal 'Startup' option.
    lblStartup.Top = Shape1(1).Top + Shape1(1).Height - 30
    lblStartup.Visible = True
    cmbStartup.Top = lblStartup.Top - 2
    cmbStartup.Visible = True
    
    Shape1(3).Top = 1.07 * Shape1(1).Height + Shape1(1).Top ' Shape enclosing hotkeys
    cmbShift.Top = Shape1(3).Top + 8
    cmbKey.Top = Shape1(3).Top + 8
    lblHotkey.Top = Shape1(3).Top + 11
    zbSetHK.Top = Shape1(3).Top + 8
    'zbTest.Top = Shape1(3).Top + 48
    zbTest.Top = Shape1(3).Top + Shape1(3).Height + Shape1(2).Top
    zbDone.Top = zbTest.Top
    zbCancel.Top = zbTest.Top
    'Me.Height = 960 + Me.ScaleY(zbCancel.Top, Me.ScaleMode, vbTwips)
    Me.Height = Me.ScaleY(zbTest.Top + 2.25 * zbTest.Height + Shape1(2).Top, Me.ScaleMode, vbTwips)
    lblPlus.Move cmbKey.left - 10, cmbKey.Top + 2
    Call CentreForm(Me)
    
End Sub

Private Sub Form_SetMode(ByRef prop As clsZenDictionary)
'Const LineHeight = 28
'Dim sngTop As Single
Static LineHeight As Single
Static LineTop As Single
Static LineHeightBig As Single

    If LineTop = 0 Then
        Rem - Initialize
        LineTop = txtFile.Top
        LineHeight = txtParameter.Top - LineTop
        LineHeightBig = LineTop - cmbAction.Top
    End If

    'sngTop = txtFile.Top
    Select Case CurClass
        Case "File", "" ' Make default mode
            Call Form_SetSize(3)
            cmdBrowse.Top = LineTop
            lblFile.Top = LineTop + 4
            lblFile.Caption = "File"
            lblFile.Visible = True
            txtFile.Visible = True
            cmdBrowse.Visible = True
            lblRunIn.Top = LineTop + LineHeight + 4
            lblRunIn.Visible = True
            cmbRunIn.Top = LineTop + LineHeight
            cmbRunIn.Visible = True
            lblParameter.Caption = "Param"
            lblParameter.Top = LineTop + 2 * LineHeight + 4
            lblParameter.Visible = True
            txtParameter.Text = prop("Param")
            txtParameter.Top = LineTop + 2 * LineHeight
            txtParameter.Visible = True
            chkBringToForeground.Top = lblParameter.Top
            chkBringToForeground.Visible = True
            txtFile.ToolTipText = "Enter the name of the shortcut or executable file for the program,.Click on the ""..."" button to select a file."
            Select Case prop("Action")
                Case "", "rundll32.exe": txtFile.Text = "Select the file/applet to be opened."
                Case Else: txtFile.Text = prop("Action")
            End Select
            'If Prop_Get("ChangeDir", Prop) <> "No" Then chkChangeDir.Value = 1 Else chkChangeDir.Value = 0
            
            Rem - Values for ChangeDir
            Rem - If ChangeDir = "No" - Stay in current dir
            Rem - If InStr(ChangeDir, "\") > 0 - Changes to the specified dir
            Rem - Else changes to App dir
            
            Rem - cmbRunIn list indices.
            Rem - 0 - Applications ' own folder
            Rem - 1 - Current folder
            Rem - 2 - Another selected folder
            Rem - 3 - <Any custoom folder that they have selected>
            Dim strDir As String
            strDir = prop("ChangeDir")
            If strDir = "No" Then
                Rem - Stay in current folder
                cmbRunIn.ListIndex = 1
            Else
                If InStr(strDir, "\") > 0 Then
                    Rem - Changes to the specified dir
                    If cmbRunIn.ListCount > 3 Then cmbRunIn.RemoveItem 3
                    cmbRunIn.AddItem strDir
                    cmbRunIn.ListIndex = 3
                Else
                    Rem - Else changes to App dir
                    cmbRunIn.ListIndex = 0
                End If
            End If
            
        Case "Search"
            Call Form_SetSize(2)
            Call Search_Init
            'txtParameter.Width = 260
            txtParameter.Move txtFile.left, LineTop + LineHeight, 260
            txtParameter.Visible = True
            lblParameter.Caption = "Address"
            lblParameter.Top = LineTop + 4 + LineHeight
            lblParameter.Visible = True
            lblFile.Caption = "Site"
            lblFile.Top = LineTop + 4
            lblFile.Visible = True
            Select Case prop("SearchString")
                Case "http://www.google.com/search?q=<Criteria>", "": cmbActItem.ListIndex = 1
                Case Else: cmbActItem.ListIndex = 0
            End Select
            cmbActItem.Move txtFile.left, LineTop
            cmbActItem.Visible = True
            
        Case "Media"
            Call Form_SetSize(2)
            cmbActItem.Visible = True
            lblFile.Caption = "Command"
            lblFile.Visible = True
            lblSpecial.Caption = "Send commands to"
            txtFile.Visible = False
            cmbSpecial.Visible = True
            lblSpecial.Visible = True
            cmbActItem.Move 83, LineTop, 246
            cmbActItem.Visible = True
            cmbActItem.Visible = True
            Call Media_Init
            txtFile.ToolTipText = "Enter the name of the folder to be opened here. Click on the ""..."" button to select a file."
        Case "Group"
            Call Form_SetSize(1)
            cmbActItem.ToolTipText = "This action opens a group of pre-defined actions."
            cmbActItem.Visible = True
            lblFile.Caption = "Action"
            lblFile.Visible = True
        Case "CPApplet"
            Call Form_SetSize(1)
            lblFile.Caption = "Action"
            lblFile.Visible = True
            txtFile.ToolTipText = "Type in the name of the applet here, or click on the '...' button to browse for an applet"
            txtFile.Visible = True
            cmdBrowse.Visible = True
            Select Case prop("Param")
                Case "": txtFile.Text = "Select the file/applet to be opened."
                Case Else: txtFile.Text = Mid$(prop("Param"), Len("shell32.dll,Control_RunDLL") + 2) 'Prop_Get("Action", Prop)
            End Select
        Case "URL"
            Call Form_SetSize(2)
            lblFile.Caption = "Address"
            lblFile.Visible = True
            cmdBrowse.Visible = True
            cmbSpecial.Visible = True
            lblSpecial.Caption = "Address type"
            cmbSpecial.ToolTipText = "Click here to browse your 'Favourites' folder for Web shortcuts/URL's."
            'cmbActItem.Visible = True
            lblSpecial.Visible = True
            txtFile.ToolTipText = "Enter the address of the Website here, or click on the ""..."" button to choose a URL from your favourites."
            txtFile.Visible = True
            
            Call URL_Init
            If (Len(prop("Action")) > 0) And (prop("Class") = "URL") Then
                txtFile.Text = prop("Action")
                Call cmbSpecial_Click
            Else
                txtFile.Text = "http://www.quinnware.com"
            End If
        Case "Folder"
            Call Form_SetSize(2)
            lblFile.Caption = "Folder"
            lblFile.Visible = True
            txtFile.Visible = True
            cmdBrowse.Visible = True
            cmbSpecial.Visible = True
            lblSpecial.Caption = "Special folder" '
            lblSpecial.Visible = True
            txtFile.ToolTipText = "Enter the name of the folder to be opened here. Click on the ""..."" button to select a file."
            cmbSpecial.ToolTipText = "Select the special folder to be opened."
            Call Folders_Init
            
            If (Len(prop("Action")) > 0) And (prop("Class") = "Folder") Then
                txtFile.Text = prop("Action")
            Else
                txtFile.Text = "Select the folder to be opened."
            End If
            cmbSpecial.Visible = True
            'a.Move 68, 100, 261
        Case "SpecialFolder"
            Call Form_SetSize(1)
            Call SpecialFolder_Init
            lblSpecial.Caption = "Special folder"
            cmbActItem.Visible = True
            cmbActItem.ToolTipText = "Open a 'Windows defined' folder e.g. The desktop folder"
            lblFile.Caption = "Folder"
            lblFile.Visible = True
        Case "SystemFolder"
            Call Form_SetSize(1)
            lblSpecial.Caption = "System folder"
            cmbActItem.Visible = True
            cmbActItem.ToolTipText = "Open a 'Windows System Folder e.g. The Recyclce bin"
            lblFile.Caption = "Folder"
            lblFile.Visible = True
        Case "KeyStrokes"
            Call Form_SetSize(1)
            lblFile.Caption = "Press these keys :"
            lblFile.Visible = True
            If FormKeyStrokes Is Nothing Then Set FormKeyStrokes = New frmKeystrokes
            txtKeyStrokes.Text = KS_GetDescription(prop("Action"))
            txtKeyStrokes.Move 134, LineTop, 112
            cmdKeyStrokes.Move 246, LineTop
            txtKeyStrokes.Visible = True
            cmdKeyStrokes.Visible = True
        Case Else             ' ZenKEY

            Call Form_SetSize(1)
            cmbActItem.Visible = True
            lblFile.Caption = "Action"
            lblFile.Visible = True
            
    End Select

End Sub


Private Sub SpecialFolder_Init()

Dim k As Integer
Dim strTemp As String

    With cmbActItem
        .Clear
        For k = 1 To 255
            strTemp = InsertSpecialFolder("%" & CStr(k) & "%")
            If Len(strTemp) > 0 Then .AddItem FBR_GetLastFolder(strTemp) & " (" & CStr(k) & ")"
        Next k
        .ListIndex = 0
    End With

End Sub

Private Sub SpecialFolder_Init()

Dim k As Integer
Dim strTemp As String

    With cmbActItem
        .Clear
        For k = 1 To 255
            strTemp = InsertSpecialFolder("%" & CStr(k) & "%")
            If Len(strTemp) > 0 Then .AddItem FBR_GetLastFolder(strTemp) & " (" & CStr(k) & ")"
        Next k
        .ListIndex = 0
    End With

End Sub
Private Sub HK_SetCombo(ByRef prop As clsZenDictionary)
        
    If Len(prop("ShiftKey")) <> 0 Then
        cmbShift.Text = HotKeys.ShiftValToStr(HotKeys.ShiftValue(prop("ShiftKey")))         ' COnversion required for compatability
    Else
        cmbShift.ListIndex = 0
    End If
    
    Dim lngKey As String
    lngKey = Val(prop("Hotkey"))
    If lngKey <> 0 Then Call HKCombo_Display(lngKey, cmbKey) Else cmbKey.ListIndex = 0

End Sub


Private Sub Search_SetCombo()
Dim k As Long, strAct As String
Dim i

    strAct = txtParameter.Text
    If Len(strAct) = 0 Then
        Rem - Default to google
        k = 0
        txtParameter.Text = Searches(0)
    Else
        Rem - Search for the next
        k = -1
        For i = 0 To UBound(Searches())
            If strAct = Prop_Get("Action", Searches(i)) Then
                k = i + 1
                Exit For
            End If
        Next
    End If
    
    cmbActItem.Tag = "CHANGING"
    If k > -1 Then 'Set to the one found
        cmbActItem.ListIndex = k
    Else 'Set to custom
        cmbActItem.ListIndex = 0
    End If
    cmbActItem.Tag = ""

End Sub

Private Sub Startup_Set(ByVal StartUp As String)
Dim Index As Long, booPrev As Boolean

    Select Case True
        Case StartUp = "0"
            Index = 1
        Case IsNumeric(StartUp)
            Rem - Remove a previosu choice
            If cmbStartup.ListCount > 3 Then Call cmbStartup.RemoveItem(3)
            Call cmbStartup.AddItem("fire this action after " & StartUp & "  seconds")
            Index = 3
        Case Else
            StartUp = vbNullString
    End Select
    
    booPrev = booLoading
    booLoading = True
    If Index < 3 And cmbStartup.ListCount > 3 Then Call cmbStartup.RemoveItem(3)
    cmbStartup.ListIndex = Index
    cmbStartup.Tag = StartUp
    booLoading = booPrev

End Sub

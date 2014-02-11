VERSION 5.00
Begin VB.Form frmWizNewItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "| ZenKEY - Configure action |"
   ClientHeight    =   4500
   ClientLeft      =   8835
   ClientTop       =   7320
   ClientWidth     =   8955
   Icon            =   "frmWizNewItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   Begin VB.TextBox txtKeyStrokes 
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
      Height          =   345
      Left            =   5040
      TabIndex        =   17
      Text            =   "Keystrokes"
      ToolTipText     =   "Use this field to specify the details of the action you wish to perform."
      Top             =   300
      Width           =   2850
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
      Height          =   345
      Left            =   5460
      TabIndex        =   14
      ToolTipText     =   "Browse for or type in the name of a folder. For 'Windows Special Folders', just enter the number."
      Top             =   3240
      Width           =   3315
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   8460
      TabIndex        =   4
      ToolTipText     =   "Click here to choose the file or program"
      Top             =   2040
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
      Left            =   5640
      TabIndex        =   2
      Text            =   "Select the file/applet to be opened."
      ToolTipText     =   "Use this field to specify the details of the action you wish to perform."
      Top             =   2040
      Width           =   2805
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
      Left            =   5640
      TabIndex        =   0
      Text            =   "New item name"
      ToolTipText     =   "Set caption on the item as it appears in the menu"
      Top             =   900
      Width           =   3150
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
      ItemData        =   "frmWizNewItem.frx":058A
      Left            =   5640
      List            =   "frmWizNewItem.frx":0591
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Select the special folder to be opened."
      Top             =   2640
      Width           =   3135
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
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "This field specifies the type of action you wish to perform. Click on the dropdown box for all your options."
      Top             =   1500
      Width           =   3150
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
      ItemData        =   "frmWizNewItem.frx":05A1
      Left            =   600
      List            =   "frmWizNewItem.frx":05A8
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select the action to be taken"
      Top             =   360
      Width           =   3150
   End
   Begin VB.CommandButton zbTest 
      Caption         =   "Test"
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
      Left            =   960
      TabIndex        =   10
      ToolTipText     =   "Fire the action, just to make sure it works..."
      Top             =   3240
      Width           =   1575
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
      TabIndex        =   11
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
      TabIndex        =   12
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
      TabIndex        =   13
      Top             =   3945
      Width           =   1455
   End
   Begin VB.CheckBox chkBringToForeground 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bring this program to the foreground if it is open and I press the Hotkey (Programs only)"
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
      Left            =   480
      TabIndex        =   16
      ToolTipText     =   $"frmWizNewItem.frx":05B8
      Top             =   2700
      Width           =   8295
   End
   Begin VB.Label lblParameter 
      Alignment       =   1  'Right Justify
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
      Left            =   3255
      TabIndex        =   15
      ToolTipText     =   "Sets how many window actions can be undone"
      Top             =   3360
      Width           =   465
   End
   Begin VB.Label lblSpecial 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special folder"
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
      TabIndex        =   9
      ToolTipText     =   "Sets how many window actions can be undone"
      Top             =   2700
      Width           =   990
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1. The name of this item should be ................................................."
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
      TabIndex        =   8
      ToolTipText     =   "This is the text that appears in the ZenKEY menu"
      Top             =   960
      Width           =   5520
   End
   Begin VB.Label lblAction 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2. The type of action I want it to perform is....................................."
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
      TabIndex        =   7
      ToolTipText     =   "This field specifies the type of action you wish to perform. Click on the dropdown box for all your options."
      Top             =   1560
      Width           =   5505
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3. The specific action I want it to perform is........................................................................"
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
      TabIndex        =   5
      ToolTipText     =   "Use this field to specify the details of the action you wish to perform."
      Top             =   2160
      Width           =   7650
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectDrag 
         Caption         =   "Select by drag and drop"
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "Look for Programs on Desktop"
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
Attribute VB_Name = "frmWizNewItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim FormKeyStrokes As frmKeystrokes
Dim booLoading As Boolean
Private CurClass As String

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
        Call Form_SetMode(ZW_NewItem)
        
    End If


End Sub


Private Sub cmbActItem_Click()

    If CurClass = "SEARCH" Then
        If cmbActItem.Tag <> "CHANGING" Then
            If cmbActItem.ListIndex > 0 Then txtParameter.Text = Prop_Get("Action", Searches(cmbActItem.ListIndex - 1))
        End If
    End If
    
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



Private Sub cmdBrowse_Click()
Dim strFName As String

    On Error GoTo ErrorTrap

    Select Case Prop_Get("Class", Actions(cmbAction.ListIndex).Action(0))
        Case "Keystrokes"
            With FormKeyStrokes
                Set .prop = zenDic("Action", txtKeyStrokes.Tag)
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
                    txtKeyStrokes.Tag = .prop("Action")
                    txtKeyStrokes.Text = KS_GetDescription(txtKeyStrokes.Tag)
                End If
            End With
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
            Call PopupMenu(mnuMain)
        
    End Select
    Exit Sub
        
ErrorTrap:
    Call ZenMB("Oops. Error " & Err.Number & ", " & Err.Description & " in sub cmdBrowse_Click", "OK")
    Err.Clear
    
End Sub







Private Sub Form_Activate()
On Error Resume Next

    txtCaption.SetFocus
    txtCaption.SelStart = 0
    txtCaption.SelLength = Len(txtCaption.Text)
    
End Sub

Private Sub Form_Load()

    cmbActItem.Move txtFile.left, txtFile.Top
    txtKeyStrokes.Move txtFile.left, txtFile.Top
    Set zbNext.Picture = zbBack.Picture
    Set zbCancel.Picture = zbBack.Picture
    Set zbTest.Picture = zbBack.Picture
        
    

    booLoading = True
    Call Actions_Load
    Call Actions_Init
    booLoading = False
        
    Rem - Set objects for Zenkey Config compatibility
    Set MainForm = Me

    CurClass = ZW_NewItem("Class")
    cmbAction.ListIndex = Actions_GetClassIndex(CurClass)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then End
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

Private Sub mnuSelectDrag_Click()
Dim It As frmWindowCapture
Dim strExe As String

    Set It = New frmWindowCapture
    Set It.CallingForm = Me
    Me.Visible = False
    If It.SelectExe(strExe) Then
        txtFile.Text = strExe
    End If
    Me.Visible = True
    Unload It
    
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

Private Sub zbBack_Click()

    Call ZW_Next("Previous")
    
End Sub

Private Sub zbCancel_Click()
    
    End
    
End Sub









Private Sub zbNext_Click()
    If Set_Action() Then Call ZW_Next("Next")
End Sub

Private Sub zbTest_Click()

    Select Case CurClass
        Case "Group"
            Call ZenMB("Sorry, but menus cannot be tested. Adding a menu adds a menu item which then contains other items.")
        Case "ZenKEY"
            Call ZenMB("Sorry, but ZenKEY actions cannot be tested")
        Case Else
            ZW_NewItem("Hwnd") = Me.hwnd
            If Set_Action Then Call TestAction(ZW_NewItem)
    End Select

End Sub






Private Sub mnuBrowse_Click()
Dim strFName  As String

    strFName = "C:"
    If FBR_GetOFName("Select the file to run or open.", strFName, "All files (*.*)") Then txtFile.Text = strFName

End Sub

Private Sub mnuDesktop_Click()
Dim strFName  As String

    
    strFName = InsertSpecialFolder("%25%")
    'If FBR_GetOFName("Select the file to run or open.", strFName, "Shortcuts (*.lnk)", "Executable files (*.exe)") Then
    If FBR_GetOFName("Select the file to run or open.", strFName, "Executable files (*.exe)") Then txtFile.Text = strFName

End Sub

Private Sub mnuMydocuments_Click()
Dim strFName  As String

    strFName = InsertSpecialFolder("%5%")
    If FBR_GetOFName("Select the file to run or open.", strFName, "All files (*.*)") Then txtFile.Text = strFName

End Sub


Private Sub mnuPrograms_Click()
Dim strFName  As String

    'strFName = InsertSpecialFolder("%38%")
    strFName = Registry.GetRegistry(HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir") ' Other does not work on 98
    If FBR_GetOFName("Select the file to run or open.", strFName, "Executable files (*.exe)", "Shortcuts (*.lnk") Then txtFile.Text = strFName

End Sub




Private Sub txtFile_Change()

    If txtFile.Tag <> "CHANGING" Then
        If CurClass = "URL" Then Call URL_SetCombo
    End If
    
End Sub



Private Function Set_Action() As Boolean
Dim strTemp As String
Dim strAction As String
Dim i As Long, j As Long
Dim k As Long
Dim strNewAct As String
Dim strClass As String

On Error GoTo ErrorTrap

    Rem - Check that the item is valid
    
    strClass = CurClass
    ZW_NewItem("Class") = strClass
    Select Case strClass
        Case "FILE"
            Rem - File to run / open
            If Len(Dir(Trim(txtFile.Text))) = 0 Then
                Call Err.Raise(vbObjectError + 1, , "Please select a program to run or a file to open.")
            End If
            strNewAct = txtFile.Text
            If Len(txtParameter.Text) > 0 Then ZW_NewItem("Param") = txtParameter.Text
            If chkBringToForeground.Value = 0 Then
                ZW_NewItem("NewInstance") = "True"
            Else
                ZW_NewItem("NewInstance") = vbNullString
            End If
            
        Case "ZenKEY"
            Rem - Check that they have not changed the caption
           Select Case Prop_Get("Action", Actions(cmbAction.ListIndex).Action(cmbActItem.ListIndex + 1))
                Case "SETAOT", "FOLLOWACTIVE", "HIDEFORM", "WindowUnderMouse", "ToggleHotkeys"
                    Call ZenMB("Please note that the caption of this item is determined by ZenKEY on the fly, and not by what you see here..", "OK")
            End Select
        Case "FOLDER"
            Rem - Open folder
            strTemp = txtFile.Text
            If Len(Dir(strTemp, vbNormal + vbHidden + vbSystem + vbVolume + vbDirectory)) = 0 Then
                Call Err.Raise(vbObjectError + 1, , "This folder '" + strTemp + "' does not appear to exist. Please select a folder to open.")
            End If
            strNewAct = Trim(strTemp)
        Case "CPAPPLET"
            Rem - Control panel applet
            If Len(Dir(InsertSpecialFolder("%SYSTEM32%\") & txtFile.Text)) = 0 Then Call Err.Raise(vbObjectError + 1, , "This Control Panel applet does not appear to lie in the Windows System folder. Please ensure the name and spelling are correct, or browse for the Applet file (*.cpl).")
            'Call Prop_Set("Action", "rundll32.exe", strAction)
            strNewAct = "rundll32.exe"
            ZW_NewItem("Param") = "shell32.dll,Control_RunDLL " & txtFile.Text
            ZW_NewItem("Class") = "File"
            
        Case "SYSTEM", "Winamp", "Windows", "ZenKEY"
            Rem - Check that they have not changed the caption
            If Prop_Get("Action", Actions(cmbAction.ListIndex).Action(cmbActItem.ListIndex + 1)) = "PREVENTSAVER" Then Call ZenMB("Please note that the caption of this item is determined by ZenKEY of the fly, and not by what you see here..", "OK")
        Case "Media"
            Select Case cmbSpecial.ListIndex
                Case 0 ' Active window
                    ZW_NewItem("Window Class") = "Active"
                Case 1 ' Sonique 2
                    ZW_NewItem("Window Class") = "Sonique2 Window Class"
                Case 2 ' Windows media player
                    ZW_NewItem("Window Class") = "WMPlayerApp"
                Case Else ' User defined window class / ' An already defined User defined window class
                    ZW_NewItem("Window Class") = cmbSpecial.Text
            End Select
        Case "URL"
            strTemp = Trim(txtFile.Text)
            If Len(strTemp) < 1 Then Call Err.Raise(vbObjectError + 1, , "Please enter the 'URL', or address of the internet resource.")
            strNewAct = strTemp
        Case "SEARCH"
            strNewAct = Trim(txtParameter.Text)
            If InStr(strNewAct, "<Criteria>") = 0 Then Call Err.Raise(vbObjectError + 1, , "In order to pass your keywords to the Search engine, your address must contain a '<Criteria>' field, where your keywords will be inserted. e.g. http....dotcom/search?q=<Criteria>")
        Case "SpecialFolder"
            Rem - Open a special folder
            i = InStr(cmbActItem.Text, "(")
            j = Val(Mid(cmbActItem.Text, i + 1))
            strNewAct = "%" & CStr(j) & "%"
        Case "KEYSTROKES"
            Rem - Simulate a series of keypresses
            strNewAct = txtKeyStrokes.Tag
            If Val(strNewAct) < 1 Then Call Err.Raise(vbObjectError + 1, , "Please record a series of Keystrokes to simulate by clicking on the '...' button above.")

    End Select
    
    If Len(strNewAct) < 1 Then strNewAct = Prop_Get("Action", Actions(cmbAction.ListIndex).Action(cmbActItem.ListIndex + 1))
    ZW_NewItem("Action") = strNewAct
    
    Rem - Check caption
    If Len(Trim(txtCaption.Text)) < 1 Then Call Err.Raise(vbObjectError + 1, , "Please enter a caption for this item.")
    ZW_NewItem("Caption") = Trim(txtCaption.Text)
        
    Set_Action = True

    Exit Function

ErrorTrap:
    Call ZenMB(Err.Description, "OK")
    Err.Clear
    
End Function


Public Sub Folders_Init()
Dim k As Integer
Dim strTemp As String

    With cmbSpecial
        .Clear
        .AddItem "<Click here>", 0
        For k = 1 To 255
            strTemp = InsertSpecialFolder("%" & CStr(k) & "%")
            If Len(strTemp) > 0 Then cmbSpecial.AddItem FBR_GetLastFolder(strTemp) & " (" & CStr(k) & ")"
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
        
        Select Case ZW_NewItem("Window Class")
            Case vbNullString: .ListIndex = 0
            Case "Sonique2 Window Class": .ListIndex = 1
            Case "WMPlayerApp": .ListIndex = 2
            Case Else
                .AddItem ZW_NewItem("Window Class"), 4
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
    
End Sub


Private Sub Search_Init()
Dim max As Long, k As Long

    'cmbActItem.Tag = "CHANGING"
    cmbActItem.Clear
    cmbActItem.AddItem "Custom"
    max = UBound(Searches())
    For k = 0 To max
        cmbActItem.AddItem Prop_Get("Caption", Searches(k))
    Next k
    cmbActItem.ListIndex = 0
    'cmbActItem.Tag = "CHANGING"
    
End Sub




Private Sub Form_SetMode(ByVal prop As clsZenDictionary)

    cmbActItem.Visible = False
    txtFile.Visible = False
    cmdBrowse.Visible = False
    lblSpecial.Visible = False
    cmbSpecial.Visible = False
    txtParameter.Visible = False
    lblParameter.Visible = False
    chkBringToForeground.Visible = False
    txtKeyStrokes.Visible = False
    
    Select Case CurClass
        Case "File", "" ' Make default mode
            txtFile.Visible = True
            cmdBrowse.Visible = True
            txtFile.ToolTipText = "Enter the name of the document or executable file for the program,.Click on the ""..."" button to select a file/program."
            Select Case prop("Action")
                Case "": txtFile.Text = "Select the file/program to open."
                Case "rundll32.exe": txtFile.Text = "Select the file/applet to be opened."
                Case Else: txtFile.Text = prop("Action")
            End Select
            chkBringToForeground.Visible = True
            lblFile.Caption = "3. The file or program I wish to open is ........................................................................."
        Case "Keystrokes"
            txtKeyStrokes.ToolTipText = "Enter the key sequence you wish to simulate. Click on the ""..."" button to record your keystrokes."
            txtKeyStrokes.Move txtFile.left, txtFile.Top
            txtKeyStrokes.Text = KS_GetDescription(prop("Action"))
            If FormKeyStrokes Is Nothing Then Set FormKeyStrokes = New frmKeystrokes
            txtKeyStrokes.Visible = True
            cmdBrowse.Visible = True
            lblFile.Caption = "3. The keystrokes I wish to simulate are ............................................................"
        Case "Search"
            Call Search_Init
            Select Case prop("SearchString")
                Case "http://www.google.com/search?q=<Criteria>", "": cmbActItem.ListIndex = 1
                Case Else: cmbActItem.ListIndex = 0
            End Select
            cmbActItem.Visible = True
            lblFile.Caption = "3. The type of search I wish to conduct is .........................................................."
        Case "Media"
            cmbActItem.Visible = True
            lblSpecial.Caption = "Send commands to"
            txtFile.Visible = False
            cmbSpecial.Visible = True
            lblSpecial.Visible = True
            'cmbActItem.Move 83, 100, 246
            cmbActItem.Visible = True
            Call Media_Init
            txtFile.ToolTipText = "Enter the name of the folder to be opened here. Click on the ""..."" button to select a file."
            lblFile.Caption = "3. The Media command I wish to issue is .................................................."
        Case "Group"
            cmbActItem.ToolTipText = "This action opens a group of pre-defined actions."
            cmbActItem.Visible = True
            lblFile.Caption = "3. The specific action I wish to perform is to ...................................................."
        Case "CPApplet"
            txtFile.ToolTipText = "Type in the name of the applet here, or click on the '...' button to browse for an applet"
            txtFile.Visible = True
            cmdBrowse.Visible = True
            Select Case prop("Param")
                Case "": txtFile.Text = "Select the file/applet to be opened."
                Case Else: txtFile.Text = Mid$(prop("Param"), Len("shell32.dll,Control_RunDLL") + 2)  'Prop_Get("Action", Prop)
            End Select
            cmbSpecial.Visible = False
            lblSpecial.Visible = False
            lblFile.Caption = "3. The Control panel applet I wish to open is ................................................................"
        Case "URL"
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
            lblFile.Caption = "3. The Internet address I wish to open is ....................................................."
            
        Case "Folder"
            txtFile.Visible = True
            cmdBrowse.Visible = True
            'cmbSpecial.Visible = True
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
            lblFile.Caption = "3. The folder I wish to open is ....................................................................."
        Case "SpecialFolder"
            Call SpecialFolder_Init
            lblSpecial.Caption = "Special folder"
            'cmbSpecial.Visible = True
            cmbActItem.Visible = True
            txtFile.ToolTipText = "Open a 'Windows defined' folder. This value is usually a registry key "
            lblFile.Caption = "3. The Special folder I wish to open is ...................................................................."
            'txtFile.Visible = True
            'lblSpecial.Visible = True
        Case Else             ' ZenKEY

            cmbActItem.Visible = True
            lblFile.Caption = "3. The specific action I wish to perform is to ............................................."
    End Select
'    If Len(Prop_Get("Caption", Prop)) > 0 Then txtCaption.Text = Prop_Get("Caption", Prop)
    
End Sub



Public Sub Init()
Dim prop As clsZenDictionary
    
    Set prop = ZW_NewItem
    If prop.IsEmpty Then Exit Sub
    
    txtCaption.Text = prop("Caption")
    Select Case CurClass
        Case "Windows", "System", "Winamp", "ZenKEY"
            'Call cmbAction_Click
            cmbActItem.ListIndex = Actions_GetActIndex(prop("Action"), cmbAction.ListIndex)
        Case "File"
            If prop("Action") = "rundll32.exe" Then
                Rem - A control panel applet
                cmbAction.ListIndex = Actions_GetClassIndex("CPAPPLET")
            Else
                Rem - A file
                txtFile.Text = prop("Action")
                txtParameter.Text = prop("Param")
                If prop("NewInstance") = "True" Then
                    chkBringToForeground.Value = 0
                Else
                    chkBringToForeground.Value = 1
                End If
                booLoading = True
                'Select Case Prop_Get("Startup", Prop)
                '    Case "Min": cmbStartup.ListIndex = 1
                '    Case "Normal": cmbStartup.ListIndex = 2
                '    Case "Max": cmbStartup.ListIndex = 3
                '    Case Else: cmbStartup.ListIndex = 0
                'End Select
                booLoading = False
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
        Case "URL"
            txtFile.Text = prop("Action")
        Case "Search"
            txtParameter.Text = prop("Action")

    End Select

    If CurClass = "Group" Then
        cmbAction.Enabled = False
        zbTest.Visible = False
    Else
        cmbAction.Enabled = True
        zbTest.Visible = True
    End If



End Sub

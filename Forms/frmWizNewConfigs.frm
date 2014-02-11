VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWizNewConfigs 
   Caption         =   "| ZenKEY - Wizard |"
   ClientHeight    =   4500
   ClientLeft      =   7635
   ClientTop       =   4950
   ClientWidth     =   8955
   Icon            =   "frmWizNewConfigs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   Begin VB.CommandButton zbMenuHK 
      Caption         =   "some keys"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   7
      ToolTipText     =   "Click here to set the Hotkey combination."
      Top             =   3225
      Width           =   1215
   End
   Begin MSComctlLib.ListView lsvGroup 
      Height          =   1575
      Left            =   2580
      TabIndex        =   6
      ToolTipText     =   "Click here to change, add or remove item."
      Top             =   1500
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CheckBox chkMenuHK 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Yes"
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
      Left            =   7965
      TabIndex        =   5
      ToolTipText     =   "Show's a pop-up menu when the keys are pressed"
      Top             =   3240
      Width           =   750
   End
   Begin VB.CheckBox chkRightClick 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Yes"
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
      Left            =   7965
      TabIndex        =   4
      ToolTipText     =   "Set this menu as the one that appears when you right click on the system tray icon."
      Top             =   3540
      Width           =   750
   End
   Begin VB.CheckBox chkEnableHK 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Yes"
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
      Left            =   7965
      TabIndex        =   3
      ToolTipText     =   "Disable or enable all the Hotkeys in this menu"
      Top             =   1200
      Value           =   1  'Checked
      Width           =   750
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
      Top             =   3945
      Width           =   1455
   End
   Begin VB.Label lblRightClick 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4. This menu should appear when I right click on the system tray icon."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Set this menu as the one that appears when you right click on the system tray icon."
      Top             =   3540
      Width           =   8115
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMenuHK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3. This menu should appear when I press"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Show's a pop-up menu when the keys are pressed"
      Top             =   3240
      Width           =   8115
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblEnableHK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1. Would you like to use the shown Hotkeys for the items in this menu?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Disable or enable all the Hotkeys in this menu"
      Top             =   1200
      Width           =   8115
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblClick 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2. Click on the items on the          right to change them."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The following items will appear in the 'XX' menu"
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
      TabIndex        =   8
      Top             =   780
      Width           =   7875
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuNothing 
         Caption         =   "Err, do nothing"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveHK 
         Caption         =   "Remove Hotkey"
      End
      Begin VB.Menu mnuEditHK 
         Caption         =   "Change Hotkey"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "Add new Item"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Item"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Remove Item"
      End
      Begin VB.Menu mnutest 
         Caption         =   "Test this Item"
      End
   End
End
Attribute VB_Name = "frmWizNewConfigs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Private Const ColDisabled = &H808080   '&HC0C0C0
Dim GroupClass() As String
Public CallingForm As Form
Dim GRP_Index As Long
Dim GRP_Start As Long
Dim GRP_End As Long
Dim booLoading As Boolean
Dim TimerAction As String
Private Sub chkEnableHK_Click()
    Call Tree_Load
    
End Sub

Private Sub chkMenuHK_Click()
    
    If Not booLoading Then
        If chkMenuHK.Value = 1 Then
            Rem - It is being enabled. Prompt for the Hokey
            ZIndex = GRP_Start
            Call HK_Selected
        Else
            ZKMenu(GRP_Start)("ShiftKey") = vbNullString
            ZKMenu(GRP_Start)("Hotkey") = vbNullString
        End If
        booLoading = True
        Call Group_ShowHK
        booLoading = False
    End If
    
End Sub

Private Sub chkRightClick_Click()
    
    If Not booLoading Then
        Dim lngPrev As Long, k As Long
        Dim max As Long
        
        Rem - Check if the right click menu is already in use....
        max = UBound(ZKMenu())
        For k = 2 To max
            If ZKMenu(k)("RightClickMenu") = "True" Then
                lngPrev = k
                Exit For
            End If
        Next k
        
        Dim booEnabled As Boolean
        booEnabled = CBool(chkRightClick.Value = 1)
        
        If (lngPrev > 0) And booEnabled Then
            If ZenMB("The 'Right Click' menu is already enabled for '" & ZKMenu(lngPrev)("Caption") & "' and it can only be enabled for one group at a time. Do you wish to set it to this Menu instead?", "Yes", "No") = 1 Then
                booLoading = True
                chkRightClick.Value = 0
                booLoading = False
                Exit Sub
            End If
            ZKMenu(lngPrev)("RightClickMenu") = vbNullString
        End If
        
        ZKMenu(GRP_Start)("RightClickMenu") = IIf(booEnabled, "True", vbNullString)
    End If

End Sub

Private Sub Form_Activate()
    If lsvGroup.Visible Then lsvGroup.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then End
End Sub




Private Sub lsvGroup_Click()
Dim lngSelected As Long
        
    lngSelected = lsvGroup.SelectedItem.Index
    Call PopupMenu(mnuFile)
    If Len(TimerAction) > 0 Then Call Timer_Action
    
    If lsvGroup.ListItems.Count > lngSelected Then lsvGroup.ListItems(lngSelected).Selected = True

    
    
End Sub



Private Sub Timer_Action()
Dim It As frmAction

    Select Case TimerAction
        Case "Edit"
                
            Rem - Ensure that they do not delete everything
            ZIndex = Val(Prop_Get("Index", lsvGroup.SelectedItem.key))
            Set It = New frmAction
            Set MainForm = Me
            With It
                Set .CallingForm = Me
                .prop = ZKMenu(ZIndex)
                .Init
                .EditIndex = ZIndex
                .Move 0.5 * (Screen.Width - .Width), 0.5 * (Screen.Height - .Height)
                .Show
                Me.Hide
                While .Visible
                    DoEvents
                Wend
                If .booValid Then
                    Set ZKMenu(ZIndex) = .prop
                    Call Tree_Load
                End If
                Me.Show
            End With
            Unload It
            Set It = Nothing
            
        Case "New"

            Rem - Ensure that they do not delete everything
            Set It = New frmAction
            With It
                Set .CallingForm = Me
                .prop = "|Class=" & GroupClass(GRP_Index) & "|Caption=Item caption|"
                .Init
                
                .Move 0.5 * (Screen.Width - .Width), 0.5 * (Screen.Height - .Height)
                .Show
                Me.Hide
                Set MainForm = Me
                While .Visible
                    DoEvents
                Wend
                If .booValid Then
                    ZIndex = Val(Prop_Get("Index", lsvGroup.SelectedItem.key)) '+ 1
                    If .prop("Class") = "Group" Then
                        Rem - Adding a group
                        Call ZenMB("Sorry, but you must please use the ZenKEY configuration panel to add groups.")
                        Exit Sub
                    Else
                        Rem - Adding a singular item
                        Call Array_Down(ZIndex, 1)
                        Set ZKMenu(ZIndex) = .prop
                        Call Group_Init(GRP_Index)
                        Call Tree_Load
                    End If
                End If
                Me.Show
            End With
            Unload It
            Set It = Nothing
    End Select
    TimerAction = vbNullString

End Sub



Private Sub zbBack_Click()

    If GRP_Index > 0 Then
        Call Group_Set
        Call Group_Init(GRP_Index - 1)
        Call Group_Display
    Else
        Call ZW_Next("Previous")
    End If
    
End Sub


Private Sub zbCancel_Click()
    End
End Sub


Public Sub Init()
    
    'Call MsgBox("Hello", vbOKOnly)
    Call GroupList_Load
    Call Group_Init(0)
    
    Dim clmX As ColumnHeader
        
    With lsvGroup
        '.View = lvwReport
        Set clmX = .ColumnHeaders.Add(, , "Items", 0.56 * .Width)
        Set clmX = .ColumnHeaders.Add(, , "Hotkeys", 0.36 * .Width)
        ' Set View property to Report.
    End With
    Call Tree_Load
    
    Set zbNext.Picture = zbBack.Picture
    Set zbCancel.Picture = zbBack.Picture
    
End Sub

Public Sub Tree_Load()
Dim k As Long
Dim strCap As String
Dim nParent As Node

    booLoading = True
    Set nParent = Nothing
    lblMessage.Caption = "The following items will appear in the '" & GroupList(GRP_Index) & "' menu"

    With lsvGroup
        .Visible = False
        .ListItems.Clear
        For k = GRP_Start + 1 To GRP_End - 1
            Rem - Add the group item
            If ZKMenu(k)("Class") = "Group" Then
                Rem - Move to the end - ignosre this group
                k = GetGroupEnd(k) + 1
            Else
                Rem - Just add to the current group
                Call Tree_AddItem(k)
            End If
        Next k
        .Visible = True
        .Refresh
        If Me.Visible Then .SetFocus
    End With
    chkRightClick.Value = IIf(ZKMenu(GRP_Start)("RIGHTCLICKMENU") = "True", 1, 0)
    Call Group_ShowHK
    
    booLoading = False
    
End Sub





Public Sub GroupList_Load()
Dim k As Long
Dim strName As String
Dim strClass As String
Dim optButton As CheckBox
Dim Count As Long

    For k = 1 To 11
        Select Case k
            Case 1
                Set optButton = CallingForm.chkPrograms ' launch programs
                strName = "My programs"
                strClass = "File"
            Case 2
                Set optButton = CallingForm.chkDocuments  ' Open documents and folder
                strName = "My documents"
                strClass = "File"
            Case 3
                Set optButton = CallingForm.chkFolders ' Open folders
                strName = "Windows Folders"
                strClass = "Folder"
            Case 4
                Set optButton = CallingForm.chkInternet ' Open internet resources
                strName = "Internet resources"
                strClass = "URL"
            Case 5
                Set optButton = CallingForm.chkControlPanel ' Enable the Infinite desktop
                strName = "Control Panel"
                strClass = "CPAPPLET"
            Case 6
                Set optButton = CallingForm.chkWinMove ' Control program windows
                strName = "My window"
                strClass = "WINDOWS"
            Case 7
                Set optButton = CallingForm.chkMedia ' Control media players
                strName = "Windows Media Commands"
                strClass = "MEDIA"
            Case 8
                Set optButton = CallingForm.chkWinamp ' Control Winamp
                strName = "Winamp Controls"
                strClass = "WINAMP"
            Case 9
                Set optButton = CallingForm.chkWinPrograms ' access Windows utils
                strName = "Windows programs"
                strClass = "File"
            Case 10
                Set optButton = CallingForm.chkSystem ' Issue Window system commands
                strName = "Windows System"
                strClass = "SYSTEM"
            Case 11
                Set optButton = CallingForm.chkIDT ' Enable the Infinite desktop
                strName = "My Infinite desktop"
                strClass = "IDT"
        End Select
        If optButton.Value = 1 Then
            ReDim Preserve GroupList(0 To Count)
            ReDim Preserve GroupClass(0 To Count)
            GroupList(Count) = strName
            GroupClass(Count) = strClass
            Count = Count + 1
        End If
    Next k
    
    Rem - Determine the settings that needto be activated.
    If CallingForm.chkAWT.Value = 1 Then
        Call Prop_Set("AWT", "Y", ZW_Settings)
    Else
        Call Prop_Set("AWT", "", ZW_Settings)
    End If
    If CallingForm.chkIDT.Value = 1 Then
        Call Prop_Set("IDT", "Y", ZW_Settings)
    Else
        Call Prop_Set("IDT", "", ZW_Settings)
    End If
    If CallingForm.chkDTM.Value = 1 Then
        Call Prop_Set("DTM", "Y", ZW_Settings)
    Else
        Call Prop_Set("DTM", "", ZW_Settings)
    End If
    
End Sub



Private Sub mnuDelete_Click()
Dim lngParent As Long
Dim lngEnd As Long
Dim booGroup As Boolean

    Rem - Ensure that they do not delete everything
    ZIndex = Val(Prop_Get("Index", lsvGroup.SelectedItem.key))
    booGroup = CBool(ZKMenu(ZIndex)("Class") = "Group")
        
    If ZIndex = GRP_Start Then ' The first item is being deleted
        Call ZenMB("Sorry, but you cannot delete this menu.", "OK")
        Exit Sub
    End If
    
    If ZKMenu(ZIndex)("Class") = "Group" Then
        If ZenMB("You are deleting a group, which will delete all the items inside the group. Are you sure you wish to do this?", "Yes", "No") = 0 Then
            Rem - Delete the group
            lngEnd = GetGroupEnd(ZIndex)
            Rem - Deletre Group
            Call Array_Up(lngEnd + 1, lngEnd - ZIndex + 1)
            Rem - Refresh the tree
            Call Group_Init(GRP_Index)
            Call Tree_Load
        End If
    Else
        Rem - Delete the item if not the only item in the group
        
        If GRP_End - GRP_Start < 3 Then
            Rem - It is the last item
            Call ZenMB("You cannot delete the last item in a group. Rather just delete the group itself?", "OK")
        Else
            Rem - Deletre Item
            Call Array_Up(ZIndex + 1, 1)
            Rem - Refresh the tree
            Call Group_Init(GRP_Index)
            Call Tree_Load
        End If
        
    End If

End Sub

Private Sub mnuEdit_Click()

    TimerAction = "Edit"

End Sub


Private Sub mnuEditHK_Click()
    
    Rem - Ensure that they do not delete everything
    ZIndex = Val(Prop_Get("Index", lsvGroup.SelectedItem.key))
    If HK_Selected Then Call Tree_Load

End Sub

Private Sub zbMenuHK_Click()

    If Not booLoading Then
        booLoading = True
        ZIndex = GRP_Start
        If HK_Selected Then Call Group_ShowHK
        booLoading = False
    End If
    
End Sub


Private Sub mnuNew_Click()

    TimerAction = "New"
    
End Sub

Private Sub zbNext_Click()
    
    
    Call Group_Set
    
    If GRP_Index >= UBound(GroupList()) Then
        Call ZW_Next("Next")
    Else
        Call Group_Init(GRP_Index + 1)
        Call Group_Display
    End If

End Sub



Private Function GetCaption(ByVal ZKIndex As Long) As String
Dim strTemp As String
Dim strShift As String
Dim strKey As String


    GetCaption = ZKMenu(ZKIndex)("Caption")
    If chkEnableHK.Value = 1 Then

        Rem - Load the ZenKey objects as we build the menu
        Rem - Add the hotkey to the caption if it has one
        strShift = ZKMenu(ZKIndex)("ShiftKey")
        strKey = ZKMenu(ZKIndex)("Hotkey")
        If Len(strShift & strKey) > 0 Then
            Rem - If either key has used, add the key names to the caption
            strTemp = HotKeys.Keyname(Val(strKey))
            If Len(strTemp) > 0 Then strKey = strTemp Else strKey = "Ext-" & strKey
            If (Len(strKey) = 0) Or (Len(strShift) = 0) Then
                GetCaption = GetCaption & ", Hotkey = " & strShift & strKey
            Else
                GetCaption = GetCaption & ", Hotkey = " & strShift & " + " & strKey
            End If
            'Call Prop_Set("Caption", Prop_Get("Caption", ZKMenu(k)) & strTemp, ZKMenu(k))
        End If
    End If

End Function

Private Sub Group_Display()
    
    If ZKMenu(GRP_Start)("DisableHK") = "True" Then
        chkEnableHK.Value = 0
    Else
        chkEnableHK.Value = 1
    End If
    Call Tree_Load
    
End Sub

Private Sub Group_Set()
    
    If chkEnableHK.Value = 0 Then
        ZKMenu(GRP_Start)("DisableHK") = "True"
    Else
        ZKMenu(GRP_Start)("DisableHK") = ""
    End If
    
End Sub

Private Sub Group_Init(ByVal Index As Long)
    
    GRP_Index = Index
    GRP_Start = GetGroup(GroupList(GRP_Index))
    GRP_End = GetGroupEnd(GRP_Start)
    
End Sub

Private Sub Tree_AddItem(ByVal Index As Long)
Dim itmNew As ListItem
                
    Set itmNew = lsvGroup.ListItems.Add()
    itmNew.Text = ZKMenu(Index)("Caption")
    itmNew.key = "|Index=" & CStr(Index) & "|"
    
    If chkEnableHK.Value = 1 Then
        Rem - Load the ZenKey objects as we build the menu
        Rem - Add the hotkey to the caption if it has one
        itmNew.SubItems(1) = HotKeys.GetCaption(ZKMenu(Index))
    Else
        itmNew.SubItems(1) = "- none -" 'HotKeys.GetCaption(ZKMenu(Index))
    End If
            
    

End Sub

Private Function HK_Selected() As Boolean
Rem - NOTE : ZIndex must be set
Dim It As frmSetHotKey

    Set It = New frmSetHotKey
    Set MainForm = Me
    With It
        Set .CallingForm = Me
        .prop = ZKMenu(ZIndex)
        .Init
        .EditIndex = ZIndex
        .Show
        Me.Hide
        While .Visible
            DoEvents
        Wend
        HK_Selected = .booValid
        If HK_Selected Then Set ZKMenu(ZIndex) = .prop
        Me.Show
    End With
    Unload It
    Set It = Nothing

End Function

Private Sub Group_ShowHK()
Dim strHK As String
    
    strHK = HotKeys.GetCaption(ZKMenu(GRP_Start))
    If Len(strHK) > 0 Then
        chkMenuHK.Value = 1
    Else
        chkMenuHK.Value = 0
        strHK = "some keys"
    End If
    
    With zbMenuHK
        .Visible = False
        .Caption = strHK
        Set Me.Font = .Font
        .Visible = True
    End With
    
    
End Sub

Private Sub mnuRemoveHK_Click()
    Rem - Ensure that they do not delete everything
    ZIndex = Val(Prop_Get("Index", lsvGroup.SelectedItem.key))
    ZKMenu(ZIndex)("Hotkey") = vbNullString
    ZKMenu(ZIndex)("ShiftKey") = vbNullString
    Call Tree_Load
    
End Sub


Private Sub mnuTest_Click()
Dim ZIndex As Integer ' So it does not clash with the real one.....

    ZIndex = Val(Prop_Get("Index", lsvGroup.SelectedItem.key))
    Select Case ZKMenu(ZIndex)("Class")
        Case "Group"
            Call ZenMB("Sorry, but menus cannot be tested. Adding a menu adds a menu item which then contains other items.")
        Case "ZenKEY"
            Call ZenMB("Sorry, but ZenKEY actions cannot be tested")
        Case Else
            Call TestAction(ZKMenu(ZIndex))
    End Select

End Sub



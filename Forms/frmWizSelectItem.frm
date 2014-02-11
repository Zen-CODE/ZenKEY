VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmZenWSelectItem 
   Caption         =   "| ZenKEY - Wizard |"
   ClientHeight    =   4500
   ClientLeft      =   5130
   ClientTop       =   2850
   ClientWidth     =   8955
   Icon            =   "frmWizSelectItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   Begin VB.TextBox txtHotkey 
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Click inside this textbox and press the Hotkeys you use to fire the item you wish to find"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtName 
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
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "Type in the beginning lettes of the item you wish to find"
      Top             =   2040
      Width           =   1935
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   4
      Top             =   3945
      Width           =   1455
   End
   Begin MSComctlLib.TreeView tvTree 
      Height          =   2475
      Left            =   2640
      TabIndex        =   6
      ToolTipText     =   "Click on the item you wish to select"
      Top             =   1200
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4366
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   265
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlTree"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTree 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSelectItem.frx":058A
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSelectItem.frx":08DC
            Key             =   "Moving"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSelectItem.frx":0C2E
            Key             =   "FolderOpen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSelectItem.frx":0F80
            Key             =   "Action"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "with a hotkey using"
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
      TabIndex        =   9
      ToolTipText     =   "Click inside this textbox and press the Hotkeys you use to fire the item you wish to find"
      Top             =   2520
      Width           =   1470
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "with the name containing"
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
      ToolTipText     =   "Type in the beginning lettes of the item you wish to find"
      Top             =   1680
      Width           =   1905
   End
   Begin VB.Label lblCriteria 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search for"
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
      ToolTipText     =   "Use the boxes below to search for an item"
      Top             =   1320
      Width           =   780
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the item you wish to delete."
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
      Left            =   3210
      TabIndex        =   5
      Top             =   780
      Width           =   2625
   End
End
Attribute VB_Name = "frmZenWSelectItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Dim lngShift As Long
Private Const ColDisabled = &H808080   '&HC0C0C0


Public Sub Tree_Load()
Dim k As Long
Dim max As Long
Dim nodX As Node
Dim strCap As String
Dim nParent As Node

    max = UBound(ZKMenu()) - 2
    Set nParent = Nothing
    
    With tvTree
        .Visible = False
        .Nodes.Clear
        For k = 2 To max
            Rem - Add the group item
            If ZKMenu(k)("Class") = "Group" Then
                Rem - Start a new group
                k = Tree_Load_Group(-1, k)
            Else
                Rem - Just add to the current group
                Set nodX = .Nodes.Add(, tvwLast, "|Index=" & CStr(k) & "|", ZKMenu(k)("Caption"))
                nodX.ForeColor = IIf(ZKMenu(k)("Disabled") = "True", ColDisabled, vbBlack)
                nodX.EnsureVisible
                nodX.Image = "Action"
            End If
        Next k
        .Visible = True
    End With
    
End Sub

Public Function Tree_Load_Group(ByVal ParentIndex As Long, ByVal StartIndex As Long) As Long
Rem - Returns the Ending index of the group

Dim k As Long
Dim max As Long
Dim nodX As Node

    Rem - Start a new group
    max = UBound(ZKMenu())
    
    With tvTree
        If ParentIndex = -1 Then
            Rem - In root menu
            Set nodX = .Nodes.Add(, tvwLast, "|Index=" & CStr(StartIndex) & "|", ZKMenu(StartIndex)("Caption"))
        Else
            Rem - In a submenu
            Set nodX = .Nodes.Add("|Index=" & CStr(ParentIndex) & "|", tvwChild, "|Index=" & CStr(StartIndex) & "|", ZKMenu(StartIndex)("Caption"))
        End If
        nodX.ForeColor = IIf(ZKMenu(StartIndex)("Disabled") = "True", ColDisabled, vbBlack)
                
        nodX.Image = "Folder"
        nodX.ExpandedImage = "FolderOpen"

    
        Rem - Nowq add all its sub items
        For k = StartIndex + 1 To max
            Rem - Add the group item
            If ZKMenu(k)("EndGroup") = "True" Then
                Set nodX = nodX.Parent
                nodX.Expanded = False
                Exit For
            End If
            If ZKMenu(k)("Class") = "Group" Then
                Rem - Add another sub group
                k = Tree_Load_Group(StartIndex, k)
            Else
                Rem - Add the item to the group
                Set nodX = .Nodes.Add("|Index=" & CStr(StartIndex) & "|", tvwChild, "|Index=" & CStr(k) & "|", ZKMenu(k)("Caption"))
                nodX.Image = "Action"
            End If
            nodX.ForeColor = IIf(ZKMenu(k)("Disabled") = "True", ColDisabled, vbBlack)
            nodX.EnsureVisible
            
        Next k
        Tree_Load_Group = k
    End With

End Function


Private Sub Form_Activate()
    tvTree.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static lngKey As Long
Static lngShift As Long
Dim zdTemp As New clsZenDictionary

    Rem - Establish the Shiftkey
    If Me.ActiveControl.Name <> txtHotkey.Name Then Exit Sub
    If Shift <> 0 Then
        Rem - Here 1 = Shift          , Windows shift = 4
        Rem - Here 4  = Alt         , Windows 1 = ALt
        Rem = Here 2 = Cntrl       , Window 2 =Control
        lngShift = 0
        If Shift >= 4 Then lngShift = 1: Shift = Shift - 4
        If Shift >= 2 Then lngShift = lngShift + 2: Shift = Shift - 2
        If Shift >= 1 Then lngShift = lngShift + 4: Shift = Shift - 4
    End If

    Rem - Establish the Hotkey
    If (KeyCode >= HK_MIN) Then lngKey = KeyCode

    If lngShift > 0 Then zdTemp("ShiftKey") = HotKeys.ShiftValToStr(lngShift)
    If lngKey > 0 Then zdTemp("Hotkey") = lngKey
    txtHotkey.Text = HotKeys.GetCaption(zdTemp)

Dim lngIndex As Long
Dim k As Long, strShift As String

    Rem - Now search the ZKMenu for a match
    strShift = zdTemp("ShiftKey")
    If lngKey + lngShift > 0 Then
        For k = UBound(ZKMenu()) To 0 Step -1
            If Val(ZKMenu(k)("Hotkey")) = lngKey Then
                If ZKMenu(k)("ShiftKey") = strShift Then
                    lngIndex = k
                    Exit For
                End If
            End If
        Next k
    End If

    Rem - Now locate it in the tree
    If lngIndex > 0 Then
        With tvTree
            For k = .Nodes.Count To 1 Step -1
                Rem - Any name starting with
                If Val(Prop_Get("Index", .Nodes(k).key)) = lngIndex Then
                    .SelectedItem = .Nodes(k)
                    .SetFocus
                    Exit Sub
                End If
            Next k
        End With
    End If
    KeyCode = 0
    Shift = 0

End Sub


Private Sub Form_Load()
    
    Set zbNext.Picture = zbBack.Picture
    Set zbCancel.Picture = zbBack.Picture

'    Call AddKeyDescrip
    Call Tree_Load
    If Mode = "NewItem" Then
        lblMessage.Caption = "Please select the item before which the new item should be inserted."
    End If
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then End
End Sub

Private Sub tvTree_DblClick()
    Call zbNext_Click
End Sub


Private Sub txtHotkey_KeyPress(KeyAscii As Integer)
    KeyAscii = 0

End Sub

Private Sub txtName_Change()
Dim max As Integer
Dim strText As String
Dim k As Integer

    strText = txtName.Text
    If Len(strText) > 0 Then
        With tvTree
            max = .Nodes.Count
            For k = 1 To max
                Rem - Any name starting with
                If left(.Nodes(k).Text, Len(strText)) = strText Then
                    .SelectedItem = .Nodes(k)
                    Exit Sub
                End If
            Next k
            
        End With
    End If
    
End Sub

Private Sub zbBack_Click()
    Call ZW_Next("Previous")
End Sub


Private Sub zbCancel_Click()
    End
End Sub


Private Sub AddKeyDescrip()
Dim strShift As String
Dim strKey As String
Dim max As Long
Dim k As Long
Dim strTemp As String

    max = UBound(ZKMenu())
    For k = 0 To max
        strShift = ZKMenu(k)("ShiftKey")
        strKey = ZKMenu(k)("Hotkey")
        If Len(strShift & strKey) > 0 Then
            Rem - If either key has used, add the key names to the caption
            strTemp = HotKeys.Keyname(Val(strKey))
            If Len(strTemp) > 0 Then strKey = strTemp Else strKey = "Ext-" & strKey
            If (Len(strKey) = 0) Or (Len(strShift) = 0) Then
                strTemp = strShift & strKey
            Else
                strTemp = strShift & " + " & strKey
            End If
            strTemp = ZKMenu(k)("Caption") & " (" & strTemp & ")"
            ZKMenu(k)("Caption") = strTemp
        End If
    Next k
    
End Sub

Private Sub zbNext_Click()
    ZIndex = CLng(Prop_Get("Index", tvTree.SelectedItem.key))
    Call ZW_Next("Next")
End Sub



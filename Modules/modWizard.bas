Attribute VB_Name = "modWizard"
Option Compare Text
Option Explicit
Private CurrentForm As Form
Public PicBack As StdPicture
Public Const HK_MIN = 32
Public Const HK_MAX = 255
Public Mode As String
Public ZIndex As Long
Private WizForms() As Form
Private WizDepth As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public ZW_NewItem As New clsZenDictionary
Public GroupList() As String ' The list of New groups names in thre "New configuration" option
Public ZW_Settings As String
Public Function Extract(ByVal Sentance As String, ByVal AfterNthSep As Long, ByVal Separator As String) As String
Rem - Pumps the pipe separated items into Items()
Dim k As Integer, intEnd As Integer

    intEnd = InStr(Sentance, Separator)
    For k = 0 To AfterNthSep - 1
        If intEnd > 0 Then
            Sentance = Mid$(Sentance, intEnd + 1)
        Else
            Sentance = vbNullString
        End If
        intEnd = InStr(Sentance, Separator)
    Next k
    intEnd = InStr(Sentance, Separator)
    If intEnd > 0 Then Extract = left$(Sentance, intEnd - 1) Else Extract = Sentance

End Function

Private Function ZW_EditItem() As Boolean

Dim ActForm As frmAction
    
    Set ActForm = New frmAction
    With ActForm
        Set .CallingForm = CurrentForm 'WizForms(WizDepth)
        Set .prop = ZKMenu(ZIndex)
        .EditIndex = ZIndex
        .Init
        Call CentreForm(ActForm)
        CurrentForm.Visible = False
        .Show
        While Not .booDone
            DoEvents
        Wend
        
        If .booValid Then Set ZKMenu(ZIndex) = .prop
        ZW_EditItem = .booValid
        
    End With
    
    Unload ActForm
    Set ActForm = Nothing


End Function



Public Sub Array_Down(ByVal Start As Long, ByVal Num As Long)
Dim k As Integer
Dim max As Integer

    max = UBound(ZKMenu())
    
    ReDim Preserve ZKMenu(max + Num)
    
    For k = max + Num To Start + Num Step -1
        Set ZKMenu(k) = ZKMenu(k - Num)
    Next k
    

End Sub

Private Function Item_GetGroupEnd(ByVal GrpIndex As Long) As Long
Dim max As Long, k As Long

    max = UBound(ZKMenu())
    k = GrpIndex + 1
    
    Do
        If ZKMenu(k)("Class") = "Group" Then
            k = Item_GetGroupEnd(k)
        ElseIf ZKMenu(k)("EndGroup") = "True" Then
            Item_GetGroupEnd = k
            Exit Function
        End If
        k = k + 1
    Loop While k < max
    Item_GetGroupEnd = k
    
End Function


Public Function ZK_Running() As Boolean

    ZK_Running = CBool(Val(Registry.GetRegistry(HKCU, "SOFTWARE\ZenCODE\ZenKEY", "WindowHandle")) > 0)
    
End Function

Public Sub ZW_Next(ByVal FormType As String)

    ReDim Preserve WizForms(WizDepth)
    Select Case FormType
        Case "Start"
            Rem - Initiliase stuff
            Call Actions_Load
            Call LoadArray(Searches(), App.Path & "\Search.ini")

            Rem - Initialise the form array
            WizDepth = 0
            Set CurrentForm = Forms(0)
            CurrentForm.Visible = False
        Case "Type"
            Rem - Select the 'Mode', whether adding, deleting, editing or seting up an entirely new config
            Set CurrentForm = New frmZenWType
            WizDepth = WizDepth + 1
        Case "Previous"
            Rem - Return to the previous from
            WizDepth = WizDepth - 1
            Unload WizForms(WizDepth + 1)
            Set CurrentForm = WizForms(WizDepth)
        Case "NewConfig"
            Rem - Setup an entirely new cofniguration
            Mode = "NewConfig"
            Call Menu_LoadINI(App.Path + "\Default.ini", True)
            Set CurrentForm = New frmWizNCOptions
            WizDepth = WizDepth + 1
        Case "Edit"
            Rem - Enter 'Edit' mode
            Rem - Sequence - |?
            Mode = "Edit"
            Call Menu_LoadINI("ZenKEY.ini", True)
            Set CurrentForm = New frmZenWSelectItem
            CurrentForm.lblMessage.Caption = "Select the item you wish to edit."
            WizDepth = WizDepth + 1
        Case "NewItem"
            Rem - Enter 'NewIterm' mode
            Rem - Sequence - | NewItem -> Hotkey -> Select item (to place before) -> Finnish
            Mode = "NewItem"
            Set CurrentForm = New frmWizNewItem
            Call CurrentForm.Init
            Call Menu_LoadINI("ZenKEY.ini", True)
            WizDepth = WizDepth + 1
        Case "Remove"
            Rem - Enter 'Remove' mode
            Rem - Sequence - | Select -> Finnish
            Mode = "Remove"
            Set CurrentForm = New frmZenWSelectItem
            Call Menu_LoadINI("ZenKEY.ini", True)
            WizDepth = WizDepth + 1
        Case "Next"
            Rem - Proceed to te next item in dialog mode
            Select Case Mode
                Case "NewConfig"
                    If WizDepth < 3 Then
                        Rem - The first screen. Let them choose the menu's they wish to use
                        Set CurrentForm = New frmWizNewConfigs
                        Set CurrentForm.CallingForm = WizForms(WizDepth)
                        Call CurrentForm.Init
                    Else
                        Rem - All done. Confirm and exit
                        Set CurrentForm = New frmZenWDone
                        CurrentForm.lblMessage.Caption = "You have finished. All is good."
                        If Not ZK_Running Then CurrentForm.chkRestart.Caption = vbNullString
                    End If
                    WizDepth = WizDepth + 1
                Case "Remove"
                    Set CurrentForm = New frmZenWDone
                    If ZKMenu(ZIndex)("Class") = "Group" Then
                        CurrentForm.lblMessage.Caption = "You are about to delete the menu '" & ZKMenu(ZIndex)("Caption") & "' and all its contents."
                    Else
                        CurrentForm.lblMessage.Caption = "You are about to delete '" & ZKMenu(ZIndex)("Caption") & "'."
                    End If
                    WizDepth = WizDepth + 1
                Case "NewItem"
                    Select Case WizDepth
                        Case 3 ' After Hotkey dialog, find out where they want to add it
                            Set CurrentForm = New frmZenWSelectItem
                            WizDepth = WizDepth + 1
                        Case 4 ' They have chosen where. Now finish up
                            Set CurrentForm = New frmZenWDone
                            CurrentForm.lblMessage.Caption = "You are about to Add a new item, '" & ZW_NewItem("Caption") & "'."
                            WizDepth = WizDepth + 1
                        Case 2 ' Show hotkey dialog, then Finnish
                            Set CurrentForm = New frmWizHotkey
                            Call CurrentForm.Init
                            WizDepth = WizDepth + 1
                    End Select
                Case "Edit"
                    CurrentForm.Visible = False
                    If Not ZW_EditItem Then
                        CurrentForm.Visible = True
                        Exit Sub
                    Else
                        Set CurrentForm = New frmZenWDone
                        CurrentForm.lblMessage.Caption = "You have edited the item, '" & ZKMenu(ZIndex)("Caption") & "'."
                        WizDepth = WizDepth + 1
                    End If
            End Select
            
            
        Case "Finnish"
            Select Case Mode
                Case "Edit"
                    ' Do nothing. The editing has been done
                Case "Remove"
                    Call ZW_Remove
                Case "NewItem"
                    Call ZW_AddNew(ZW_NewItem)
                Case "NewConfig"
                    Call ZW_ProcessINI
            End Select
            Call SaveToINI
            Rem - See if we should restart ZenKEY
            If CurrentForm.chkRestart.Value = 1 Then Call ZK_Restart
            
            
            End
    End Select

    ReDim Preserve WizForms(0 To WizDepth)
    Set WizForms(WizDepth) = CurrentForm
    Set CurrentForm.Picture = PicBack
    Call CentreForm(CurrentForm)
    CurrentForm.Visible = True
    If WizDepth > 0 Then WizForms(WizDepth - 1).Visible = False
        
End Sub
Private Sub SaveToINI()
Dim intFNum As Integer
Dim k As Integer
Dim intItemMax As Integer
    
    On Error GoTo ErrorTrap
    
    Rem ======================     Open file and start loading
    Rem - Initialise variables
    intFNum = FreeFile
    intItemMax = UBound(ZKMenu())
    
    Open settings("SavePath") & "\Zenkey.ini" For Output As #intFNum
            Print #intFNum, "====================== Zenkey initialisation file ======================"
            'For intSection = 0 To SecIndexes
            For k = 2 To intItemMax - 2
                Rem - Initialise for a new section
                If ZKMenu(k)("Class") = "Group" Then Print #intFNum, "====================== " & ZKMenu(k)("Caption") & " ======================"
                Print #intFNum, ZKMenu(k).ToProp
            Next k
    Close #intFNum
    Exit Sub
    
ErrorTrap:
    Call MsgBox("Unable to save settings. Please ensure you have appropriate permissions (Sub SaveToINI).")
    Err.Clear
End Sub




Public Sub ZW_Remove()
Dim lngParent As Long
Dim lngEnd As Long
Dim strProp As String
Dim booGroup As Boolean

    Rem - Ensure that they do not delete everything
    booGroup = CBool(ZKMenu(ZIndex)("Class") = "Group")
        
    If ZIndex = 2 Then ' The first item is being deleted
        If booGroup Then
            lngEnd = Item_GetGroupEnd(ZIndex)
            If lngEnd > UBound(ZKMenu()) - 3 Then
                Call ZenMB("Sorry, but you cannot delete the last item in ZenKEY. Otherwise, why bother, really!", "OK")
                Exit Sub
            End If
        ElseIf UBound(ZKMenu()) < 5 Then
            Call ZenMB("Sorry, but you cannot delete the last item in ZenKEY. Otherwise, why bother, really!", "OK")
            Exit Sub
        End If
    End If
    
    If ZKMenu(ZIndex)("Class") = "Group" Then
        If ZenMB("You are deleting a menu, which will delete all the items inside the menu. Are you sure you wish to do this?", "Yes", "No") = 0 Then
            Rem - Delete the group
            lngEnd = Item_GetGroupEnd(ZIndex)
            Rem - Deletre Group
            Call Array_Up(lngEnd + 1, lngEnd - ZIndex + 1)
        End If
    Else
        Rem - Delete the item if not the only item in the group
        lngParent = Item_GetGroup(ZIndex)
        lngEnd = Item_GetGroupEnd(lngParent)
        
        If lngEnd - lngParent < 3 Then
            Rem - It is the last item
            Call ZenMB("You cannot delete the last item in a menu. Rather just delete the menu itself?", "OK")
        Else
            Rem - Deletre Item
            Call Array_Up(ZIndex + 1, 1)
        End If
    End If
    
End Sub
Private Function Item_GetGroup(ByVal ItemIndex As Long) As Long
Dim Min As Long, k As Long

    Min = -1
    k = ItemIndex - 1
    
    Do
        If ZKMenu(k)("EndGroup") = "True" Then
            k = Item_GetGroup(k)
            k = k - 1
        ElseIf ZKMenu(k)("Class") = "Group" Then
            Item_GetGroup = k
            Exit Function
        End If
        k = k - 1
    Loop While k > Min
    Item_GetGroup = Min
    
End Function


Public Sub Array_Up(ByVal Start As Long, ByVal Num As Long)
Dim k As Integer
Dim max As Integer

    max = UBound(ZKMenu())
    
    For k = Start To max
        Set ZKMenu(k - Num) = ZKMenu(k)
    Next k
    
    ReDim Preserve ZKMenu(max - Num)
    
    
End Sub

Private Sub ZW_AddNew(ByVal zdAct As clsZenDictionary)
    
    If zdAct("Class") = "Group" Then
        Rem - Adding a group
        Call Array_Down(ZIndex, 3)
        Set ZKMenu(ZIndex) = zdAct
        Set ZKMenu(ZIndex + 1) = zenDic("Class", "ZenKEY", "Action", "About", "Caption", "Item in new menu")
        Set ZKMenu(ZIndex + 2) = zenDic("ENDGROUP", "True")
        
    Else
        Rem - Adding a singular item
        Call Array_Down(ZIndex, 1)
        Set ZKMenu(ZIndex) = zdAct
    End If

End Sub


Private Sub ZW_ApplyConfig()
Dim k As Long
Dim i As Long
Dim max As Long
Dim lngEnd As Long
    
    Rem - Remove the Hotkeys for all the Groups where they have chosen to disable
    max = UBound(ZKMenu())
    For k = 2 To max
        If ZKMenu(k)("Class") = "Group" Then
            If ZKMenu(k)("DisableHK") = "True" Then
                lngEnd = GetGroupEnd(k)
                For i = k + 1 To lngEnd
                    ZKMenu(i)("ShiftKey") = vbNullString
                    ZKMenu(i)("Hotkey") = vbNullString
                Next i
                k = lngEnd
            Else
                k = GetGroupEnd(k)
            End If
        End If
    Next k

    

End Sub
Public Function GetGroupEnd(ByVal GrpIndex As Long) As Long
Dim max As Long, k As Long

    max = UBound(ZKMenu())
    k = GrpIndex + 1
    
    Do
        If ZKMenu(k)("Class") = "Group" Then
            k = GetGroupEnd(k)
        ElseIf ZKMenu(k)("EndGroup") = "True" Then
            GetGroupEnd = k
            Exit Function
        End If
        k = k + 1
    Loop While k < max
    GetGroupEnd = k
    
End Function


Private Sub ZW_ProcessINI()
Dim max As Long
Dim k As Long
Dim colItems As Collection

    Rem ==========================================================
    Rem - Save the menu items
    Rem ==========================================================
    Set colItems = New Collection
    With colItems
        .Add "My Programs", "My Programs"
        .Add "My Documents", "My Documents"
        .Add "My window", "My window"
        .Add "Control Panel", "Control Panel"
        .Add "My Infinite desktop", "My Infinite desktop"
        .Add "Internet resources", "Internet resources"
        .Add "Winamp controls", "Winamp controls"
        .Add "Windows Folders", "Windows Folders"
        .Add "Windows Programs", "Windows Programs"
        .Add "Windows Media Commands", "Windows Media Commands"
        .Add "Windows System", "Windows System"
        .Add "Windows System folders", "Windows System folders"
        .Add "Windows Utilities", "Windows Utilities"
        .Add "Windows Management", "Windows Management"
    
        Rem - Now remove the items that are not to be removed
        max = UBound(GroupList())
        For k = 0 To max
            Select Case GroupList(k)
                Case "My programs", "My documents", "Internet resources", "My window", "Windows System", "Windows Media commands"
                    .Remove GroupList(k)
                Case "Winamp Controls", "My Infinite desktop", "Control Panel"
                    .Remove GroupList(k)
                Case "Windows Folders"
                    .Remove "Windows Folders"
                    .Remove "Windows System folders"
                Case "Windows programs"
                    .Remove "Windows Programs"
                    .Remove "Windows Utilities"
                    .Remove "Windows Management"
            End Select
        Next k
    
        Rem - Now we are left only with items that should be remove. Take em out!
        Dim lngEnd As Long
        While .Count > 0
            k = GetGroup(.Item(1))
            lngEnd = Item_GetGroupEnd(k)
            Call Array_Up(lngEnd + 1, lngEnd - k + 1)
            .Remove 1
        Wend
        
        Rem - Now strip out the Hotkeys that have been disabled by group
        Dim i As Long
        max = UBound(ZKMenu())
        For k = 0 To max
            If ZKMenu(k)("Class") = "Group" Then
                If ZKMenu(k)("DisableHK") = "True" Then
                    ZKMenu(i)("DisableHK") = vbNullString
                    lngEnd = GetGroupEnd(k)
                    For i = k + 1 To lngEnd
                        ZKMenu(i)("Hotkey") = vbNullString
                        ZKMenu(i)("ShiftKey") = vbNullString
                    Next i
                End If
            End If
        Next k
    End With
    
    Rem ==========================================================
    Rem - Save the settings
    Rem ==========================================================
    If Prop_Get("IDT", ZW_Settings) = "Y" Then
        Rem- Enable the Infinite desktop
        settings("IDT_Enable") = "Y"
        settings("IDT_VisOnExit") = "Y"
        settings("DTM_Position") = "5"
        settings("DTM_Refresh") = "2"
        settings("DTM_Show") = "5"
        settings("DTM_Layer") = "1"
    Else
        settings("IDT_Enable") = ""
    End If
    
    If Prop_Get("AWT", ZW_Settings) = "Y" Then
        Rem - Enable Auto-window transparency
        settings("AutoTrans") = "True"
        settings("AWTDepth") = "1"
        settings("TransActive") = "Opaque"
        settings("TransInactive") = "30%"
    Else
        settings("AutoTrans") = ""
    End If
    
    If Prop_Get("DTM", ZW_Settings) = "Y" Then
        Rem- Enable the Desktop map
        settings("DTM_Position") = "2"
        settings("DTM_Show") = "6"
    Else
        settings("DTM_Position") = "6"
    End If
    Call settings.ToINI("Settings.ini")
    
End Sub
Public Function GetGroup(ByVal GrpName As String) As Long
Dim k As Long
Dim max As Long

    max = UBound(ZKMenu()) - 2
    For k = 2 To max
        If ZKMenu(k)("Class") = "Group" Then
            If ZKMenu(k)("Caption") = GrpName Then
                GetGroup = k
                Exit Function
            End If
        End If
    Next k
    Call MsgBox("Unable to find the menu '" & GrpName & "'.", vbInformation)
        
End Function


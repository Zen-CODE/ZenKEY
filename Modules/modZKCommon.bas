Attribute VB_Name = "modZKCommon"
Option Explicit
Option Compare Text
Public ZKMenu() As clsZenDictionary
Rem - For Optimization
#If ZKCONFIG <> 1 Then
    #If ZenWiz <> 1 Then
        Rem - For ZenKEY
        Public ZK_Win As clsWindows
        Public ZK_DTM As frmDesktopMap
        Public ZK_IDT As clsIDT
    #Else
        Rem - For ZenWizard
        Public ZK_Win As Object
        Public ZK_DTM As Object
        Public ZK_IDT As Object
    #End If
#Else
    Rem - For ZenConfig
    Public ZK_Win As Object
    Public ZK_DTM As Object
#End If

Public Const ZenKEYCap = "   ---   ZenKEY   ---   "
Public booKill As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Rem ======================== For special folder ====================
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public settings As New clsZenDictionary
Public progInfo As New clsZenDictionary
Rem ====================== Variables for standard colouring ====================
Public Declare Function GetTickCount& Lib "kernel32" ()
Public WIN_ZKAction As Long ' For passing messages to active instance
Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long ' For XP theming
Public COL_Zen As OLE_COLOR

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
Public Sub TileMe(ByRef DaForm As Form, ByRef Pic As StdPicture)
Dim sngWidth As Single, sngHeight As Single
Dim k As Single, i As Single
    
    With DaForm
        .AutoRedraw = True
        Set .Picture = Pic
        sngWidth = .ScaleX(.Picture.Width, vbHimetric, .ScaleMode)
        sngHeight = .ScaleY(.Picture.Height, vbHimetric, .ScaleMode)
        For i = 0 To .ScaleHeight Step sngHeight
            For k = 0 To .ScaleWidth Step sngWidth
                .PaintPicture .Picture, k, i
            Next k
        Next i
        .AutoRedraw = False
    End With
    
        
End Sub



Public Function ZenMB(ByVal Message As String, ParamArray Caps()) As Long
Rem - If no parameters are passed, no buttons, disappears in 2.5 secs
Rem - Else puts each parameter in a button, and returns the caption they clicked
Dim k As Long, frmMB As New frmMB

On Error GoTo ErrorTrap

    If Forms.Count < 10 Then
        With frmMB
            .lblTop = Message
            If Not (settings("HideQuotes") = "True") Then
                .lblBot = Get_Koan
            Else
                .lblBot = vbNullString
            End If
            .lngButtons = UBound(Caps())
            .Initialise
            If .lngButtons < 0 Then
                Rem - Don't show modal..
                .Show
            Else
                Rem - Show modally
                For k = .lngButtons To 0 Step -1
                    .zbButton(k).Caption = Caps(k)
                Next k
                Rem - Buttons. Set result
                .Show vbModal
                ZenMB = .lngButtons
            End If
        End With
    End If
    Exit Function
    
ErrorTrap:
    Call MsgBox("Error " & CStr(Err.Number) & ", " & Err.Description, vbQuestion)
    Err.Clear
End Function




Public Sub Menu_LoadINI(ByVal FileName As String, Optional LoadDisabled As Boolean = False)
Dim lngFNum As Long
Dim lngIndex As Long
Dim strSection As String
Dim strTemp As String
Dim strLine As String
Dim booSkip As Boolean

On Error GoTo ErrorTrap
    
    Rem ======================     Open file and start loading
    Rem - Initialise variables
    lngFNum = FreeFile
    lngIndex = -1
    
    If InStr(FileName, "\") = 0 Then FileName = settings("SavePath") & "\" & FileName ' Add full path if none given
    Open FileName For Input As #lngFNum
    
        While Not EOF(lngFNum)
            Line Input #lngFNum, strSection ' Read section heading
            If Not (left$(strSection, 1) = "=") Then
            
                booSkip = False
                If Not LoadDisabled Then
                    booSkip = CBool(Prop_Get("Disabled", strSection) = "True")
                End If
                Select Case Prop_Get("Class", strSection)
                    Case "Group"
                        Rem - Determine whether the group should be loaded or not
                        If booSkip Then Call p_SkipGroup(lngFNum)
                    Case "ZenKEY"
                        Rem - These captions have been misleading previously, so correct them and set their state
                        Rem - for the menu (Tick or Cross)
                        Select Case Prop_Get("Action", strSection)
                            Case "FOLLOWACTIVE": Call Prop_Set("Menu", IIf(settings("FollowActive") = "True", "T", "C"), strSection)
                            Case "HIDEFORM": Call Prop_Set("Menu", IIf(settings("HideForm") = "True", "T", "C"), strSection)
                            Case "WindowUnderMouse"
                                Call Prop_Set("Caption", "Act on Window under mouse (else Active window)", strSection)
                                Call Prop_Set("Menu", IIf(settings("WindowUnderMouse") = "True", "T", "C"), strSection)
                            Case "SETAOT"
                                Call Prop_Set("Menu", IIf(settings("AOT") = "True", "T", "C"), strSection)
                            Case "TOGGLEAWT"
                                Call Prop_Set("Caption", "Auto-window transparency", strSection)
                                Call Prop_Set("Menu", IIf(settings("AutoTrans") = "True", "T", "C"), strSection)
                            Case "TOGGLEHOTKEYS"
                                Call Prop_Set("Caption", "Unload Hotkeys", strSection)
                                Call Prop_Set("Menu", "C", strSection)
                            Case "TOGGLEHKEX"
                                Call Prop_Set("Caption", "Unload Hotkeys but not this one", strSection)
                                Call Prop_Set("Menu", "C", strSection)
                        End Select
                End Select
                                
                Rem - Load the item if we can should
                If Not booSkip Then
                    lngIndex = lngIndex + 1
                    ReDim Preserve ZKMenu(0 To lngIndex + 2)
                    Set ZKMenu(lngIndex + 2) = New clsZenDictionary
                    Call ZKMenu(lngIndex + 2).FromProp(strSection)
                End If ' OS Compliant
            End If ' f Not (Left$(strSection, 1) = "=") Then
            
        Wend
            
        Rem - Add the fixed tems
        ReDim Preserve ZKMenu(0 To lngIndex + 4)
        Set ZKMenu(0) = zenDic("Class", "Zenkey", "Action", "ABOUT", "Caption", ZenKEYCap)
        Set ZKMenu(1) = zenDic("Class", "Zenkey", "Caption", "-")
        Set ZKMenu(lngIndex + 3) = zenDic("Class", "ZenKEY", "Caption", "-")
        Set ZKMenu(lngIndex + 4) = zenDic("Class", "ZenKEY", "Action", "EXIT", "Caption", "   ---   Exit   ---   ")
        
    Close #lngFNum
    
    Exit Sub
    
ErrorTrap:

    With Err
        Select Case .Number
            Case vbObjectError + 1: Call ZenMB(.Description, "OK")
            Case Else: Call ZenMB("Error " & CStr(.Number) & ", " & .Description, "OK")
        End Select
    End With
    End

End Sub











Public Sub ZenErr(ByVal ActionStr As String)
Dim strErr As String

    strErr = "Oops. " & Err.Description & " (Error " & CStr(Err.Number) & ") in " & Err.Source
    Call ZenMB(strErr, "OK")
    
    Dim FNum As Integer
    FNum = FreeFile
    Open settings("SavePath") & "\ZKError.log" For Append As #FNum
        Print #FNum, "------------------------ Zenkey error -----------------------"
        Print #FNum, CStr(Now) & ", "; "Version " & App.Major & "." & App.Minor & "." & App.Revision
        Print #FNum, strErr
        Print #FNum, "ActionString = " & ActionStr
    Close #FNum

End Sub

Public Function Get_Koan() As String
Dim FNum As Long, NumEntries As Long
Dim strTemp As String, k As Long
Static Num As Long
On Error Resume Next

    FNum = FreeFile
    Open App.Path & "\Quotes\" & settings("Quotes") & ".txt" For Input As #FNum
    Input #FNum, strTemp
    NumEntries = CLng(Val(strTemp))
    If Num = 0 Then
        Randomize
        Num = CLng(Rnd * NumEntries + 0.5) '+ 1
    Else
        Num = (Num Mod NumEntries)
        Num = Num + 1
    End If
    For k = 1 To Num
        Line Input #FNum, strTemp
    Next k
    Close #FNum
    Get_Koan = Trim$(strTemp)
    If Len(Get_Koan) = 0 Then Get_Koan = "If the '" & App.Path & "\" & settings("Quotes") & ".txt' file cannot be opened, then there can be no witty expressions.."

End Function

Public Function InsertSpecialFolder(ByVal Path As String) As String
Dim lngPos As Long
Dim strRest As String
Dim strSpecial As String
Dim r As Long
Dim IDL As ITEMIDLIST
Dim csidl_Item As Long

    Rem - Remove first the second
    lngPos = InStr(Path, "%")
    If lngPos > 0 Then
        Path = Mid$(Path, lngPos + 1)
        lngPos = InStr(Path, "%")
        If lngPos > 0 Then
            strRest = Mid$(Path, lngPos + 1)
            Path = left(Path, lngPos - 1)
            strSpecial = Path
        End If
    Else
        InsertSpecialFolder = Path
        Exit Function
    End If

    'Select Case UCase(strSpecial)
    Select Case True
        Case IsNumeric(strSpecial)
            csidl_Item = CLng(Val(strSpecial))
            r = SHGetSpecialFolderLocation(0, csidl_Item, IDL)
            'If r = 0 Then
            If r = 0 Then
                Rem - No error. Create a buffer
                strSpecial = Space$(512)
                r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strSpecial$)
                InsertSpecialFolder = left$(strSpecial, InStr(strSpecial, Chr$(0)) - 1) & strRest
            Else
                InsertSpecialFolder = ""
            End If
        Case left$(strSpecial, 7) = "APPPATH"
            InsertSpecialFolder = App.Path & strRest
            Exit Function
        Case left$(strSpecial, 7) = "WINDOWS"
            Rem - Create a buffer string
            strSpecial = Space(255)
            InsertSpecialFolder = left$(strSpecial, GetWindowsDirectory(strSpecial, 255)) & strRest
        Case left$(strSpecial, 6) = "SYSTEM" ', "SYSTEM32"
            strSpecial = Space(255)
            InsertSpecialFolder = left$(strSpecial, GetSystemDirectory(strSpecial, 255)) & strRest
        Case Else
            InsertSpecialFolder = vbNullString
    End Select

End Function










Private Sub p_SkipGroup(ByVal FNum As String)
Dim strTemp As String
    
    Do
        Input #FNum, strTemp
        If Prop_Get("Class", strTemp) = "Group" Then Call p_SkipGroup(FNum)
    Loop While Not CBool(Prop_Get("ENDGROUP", strTemp) = "True")


End Sub

Public Function SpecialFolderCaption(ByVal CSL_ID As Long) As String
    
    SpecialFolderCaption = InsertSpecialFolder("%" & CStr(CSL_ID) & "%")
    If Len(SpecialFolderCaption) > 1 Then SpecialFolderCaption = FBR_GetLastFolder(SpecialFolderCaption) & " (" & CStr(CSL_ID) & ")"
    
End Function

Public Sub Init_ZK()
Dim strSavePath As String
Dim Registry As New clsRegistry
    

    Rem - Okay, initialize the storage path
    Select Case Registry.GetRegistry(HKLM, "Software\ZenCODE\ZenKEY", "UserPath")
        Case "1"
            strSavePath = InsertSpecialFolder("%26%") + "\ZenCODE"  ' Application data
            If Len(Dir(strSavePath, vbDirectory)) = 0 Then Call MkDir(strSavePath)
            strSavePath = strSavePath + "\ZenKEY"
            If Len(Dir(strSavePath, vbDirectory)) = 0 Then Call MkDir(strSavePath)
        Case Else
            strSavePath = App.Path
    End Select
    
    Rem - Initialise settings
    If Len(Dir(strSavePath & "\Settings.ini")) = 0 Then Call FileCopy(App.Path & "\SetDef.ini", strSavePath & "\Settings.ini")
    If Len(Dir(strSavePath & "\ZenKEY.ini")) = 0 Then
        Call FileCopy(App.Path & IIf(IsWindows7OrBetter, "\Default_Complete7.ini", "\Default_Complete.ini"), strSavePath & "\ZenKEY.ini")
    End If
    
    Call settings.FromINI(strSavePath & "\Settings.ini")
    If Len(Dir(strSavePath & "\ProgInfo.ini")) > 0 Then Call progInfo.FromINI(strSavePath & "\ProgInfo.ini")
    settings("SavePath") = strSavePath
    COL_Zen = RGB(46, 136, 255)
            
End Sub

Public Sub INI_LoadFiles(ByVal FileName As String, ByRef colFiles As Collection, ByRef IsExe As Boolean)
Rem - Returns the number of files loaded from the ini

    Rem - Load the Omitted window list
    If Len(Dir(settings("SavePath") & "\" & FileName)) > 0 Then
        Dim FNum As Long, strTemp As String
        
        FNum = FreeFile
        Open settings("SavePath") & "\" & FileName For Input As #FNum
            While Not EOF(FNum)
                Input #FNum, strTemp
                Rem - Keep compatibility with previously used GetFileTitle, which left out the file extension if 'Hide known file extentions" was enabled
                Rem - Default to "exe" files
                If IsExe Then
                    If InStr(strTemp, ".") = 0 Then strTemp = strTemp & ".exe"
                End If
                colFiles.Add strTemp
            Wend
        Close #FNum
    End If

End Sub



Public Function zenDic(ParamArray KeyValPairs()) As clsZenDictionary
Dim zDic As New clsZenDictionary, k As Long
Dim max As Long

    max = UBound(KeyValPairs())
    For k = 0 To max Step 2
        zDic(CStr(KeyValPairs(k))) = KeyValPairs(k + 1)
    Next k
    Set zenDic = zDic

End Function

Public Function SetFocusToLastActive(Optional startIndex = 0) As Boolean
Rem - Set focus to the last active app that is on the current desktop
Dim k As Long, hwnd As Long

    For k = startIndex To 15
        hwnd = ActiveWindow(k)
        If IsWindowVisible(hwnd) Then
            If IsOnDesktop(hwnd) Then
                Call SetForegroundWindow(hwnd)
                SetFocusToLastActive = True
                Exit For
            End If
        End If
    Next k

End Function

Attribute VB_Name = "modZKConfig"
Option Explicit
Option Compare Text
Rem =======================================  Region calls
'Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Public Declare Function SetWindowRgn Lib "user32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Rem ======================================= For the splash screen  ===================================
Public StartSection As String
Public Registry As New clsRegistry
Public HotKeys As New clsHotkey
Public MainForm As Form
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Public Searches() As String
Private Const HKSep = "                                            -"
Public Const ZK_TransWarn = "Before this setting is enabled, there are a few things you should know." & vbCr & vbCr & _
"1. Transparency is a real performance hog. A good graphics card is almost essential for this feature to work nicely." & vbCr & _
"2. Some program windows, especially video/media applications, do not display properly whilst semi-transparent. Black squares in a window indicate problems..." & vbCr & _
"3. If you would like to prevent these or any other windows from being affected, please add them to the 'Excluded applications' list." & vbCr & _
"4. There are various options available in the ZenKEY Configuration Utility if you experience other display anomalies."


Public Sub ZK_Restart()
Const WM_CLOSE = &H10
Dim lngHandle As Long
Dim lngStart As Long
Const MaxWait = 2500
Dim booLoop As Boolean
    lngHandle = Val(Registry.GetRegistry(HKCU, "SOFTWARE\ZenCODE\ZenKEY", "WindowHandle"))
    Rem - If there is an instance, force it to close and then restart it!
    If lngHandle <> 0 Then
        Call PostMessage(lngHandle, WM_CLOSE, 0&, ByVal 0&)
        lngStart = GetTickCount
        booLoop = True
        Do
            DoEvents
            booLoop = CBool(Len(Registry.GetRegistry(HKCU, "SOFTWARE\ZenCODE\ZenKEY", "WindowHandle")) > 0)
            If booLoop Then booLoop = CBool(GetTickCount - lngStart < MaxWait)
        Loop While booLoop
    End If
    DoEvents
    Call Shell(App.Path & "\ZenKEY.exe RESTART", vbNormalFocus)

End Sub
Private Function Extract(ByVal Sentance As String, ByVal AfterNthSep As Long, ByVal Separator As String) As String
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
Public Function KS_GetDescription(ByVal KeySequence As String) As String
Dim KeyCount As Long
Dim k As Long
Dim lngKeyVal As Long

    Rem - Format of action
    Rem - KeyCount>KeyPressed1>KeyDown>KeyPressed2>KeyDown>...[-ms Pause>False].....
    KS_GetDescription = vbNullString
    
    
    KeyCount = Val(KeySequence)
    Rem - No need to load everything if it's just for the description
    If KeyCount > 10 Then KeyCount = 10
    For k = 0 To KeyCount - 1
        lngKeyVal = Val(Extract(KeySequence, 2 * k + 1, ">"))
        If lngKeyVal < 0 Then
            Rem - Paused
            KS_GetDescription = KS_GetDescription & ", Pause"
        'ElseIf CBool("Y" = Extract(KeySequence, 2 * k + 2, ">")) Then
            Rem - Key pressed
            'KS_GetDescription = KS_GetDescription & ", " & HotKeys.Keyname(lngKeyVal)
        ElseIf Val(Extract(KeySequence, 2 * k + 2, ">")) Mod 2 = 0 Then
            Rem - Use encryption
            KS_GetDescription = KS_GetDescription & ", " & HotKeys.Keyname(lngKeyVal - 71 - k)
        Else
            'keybd_event lngKeyVal, 0, KEYEVENTF_KEYUP, 0   ' release H
        End If
    Next k
    If Len(KS_GetDescription) > 0 Then
        KS_GetDescription = Mid(KS_GetDescription, 3)
        If KeyCount = 10 Then KS_GetDescription = KS_GetDescription & "..."
        
    Else
        KS_GetDescription = "No keystokes"
    End If
    
    
End Function
Public Sub TestAction(ByRef prop As clsZenDictionary)
Dim strProp As String

    strProp = prop.ToProp
    If Prop_Get("Class", strProp) = "KEYSTROKES" Then
        Call ShellExe(App.Path & "\ZenKP.exe", strProp)
    Else
        WIN_ZKAction = RegisterWindowMessage(ByVal "ZenKEY Action")
        If Not SendToZenKEY(strProp) Then
            Call ShellExe(App.Path & "\ZenKEY.exe", strProp)
        End If
    End If
    
End Sub
Rem - For activating a previous instance
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As
'Public Const SW_NORMAL = 1
Public Sub FloodBack(DestObj As Object, BackColor As Long, ForeColor As Long, GradStyle As Integer, X As Long, Y As Long)

'Paints a gradated background, fading from one color into another
'Sample Call : GradateBackground Me, &H400000, &HFF0000, 0, 0, 0
'GradStyle Modes:
'0 - vertical
'1 - circular from center
'2 - horizontal
'3 - ellipse from upper left
'4 - ellipse from upper right
'5 - ellipse from lower right
'6 - ellipse from lower left
'7 - ellipse from upper center
'8 - ellipse from right center
'9 - ellipse from lower center
'10 - ellipse from left center
'11 - ellipse from x,y - twips

Dim foo As Integer, foobar As Integer
Dim DestWidth As Integer, DestHeight As Integer, DestMode As Integer
Dim StartPnt As Integer, EndPnt As Integer, DrawHeight As Double, DrawWidth As Double
Dim dblG As Double, dblR As Double, dblB As Double
Dim addg As Double, addr As Double, addb As Double
Dim Mask As Long, Mask2 As Long, colorstep As Integer
Dim bckR As Double, bckG As Double, bckB As Double
Dim Linecolor As Long, PixelStep As Long, LineHeight As Integer
Dim PixelCount As Integer, aspect As Single
Dim CenterX As Long, CenterY As Long
On Error Resume Next

Screen.MousePointer = 11

'set up rgb bitmask
Mask = 255
Mask = Mask ^ 2
Mask2 = 255
Mask2 = 255 ^ 3

'Init dimensions in twips, set backcolor, set modes
DestMode = DestObj.ScaleMode
DestObj.ScaleMode = 1
DestHeight = DestObj.ScaleHeight
DestWidth = DestObj.ScaleWidth
DestObj.BackColor = BackColor
DestObj.AutoRedraw = True
DestObj.DrawStyle = 5 'transparent
DestObj.DrawMode = 13 'CopyPen

'solid offset
Select Case GradStyle
Case 2 'Horizontal
    StartPnt = DestWidth * 0.05
    EndPnt = DestWidth * 0.95
Case Else
    StartPnt = DestHeight * 0.05
    EndPnt = DestHeight * 0.95
    Select Case GradStyle
    Case 3 'ellipse from upper left
        CenterX = 0
        CenterY = 0
    Case 4 'ellipse from upper right
        CenterX = DestWidth
        CenterY = 0
    Case 5 'ellipse from lower right
        CenterX = DestWidth
        CenterY = DestHeight
    Case 6 'ellipse from lower left
        CenterX = 0
        CenterY = DestHeight
    Case 7 'ellipse from upper center
        CenterX = DestWidth / 2
        CenterY = 0
    Case 8 'ellipse from right center
        CenterX = DestWidth
        CenterY = DestHeight / 2
    Case 9 'ellipse from lower center
        CenterX = DestWidth / 2
        CenterY = DestHeight
    Case 10 'ellipse from left center
        CenterX = 0
        CenterY = DestHeight / 2
    Case 11 'ellipse from x,y - twips
        CenterX = X
        CenterY = Y
    End Select
End Select
aspect = DestHeight / DestWidth

Select Case GradStyle
Case 0
    DrawHeight = EndPnt - StartPnt
Case 1
    DrawHeight = Sqr((DestHeight / 2) ^ 2 + (DestWidth / 2) ^ 2)
Case 2
    DrawWidth = EndPnt - StartPnt
Case 3, 4, 5, 6
    DrawHeight = Sqr((DestHeight) ^ 2 + (DestWidth) ^ 2)
Case 7, 8, 9, 10
    If DestHeight >= DestWidth Then
        DrawHeight = DestHeight
    Else
        DrawHeight = DestWidth
    End If
Case 11
    DrawHeight = CenterX
    If Sqr(CenterY ^ 2 + CenterX ^ 2) > DrawHeight Then DrawHeight = Sqr(CenterY ^ 2 + CenterX ^ 2)
    If Sqr(CenterY ^ 2 + (DestWidth - CenterX) ^ 2) > DrawHeight Then DrawHeight = Sqr(CenterY ^ 2 + (DestWidth - CenterX) ^ 2)
    If Sqr((DestHeight - CenterY) ^ 2 + (DestWidth - CenterX) ^ 2) > DrawHeight Then DrawHeight = Sqr((DestHeight - CenterY) ^ 2 + (DestWidth - CenterX) ^ 2)
    If Sqr((DestHeight - CenterY) ^ 2 + CenterX ^ 2) > DrawHeight Then DrawHeight = Sqr((DestHeight - CenterY) ^ 2 + CenterX ^ 2)
    'DrawHeight = DrawHeight * .9
End Select
dblR = CDbl(BackColor And &HFF)
dblG = CDbl(BackColor And &HFF00&) / 255
dblB = CDbl(BackColor And &HFF0000) / &HFF00&
bckR = CDbl(ForeColor And &HFF&)
bckG = CDbl(ForeColor And &HFF00&) / 255
bckB = CDbl(ForeColor And &HFF0000) / &HFF00&
If GradStyle = 2 Then
    addr = (bckR - dblR) / (DrawWidth / Screen.TwipsPerPixelY)
    addg = (bckG - dblG) / (DrawWidth / Screen.TwipsPerPixelY)
    addb = (bckB - dblB) / (DrawWidth / Screen.TwipsPerPixelY)
Else
    addr = (bckR - dblR) / (DrawHeight / Screen.TwipsPerPixelY)
    addg = (bckG - dblG) / (DrawHeight / Screen.TwipsPerPixelY)
    addb = (bckB - dblB) / (DrawHeight / Screen.TwipsPerPixelY)
End If

DestObj.Cls

PixelStep = Screen.TwipsPerPixelY
LineHeight = PixelStep * 2
Select Case GradStyle
Case 0 'Vertical
    For foo = 1 To DrawHeight Step PixelStep
        dblR = dblR + addr
        dblG = dblG + addg
        dblB = dblB + addb
        If dblR > 255 Then dblR = 255
        If dblG > 255 Then dblG = 255
        If dblB > 255 Then dblB = 255
        If dblR < 0 Then dblR = 0
        If dblG < 0 Then dblG = 0
        If dblB < 0 Then dblB = 0
        Linecolor = RGB(dblR, dblG, dblB)
        DestObj.Line (0, foo + StartPnt)-(DestWidth, foo + StartPnt + LineHeight), Linecolor, BF
    Next foo
    For foo = EndPnt To DestHeight Step PixelStep
        DestObj.Line (0, foo)-(DestWidth, foo + LineHeight), ForeColor, BF
    Next foo
Case 2 'horizontal
    For foo = 1 To DrawWidth Step PixelStep
        dblR = dblR + addr
        dblG = dblG + addg
        dblB = dblB + addb
        If dblR > 255 Then dblR = 255
        If dblG > 255 Then dblG = 255
        If dblB > 255 Then dblB = 255
        If dblR < 0 Then dblR = 0
        If dblG < 0 Then dblG = 0
        If dblB < 0 Then dblB = 0
        Linecolor = RGB(dblR, dblG, dblB)
        DestObj.Line (foo + StartPnt, 0)-(foo + StartPnt + LineHeight, DestHeight), Linecolor, BF
    Next foo
    For foo = EndPnt To DestWidth Step PixelStep
        DestObj.Line (foo, 0)-(foo + LineHeight, DestHeight), ForeColor, BF
    Next foo
Case 1 'circular
    Screen.MousePointer = 11
    DestObj.FillStyle = 0
    PixelCount = 5
    PixelStep = PixelStep * -1 * PixelCount
    For foo = DrawHeight To 1 Step PixelStep
        dblR = dblR + (addr * PixelCount)
        dblG = dblG + (addg * PixelCount)
        dblB = dblB + (addb * PixelCount)
        If dblR > 255 Then dblR = 255
        If dblG > 255 Then dblG = 255
        If dblB > 255 Then dblB = 255
        If dblR < 0 Then dblR = 0
        If dblG < 0 Then dblG = 0
        If dblB < 0 Then dblB = 0
        Linecolor = RGB(dblR, dblG, dblB)
        DestObj.FillColor = Linecolor
        DestObj.Circle (DestWidth / 2, DestHeight / 2), foo, Linecolor, , , aspect
    Next foo
    Screen.MousePointer = 0
Case Else 'elliptical from various points
    DestObj.FillStyle = 0
    PixelCount = 5
    PixelStep = PixelStep * -1 * PixelCount
    For foo = DrawHeight To 1 Step PixelStep
        dblR = dblR + (addr * PixelCount)
        dblG = dblG + (addg * PixelCount)
        dblB = dblB + (addb * PixelCount)
        If dblR > 255 Then dblR = 255
        If dblG > 255 Then dblG = 255
        If dblB > 255 Then dblB = 255
        If dblR < 0 Then dblR = 0
        If dblG < 0 Then dblG = 0
        If dblB < 0 Then dblB = 0
        Linecolor = RGB(dblR, dblG, dblB)
        DestObj.FillColor = Linecolor
        DestObj.Circle (CenterX, CenterY), foo, Linecolor, , , aspect
    Next foo
End Select
DestObj.ScaleMode = DestMode
Screen.MousePointer = 0

End Sub


Public Sub Swap(ByRef Item1 As Variant, ByRef Item2 As Variant)
Dim Temp As Variant
    
    Temp = Item2
    Item2 = Item1
    Item1 = Temp
    
End Sub







Public Sub Main()

    Call Actions_Load
    Call LoadArray(Searches(), App.Path & "\Search.ini")

    Rem - A command line has been passed. Disect it and start in the appropriate section
    StartSection = Command$
    Rem - Initailise settings (MB Defaults in "SetDef.ini")
    Call Init_ZK
    
    #If ZenWiz <> 1 Then
        Select Case True
            Case InStr(StartSection, "ABOUT") > 0, InStr(StartSection, "SPLASH") > 0
                Set MainForm = New frmAbout
            Case Else
                Set MainForm = New frmZKConfig
        End Select
    MainForm.Display
    #End If
    
End Sub
Public Function HKIsOkay(ByVal strShift As String, ByVal strHK As String, ByVal EditIndex As Long) As Boolean
Rem - Return True if the hotkey combincation is acceptable, False otherwise

    If strShift = "Win" Then
        Call ZenMB("Sorry, but all 'Win + key' combinations are reserved by Windows for use by the operating system.", "OK")
        Exit Function
    ElseIf Len(strHK) > 0 And Len(strShift) > 0 Then
        Dim i As Long
        If p_HKIsDuplicated(strShift, strHK, EditIndex, i) Then
            Dim strMsg As String
            strMsg = "It seems that this Hotkey combination is already being used for the '" & ZKMenu(i)("Caption") & "'' item (Class " & ZKMenu(i)("Class") & ")." & _
                " Please note that if both of these items are enabled, only one of them will be assigned the Hotkey and the other will give you a rude error message." & _
                vbCr & vbCr & "Do you wish to continue anyway?"
            If 1 = ZenMB(strMsg, "Yes", "No") Then Exit Function
        End If
    ElseIf Len(strShift) > 0 Then
        If ZenMB("You have chosen to use only 'modifier' keys (Control, Alt, Shift) as a 'Hotkey'. This may interfere with normal keyboard operations. Are you sure you wish to proceed?", "Yes", "No") = 1 Then Exit Function
    ElseIf Len(strHK) > 0 Then
        If ZenMB("You have chosen to use only one key as a 'Hotkey'. This may interfere with normal keyboard operations. Are you sure you wish to proceed?", "Yes", "No") = 1 Then Exit Function
    End If
    HKIsOkay = True

End Function
Private Function p_HKIsDuplicated(ByVal strShift As String, ByVal strKey As String, ByVal ItemIndex As Long, ByRef DupeIndex As Long) As Boolean
Dim intMax As Integer
Dim k As Integer

    p_HKIsDuplicated = False
    
    intMax = UBound(ZKMenu())
    For k = 0 To intMax
        If strKey = Val(ZKMenu(k)("Hotkey")) Then
            If HotKeys.ShiftValue(strShift) = HotKeys.ShiftValue(ZKMenu(k)("ShiftKey")) Then
                If k <> ItemIndex Then
                    p_HKIsDuplicated = True
                
                    Rem - If in the wizard, check that they are not in the alternate winamp / windows media sections
                    If App.ProductName = "ZenKEY Wizard" Then
                        Dim lngGroup1 As Long, lngGroup2 As Long
                        Dim strName1 As String, strName2 As String
                        
                        lngGroup1 = Item_GetGroup(k)
                        lngGroup2 = Item_GetGroup(ItemIndex)
                        If lngGroup1 > 0 And lngGroup2 > 0 Then
                            strName1 = ZKMenu(lngGroup1)("Caption")
                            strName2 = ZKMenu(lngGroup2)("Caption")
                            Select Case strName1
                                Case "Windows Media commands"
                                    If strName2 = "Winamp Controls" Then p_HKIsDuplicated = False
                                Case "Winamp Controls"
                                    If strName2 = "Windows Media commands" Then p_HKIsDuplicated = False
                            End Select
                        End If
                    End If
                    Rem - --------------------------------------------------------------------------------------------------------------------------------
                    
                    If p_HKIsDuplicated Then
                        DupeIndex = k
                        Exit Function
                    End If
                End If
            End If
        End If
    Next k
    

End Function


Public Function Item_GetGroup(ByVal ItemIndex As Long) As Long
Dim Min As Long, k As Long

    Min = -1
    k = ItemIndex - 1
    
    If k > 0 Then
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
    End If
    Item_GetGroup = Min
    
End Function
Public Sub CentreForm(ByRef It As Form)

    With It
        .Move (Screen.Width - .ScaleX(.ScaleWidth, .ScaleMode, vbTwips)) / 2, (Screen.Height - .ScaleY(.ScaleHeight, .ScaleMode, vbTwips)) / 2
    End With

End Sub



Public Sub Action_Fired(ByVal ID As Long)
Rem - Just to allow the Hotkey class to shared

End Sub

Public Function ZK_GetObject()
    Rem - Stub
End Function

Public Sub LoadArray(ByRef TheArray() As String, ByVal FileName As String)

    If Len(Dir(FileName)) > 0 Then
        Dim lngCount As Long, FNum As Long
        
        FNum = FreeFile
        Open FileName For Input As #FNum
        While Not EOF(FNum)
            ReDim Preserve TheArray(0 To lngCount)
            Line Input #FNum, TheArray(lngCount)
            lngCount = lngCount + 1
        Wend
        
    End If

End Sub

Public Sub HKCombo_Init(ByRef It As ComboBox)
Dim k As Long
Const HK_MIN = 19
Const HK_NumPad5 = 12
Const HK_MAX = 255
    
    Dim strTemp As String
    With It
        .Clear
        .AddItem "<None>"
        .AddItem Trim(HotKeys.Keyname(9)) & HKSep & CStr(9)
        .AddItem Trim(HotKeys.Keyname(HK_NumPad5)) & HKSep & CStr(HK_NumPad5)
        For k = HK_MIN To HK_MAX
            strTemp = Trim(HotKeys.Keyname(k))
            If Len(strTemp) = 0 Then strTemp = "Extra"
    
            Rem - Check that the key as a description
            If Len(strTemp) > 0 Then
                Select Case k
                    Case 91, 160, 162 To 165
                        Rem - Do not add any illegal keys
                        Rem - Window = 91, Shift = 160, Control = 162, 163
                        Rem - Alt = 164, 165
                    Case 96 To 105, 110
                        Rem - With Numlock
                        Rem - Numpd 9 = 105, Numpad 8 = 104, Numpad 7 = 103
                        Rem - Numpd 4 = 100, Numpad 5 = 101, Numpad 7 = 102
                        Rem - Numpd 1 = 97, Numpad 2 = 98, Numpad 3 = 99
                        Rem- Numpad . = 110, Numpad 0 = 96
                        .AddItem strTemp & " -NL" & HKSep & CStr(k)
                    Case Else
                        Rem - Add the key to the combobox and to the array
                        .AddItem strTemp & HKSep & CStr(k)
                End Select
            End If
        Next k
        .ListIndex = 0
    End With

End Sub
Public Sub HKCombo_Display(ByVal KeyCode As Long, ByRef It As ComboBox)
    
    On Error Resume Next
    Dim strTemp As String

    strTemp = HotKeys.Keyname(KeyCode)
    If Len(strTemp) = 0 Then strTemp = "Extra"
    It.Text = strTemp & HKSep & CStr(KeyCode)
    If Err.Number <> 0 Then
        Rem - It must be a special key
        Select Case KeyCode
            Case 91, 160, 162 To 165
                Rem - Do not add any illegal keys
                Rem - Window = 91, Shift = 160, Control = 162, 163
                Rem - Alt = 164, 165
                Exit Sub
            Case 96 To 105, 110
                Rem - With Numlock
                Rem - Numpd 9 = 105, Numpad 8 = 104, Numpad 7 = 103
                Rem - Numpd 4 = 100, Numpad 5 = 101, Numpad 7 = 102
                Rem - Numpd 1 = 97, Numpad 2 = 98, Numpad 3 = 99
                Rem- Numpad . = 110, Numpad 0 = 96
                strTemp = strTemp & " -NL"
        End Select
        It.AddItem strTemp & HKSep & CStr(KeyCode)
        It.ListIndex = It.ListCount - 1
    End If

    On Error GoTo 0

End Sub


Public Function HKCombo_GetValue(ByRef It As ComboBox) As String
Dim i As Long

    i = InStr(It.Text, HKSep)
    HKCombo_GetValue = Mid(It.Text, i + Len(HKSep))
    

End Function

Public Function SelectFileDlg(ByRef CallingForm As Form, ByRef FileName As String, Optional selClass As Boolean = False) As Boolean
Dim It As frmWindowCapture
Dim strFName As String

    CallingForm.Enabled = False
    Set It = New frmWindowCapture
    If selClass Then
        Rem - Select a class
        Set It.CallingForm = CallingForm
        SelectFileDlg = It.SelectClass(strFName)
    Else
        Call It.PopupMenu(It.mnuMain)
        Select Case It.SelMode
            Case "Browse"
                Rem - Just let them browse
                strFName = "C:"
                SelectFileDlg = FBR_GetOFName("Select the file to run or open.", strFName, "All files (*.*)")
    
            Case "MyDocuments"
                Rem - Look in My documents
                strFName = InsertSpecialFolder("%5%")
                SelectFileDlg = FBR_GetOFName("Select the file to run or open.", strFName, "All files (*.*)")
            
            Case "ProgramFiles"
                Rem - Look in program files
                strFName = InsertSpecialFolder("%38%")
                SelectFileDlg = FBR_GetOFName("Select the file to run or open.", strFName, "Executable files (*.exe)", "Shortcuts (*.lnk")
                
            Case "Drag"
                Rem - Select by drag and drop
                strFName = FileName
                Set It.CallingForm = CallingForm
                SelectFileDlg = It.SelectExe(strFName)
        End Select
    End If
    If SelectFileDlg Then FileName = strFName
    CallingForm.Enabled = True
    Unload It

End Function


Attribute VB_Name = "modZenKP"
Option Explicit
'Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Private Declare Function GetTickCount& Lib "kernel32" ()
'Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Type KEYBDINPUT
  wVk As Integer
  wScan As Integer
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type
Private Const INPUT_KEYBOARD = 1
Private Type GENERALINPUT
    dwType As Long
    xi(0 To 23) As Byte
End Type
Public Declare Function GetTickCount& Lib "kernel32" ()
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public MainForm As Variant
Public WIN_Shift As Long
Public WIN_Active As Long
Public AWT_LastTrans As Long
Rem ====================== Variables for standard colouring ====================
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
'Public Function Prop_Get(ByVal PropName As String, ByVal PropString As String) As String
'Dim lngPos As Long
'Dim lngEnd As Long
'Const strSep As String = "|"'

'
 '   lngPos = InStr(1, PropString, strSep & PropName & "=", vbTextCompare)
'    If lngPos > 0 Then
'        lngEnd = InStr(Mid$(PropString, lngPos + 1), strSep)
'        If lngEnd = 0 Then lngEnd = 2 ^ 15
'        Prop_Get = Mid$(PropString, lngPos + 2 + Len(PropName), lngEnd - Len(PropName) - 2)
'    Else
'        Prop_Get = vbNullString
'    End If
    
'End Function
 

Private Sub WaitForRelease(ByVal KeyValue As Long)
    
    While KeyValue > 0
        If GetAsyncKeyState(KeyValue) = 0 Then KeyValue = 0 Else DoEvents
    Wend

End Sub

Public Sub Main()
Dim Prop As String ', Registry As New clsRegistry
    
    Prop = Command$
    If Len(Prop) > 0 Then
        If Len(Prop_Get("Startup", Prop)) > 0 Then
            Call Timer(Prop)
        Else
            Call DoKeypress(Prop)
        End If
    End If
    
End Sub







Private Sub DoKeypress(ByVal Prop As String)
Dim strTemp As String, lngVal As Long
Const KEYEVENTF_KEYUP = &H2
Dim lngTick As Long

    lngVal = Val(Prop_Get("Hotkey", Prop))
    strTemp = Prop_Get("ShiftKey", Prop)
    
    Dim FNum As Long
    FNum = FreeFile
    If lngVal > 0 Then Call WaitForRelease(lngVal)
    If InStr(strTemp, "Alt") > 0 Then Call WaitForRelease(18)
    If InStr(strTemp, "Ctrl") > 0 Then Call WaitForRelease(17) 'Const VK_CONTROL = &H11
    If InStr(strTemp, "Shift") > 0 Then Call WaitForRelease(16) 'Const VK_SHIFT = &H10
    Rem - Wait until the keys are released before firing te keypresses

    Rem - Format of action
    Rem - KeyCount>KeyPressed1>KeyDown>KeyPressed2>KeyDown>...[-ms Pause>False].....
    
    Rem - Encryption scheme -----
    Rem - KeyCount>KeyPressed1+71+1>Odd = KeyDown, Even = KeyUp>KeyPressed2+71+ 2>Odd = KeyDown, Even = KeyUp>...[-ms Pause>False].....
    Rem - KeyCount - No change
    Rem - KeyPressed - Key value = KeyValue + 70 + KeyNumber
    Rem - KeyDown - Even = "Y", Odd = "N"
    Rem - Pause - No Change
    Rem - Encryption scheme -----
    
    Dim strAction As String
    Dim KeyCount As Long
    Dim lngKeyVal As Long
    
    strAction = Prop_Get("Action", Prop)
    KeyCount = Val(strAction)
    
    Dim k As Long
    For k = 0 To KeyCount - 1
        lngKeyVal = Val(Extract(strAction, 2 * k + 1, ">"))
        If lngKeyVal < 0 Then
            lngTick = GetTickCount
            While GetTickCount - lngTick < Abs(lngKeyVal)
                DoEvents
            Wend
        Else
            Rem - Previous
'            lngKeyVal = lngKeyVal - 71 - k
'            If CBool(Val(Extract(strAction, 2 * k + 2, ">")) Mod 2 = 0) Then
'                keybd_event lngKeyVal, 0, 0, 0 ' press
'            Else
'                keybd_event lngKeyVal, 0, KEYEVENTF_KEYUP, 0   ' release
'            End If
            Rem - Previous
            
            Dim GInput(0 To 0) As GENERALINPUT
            Dim KInput As KEYBDINPUT

            GInput(0).dwType = INPUT_KEYBOARD
            KInput.wVk = lngKeyVal - 71 - k
                
            If CBool(Val(Extract(strAction, 2 * k + 2, ">")) Mod 2 = 0) Then
                KInput.dwFlags = 0 ' press
            Else
                KInput.dwFlags = KEYEVENTF_KEYUP ' release
            End If
            CopyMemory GInput(0).xi(0), KInput, Len(KInput)
            Call SendInput(1, GInput(0), Len(GInput(0)))
        End If
        
' Previous - no encryption
'        If lngKeyVal < 0 Then
'            lngTick = GetTickCount
'            While GetTickCount - lngTick < Abs(lngKeyVal)
'                DoEvents
'            Wend
'        ElseIf CBool("Y" = Extract(strAction, 2 * k + 2, ">")) Then
'            keybd_event lngKeyVal, 0, 0, 0
'        Else
'            keybd_event lngKeyVal, 0, KEYEVENTF_KEYUP, 0   ' release H
'        End If
    Next k

End Sub

Private Sub Timer(ByVal Prop As String)
Dim lngStart As Long, lngTime As Long

    lngTime = 1000 * Val(Prop_Get("Startup", Prop))
    lngStart = GetTickCount
    While GetTickCount - lngStart < lngTime
        DoEvents
    Wend
    
    If Prop_Get("Class", Prop) = "KEYSTROKES" Then
        Call DoKeypress(Prop)
    Else
        Rem - Now pass it back to ZenKEY
        Rem - taken from SendToZenKEY in config
        Dim lngHWnd As String
        Dim Registry As New clsRegistry, WIN_ZKAction As Long
        WIN_ZKAction = RegisterWindowMessage(ByVal "ZenKEY Action")
    
        lngHWnd = Val(Registry.GetRegistry("HKCU", "SOFTWARE\ZenCODE\ZenKEY", "WindowHandle"))
        If Len(lngHWnd) > 0 Then
            Call Registry.SetRegistry("HKCU", "SOFTWARE\ZenCODE\ZenKEY", "Action", Prop)
            Rem - Now notify the running ZenKEY of the command
            Call PostMessage(lngHWnd, WIN_ZKAction, 0, 0)
        End If
    End If

End Sub


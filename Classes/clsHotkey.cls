VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHotkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public booLoaded As Boolean
Private Const PM_REMOVE = &H1
Private Const PM_NOREMOVE = &H0
Private Const WM_HOTKEY = &H312
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type Msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Rem -------------------------------------------------- For obtaining & deleting ID's for Horkey
Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Private lngKeyCount As Long
Private ID_External() As Long
Private ID_Internal() As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Rem ======================================= For determinuig the names of keys ===================================
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Long) As Long
Private lngPreserve As Long
Rem ======================== For message box ====================
Public Function Keyname(ByVal KeyCode As Long) As String
Dim lngSC As Long 'Scan code
Dim lngVK As Long 'Virtual key
Dim KName As String * 255
Dim intPos As Integer

    Rem - Get the virtual key, then the scancode
    lngVK = KeyCode
    lngSC = MapVirtualKey(lngVK, 0)
    
    Rem - Get the  key name
    Rem - Note: 44 ("Print screen = SysRq")
    Rem - Note: 19 ("Break")
    Call GetKeyNameText(lngSC * &H10000, KName, 255)
    intPos = InStr(KName, Chr$(0))
    Select Case intPos
        Case Is > 1
            Keyname = left$(KName, intPos - 1)
        Case 1
            Rem - Catch any missing descriptions here!
            Select Case KeyCode
                Case 91 ' Windows key
                    Keyname = "Windows"

            End Select
    End Select
    
End Function
Public Function ShiftValue(ByRef Keys As String) As Long
Rem - Const MOD_ALT = &H1, Const MOD_CONTROL = &H2
Rem - Const MOD_SHIFT = &H4, MOD_WIN = &H8

    ShiftValue = 0
    If InStr(Keys, "Alt") > 0 Then ShiftValue = 1
    If InStr(Keys, "Ctrl") > 0 Then ShiftValue = ShiftValue + 2
    If InStr(Keys, "Shift") > 0 Then ShiftValue = ShiftValue + 4
    If InStr(Keys, "Win") > 0 Then ShiftValue = ShiftValue + 8 ' 4000
    
End Function
Public Function ShiftValToStr(ByVal ShiftState As Long) As String
Dim strTemp As String
        
    'ShiftSate = ShiftState Mod 8
    If ShiftState >= 8 Then strTemp = " + Win": ShiftState = ShiftState - 8
    If ShiftState >= 4 Then strTemp = strTemp & " + Shift": ShiftState = ShiftState - 4
    If ShiftState >= 2 Then strTemp = strTemp & " + Ctrl": ShiftState = ShiftState - 2
    If ShiftState >= 1 Then strTemp = strTemp & " + Alt"
    
    If Len(strTemp) > 0 Then strTemp = Mid(strTemp, 4)
    ShiftValToStr = strTemp
        
End Function
Public Function AddHotkey(ByVal ID As Long, ByVal Shiftkeys As String, ByVal HotKey As Long, ByVal DoNotUnload As Boolean) As Boolean
Dim lngShiftkeys As Long
Dim lngID As Long
Dim lngRet As Long

    Rem - Now see if we can load the Hotkey
    lngShiftkeys = ShiftValue(Shiftkeys)
    lngID = GlobalAddAtom(App.ExeName & " ID " & CStr(lngKeyCount))
    If lngID <> 0 Then
        lngRet = RegisterHotKey(MainForm.hwnd, lngID, lngShiftkeys, HotKey)
        AddHotkey = CBool(lngRet > 0)
        If AddHotkey Then
            ReDim Preserve ID_Internal(0 To lngKeyCount)
            ReDim Preserve ID_External(0 To lngKeyCount)
            ID_External(lngKeyCount) = ID
            ID_Internal(lngKeyCount) = lngID
            
            
            Rem - If it should be preserved, move it to the beginning of the array
            If DoNotUnload Then
                Call Swap(ID_External(lngKeyCount), ID_External(lngPreserve))
                Call Swap(ID_Internal(lngKeyCount), ID_Internal(lngPreserve))
                lngPreserve = lngPreserve + 1
            End If
            
            lngKeyCount = lngKeyCount + 1
        End If
    End If

End Function

Public Sub Unload()
Dim k As Long
    
    Rem - Unregister hotkeys
    For k = lngPreserve To lngKeyCount - 1
        Call UnregisterHotKey(MainForm.hwnd, ID_Internal(k))
        Call GlobalDeleteAtom(ID_Internal(k))
    Next k
    lngKeyCount = lngPreserve
    booLoaded = False
 
End Sub


Public Sub WaitForMessages()
Dim Message As Msg
Dim lngKey As Long
    
    Rem - Loop until bCancel is set to True
    Do While Not booKill
        Rem - wait for a message
        Call WaitMessage
        Rem- check if its a HOTKEY-message
        If Not booKill Then
            If PeekMessage(Message, MainForm.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
                Rem ====== Check thorugh each Hotkey
                For lngKey = 0 To lngKeyCount - 1
                    If Message.wParam = ID_Internal(lngKey) Then
                        Call MainForm.DoAction(zenDic("Action", "FIREACTION", "Index", ID_External(lngKey)))
                    End If
                Next lngKey
            End If ' Peekmessage
        End If
        Rem - Let the operating system process other events
        DoEvents
    Loop

End Sub



Public Function GetCaption(ByRef prop As clsZenDictionary) As String
Dim strShift As String
Dim strKey As String
Dim strTemp As String
        
    strShift = prop("ShiftKey")
    strKey = prop("Hotkey")
    If Len(strKey) > 0 Then
        strTemp = HotKeys.Keyname(Val(strKey))
        If Len(strTemp) > 0 Then strKey = strTemp Else strKey = "Ext-" & strKey
    End If
    
    If Len(strShift & strKey) > 0 Then
        If (Len(strKey) = 0) Or (Len(strShift) = 0) Then
            GetCaption = strShift & strKey
        Else
            GetCaption = strShift & " + " & strKey
        End If
    Else
        GetCaption = vbNullString
    End If

End Function


Private Sub Swap(ByRef Num1 As Long, Num2 As Long)
Dim lngTemp As Long

    lngTemp = Num1
    Num1 = Num2
    Num2 = lngTemp

End Sub


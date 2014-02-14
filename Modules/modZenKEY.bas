Attribute VB_Name = "modZenKEY"
Option Explicit
Rem - Declare the API-Functions  for getting the mouse pos & Window we are about to mount
Public cMenu As New clsMenu
Public Const MNU_Start = 1000
Public Hotkeys As New clsHotkey
Public MainForm As frmZenKEY
Public Registry As New clsRegistry
Public AWT_LastTrans As Long ' Last window automatically made transparent
Public AWT_Depth As Long
Public SET_FollowActive As Boolean
Public SET_Layer As Long
Public SET_Trans As Long
Public SET_ZenBar As Boolean

Rem - For Optimization
Public WIN_Shift As Long
Public WIN_Active As Long
Public Sub Main()
    Rem - For direct object testing
    'Call ZK_GetObject("Search").DoAction("|SearchString=http://www.google.com/search?q=<Criteria>|")
    'Call ZK_GetObject("KEYSTROKES").DoAction("|Class=KEYSTROKES|Action=16>-5000>1>-5000>5>90>2>92>0>41>6>142>7>94>5>96>1>-5000>3>-5000>7>98>0>100>0>153>0>154>3>103>1>103>5>|Caption=Delayed startup|")
    
'|Class=SYSTEM|Action=RUNDIALOG|Caption=Run dialog|
'|Class=SYSTEM|Action=CTRMOUSE|Caption=Center the mouse|Hotkey=36|ShiftKey=Ctrl|
'|Class=SYSTEM|Action=CTRMOUSEACTIVE|Caption=Center the mouse on the active control|Hotkey=35|ShiftKey=Ctrl|
'|Class=KEYSTROKES|Action=16>-5000>1>-5000>5>90>2>92>0>141>6>142>7>94>5>96>1>-5000>3>-5000>7>98>0>100>0>153>0>154>3>103>1>103>5>|Caption=Delayed startup|
'|Caption=Open Control Panel|Action=,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}|Class=SystemFolder|

    Rem - GetDesktopWindow SystemParametersInfo(SPI_GETWORKAREA...) only work for 1 screen, not for multiple
    DTP_Handle = FindWindow("PROGMAN", vbNullString)
    Call GetWindowRect(DTP_Handle, DTP_Area)

    #If IDE = 1 Then
        Call MsgBox("IDE mode = 1")
    #End If
    #If LOGMODE > 0 Then
        Call LOG_Open
        Call LOG_Write("=================================================")
        Call LOG_Write("ZenKEY started - " & Format(Now, "Long date") & ", " & Format(Now, "Long Time"))
        Call LOG_Write("=================================================")
        
        Rem - Use the desktop
        Call LOG_Write("Desktop area - " & CStr(DTP_Area.left) & ", " & CStr(DTP_Area.Top) & ", " & CStr(DTP_Area.Right) & ", " & CStr(DTP_Area.Bottom))
        
    #End If

    Rem - Register the message to notify ZenKEY of a pending action.
    WIN_ZKAction = RegisterWindowMessage(ByVal "ZenKEY Action")

    Call Init_ZK
    Set MainForm = New frmZenKEY
    Call MainForm.Vars_Init
    Call MainForm.Load_Graphics
    Call MainForm.SetFormRegion
    
    Dim bRun As Boolean
    Select Case Command$
        Case ""
            If App.PrevInstance Then
                Call ZenMB("Hmm.. ZenKEY is already running. Click on the icon in the System tray (usually the bottom right of screen) or try pressing 'Alt + Space' to access the program. ZenKEY? Click? ", "OK")
            Else
                If settings("ShowSplash") <> "N" Then Call Shell(App.Path & "\ZkConfig.exe SPLASH", vbNormalFocus)
                bRun = True
            End If
        Case "RESTART"
            bRun = True
        Case Else
            'If CBool(Len(Prop_Get("Action", Command$)) > 0) Then
                If Not SendToZenKEY(Command$) Then
                    Rem - We have been sent a command to execute with no previous instance.
                    ReDim ZKMenu(0)
                    Dim zDic As New clsZenDictionary
                    Call zDic.FromProp(Command$)
                    Call ZK_GetObject(Prop_Get("Class", Command$)).DoAction(zDic)
                    Unload MainForm
                End If
            'End If
    End Select
    
    Rem - DoEvents to try to prevent corrupt drawing on startup....
    If bRun Then
        DoEvents
        Call MainForm.Initialise
        If DTM_Enabled Then Call ZK_DTM.DrawWindows
        If SET_ZenBar Then Call Icon_Layout
        Call Icon_MakeVis
        Call Hotkeys.WaitForMessages
    End If
    
End Sub

Public Function ZK_GetObject(ByVal TypeName As String) As Object
Static cWinamp As clsWinamp
Static cFile As clsFile
Static cFolder As clsFolder
Static cSystem As clsSystem
Static cMedia As clsMedia
Static cSearch As frmSearch
Static cKeyStrokes As clsKeyStroke

    'TODO: Repalce with new ?
    Select Case UCase(TypeName)
        Case "WINAMP"
            If cWinamp Is Nothing Then Set cWinamp = New clsWinamp
            Set ZK_GetObject = cWinamp
        Case "WINDOWS", "WINDOWSEL"
            Set ZK_GetObject = ZK_Win
        Case "FILE", "URL"
            If cFile Is Nothing Then Set cFile = New clsFile
            Set ZK_GetObject = cFile
        Case "MEDIA"
            If cMedia Is Nothing Then Set cMedia = New clsMedia
            Set ZK_GetObject = cMedia
        Case "FOLDER", "SPECIALFOLDER", "SYSTEMFOLDER"
            If cFolder Is Nothing Then Set cFolder = New clsFolder
            Set ZK_GetObject = cFolder
        Case "SYSTEM"
            If cSystem Is Nothing Then Set cSystem = New clsSystem
            Set ZK_GetObject = cSystem
        Case "SEARCH"
            If cSearch Is Nothing Then Set cSearch = New frmSearch
            Set ZK_GetObject = cSearch
        Case "KEYSTROKES"
            If cKeyStrokes Is Nothing Then Set cKeyStrokes = New clsKeyStroke
            Set ZK_GetObject = cKeyStrokes
        Case "IDT"
            If ZK_IDT Is Nothing Then Set ZK_IDT = New clsIDT
            Set ZK_GetObject = ZK_IDT
        Case "DTM"
            If ZK_DTM Is Nothing Then Set ZK_DTM = New frmDesktopMap
            Set ZK_GetObject = ZK_DTM
        Case Else '"ZENKEY"
            Set ZK_GetObject = MainForm 'Forms(0)
    End Select

End Function




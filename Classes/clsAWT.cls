VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAWT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private colTransActive As Collection
Rem ------------------- For the Auto-Window transparency --------------
Private colAWTOmit As Collection
Private colAWTOmitClasses As Collection
Private AWT_Index As Long
Private AWT_History() As Long
Private lngActiveTrans As Long, lngInactiveTrans As Long
Private lngTaskbar As Long
Private booTBSwitch As Boolean
Private booATWUndoTB As Boolean


Private Sub Class_Initialize()
  
    Set colTransActive = New Collection
    Set colAWTOmit = New Collection
    Set colAWTOmitClasses = New Collection
  
    Rem - Enable the Auto-window transparency feature

    Rem - Add Zenkey config so it controls its own transparency
    Call INI_LoadFiles("OmitList.ini", colAWTOmit, True)
    colAWTOmit.Add "zkconfig.exe"
    colAWTOmit.Add "ZenKEY.exe"
    colAWTOmit.Add "ZenDim.exe"
    colAWTOmit.Add "dwm.exe"
    
    Rem - Add the classes to ignore
    Call INI_LoadFiles("OmitListClasses.ini", colAWTOmitClasses, False)
    Rem - Progman = Parent Desktop window classname = SHELLDLL_DefView
    colAWTOmitClasses.Add "PROGMAN" ' Desktop
    colAWTOmitClasses.Add "SHELLDLL_DefView" ' System tray
    'colAWTOmitClasses.Add "MSTaskSwWClass" ' Taskbar
    'colAWTOmitClasses.Add "Shell_TrayWnd" '  Tray
    'colAWTOmitClasses.Add "SysPager" '  Tray
    'colAWTOmitClasses.Add "TrayNotifyWnd" ' Tray
    If IsWindows8 Then
        Call colAWTOmitClasses.Add("WorkerW")
        Call colAWTOmitClasses.Add("Shell_CharmWindow")
        Call colAWTOmitClasses.Add("ImmersiveSplashScreenWindowClass")
        Call colAWTOmitClasses.Add("Windows.UI.Core.CoreWindow")
        Call colAWTOmitClasses.Add("ImmersiveSwitchList")
        Call colAWTOmitClasses.Add("ImmersiveLauncher")
        Call colAWTOmitClasses.Add("SearchPane")
        Call colAWTOmitClasses.Add("MetroGhostWindow")
        Call colAWTOmitClasses.Add("Snapped Desktop")
        Call colAWTOmitClasses.Add("ImmersiveGutter")
    End If
    
    ReDim AWT_History(0 To AWT_Depth - 1)
    Rem - Set active/inactive transparency. Anything less than 1 is opaque
    lngInactiveTrans = Val(settings("TransInactive"))
    lngActiveTrans = Val(settings("TransActive"))
    
    Rem - Prevent taskbar transparency when exlcuded apps are active (prevent media window blank).
    booATWUndoTB = CBool(settings("ATWUndoTB") <> "N")
    
End Sub





Public Sub SetAutoTrans(ByVal NewHWnd As Long)
Dim lngLastHwnd  As Long
Dim booNewLegal As Boolean
Dim booOldLegal As Boolean
        
    Rem - Get the last handle, and make sure we are not working with the same window.
    Dim k As Long
    For k = 1 To AWT_Depth
        If (AWT_History(k - 1) = NewHWnd) Then Exit Sub
    Next k
    
    AWT_Index = (AWT_Index + 1) Mod AWT_Depth
    lngLastHwnd = AWT_History(AWT_Index)
    If IsLegalWindow(lngLastHwnd) Then
        Rem - Set the inactive transparency
        booOldLegal = True
        Call AWT_SetTrans(lngLastHwnd, lngInactiveTrans)
        AWT_LastTrans = lngLastHwnd
    End If
    
    If IsLegalWindow(NewHWnd) Then
        Rem - Set active transparency
        booNewLegal = True
        Call AWT_SetTrans(NewHWnd, lngActiveTrans)
    End If
    
    Rem - Maintain a list of windows so that we can undo
    AWT_History(AWT_Index) = NewHWnd ' For active window history list

    Rem - FIX 2 : Media player screen blanks when taskbar transparent.
    Rem - The moment an excluded app has got focus, undo taskbar transparency
    If booATWUndoTB Then
        If Not booNewLegal Then
            If CBool(settings("ATWTaskbar") <> "N") Then
                If lngTaskbar <> 0 Then
                    Call AWT_SetTrans(lngTaskbar, -1)
                    booTBSwitch = True
                End If
            End If
        Else
            If booTBSwitch Then
                Call AWT_SetTrans(lngTaskbar, lngInactiveTrans)
                booTBSwitch = False
            End If
        End If
    End If
    
End Sub

Private Function IsLegalWindow(ByVal hwnd As Long) As Boolean
Rem - Check if the window is in the Omit list
Dim clsName As String

    Rem - Check it is not the desktop window being made transparent...
    IsLegalWindow = True
    clsName = ClassName(hwnd)
    If Len(clsName) = 0 Then
            Rem - hwnd = 0 or the mysterious window handle that cannot be identified?
            IsLegalWindow = False
    ElseIf clsName = "SHELL_TRAYWND" Then
            Rem - SHELL_TRAYWND = Taskbar
            IsLegalWindow = CBool(settings("ATWTaskbar") <> "N")
            lngTaskbar = hwnd
    ElseIf dhc_Contains(colAWTOmitClasses, clsName) Then
            IsLegalWindow = False
    End If
    
    If IsLegalWindow Then
        Dim k As Long
        
        Rem - Check if we should avoid already transparent windows
        If IsLayered(hwnd) Then
            If Not dhc_Contains(colTransActive, hwnd) Then
                IsLegalWindow = False
                Exit Function
            End If
        End If
    
        Rem - Okay, check that it is not an omitted window, either by handle it name
        If colAWTOmit.Count > 0 Then
            If dhc_Contains(colAWTOmit, GetFileName(GetExeFromHandle(hwnd))) Then
                IsLegalWindow = False
            ElseIf dhc_Contains(colAWTOmit, hwnd) Then
                IsLegalWindow = False
            End If
        End If
    End If
    
    
End Function

Private Function IsLayered(ByVal lngHWnd As Long) As Boolean
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Dim Ret As Long
    
        Rem - Set the window style to 'Layered'
        Ret = GetWindowLong(lngHWnd, GWL_EXSTYLE)
        IsLayered = CBool(Ret And WS_EX_LAYERED)
        
End Function


Private Sub AWT_SetTrans(ByVal hwnd As Long, ByVal lngTrans As Long)

    Call ZK_Win.SetTrans(hwnd, lngTrans)
    If lngTrans < 1 Then
        Rem - Transparency below 1 removes layering
        Call dhc_Remove(colTransActive, hwnd)
    Else
        Call dhc_Add(colTransActive, hwnd)
    End If
        
End Sub

Public Sub AWT_Flush()
Dim k As Long

    For k = colTransActive.Count To 1 Step -1
        Call ZK_Win.SetTrans(colTransActive(k), -1)
    Next k
    
End Sub


Public Sub AWT_OmitList(ByVal Handle As Long, ByVal Add As Boolean)
On Error Resume Next

    If Add Then
        Rem - If we made it transparent, then undo it.
        If dhc_Contains(colTransActive, Handle) Then
            Call ZK_Win.SetTrans(Handle, -1)
            Call dhc_Remove(colTransActive, Handle)
        End If
        colAWTOmit.Add Handle
    Else
        Rem - Remove
        Call dhc_Remove(colAWTOmit, Handle)
    End If
    
End Sub


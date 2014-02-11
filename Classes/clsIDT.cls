VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private colIDTOmit As Collection
Public Function IsUseable(ByVal hwnd As Long) As Boolean
Rem - Check if the window is in the Omit list
Dim strFile As String, k As Long
Dim item As Variant

    IsUseable = True
        
    Rem - As there is always at least 1 item in the list, always loop
    strFile = GetFileName(GetExeFromHandle(hwnd))
    For k = colIDTOmit.Count To 1 Step -1
        item = colIDTOmit.item(k)
        If VarType(item) = vbString Then
            If strFile = item Then IsUseable = False
        Else
            If hwnd = item Then IsUseable = False
        End If
        If Not IsUseable Then Exit Function
    Next

End Function



Public Sub CenterOnWindow(ByRef hwnd As Long)
Dim wp As WINDOWPLACEMENT
Dim lngX As Long, lngY As Long

    wp.Length = Len(wp)
    Call GetWindowPlacement(hwnd, wp)

    'WMV_X = 0.5 * MainForm.ScaleX(Screen.Width, vbTwips, vbPixels) - 0.5 * (wp.rcNormalPosition.Left + wp.rcNormalPosition.Right)
    'WMV_Y = 0.5 * MainForm.ScaleY(Screen.Height, vbTwips, vbPixels) - 0.5 * (wp.rcNormalPosition.Top + wp.rcNormalPosition.Bottom)
    'WMV_TotX = WMV_TotX + WMV_X
    'WMV_TotY = WMV_TotY + WMV_Y
    
    lngX = 0.5 * MainForm.ScaleX(Screen.Width, vbTwips, vbPixels) - 0.5 * (wp.rcNormalPosition.left + wp.rcNormalPosition.Right)
    lngY = 0.5 * MainForm.ScaleY(Screen.Height, vbTwips, vbPixels) - 0.5 * (wp.rcNormalPosition.Top + wp.rcNormalPosition.Bottom)
    
    Rem - Set the desktop area for the movewindow command to use
    Call MoveAllWindows(lngX, lngY, True)

End Sub
Public Sub DoAction(ByVal prop As clsZenDictionary)
Dim strAction As String, k As Long
Dim lngX As Long, lngY As Long

    On Error GoTo ErrorTrap:

    strAction = prop("Action")
    If Not p_IsLegal(strAction) Then Exit Sub
    Select Case strAction
        Case "RETHOME" ' Return to original desktop position
            lngX = -WMV_TotX
            lngY = -WMV_TotY
            Call MoveAllWindows(lngX, lngY, True)
        Case "BRINGTODT" ' Bring window onto desktop
            Dim WinPlace As WINDOWPLACEMENT
            Dim rtn As Long
            
            WinPlace.Length = Len(WinPlace)
            rtn = GetWindowPlacement(ActiveWindow(0), WinPlace)
            If rtn <> 0 Then
                WinPlace.showCmd = SW_NORMAL
                Call SetWindowPlacement(ActiveWindow(0), WinPlace)
            End If
        Case "DTCACTIVE" ' Center desktop of the active window
            Call CenterOnWindow(ActiveWindow(0))
        Case "MakeAllVisible" ' On exit, ensure all windows are visible
            IDT_Enabled = False
            Call MoveAllWindows(0, 0, True)
            IDT_Enabled = True
        Case Else ' "DTUP"
            Rem - Must be a Desktop map movement
            Select Case strAction
                Case "DTUP" ' ' Move one desktop up
                    lngY = -(DTP_Area.Bottom - DTP_Area.Top): lngX = 0
                Case "DTHALFUP" ' Move half a desktop up
                    lngY = -0.5 * (DTP_Area.Bottom - DTP_Area.Top): lngX = 0
                Case "DTDOWN" ' Move one desktop down
                    lngY = (DTP_Area.Bottom - DTP_Area.Top): lngX = 0
                Case "DTHALFDOWN" ' Move half a desktop down
                    lngY = 0.5 * (DTP_Area.Bottom - DTP_Area.Top): lngX = 0
                Case "DTLEFT" ' Move one desktop left
                    lngY = 0: lngX = -(DTP_Area.Right - DTP_Area.left)
                Case "DTHALFLEFT" 'Move half a desktop left
                    lngY = 0: lngX = -0.5 * (DTP_Area.Right - DTP_Area.left)
                Case "DTRIGHT" ' Move one desktop right
                    lngY = 0: lngX = (DTP_Area.Right - DTP_Area.left)
                Case "DTHALFRIGHT" ' Move half a desktop right
                    lngY = 0: lngX = 0.5 * (DTP_Area.Right - DTP_Area.left)
                Case Else
                    ' Delegate to the Desktop map to preserve compatibility with DTM and ID being one IDT class
                    Call ZK_GetObject("DTM").DoAction(prop)
                    Exit Sub
            End Select
            Call MoveAllWindows(lngX, lngY, True)
            If IDT_AutoFocus Then
                DoEvents
                If Not SetFocusToLastActive Then
                    ' If we could not find the last active, try the one under the cursor
                    k = WindowFromCursor
                    If k > 0 Then Call SetForegroundWindow(k)
                End If
            End If
    End Select
    Exit Sub
    
ErrorTrap:
    Call ZenErr("Error " & CStr(Err.Number) & ", " & Err.Description & " in clsIDT")
    Resume Next
End Sub


Public Sub OmitList(ByVal Handle As Long, ByVal Add As Boolean)
On Error Resume Next

    If Add Then
        Rem - Add to the list
        Call colIDTOmit.Add(Handle)
    Else
        Rem - Remove
        Call dhc_Remove(colIDTOmit, Handle)
    End If
    
End Sub
Private Function p_IsLegal(ByVal strAction As String) As Boolean
    
    p_IsLegal = True
    Select Case strAction
        Case "BRINGTODT", "DTCACTIVE", "RETHOME", "DTUP", "DTHALFUP", "DTDOWN", _
                "DTHALFDOWN", "DTLEFT", "DTHALFLEFT", "DTRIGHT", "DTHALFRIGHT", "MakeAllVisible"
            If Not IDT_Enabled Then
                Call ZenMB("Sorry, but the 'Infinite Desktop' must be enabled in order to perform this function.")
                p_IsLegal = False
            End If
    End Select

End Function

Private Sub Class_Initialize()

    Set colIDTOmit = New Collection
    Rem - Always exclud ZenKEY form desktop moves
    Call INI_LoadFiles("IDTOmit.ini", colIDTOmit, True)
    colIDTOmit.Add "ZenKEY.exe"

End Sub



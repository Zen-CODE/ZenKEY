Attribute VB_Name = "modIcon"
Option Compare Text
Option Explicit
Public ICN_Forms As Collection
Public Const ICN_Height As Long = 44
Public Const ICN_HGap As Long = 2
Public Const ICN_VGap As Long = 4
Public Const ICN_TLeft As Long = 34
Public Const ICN_LineGap As Long = 16
Private booWinMode As Boolean
Public Function Icon_FlushExe(ByVal ExeName As String) As Long
Rem - For when the Exe is launched/brought to front, show any instances that are iconized.
Dim k As Long, it As frmIconify

    Call Icon_Reacquire
    k = p_GetIndex(ExeName)
    If k > 0 Then
        Set it = ICN_Forms(k)
        ' We only show it if iconized to the ShowExeWin can switch between multiple application windows when not iconized
        If it.booIsIconized And it.WinList.Count > 0 Then ' If it is iconized, show it
            Call it.ShowWin
            Icon_FlushExe = it.WinList.item(1)
        End If
        If Not it.mnuPreserve.Checked Then Call it.IconClose(SET_ZenBar)
    End If
    
End Function
Public Sub Icon_FlushDead()
Rem - Remove all icons with dead window links
Dim k As Long, it As frmIconify

    For k = ICN_Forms.Count To 1 Step -1
        If ICN_Forms(k).WinList.Count > 0 Then
            If IsWindow(ICN_Forms(k).WinList(1)) = 0 Then
                Call ICN_Forms(k).IconClose(SET_ZenBar)
            End If
        Else
            Call ICN_Forms(k).IconClose(SET_ZenBar)
        End If
    Next k
    
End Sub

Public Sub Icon_Load(pInfo As clsZenDictionary)
Dim strFileName As String, k As Integer
            
    strFileName = pInfo("Icon1")
    Do While Len(strFileName) > 0
        Rem - Create a new from and Iconify!
        Dim it As frmIconify
        Set it = New frmIconify
        Call it.Init(strFileName)
        k = k + 1
        strFileName = pInfo("Icon" & CStr(k + 1))
        If Len(pInfo("ICN_Forecolor")) > 0 Then it.ForeColor = Val(pInfo("ICN_Forecolor"))
    Loop
    Call Icon_Reacquire

End Sub


Public Sub Icon_Layout()
#If (ZKCONFIG <> 1) And (ZenWiz <> 1) Then
Rem ===================================================================
Rem - Align with the Map
Rem ===================================================================
Rem - 1. Initialise settings
Dim VGap As Single
Dim sngLeft As Single, sngTop As Single
Dim sngWidth As Single, sngHeight As Single
Dim k As Long, it As Form

    If ICN_Forms.Count < 1 Then Exit Sub

    VGap = MainForm.ScaleY(10, vbPixels, vbTwips)
    sngHeight = MainForm.ScaleX(ICN_Height, vbPixels, vbTwips)
    If DTM_Enabled Then Call ICN_Forms(1).Move(ZK_DTM.left, ZK_DTM.Top + ZK_DTM.Height + VGap, ZK_DTM.Width, sngHeight)
    If ICN_Forms.Count > 1 Then
        Rem - Layout the things
        Set it = ICN_Forms(1)
        For k = 2 To ICN_Forms.Count
            ICN_Forms(k).Move it.left, it.Top + sngHeight + VGap, it.Width, sngHeight
            Set it = ICN_Forms(k)
        Next k
    End If
#End If
End Sub



Public Sub Icon_Make(ByVal hwnd As Long, ByVal Iconify As Boolean)
Dim strFileName As String
Dim k As Long

    Rem - Check to see if the exe belongs to an already existing icon from
    strFileName = GetExeFromHandle(hwnd)
    Call Icon_Reacquire
    k = p_GetIndex(strFileName)
    If k > 0 Then
        If Iconify Then
            If ICN_Forms(k).booIsIconized Then
                Call ICN_Forms(k).UnIconify
            Else
                Call ICN_Forms(k).Iconify
            End If
        End If
    Else
        Rem - No icon exists. Create a new form and Iconify!
        Dim it As frmIconify
        Set it = New frmIconify
        With it
            Call .Init(GetExeFromHandle(hwnd))
            If booWinMode Then Call .WinList.Add(hwnd) Else Call Icon_Reacquire
            If Iconify Then
                Call .Iconify
                Call SetWinPos(.hwnd, SET_Layer, False)
            End If
        End With
        If SET_ZenBar Then Call Icon_Layout
    End If

End Sub

Public Sub Icon_Remove(ByVal hwnd As Long)
Dim strFileName As String
Dim k As Long

    Rem - Check to see if the exe belongs to an already existing icon from
    strFileName = GetExeFromHandle(hwnd)
    k = p_GetIndex(strFileName)
    If k > 0 Then
        Dim it As frmIconify
        Set it = ICN_Forms(k)
        Call it.IconClose(SET_ZenBar)
    End If
        
End Sub

Public Sub Icon_RemoveAll(Optional ByRef SaveDic As clsZenDictionary = Nothing)
Dim k As Long, lngCounter As Long
Dim it As frmIconify, bSave As Boolean

    bSave = Not (SaveDic Is Nothing)
    
    Rem - First save the forecolour if appropriate
    If bSave Then
        If ICN_Forms.Count > 0 Then SaveDic("ICN_Forecolor") = CStr(ICN_Forms(1).ForeColor)
    End If

    Rem - Check to see if the exe belongs to an already existing icon from
    For k = ICN_Forms.Count To 1 Step -1
        Set it = ICN_Forms(k)
        If bSave Then
            If it.mnuPreserve.Checked Then
                lngCounter = lngCounter + 1
                SaveDic("Icon" & CStr(lngCounter)) = it.FileName
            End If
        End If
        Call it.IconClose(False)
    Next k
    If bSave Then SaveDic("Icon" & CStr(lngCounter + 1)) = ""
    
End Sub



Public Function Icon_UpdateAll()
Dim k As Long

    For k = ICN_Forms.Count To 1 Step -1
        If ICN_Forms(k).WinList.Count > 0 Then Call ICN_Forms(k).Refresh
    Next k
    
End Function



Public Sub Icon_Reacquire()
Rem ---------------------------------------------------
Rem - Pump all the window information into the icons
Rem ---------------------------------------------------

    Dim k As Long
    booWinMode = CBool(Len(settings("ICN_WinMode")) = 0)
    For k = ICN_Forms.Count To 1 Step -1
        If Not ICN_Forms(k).booIsIconized Then Call dhc_RemoveAll(ICN_Forms(k).WinList)
    Next k
    Call EnumWindows(AddressOf EnumWinList, ByVal 0&)

End Sub
Private Function EnumWinList(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    
    Rem - continue enumeration
    EnumWinList = True
    If IsWindowVisible(hwnd) Then
        Rem - Use on error as Win NT does not support this call
        Dim lngRet As Long
        Const GWL_STYLE = (-16)
        Const WS_CAPTION = &HC00000

        lngRet = GetWindowLong(hwnd, GWL_STYLE)
        If (lngRet And WS_CAPTION) Then
            On Error Resume Next
            Dim strApp As String
            Dim k As Long, it As frmIconify
            strApp = GetFileName(GetExeFromHandle(hwnd))
            For k = ICN_Forms.Count To 1 Step -1
                Rem - Match according to owner handle
                Set it = ICN_Forms(k)
                If strApp = it.ExeName Then
                    If Not it.booIsIconized Then
                        If it.WinList.Count < 1 Or (Not booWinMode) Then Call it.WinList.Add(hwnd)
                    End If
                    Exit Function
                End If
            Next k
        End If
    End If
    
End Function



Private Function p_GetIndex(ByVal FName As String) As Long
Dim k As Long
    
    For k = ICN_Forms.Count To 1 Step -1
        If ICN_Forms(k).FileName = FName Then
            p_GetIndex = k
            Exit Function
        End If
    Next k

End Function

Public Sub Icon_MakeVis()
Dim k As Long
        
    For k = ICN_Forms.Count To 1 Step -1
        Call SetWinPos(ICN_Forms(k).hwnd, SET_Layer, False)
    Next k

End Sub

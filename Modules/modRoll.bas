Attribute VB_Name = "modRoll"
Option Explicit
Const ROL_Normal As Long = 0
Const ROL_Rolled As Long = 1
Public Type RolledWindow
    Up As Boolean
    hwnd As Long
    Top As Long
    Height As Long
    State As Long
    'Tries As Long
End Type
Public ROL_Windows() As RolledWindow
Public ROL_Count As Long
Private lngLastActive As Long
Public ROL_Check As Boolean
Public Function ROL_GetIndex(ByVal lngWin As Long) As Long
Dim k As Long

    For k = 1 To ROL_Count
        Rem - Window specified- return only  a perfect match
        If lngWin = ROL_Windows(k).hwnd Then
            ROL_GetIndex = k
            Exit Function
        End If
    Next k

End Function


Public Sub ROL_RestoreAll()
Rem - Unrollup all rolled windows.
Dim k As Long, RecTan As RECT

    For k = ROL_Count To 1 Step -1
        Rem - Restore window to original height
        With ROL_Windows(k)
            Call GetWindowRect(.hwnd, RecTan)
            Rem - Set this information so windows automatically unrolled can be re-rolled
            RecTan.Top = .Top
            RecTan.Bottom = RecTan.Top + .Height
            Call PlaceWindow(.hwnd, RecTan)
        End With
    Next k
    Rem - Seeings as this is only called on CloseApp, no need to Erase the array?
    'Erase ROL_Windows()

End Sub


Public Sub ROL_Focus(ByVal lngHWnd As Long)
Dim k As Long

    If lngLastActive <> lngHWnd Then
        Rem - A new window is active. See if we should rollup the last one
        k = ROL_GetIndex(lngLastActive)
        If k > 0 Then
            If ROL_Windows(k).State <> ROL_Rolled Then
                Call ROL_Toggle(lngLastActive, ROL_Windows(k).Up)
            End If
        End If
    End If
    
    Rem - Unroll any rolled up windows with this handle.
    k = ROL_GetIndex(lngHWnd)
    If k > 0 Then
        If ROL_Windows(k).State <> ROL_Normal Then
            Call ROL_Toggle(lngHWnd, ROL_Windows(k).Up)
        'ElseIf IsIconic(lngHWnd) Then
        '    ROL_Check = False
        Else
            Dim rctR As RECT
            Call GetWindowRect(lngHWnd, rctR)
            
            #If LOGMODE > 0 Then
                Call LOG_Write("Checking " & GetExeFromHandle(lngHWnd) & " (" & CStr(lngHWnd) & ") - Req height = " & CStr(ROL_Windows(k).Height) & ", Actual = " & CStr(rctR.Bottom - rctR.Top) & " - " & CStr(Now))
            #End If
            Rem - To counter rare glitch where window does not unroll, try a few times
            Rem - NOTE: This code does not seem to work inside of PlaceWindow, only after
            
            
            If rctR.Bottom - rctR.Top + 1 < ROL_Windows(k).Height Then
                #If LOGMODE > 0 Then
                    Call LOG_Write(" Failure detected - " & CStr(Now))
                #End If
                rctR.Top = ROL_Windows(k).Top
                rctR.Bottom = rctR.Top + ROL_Windows(k).Height
                Call PlaceWindow(lngHWnd, rctR)
            Else
                ROL_Check = False
            End If
                            
        End If
    Else
        ROL_Check = False
    End If
    
    lngLastActive = lngHWnd

End Sub

Public Function ROL_Toggle(ByVal lngWin As Long, ByVal booUp As Boolean) As String
Const SM_CYMIN = 29 'Minimum height of window
Const RollMax = 35
Dim lngIndex As Long
Dim lngMin As Long
Dim RecTan As RECT

    lngIndex = ROL_GetIndex(lngWin)  ' Get the index of the Exe
    If GetWindowRect(lngWin, RecTan) <> 1 Then
        Rem - Failed to get window
        ROL_Toggle = "Unable to determine this window's placement for some reason. How strange?"
        If lngIndex > 0 Then
            Call ROL_Remove(lngIndex)
            lngIndex = 0
        End If
    ElseIf lngIndex < 1 Then
        Rem - Not in the window list.
        If RecTan.Bottom - RecTan.Top > RollMax Then
            Rem - Store the values if automatically unrolled.
            ROL_Count = ROL_Count + 1
            ReDim Preserve ROL_Windows(1 To ROL_Count)
            lngIndex = ROL_Count
            ROL_Windows(ROL_Count).hwnd = lngWin
        Else
            ROL_Toggle = "Sorry, but this window is too small to rollup."
        End If
    End If
    
    If IsIconic(lngWin) Then
        ROL_Toggle = "Sorry, but cannot performs rolling on minimized windows."
        lngIndex = 0
    End If
    
    If lngIndex > 0 Then
        
        With ROL_Windows(lngIndex)
            Rem - It is in the window list. Roll if up if it is not in that state, otherwise restores it
            If .State = ROL_Rolled Then
                RecTan.Top = .Top
                RecTan.Bottom = RecTan.Top + .Height
                .State = ROL_Normal
                '#If LOGMODE > 0 Then
                '    Call LOG_Write("Restore - " & GetExeFromHandle(lngWin) & " (" & CStr(lngWin) & ") - Height = " & CStr(RecTan.Bottom - RecTan.Top) & " - " & CStr(Now))
                '#End If
            Else
                '#If LOGMODE > 0 Then
                '    Call LOG_Write("Rollup - " & GetExeFromHandle(lngWin) & " (" & CStr(lngWin) & ") - Height = " & CStr(RecTan.Bottom - RecTan.Top) & " - " & CStr(Now))
                '#End If
                lngMin = GetSystemMetrics(SM_CYMIN)
                .Up = booUp
                .Height = RecTan.Bottom - RecTan.Top
                .Top = RecTan.Top
                If Not booUp Then RecTan.Top = RecTan.Bottom - lngMin ' Rolldown
                RecTan.Bottom = RecTan.Top + lngMin
                .State = ROL_Rolled
            End If
        End With
        Call PlaceWindow(lngWin, RecTan)
        
    End If
    ROL_Check = CBool(ROL_Count > 0)
        
End Function
Public Sub ROL_Remove(ByVal Index As Long)
    
    Rem - Now remove it the Item
    If Index < ROL_Count Then ROL_Windows(Index) = ROL_Windows(ROL_Count)
    ROL_Count = ROL_Count - 1
    If ROL_Count > 0 Then
        ReDim Preserve ROL_Windows(1 To ROL_Count)
    Else
        Erase ROL_Windows()
    End If
    ROL_Check = CBool(ROL_Count > 0)
    
End Sub
Public Sub ROL_RemoveNorm(ByVal lngHWnd As Long)
Rem - Remove the window from the rolled collection only if it is in a normal state.
Dim lngIndex As Long
    
    lngIndex = ROL_GetIndex(lngHWnd)
    If lngIndex > 0 Then
        If ROL_Windows(lngIndex).State = ROL_Normal Then Call ROL_Remove(lngIndex)
    End If
    
End Sub


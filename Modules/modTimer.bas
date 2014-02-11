Attribute VB_Name = "modTimer"
Option Explicit
#If False Then
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Type ZenTimer
    Handle As Long
    Object As Object
    Prop As String
End Type
Private ZenTimers() As ZenTimer
Private TimerCount As Long

Public Sub Timer_Set(ByVal MilliSecs As Long, ByRef clsObject As Object, ByVal Prop As String)
    
    ReDim Preserve ZenTimers(0 To TimerCount)
    With ZenTimers(TimerCount)
        .Prop = Prop
        Set .Object = clsObject
        .Handle = SetTimer(0, 0, MilliSecs, AddressOf Timer_Proc)
    End With
    TimerCount = TimerCount + 1


End Sub

Private Sub Timer_Proc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
Dim k As Long, booFound As Boolean

    For k = 0 To TimerCount - 1
        If ZenTimers(k).Handle = uElapse Then Exit For
    Next k
    
    If k < TimerCount Then
        
        Rem - Fire the event
        With ZenTimers(k)
            Call KillTimer(0, .Handle)
            Call .Object.DoAction(.Prop)
        End With
        
        Rem - Remove the timer from the timer array
        If k < TimerCount - 1 Then ZenTimers(k) = ZenTimers(TimerCount - 1)
        TimerCount = TimerCount - 1
        If TimerCount > 0 Then ReDim Preserve ZenTimers(TimerCount - 1)
        
    End If
    
End Sub
#End If

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Timers() As frmTimer
Private lngTimers As Long
Public Sub DoAction(ByVal Action As String)
Dim It As frmTimer

    Select Case True
        Case UCase(Action) = "ADD"
            Set It = New frmTimer
            It.Init
            It.Show vbModal
            If Len(It.StateString) > 0 Then
                Rem - Add the item to the menu
                Dim lngSize As Long, lngItem As Long
                
                lngItem = (MNU_ID - ID_START) \ ID_MODVALUE
                'lngSubItem = (MNU_ID - ID_START) Mod ID_MODVALUE - 1
                lngSize = UBound(ZKMenu(lngItem).SubItems())
                ReDim Preserve ZKMenu(lngItem).SubItems(0 To lngSize + 1)
                With ZKMenu(lngItem).SubItems(lngSize + 1)
                    .Action = It.StateString ' "TIMER"
                    Set .Parent = Me
                    .Caption = Extract(.Action, 0, "*")
                    .Enabled = True
                    .ID = ID_START + lngItem + lngSize + 2
                    Rem - Add menu item for it
                    Call DynaMenu.AddMenuItem(.Caption, "Text", .ID, 2, 0)
                    If IconMode = 3 Then
                        Call DynaMenu.SetItemProp("PICTURE", Forms(0).GetPicbox(.Caption), .ID)
                    ElseIf IconMode = 2 Then
                        Call DynaMenu.SetItemProp("PICTURE", Forms(0).picSubmenu(lngSize Mod 4), .ID)
                    End If
                    It.p_MNU_ID = .ID
                End With
                Rem - Add the item to the tiemr array
                ReDim Preserve Timers(0 To lngTimers)
                Set Timers(lngTimers) = It
                lngTimers = lngTimers + 1
                
            End If
        Case Else ' Left$(Action, 5) = "TIMER"
            Set It = Timers((MNU_ID - ID_START) Mod ID_MODVALUE - 2)
            It.Show vbModal
            
    End Select
End Sub

Public Sub RemoveMe()

End Sub

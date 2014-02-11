VERSION 5.00
Begin VB.Form frmDimmer 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4290
   ClientLeft      =   900
   ClientTop       =   765
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDimmer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmDimmer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngStartX As Long, lngStartY As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call CloseApp
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CloseApp
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If lngStartY = 0 Then
        Rem - First call. Record position.
        lngStartX = X
        lngStartY = Y
    Else
        If Abs(X - lngStartX) > 1 Or Abs(Y - lngStartY) > 1 Then Call CloseApp
    End If
    
End Sub






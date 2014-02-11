VERSION 5.00
Begin VB.Form frmSystray 
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   4290
   ClientTop       =   5565
   ClientWidth     =   2730
   Icon            =   "frmSystray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   150
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   182
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "frmSystray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public OwnerForm As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Const SW_HIDE = 0
'Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
'Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Index As Long
Private Sub Form_Load()
    
    If ST_Count > -1 Then
        ReDim Preserve ST_TrayForms(0 To ST_Count)
        Set ST_TrayForms(ST_Count) = Me
        Index = ST_Count
        ST_Count = ST_Count + 1
    End If
    
    'Call MsgBox("Tray created, Count = " & CStr(ST_Count))
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const WM_RBUTTONUP = &H205
Const WM_LBUTTONUP = &H202
Dim Msg As Single
Dim CurPos As POINTAPI
    
    Msg = Me.ScaleX(X, vbPixels, vbTwips) / Screen.TwipsPerPixelX
    If (Msg = WM_LBUTTONUP) Or (Msg = WM_RBUTTONUP) Then Unload Me

End Sub



Public Sub SendToTray()
Dim RetVal As Long
Dim ModuleName As String
Dim HIcon As Long
Dim strTray As String

    ModuleName = GetExeFromHandle(OwnerForm)
    HIcon = ExtractIcon(Me.hwnd, ModuleName, 0)
    
    Rem - Set the Icon text to the form caption
    strTray = String(200, Chr$(0))
    RetVal = GetWindowText(OwnerForm, strTray, 200)
    If RetVal > 0 Then strTray = left$(strTray, InStr(strTray, Chr$(0)) - 1)
    
    
    Call Systray_Add(Me, HIcon, strTray)
    Call ShowWindow(OwnerForm, SW_HIDE)
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Me.Visible = False
    Call Systray_Del(Me)
    Call SetWinPos(OwnerForm, HWND_TOP, True)
    
    Rem - Handle the arrang changing here
    Set ST_TrayForms(Index) = Nothing
    ST_Count = ST_Count - 1
    If Index < ST_Count Then
        Set ST_TrayForms(Index) = ST_TrayForms(ST_Count)
        ST_TrayForms(Index).Index = Index
    End If
    
    If ST_Count > 0 Then
        ReDim Preserve ST_TrayForms(0 To ST_Count - 1)
    Else
        Erase ST_TrayForms()
    End If
    
End Sub






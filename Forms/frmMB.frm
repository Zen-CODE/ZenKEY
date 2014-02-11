VERSION 5.00
Begin VB.Form frmMB 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ZenKEY Message Box"
   ClientHeight    =   1545
   ClientLeft      =   9915
   ClientTop       =   6600
   ClientWidth     =   6075
   ClipControls    =   0   'False
   Icon            =   "frmMB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6075
   Begin VB.CommandButton zbButton 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblTop"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   5835
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblBot"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   5835
      WordWrap        =   -1  'True
   End
   Begin VB.Image imiZen 
      Height          =   180
      Left            =   2940
      Picture         =   "frmMB.frx":058A
      Top             =   540
      Width           =   180
   End
   Begin VB.Shape Shape 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   1
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   2475
   End
   Begin VB.Shape Shape 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   0
      Left            =   300
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   2475
   End
End
Attribute VB_Name = "frmMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lngButtons As Long








Public Sub SetGraphics()
    Rem - Set the volour of the Settings boxes
    
    With Me
        Call TileMe(Me, LoadPicture(App.Path & "\Help\clouds.jpg"))
    
        Dim sngRight As Single
        
        sngRight = .ScaleWidth - 2 * .ScaleX(.DrawWidth, vbPixels, vbTwips)
        Me.Line (0, 0)-(sngRight, Me.ScaleHeight - 2 * .ScaleX(.DrawWidth, vbPixels, vbTwips)), , B
        If lngButtons < 1 Then
            Const wide = 200, high = 200
            Me.Line (sngRight, 0)-Step(-wide, high)
            Me.Line -Step(wide, 0)
            Me.Line -Step(-wide, -high)
            Me.Line -Step(0, high)
        End If
    End With

End Sub

Private Sub Form_Activate()
On Error Resume Next

    If Me.Visible Then Call zbButton(zbButton.UBound).SetFocus
    
End Sub

Private Sub imiZen_Click()
    If lngButtons < 1 Then Call ZBButton_Click(0)

End Sub

Private Sub lblBot_Click()
    If lngButtons < 1 Then Call ZBButton_Click(0)
End Sub


Private Sub ZBButton_Click(Index As Integer)
        
    lngButtons = Index
    Unload Me

End Sub



Public Sub Initialise()
Dim k As Integer

    Rem - Do the colours
    Me.AutoRedraw = True
    
    lblTop.ForeColor = COL_Zen
    lblBot.ForeColor = COL_Zen

    Me.ForeColor = vbBlack
    Shape(0).Top = lblTop.Top + lblTop.Height + 250 '19
    Shape(1).Top = Shape(0).Top
    If Len(lblBot.Caption) > 0 Then
        lblBot.Top = Shape(0).Top + 350
    Else
        lblBot.Top = Shape(0).Top - 40
    End If
    imiZen.Top = Shape(0).Top - 80 '4
    
    Select Case lngButtons
        Case Is < 1
            Rem - No buttons
            zbButton(0).Top = lblBot.Top + lblBot.Height + 150 '19
    
        Case Else
            Rem - 2/3 buttons
            For k = 1 To lngButtons
                Load zbButton(k)
            Next k
            Dim lngDivFac As Long
            If lngButtons > 2 Then lngDivFac = 2 Else lngDivFac = lngButtons
            For k = 0 To lngButtons
                zbButton(k).left = -0.1 * Me.ScaleWidth + ((k Mod 3) + 1) * (1.2 * Me.ScaleWidth / (lngDivFac + 2)) - zbButton(k).Width * 0.5
                zbButton(k).Top = lblBot.Top + lblBot.Height + 180 + (k \ 3) * zbButton(0).Height * 1.2
                zbButton(k).Visible = True
            Next k
            
    End Select
    
    With Me
        .Height = .ScaleY(zbButton(zbButton.UBound).Top + zbButton(0).Height + 200, .ScaleMode, vbTwips) + 400
        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2, 6200
    End With
    Call SetGraphics
    'Me.Move 0.5 * (Screen.Width - Me.Width), 0.5 * (Screen.Height - Me.Height)

End Sub



Private Sub lblTop_Click()
    If lngButtons < 1 Then Call ZBButton_Click(0)
End Sub





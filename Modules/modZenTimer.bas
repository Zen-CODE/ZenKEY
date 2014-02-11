Attribute VB_Name = "modZenTimer"
Option Explicit

Public MainForm As frmTimer
Public Sub FloodBack(DestObj As Object, BackColor As Long, ForeColor As Long, GradStyle As Integer, X As Long, Y As Long)

'Paints a gradated background, fading from one color into another
'Sample Call : GradateBackground Me, &H400000, &HFF0000, 0, 0, 0
'GradStyle Modes:
'0 - vertical
'1 - circular from center
'2 - horizontal
'3 - ellipse from upper left
'4 - ellipse from upper right
'5 - ellipse from lower right
'6 - ellipse from lower left
'7 - ellipse from upper center
'8 - ellipse from right center
'9 - ellipse from lower center
'10 - ellipse from left center
'11 - ellipse from x,y - twips

Dim foo As Integer, foobar As Integer
Dim DestWidth As Integer, DestHeight As Integer, DestMode As Integer
Dim StartPnt As Integer, EndPnt As Integer, DrawHeight As Double, DrawWidth As Double
Dim dblG As Double, dblR As Double, dblB As Double
Dim addg As Double, addr As Double, addb As Double
Dim Mask As Long, Mask2 As Long, colorstep As Integer
Dim bckR As Double, bckG As Double, bckB As Double
Dim Linecolor As Long, PixelStep As Long, LineHeight As Integer
Dim PixelCount As Integer, aspect As Single
Dim CenterX As Long, CenterY As Long
On Error Resume Next

Screen.MousePointer = 11

'set up rgb bitmask
Mask = 255
Mask = Mask ^ 2
Mask2 = 255
Mask2 = 255 ^ 3

'Init dimensions in twips, set backcolor, set modes
DestMode = DestObj.ScaleMode
DestObj.ScaleMode = 1
DestHeight = DestObj.ScaleHeight
DestWidth = DestObj.ScaleWidth
DestObj.BackColor = BackColor
DestObj.AutoRedraw = True
DestObj.DrawStyle = 5 'transparent
DestObj.DrawMode = 13 'CopyPen

'solid offset
Select Case GradStyle
Case 2 'Horizontal
    StartPnt = DestWidth * 0.05
    EndPnt = DestWidth * 0.95
Case Else
    StartPnt = DestHeight * 0.05
    EndPnt = DestHeight * 0.95
    Select Case GradStyle
    Case 3 'ellipse from upper left
        CenterX = 0
        CenterY = 0
    Case 4 'ellipse from upper right
        CenterX = DestWidth
        CenterY = 0
    Case 5 'ellipse from lower right
        CenterX = DestWidth
        CenterY = DestHeight
    Case 6 'ellipse from lower left
        CenterX = 0
        CenterY = DestHeight
    Case 7 'ellipse from upper center
        CenterX = DestWidth / 2
        CenterY = 0
    Case 8 'ellipse from right center
        CenterX = DestWidth
        CenterY = DestHeight / 2
    Case 9 'ellipse from lower center
        CenterX = DestWidth / 2
        CenterY = DestHeight
    Case 10 'ellipse from left center
        CenterX = 0
        CenterY = DestHeight / 2
    Case 11 'ellipse from x,y - twips
        CenterX = X
        CenterY = Y
    End Select
End Select
aspect = DestHeight / DestWidth

Select Case GradStyle
Case 0
    DrawHeight = EndPnt - StartPnt
Case 1
    DrawHeight = Sqr((DestHeight / 2) ^ 2 + (DestWidth / 2) ^ 2)
Case 2
    DrawWidth = EndPnt - StartPnt
Case 3, 4, 5, 6
    DrawHeight = Sqr((DestHeight) ^ 2 + (DestWidth) ^ 2)
Case 7, 8, 9, 10
    If DestHeight >= DestWidth Then
        DrawHeight = DestHeight
    Else
        DrawHeight = DestWidth
    End If
Case 11
    DrawHeight = CenterX
    If Sqr(CenterY ^ 2 + CenterX ^ 2) > DrawHeight Then DrawHeight = Sqr(CenterY ^ 2 + CenterX ^ 2)
    If Sqr(CenterY ^ 2 + (DestWidth - CenterX) ^ 2) > DrawHeight Then DrawHeight = Sqr(CenterY ^ 2 + (DestWidth - CenterX) ^ 2)
    If Sqr((DestHeight - CenterY) ^ 2 + (DestWidth - CenterX) ^ 2) > DrawHeight Then DrawHeight = Sqr((DestHeight - CenterY) ^ 2 + (DestWidth - CenterX) ^ 2)
    If Sqr((DestHeight - CenterY) ^ 2 + CenterX ^ 2) > DrawHeight Then DrawHeight = Sqr((DestHeight - CenterY) ^ 2 + CenterX ^ 2)
    'DrawHeight = DrawHeight * .9
End Select
dblR = CDbl(BackColor And &HFF)
dblG = CDbl(BackColor And &HFF00&) / 255
dblB = CDbl(BackColor And &HFF0000) / &HFF00&
bckR = CDbl(ForeColor And &HFF&)
bckG = CDbl(ForeColor And &HFF00&) / 255
bckB = CDbl(ForeColor And &HFF0000) / &HFF00&
If GradStyle = 2 Then
    addr = (bckR - dblR) / (DrawWidth / Screen.TwipsPerPixelY)
    addg = (bckG - dblG) / (DrawWidth / Screen.TwipsPerPixelY)
    addb = (bckB - dblB) / (DrawWidth / Screen.TwipsPerPixelY)
Else
    addr = (bckR - dblR) / (DrawHeight / Screen.TwipsPerPixelY)
    addg = (bckG - dblG) / (DrawHeight / Screen.TwipsPerPixelY)
    addb = (bckB - dblB) / (DrawHeight / Screen.TwipsPerPixelY)
End If

DestObj.Cls

PixelStep = Screen.TwipsPerPixelY
LineHeight = PixelStep * 2
Select Case GradStyle
Case 0 'Vertical
    For foo = 1 To DrawHeight Step PixelStep
        dblR = dblR + addr
        dblG = dblG + addg
        dblB = dblB + addb
        If dblR > 255 Then dblR = 255
        If dblG > 255 Then dblG = 255
        If dblB > 255 Then dblB = 255
        If dblR < 0 Then dblR = 0
        If dblG < 0 Then dblG = 0
        If dblB < 0 Then dblB = 0
        Linecolor = RGB(dblR, dblG, dblB)
        DestObj.Line (0, foo + StartPnt)-(DestWidth, foo + StartPnt + LineHeight), Linecolor, BF
    Next foo
    For foo = EndPnt To DestHeight Step PixelStep
        DestObj.Line (0, foo)-(DestWidth, foo + LineHeight), ForeColor, BF
    Next foo
Case 2 'horizontal
    For foo = 1 To DrawWidth Step PixelStep
        dblR = dblR + addr
        dblG = dblG + addg
        dblB = dblB + addb
        If dblR > 255 Then dblR = 255
        If dblG > 255 Then dblG = 255
        If dblB > 255 Then dblB = 255
        If dblR < 0 Then dblR = 0
        If dblG < 0 Then dblG = 0
        If dblB < 0 Then dblB = 0
        Linecolor = RGB(dblR, dblG, dblB)
        DestObj.Line (foo + StartPnt, 0)-(foo + StartPnt + LineHeight, DestHeight), Linecolor, BF
    Next foo
    For foo = EndPnt To DestWidth Step PixelStep
        DestObj.Line (foo, 0)-(foo + LineHeight, DestHeight), ForeColor, BF
    Next foo
Case 1 'circular
    Screen.MousePointer = 11
    DestObj.FillStyle = 0
    PixelCount = 5
    PixelStep = PixelStep * -1 * PixelCount
    For foo = DrawHeight To 1 Step PixelStep
        dblR = dblR + (addr * PixelCount)
        dblG = dblG + (addg * PixelCount)
        dblB = dblB + (addb * PixelCount)
        If dblR > 255 Then dblR = 255
        If dblG > 255 Then dblG = 255
        If dblB > 255 Then dblB = 255
        If dblR < 0 Then dblR = 0
        If dblG < 0 Then dblG = 0
        If dblB < 0 Then dblB = 0
        Linecolor = RGB(dblR, dblG, dblB)
        DestObj.FillColor = Linecolor
        DestObj.Circle (DestWidth / 2, DestHeight / 2), foo, Linecolor, , , aspect
    Next foo
    Screen.MousePointer = 0
Case Else 'elliptical from various points
    DestObj.FillStyle = 0
    PixelCount = 5
    PixelStep = PixelStep * -1 * PixelCount
    For foo = DrawHeight To 1 Step PixelStep
        dblR = dblR + (addr * PixelCount)
        dblG = dblG + (addg * PixelCount)
        dblB = dblB + (addb * PixelCount)
        If dblR > 255 Then dblR = 255
        If dblG > 255 Then dblG = 255
        If dblB > 255 Then dblB = 255
        If dblR < 0 Then dblR = 0
        If dblG < 0 Then dblG = 0
        If dblB < 0 Then dblB = 0
        Linecolor = RGB(dblR, dblG, dblB)
        DestObj.FillColor = Linecolor
        DestObj.Circle (CenterX, CenterY), foo, Linecolor, , , aspect
    Next foo
End Select
DestObj.ScaleMode = DestMode
Screen.MousePointer = 0

End Sub


Public Sub Main()
Dim It As New frmTimer

    It.Show

End Sub

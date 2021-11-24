Attribute VB_Name = "ShowFrm"
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Sub Main()
frmShow.WindowState = 2
frmShow.Show
SetCursorPos Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY
End Sub

Public Function SpeedLessMove(mObj As Control, mToLeft As Single, mToTop As Single, Speed As Long, cTimer As Timer) As Long
Dim ml As Single, mt As Single
On Error Resume Next
If mObj.Left < mToLeft Then
    ml = mToLeft - mObj.Left
    ml = ml / Speed
    mObj.Left = mObj.Left + ml
ElseIf mObj.Left > mToLeft Then
    ml = mObj.Left - mToLeft
    ml = ml / Speed
    mObj.Left = mObj.Left - ml
End If

If mObj.Top < mToTop Then
    mt = mToTop - mObj.Top
    mt = mt / Speed
    mObj.Top = mObj.Top + mt
ElseIf mObj.Top > mToTop Then
    mt = mObj.Top - mToTop
    mt = mt / Speed
    mObj.Top = mObj.Top - mt
End If

If Round(ml) = 0 And Round(mt) = 0 Then mObj.Left = mToLeft: mObj.Top = mToTop: cTimer.Enabled = False
End Function

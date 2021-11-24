VERSION 5.00
Begin VB.Form frmShow 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer tmrChange 
      Interval        =   5000
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tmrDraw 
      Index           =   0
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Shape shpDot 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ToLeft(16) As Single, ToTop(16) As Single, Cnt As Integer, TimeSwi As Integer

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Randomize
For i = 1 To 16
Load tmrDraw(i)
Load shpDot(i)
shpDot(i).Visible = True
shpDot(i).Move Screen.Width * Rnd, Screen.Height * Rnd
tmrDraw(i).Enabled = False
Next i
tmrChange_Timer
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 0 Or Abs(X \ Screen.TwipsPerPixelX - Screen.Width \ Screen.TwipsPerPixelX) > 3 Or _
Abs(Y \ Screen.TwipsPerPixelY - Screen.Height \ Screen.TwipsPerPixelY) > 3 Then Unload Me
End Sub

Private Sub tmrChange_Timer()
Dim a As Integer
Cnt = Cnt + 1
For a = 1 To 16
If Cnt = 2 Then TimeSwi = 1
If Cnt = 5 Then TimeSwi = 2: Cnt = 0
MoveTo a, Screen.Width * Rnd, Screen.Height * Rnd
Next a
For a = 1 To 16
tmrDraw(a).Enabled = True
Next a
End Sub

Private Sub tmrDraw_Timer(Index As Integer)
If TimeSwi = 1 Then ToTime: TimeSwi = 0
If TimeSwi = 2 Then ToClock: TimeSwi = 0
SpeedLessMove shpDot(Index), ToLeft(Index), ToTop(Index), 10, tmrDraw(Index)
End Sub

Private Sub MoveTo(Index As Integer, ToL As Single, ToT As Single)
ToLeft(Index) = ToL
ToTop(Index) = ToT
End Sub

Private Sub ToTime()
Dim tNow As String, Dl As Single, Dt As Single, Dc As Single, nInx As Integer, c As Integer, n As Integer
tNow = Format(Time, "hhmm")
Dt = (Me.ScaleHeight - shpDot(0).Height * 3) / 2
Dc = shpDot(0).Height
For c = 1 To 4
nInx = (c - 1) * 4
Dl = (Me.ScaleWidth - shpDot(0).Width * 21) / 2 + (c - 1) * Dc * 6
Select Case Mid(tNow, c, 1)
Case "0"
For n = nInx + 1 To nInx + 4
MoveTo n, 0, 0
Next n
Case "1"
MoveTo nInx + 1, Dl, Dt
For n = nInx + 2 To nInx + 4
MoveTo n, 0, 0
Next n
Case "2"
MoveTo nInx + 1, 0, 0
MoveTo nInx + 4, 0, 0
MoveTo nInx + 2, Dl + Dc * 3, Dt
MoveTo nInx + 3, Dl, Dt + Dc * 3
Case "3"
MoveTo nInx + 1, 0, 0
MoveTo nInx + 3, 0, 0
MoveTo nInx + 2, Dl + Dc * 3, Dt
MoveTo nInx + 4, Dl + Dc * 3, Dt + Dc * 3
Case "4"
MoveTo nInx + 1, Dl, Dt
MoveTo nInx + 2, 0, 0
MoveTo nInx + 3, Dl, Dt + Dc * 3
MoveTo nInx + 4, Dl + Dc * 3, Dt + Dc * 3
Case "5"
MoveTo nInx + 1, Dl, Dt
MoveTo nInx + 2, 0, 0
MoveTo nInx + 3, 0, 0
MoveTo nInx + 4, Dl + Dc * 3, Dt + Dc * 3
Case "6"
MoveTo nInx + 1, Dl, Dt
MoveTo nInx + 2, Dl + Dc * 3, Dt
MoveTo nInx + 3, Dl, Dt + Dc * 3
MoveTo nInx + 4, 0, 0
Case "7"
MoveTo nInx + 1, Dl, Dt
MoveTo nInx + 2, Dl + Dc * 3, Dt
MoveTo nInx + 3, 0, 0
MoveTo nInx + 4, Dl + Dc * 3, Dt + Dc * 3
Case "8"
MoveTo nInx + 1, Dl, Dt
MoveTo nInx + 2, Dl + Dc * 3, Dt
MoveTo nInx + 3, Dl, Dt + Dc * 3
MoveTo nInx + 4, Dl + Dc * 3, Dt + Dc * 3
Case "9"
MoveTo nInx + 1, 0, 0
MoveTo nInx + 2, Dl + Dc * 3, Dt
MoveTo nInx + 3, Dl, Dt + Dc * 3
MoveTo nInx + 4, Dl + Dc * 3, Dt + Dc * 3
End Select
Next c
End Sub

Private Sub ToClock()
Dim b As Integer, Minute As Integer, Hours As Integer, BaseX As Integer, BaseY As Integer, R As Integer
Const PI = 3.1415926

    BaseX = Me.ScaleWidth / 2
    BaseY = Me.ScaleHeight / 2
    
    R = IIf(BaseX > BaseY, BaseY * 0.8, BaseY * 0.8)

    For b = 0 To 360 Step 30
        MoveTo b / 30 + 1, CSng(BaseX + (R - 3) * Sin(b * PI / 180)), CSng(BaseY - (R - 3) * Cos(b * PI / 180))
    Next b

    Minute = DatePart("n", Time)
    Hours = DatePart("h", Time)
    
    If Hours > 12 Then
        Hours = Hours - 12
    End If

    MoveTo 14, CSng(BaseX + R * 0.8 * Sin(Minute * PI / 30)), CSng(BaseY - R * 0.8 * Cos(Minute * PI / 30))
    MoveTo 15, CSng(BaseX + R * 0.6 * Sin((Hours + Minute / 60) * PI / 6)), CSng(BaseY - R * 0.6 * Cos((Hours + Minute / 60) * PI / 6))
    MoveTo 16, CSng(BaseX), CSng(BaseY)
End Sub

VERSION 5.00
Begin VB.Form s 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFF00&
   Caption         =   "Ghetto Tanks"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   472
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   5400
      Top             =   3240
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   3015
      TabIndex        =   16
      Top             =   600
      Width           =   3015
      Begin VB.Label hp11 
         Caption         =   "Health:"
         Height          =   195
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   855
      End
      Begin VB.Label hp22 
         Caption         =   "Health:"
         Height          =   195
         Left            =   2040
         TabIndex        =   17
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Part"
      Height          =   195
      Left            =   6480
      TabIndex        =   13
      Top             =   600
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   7095
      Begin VB.OptionButton Option5 
         Caption         =   "LandFill"
         Height          =   255
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Shotgun"
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Jackhammer"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nuke"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Missile"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Cluster"
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "3"
         Height          =   255
         Index           =   0
         Left            =   6600
         TabIndex        =   15
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label5 
         Caption         =   "3"
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   11
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "3"
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   9
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "3"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "3"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "3"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "3"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   22
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "3"
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   21
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label5 
         Caption         =   "3"
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   20
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "3"
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   19
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3720
      Top             =   2040
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      Height          =   255
      Left            =   4800
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   200
      X2              =   208
      Y1              =   200
      Y2              =   176
   End
   Begin VB.Label lblStrength 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Strength = 100"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   2640
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      Height          =   255
      Left            =   3000
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CChange As Boolean
Dim OldX As Long
Dim OldY As Long
Dim WwW As Long
Dim WwW2 As Long

Dim HP1 As Long
Dim HP2 As Long

Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_ASYNC = &H1


Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim L(99999) As erf
Dim L2(99999) As erf
Dim W(9999) As erf2
Dim w2(9999) As erf2
Dim P(9999) As erf2
Dim P2(9999) As erf2
Private Type Pos
    X As Integer
    Y As Integer
End Type
Private Type erf
 Y As Integer
 Color As Long
End Type
Private Type erf2
 Grow As Integer
 Damage As Integer
 X As Integer
 Y As Integer
 OldX As Integer
 OldY As Integer
 XF As Integer
 YF As Integer
 Color As Long
 type As Long
 tag As Long
 active As Boolean
 Size As Long
 Life As Integer
 Grav As Integer
End Type
Dim Movement As Integer
Dim Turn As Boolean
Private Sub Command1_Click()
Turn = True
End Sub

Private Sub Form_Load()
Dim Temp As Integer
'Picks which layout
    Temp = MsgBox("Grass - Dirt - Mars", vbYesNoCancel, "?")
    If Temp = 7 Then
        Me.BackColor = RGB(40, 40, 40)
    End If
    If Temp = 2 Then
        Me.BackColor = RGB(100, 0, 0)
    End If
    
'Sets defaults
Movement = 50
HP1 = 100
HP2 = 100
Me.Height = Screen.Height
Me.Width = Screen.Width
Turn = True
Me.Cls
Me.Picture = Nothing
Me.Refresh
Me.AutoRedraw = True

'Sets wheel positions
For i = 0 To 3
    W(i).XF = 0
    W(i).YF = 0
    W(i).X = i * 10 + 100
    W(i).Y = 2
    W(i).Color = RGB(125, 125, 125)
    w2(i).XF = 0
    w2(i).YF = 0
    w2(i).X = i * 10 + Me.ScaleWidth - 200
    w2(i).Y = 2
    w2(i).Color = RGB(125, 125, 125)
Next

L(0).Y = Me.ScaleHeight / 4 * 3
Randomize
'Me.Caption = Temp (DEBUG)

'Sets up ground
For i = 1 To Me.Width
    If Temp = 6 Then
        L(i).Y = L(i - 1).Y + Int(Rnd * 4) - Int(Rnd * 4)
        L(i).Color = RGB(0, 160 + Int(Rnd * 25), 0)
    Else
    If Temp = 7 Then
        L(i).Y = L(i - 1).Y + Int(Rnd * 6) - Int(Rnd * 6)
        L(i).Color = RGB(30 + Int(Rnd * 25) - Int(Rnd * 25), 30 + Int(Rnd * 25) - Int(Rnd * 25), 0)
    Else
    If Temp = 2 Then
        L(i).Y = L(i - 1).Y + Int(Rnd * 8) - Int(Rnd * 8)
        L(i).Color = RGB(160 + Int(Rnd * 25), 0, 0)
    End If
    End If
    End If
Next

For i = 0 To 10
    If Temp = 6 Then
        ff = Int(Rnd * Screen.Width)
        Boom Int(ff), 0, 10, 15, 0, 0, 1
    End If
    If Temp = 7 Then
        ff = Int(Rnd * Screen.Width)
        Boom Int(ff), 0, 20, 20, 0, 0, 1
End If
    If Temp = 2 Then
        ff = Int(Rnd * Screen.Width)
        Boom Int(ff), 0, 30, 30, 0, 0, 1
        ff = Int(Rnd * Screen.Width)
        Boom Int(ff), 0, 30, 30, 0, 0, 1
    End If
Next
For i = 2 To Me.Width - 8 Step 1
    L(i).Y = (L(i - 1).Y + L(i).Y + L(i + 1).Y + L(i - 2).Y + L(i + 2).Y) / 5
Next
Me.AutoRedraw = True
For i = 2 To Me.Width
    Me.DrawWidth = 3
    Me.Line (i - 2, L(i).Y)-(i - 2, Me.ScaleHeight), L(i).Color
Next
    Me.Refresh
    Me.AutoRedraw = False
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Movement = 50
    Turn = Not Turn
    If Turn = True Then
        Shape1.Visible = True
    Else
        Shape2.Visible = True
    End If
Line1.Visible = True
lblStrength.Visible = True
OldX = W(0).X 'x
OldY = W(0).Y 'y
If Turn = True Then
    Line1.X1 = W(0).X + 20
    Line1.Y1 = W(0).Y - 30
Else
    Line1.X1 = w2(0).X + 20
    Line1.Y1 = w2(0).Y - 30
End If
Line1.X2 = X
Line1.Y2 = Y
If Turn = True Then
    Shape1.Left = Line1.X2 - Shape1.Width / 2
    Shape1.Top = Line1.Y2 - Shape1.Height / 2
Else
    Shape2.Left = Line1.X2 - Shape1.Width / 2
    Shape2.Top = Line1.Y2 - Shape1.Height / 2
End If
Dim Pos1 As Pos
Dim Pos2 As Pos
Pos1.X = Line1.X2
Pos2.X = Line1.X1
Pos1.Y = Line1.Y2
Pos2.Y = Line1.Y1

lblStrength.Move Shape1.Left, Shape1.Top - lblStrength.Height
lblStrength.Caption = "Strength: " & ((Abs(Line1.X1 - Line1.X2) + Abs(Line1.Y1 - Line1.Y2)) / 2 & "| Angle " & GetAngle(Pos1, Pos2))
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 0 Then
    Line1.X2 = X
    Line1.Y2 = Y
If 1 = 1 Then '((Abs(Line1.X1 - Line1.X2) + Abs(Line1.Y1 - Line1.Y2)) / 2) < 150 Then
If Turn = True Then
    Line1.X1 = W(0).X + 20
    Line1.Y1 = W(0).Y - 30
Else
    Line1.X1 = w2(0).X + 20
    Line1.Y1 = w2(0).Y - 30
End If
Line1.X2 = X
Line1.Y2 = Y
'Shape1.Left = Line1.X2 - Shape1.Width / 2
'Shape1.Top = Line1.Y2 - Shape1.Height / 2
If Turn = True Then
    Shape1.Left = Line1.X2 - Shape1.Width / 2
    Shape1.Top = Line1.Y2 - Shape1.Height / 2
Else
    Shape2.Left = Line1.X2 - Shape1.Width / 2
    Shape2.Top = Line1.Y2 - Shape1.Height / 2
End If
If Turn = True Then
    lblStrength.Move Shape1.Left, Shape1.Top - lblStrength.Height
Else
    lblStrength.Move Shape2.Left, Shape2.Top - lblStrength.Height
End If
Dim Pos1 As Pos
Dim Pos2 As Pos
Pos1.X = Line1.X2
Pos2.X = Line1.X1
Pos1.Y = Line1.Y2
Pos2.Y = Line1.Y1
lblStrength.Caption = "Strength: " & ((Abs(Line1.X1 - Line1.X2) + Abs(Line1.Y1 - Line1.Y2)) / 2 & "| Angle " & GetAngle(Pos1, Pos2))
End If
End If
End Sub

Private Sub NewMissle(X As Integer, Y As Integer, Color As Long, Typee As Integer, XF As Integer, YF As Integer)
'Creates a new missle
For i = 0 To 99
If P(i).active = False Then
P(i).X = X
P(i).Y = Y
P(i).XF = XF
P(i).YF = YF
P(i).type = Typee
P(i).Color = Color
P(i).active = True
If Typee = 1 Then
P(i).Damage = 10
End If
If Typee = 2 Then
P(i).Damage = 30
End If
If Typee = 3 Then
P(i).Damage = 10
End If
If Typee = 4 Then
P(i).Damage = 10
End If
If Typee = 5 Then
P(i).Damage = 5
End If
If Typee = 6 Then
P(i).Damage = 20
End If
Exit Sub
End If
Next
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call sndPlaySound(ByVal App.Path & "/boom.wav", SND_ASYNC)
Dim TypeR As Integer

TypeR = 1
If Turn = True Then
'Zorders which ammo player thing ;)
    z = 0
    Label2(0).ZOrder 0
    Label3(0).ZOrder 0
    Label4(0).ZOrder 0
    Label5(0).ZOrder 0
    Label6(0).ZOrder 0
Else
    z = 1
    Label2(1).ZOrder 0
    Label3(1).ZOrder 0
    Label4(1).ZOrder 0
    Label5(1).ZOrder 0
    Label6(1).ZOrder 0
End If

Line1.Visible = False
lblStrength.Visible = False
If Option2.Value = True Then
If Label2(z).Caption > 0 Then
TypeR = 2
Label2(z).Caption = Label2(z).Caption - 1
Else
Option1.Value = True
End If
End If
If Option3.Value = True Then
If Label3(z).Caption > 0 Then
TypeR = 3
Label3(z).Caption = Label3(z).Caption - 1
Else
Option1.Value = True
End If
End If
If Option4.Value = True Then
If Label4(z).Caption > 0 Then
TypeR = 4
Label4(z).Caption = Label4(z).Caption - 1
Else
Option1.Value = True
End If
End If
If Option5.Value = True Then
If Label5(z).Caption > 0 Then
TypeR = 5
Label5(z).Caption = Label5(z).Caption - 1
Else
Option1.Value = True
End If
End If
If Option6.Value = True Then
If Label6(z).Caption > 0 Then
TypeR = 6
Label6(z).Caption = Label6(z).Caption - 1
Else
Option1.Value = True
End If
End If
'Turn = False
WwW = 6
If Turn = True Then
NewMissle W(0).X + 20, W(0).Y - 10, vbRed, TypeR, (Line1.X2 - Line1.X1) / 10, (Line1.Y2 - Line1.Y1) / 10
If TypeR = 4 Then
NewMissle W(0).X + 20, W(0).Y - 10, vbRed, TypeR, (Line1.X2 - Line1.X1) / 10 + Int(Rnd * 10) - Int(Rnd * 10), (Line1.Y2 - Line1.Y1) / 10 + Int(Rnd * 10) - Int(Rnd * 10)
NewMissle W(0).X + 20, W(0).Y - 10, vbRed, TypeR, (Line1.X2 - Line1.X1) / 10 + Int(Rnd * 10) - Int(Rnd * 10), (Line1.Y2 - Line1.Y1) / 10 + Int(Rnd * 10) - Int(Rnd * 10)
NewMissle W(0).X + 20, W(0).Y - 10, vbRed, TypeR, (Line1.X2 - Line1.X1) / 10 + Int(Rnd * 10) - Int(Rnd * 10), (Line1.Y2 - Line1.Y1) / 10 + Int(Rnd * 10) - Int(Rnd * 10)
NewMissle W(0).X + 20, W(0).Y - 10, vbRed, TypeR, (Line1.X2 - Line1.X1) / 10 + Int(Rnd * 10) - Int(Rnd * 10), (Line1.Y2 - Line1.Y1) / 10 + Int(Rnd * 10) - Int(Rnd * 10)
End If
Else
NewMissle w2(0).X + 20, w2(0).Y - 10, vbRed, TypeR, (Line1.X2 - Line1.X1) / 10, (Line1.Y2 - Line1.Y1) / 10
If TypeR = 4 Then
NewMissle w2(0).X + 20, w2(0).Y - 10, vbRed, TypeR, (Line1.X2 - Line1.X1) / 10 + Int(Rnd * 10) - Int(Rnd * 10), (Line1.Y2 - Line1.Y1) / 10 + Int(Rnd * 10) - Int(Rnd * 10)
NewMissle w2(0).X + 20, w2(0).Y - 10, vbRed, TypeR, (Line1.X2 - Line1.X1) / 10 + Int(Rnd * 10) - Int(Rnd * 10), (Line1.Y2 - Line1.Y1) / 10 + Int(Rnd * 10) - Int(Rnd * 10)
NewMissle w2(0).X + 20, w2(0).Y - 10, vbRed, TypeR, (Line1.X2 - Line1.X1) / 10 + Int(Rnd * 10) - Int(Rnd * 10), (Line1.Y2 - Line1.Y1) / 10 + Int(Rnd * 10) - Int(Rnd * 10)
NewMissle w2(0).X + 20, w2(0).Y - 10, vbRed, TypeR, (Line1.X2 - Line1.X1) / 10 + Int(Rnd * 10) - Int(Rnd * 10), (Line1.Y2 - Line1.Y1) / 10 + Int(Rnd * 10) - Int(Rnd * 10)
End If

End If
End Sub

Private Sub Boom(X As Integer, Y As Integer, Size As Integer, dig As Integer, Dam As Integer, Optional Invert As Integer, Optional NoDraw As Integer)
If X < 1 Then Exit Sub
Call sndPlaySound(ByVal App.Path & "/boom2.wav", SND_ASYNC)
On Error Resume Next
If X + Size / 2 > w2(0).X And X - Size / 2 < w2(3).X Then
HP2 = HP2 - Dam
Call sndPlaySound(ByVal App.Path & "/hit.wav", SND_ASYNC)
End If
If X + Size / 2 > W(0).X And X - Size / 2 < W(3).X Then
HP1 = HP1 - Dam
Call sndPlaySound(ByVal App.Path & "/hit.wav", SND_ASYNC)
End If
Me.DrawWidth = 1
For i = 0 To Screen.Width
L2(i).Y = L(i).Y
Next
For i = 0 To Size
If Invert = 1 Then
L(X + i).Y = L(X + i).Y - dig
Else
L(X + i).Y = L(X + i).Y + dig - i
End If
Next
For i = 1 To Size
If Invert = 1 Then
L(X - i).Y = L(X - i).Y - dig
Else
L(X - i).Y = L(X - i).Y + dig - i
End If
Next
Me.AutoRedraw = False
'Me.Cls
'Me.Picture = Nothing
Me.Refresh
Me.AutoRedraw = True

If NoDraw <> 1 Then
For i = 2 To Me.Width Step 1
If L2(i).Y <> L(i).Y Then
L2(i).Y = (L2(i - 1).Y + L2(i).Y + L2(i + 1).Y + L2(i - 2).Y + L2(i + 2).Y) / 5
Me.Line (i - 2, 0)-(i - 2, Me.ScaleHeight), Me.BackColor
Me.Line (i - 2, L(i).Y)-(i - 2, Me.ScaleHeight), L(i).Color
End If
Next
End If
Me.AutoRedraw = False
End Sub

Private Sub NewPart(Sizee As Integer, X As Integer, Y As Integer, OX As Integer, OY As Integer, Grow As Integer, Color As Long, XF As Long, YF As Long, Life As Integer, Grav As Integer)
For i = 0 To 900
If P2(i).active = False Then
P2(i).active = True
P2(i).Size = Sizee
P2(i).X = X
P2(i).Y = Y
P2(i).XF = XF
P2(i).YF = YF
P2(i).OldX = OX
P2(i).OldY = OY
P2(i).Grow = Grow
P2(i).Color = Color
P2(i).Life = Life
P2(i).Grav = Grav
Exit For
End If
Next
End Sub

Private Sub Timer1_Timer()
If HP1 < 0 Then
MsgBox "Player 2 Wins!"
HP1 = 0
End
End If
If HP2 < 0 Then
MsgBox "Player 1 Wins!"
HP2 = 0
End
End If
L(2).Y = 999
If hp11.Caption <> "Health: " & HP1 Then
hp11.Caption = "Health: " & HP1
End If
If hp22.Caption <> "Health: " & HP2 Then
hp22.Caption = "Health: " & HP2
End If
If CChange = True Then
Me.Refresh
End If
CChange = False

If HP1 > 0 Then
If GetAsyncKeyState(vbKeyRight) Then
Movement = Movement - 1
CChange = True
If Movement > 0 Then
If Turn = False Then
For i = 0 To 3
If L(W(3).X - 2).Y - 10 < L(W(3).X + 2).Y Then
W(i).XF = W(i).XF + 2
End If
If W(i).Y + W(i).YF > L(W(i).X - 2).Y Then
W(i).Y = L(W(i).X - 2).Y - 1
End If
Next
Else
For i = 0 To 3
If L(w2(3).X - 2).Y - 10 < L(w2(3).X + 2).Y Then
w2(i).XF = w2(i).XF + 2
End If
If w2(i).Y + w2(i).YF > L(w2(i).X - 2).Y Then
w2(i).Y = L(w2(i).X - 2).Y - 1
End If
Next
End If
End If
End If
If GetAsyncKeyState(vbKeyLeft) Then
Movement = Movement - 1
CChange = True
If Movement > 0 Then
If Turn = False Then
For i = 0 To 3
On Error Resume Next
If L(W(0).X - 2).Y - 8 < L(W(0).X - 4).Y Then
W(i).XF = W(i).XF - 2
Else
W(i).XF = W(i).XF + 1
End If
If W(i).Y + W(i).YF > L(W(i).X - 2).Y Then
W(i).Y = L(W(i).X - 2).Y - 1
End If
Next
Else
For i = 0 To 3
On Error Resume Next
If L(w2(0).X - 2).Y - 8 < L(w2(0).X - 4).Y Then
w2(i).XF = w2(i).XF - 2
Else
w2(i).XF = w2(i).XF + 1
End If
If w2(i).Y + w2(i).YF > L(w2(i).X - 2).Y Then
w2(i).Y = L(w2(i).X - 2).Y - 1
End If
Next
End If
End If
End If
End If

For i = 0 To 3
W(i).X = W(i).X + W(i).XF / 2
W(i).XF = W(i).XF * 0.5

If i = 1 Then
If W(i).Y - 20 > W(0).Y Then
W(i).Y = W(0).Y + 20
W(i).YF = 0 'W(i).Yf / 2
W(i).Y = (W(i).Y + W(i + 1).Y) / 2
End If
End If
If i = 2 Then
If W(i).Y - 20 > W(3).Y Then
W(i).Y = W(3).Y + 20
W(i).YF = 0 'W(i).Yf / 2
End If
End If

On Error Resume Next
If W(i).Y + W(i).YF < L(W(i).X - 2).Y Then
W(i).YF = W(i).YF + 1
W(i).Y = W(i).Y + W(i).YF
CChange = True
Else
If W(i).YF > 3 Then
W(i).Y = L(W(i).X - 2).Y - 1
W(i).YF = -W(i).YF * 0.5
CChange = True
End If
End If
Me.DrawWidth = 8
Me.Line (W(i).X, W(i).Y)-(W(i).X, W(i).Y), W(i).Color
'
w2(i).X = w2(i).X + w2(i).XF / 2
w2(i).XF = w2(i).XF * 0.5

If w2(i).Y + w2(i).YF < L(w2(i).X - 2).Y Then
w2(i).YF = w2(i).YF + 1
w2(i).Y = w2(i).Y + w2(i).YF
Else
If w2(i).YF > 3 Then
w2(i).Y = L(w2(i).X - 2).Y - 1
w2(i).YF = -w2(i).YF * 0.5
End If
End If
Me.DrawWidth = 8
Me.Line (w2(i).X, w2(i).Y)-(w2(i).X, w2(i).Y), w2(i).Color

'
Next
Me.DrawWidth = 1

For i = 0 To 2
For z = 0 To 2
If HP1 > 0 Then
Me.Line (W(i).X, W(i).Y)-(W(i + 1).X, W(i + 1).Y), vbBlack
End If
Me.Line (w2(i).X, w2(i).Y)-(w2(i + 1).X, w2(i + 1).Y), vbBlack
Next
Next
If HP1 > 0 Then
Me.Line (W(0).X, W(0).Y)-(W(0).X, W(0).Y - 15), vbBlack
Me.Line (W(3).X, W(3).Y)-(W(3).X, W(3).Y - 15), vbBlack
Me.Line (W(0).X, W(0).Y - 5)-(W(3).X, W(3).Y - 5), vbBlack
End If
Me.Line (w2(0).X, w2(0).Y)-(w2(0).X, w2(0).Y - 15), vbBlack
Me.Line (w2(3).X, w2(3).Y)-(w2(3).X, w2(3).Y - 15), vbBlack
Me.Line (w2(0).X, w2(0).Y - 5)-(w2(3).X, w2(3).Y - 5), vbBlack
For F = 0 To 15 Step 2
If HP1 > 0 Then
Me.Line (W(0).X, W(0).Y - F)-(W(3).X, W(3).Y - F), RGB(122, 122, 122)
End If

Me.Line (w2(0).X, w2(0).Y - F)-(w2(3).X, w2(3).Y - F), RGB(122, 122, 122)
Next
F = 15
If HP1 > 0 Then
Me.Line (W(0).X, W(0).Y - F)-(W(3).X, W(3).Y - F), vbBlack
End If
Me.Line (w2(0).X, w2(0).Y - F)-(w2(3).X, w2(3).Y - F), vbBlack

'dddddddddddddd
For i = 0 To 300
If P2(i).active = True Then
CChange = True
Me.DrawWidth = P2(i).Size
If Check1.Value = 1 Then
Me.Line (P2(i).X, P2(i).Y)-(P2(i).OldX, P2(i).OldY), P2(i).Color
End If
P2(i).OldX = P2(i).X
P2(i).OldY = P2(i).Y
P2(i).Life = P2(i).Life - 1
P2(i).Size = P2(i).Size + P2(i).Grow
P2(i).X = P2(i).X + P2(i).XF
P2(i).YF = P2(i).YF + P2(i).Grav
P2(i).Y = P2(i).Y + P2(i).YF
If P2(i).Life <= 0 Then
P2(i).active = False
End If
End If
Next

For i = 0 To 3
If Abs(W(i).YF) > 3 Or Abs(W(i).XF) > 3 Then
CChange = True
End If
If Abs(w2(i).YF) > 3 Or Abs(w2(i).XF) > 3 Then
CChange = True
End If
Next
For i = 0 To 90
If P(i).active = True Then
ttt = 100 + Int(Rnd * 40)
NewPart 5, P(i).X + Int(Rnd * 10) - Int(Rnd * 10), P(i).Y + Int(Rnd * 10) - Int(Rnd * 10), P(i).X, P(i).Y, 7, RGB(ttt, ttt, ttt), 0, 0, 8, 0
CChange = True
Me.DrawWidth = 3
P(i).YF = P(i).YF + 1
P(i).OldX = P(i).X
P(i).OldY = P(i).Y
P(i).X = P(i).X + P(i).XF
P(i).Y = P(i).Y + P(i).YF
On Error Resume Next
If P(i).Y + P(i).YF > L(P(i).X - 2).Y Then
P(i).active = False
If P(i).type <> 5 Then
If P(i).X > 0 Then
NewPart 40, P(i).X, P(i).Y + 20, P(i).X, P(i).Y + 20, 5, RGB(200 + Int(Rnd * 10), 0, 0), 0, 0, 30, 0
NewPart 20, P(i).X, P(i).Y + 20, P(i).X, P(i).Y + 20, 5, RGB(190 + Int(Rnd * 50), 30 + Int(Rnd * 50), 0), 0, 0, 30, 0
NewPart 5, P(i).X, P(i).Y + 20, P(i).X, P(i).Y + 20, 4, RGB(200 + Int(Rnd * 55), 200 + Int(Rnd * 55), 0), 0, 0, 30, 0
End If
End If
If P(i).X > 0 Then
NewPart 8 + Int(Rnd * 3) - Int(Rnd * 3), P(i).X, P(i).Y, P(i).X, P(i).Y, 0, L(P(i).X - 2).Color, Int(Rnd * 10) - Int(Rnd * 10), -10 + -Int(Rnd * 10), 30 + Int(Rnd * 6) - Int(Rnd * 6), 1
NewPart 8 + Int(Rnd * 3) - Int(Rnd * 3), P(i).X, P(i).Y, P(i).X, P(i).Y, 0, L(P(i).X - 2).Color, Int(Rnd * 10) - Int(Rnd * 10), -10 + -Int(Rnd * 10), 30 + Int(Rnd * 6) - Int(Rnd * 6), 1
NewPart 8 + Int(Rnd * 3) - Int(Rnd * 3), P(i).X, P(i).Y, P(i).X, P(i).Y, 0, L(P(i).X - 2).Color, Int(Rnd * 10) - Int(Rnd * 10), -10 + -Int(Rnd * 10), 30 + Int(Rnd * 6) - Int(Rnd * 6), 1
NewPart 8 + Int(Rnd * 3) - Int(Rnd * 3), P(i).X, P(i).Y, P(i).X, P(i).Y, 0, L(P(i).X - 2).Color, Int(Rnd * 10) - Int(Rnd * 10), -10 + -Int(Rnd * 10), 30 + Int(Rnd * 6) - Int(Rnd * 6), 1
NewPart 8 + Int(Rnd * 3) - Int(Rnd * 3), P(i).X, P(i).Y, P(i).X, P(i).Y, 0, L(P(i).X - 2).Color, Int(Rnd * 10) - Int(Rnd * 10), -10 + -Int(Rnd * 10), 30 + Int(Rnd * 6) - Int(Rnd * 6), 1
End If
If P(i).type = 1 Then
Boom P(i).X - 2, P(i).Y, 30, 35, P(i).Damage
End If
If P(i).type = 2 Then
Boom P(i).X - 2, P(i).Y, 40, 60, P(i).Damage
End If
If P(i).type = 3 Then
Boom P(i).X - 2, P(i).Y, 10, 20, P(i).Damage
If P(i).tag < 3 Then
P(i).active = True
P(i).YF = -P(i).YF * 1
P(i).XF = 0
P(i).tag = P(i).tag + 1
Else
P(i).tag = 0
End If
End If
If P(i).type = 4 Then
Boom P(i).X - 2, P(i).Y, 20, 20, P(i).Damage
End If
If P(i).type = 5 Then
Boom P(i).X - 2, P(i).Y, 10, 60, P(i).Damage, 1
End If
If P(i).type = 0 Then
Boom P(i).X - 2, P(i).Y, 30, 350, P(i).Damage
End If
If P(i).type = 6 Then
Boom P(i).X - 2, P(i).Y, 30, 30, P(i).Damage
NewMissle P(i).X, P(i).Y - 10, vbRed, 1, Int(Rnd * 20) - Int(Rnd * 20), -5 + Int(Rnd * 10) - Int(Rnd * 10)
NewMissle P(i).X, P(i).Y - 10, vbRed, 1, Int(Rnd * 20) - Int(Rnd * 20), -5 + Int(Rnd * 10) - Int(Rnd * 10)
NewMissle P(i).X, P(i).Y - 10, vbRed, 1, Int(Rnd * 20) - Int(Rnd * 20), -10 + Int(Rnd * 10) - Int(Rnd * 10)
End If
End If
Me.DrawWidth = 3
Me.Line (P(i).X, P(i).Y)-(P(i).OldX, P(i).OldY), P(i).Color
End If
Next


End Sub


Private Function GetAngle(Pos As Pos, CenterPos As Pos) As Integer
    Dim intA As Integer, intB%, intC%
    Dim PI As Double
    PI = Atn(1) * 4
    intB = Abs(CenterPos.X - Pos.X)
    intC = Abs(CenterPos.Y - Pos.Y)
    If intB <> 0 Then
        GetAngle = Atn(intC / intB) * 180 / PI
    End If
    If Pos.X < CenterPos.X Then
        If Pos.Y = CenterPos.Y Then GetAngle = 180
        If Pos.Y < CenterPos.Y Then
            GetAngle = 180 - GetAngle
        End If
        If Pos.Y > CenterPos.Y Then
            GetAngle = 180 + GetAngle
        End If
    End If
    If Pos.X > CenterPos.X Then
        If Pos.Y > CenterPos.Y Then
            GetAngle = 360 - GetAngle
        End If
    End If
    If Pos.X = CenterPos.X Then
        If Pos.Y < CenterPos.Y Then
            GetAngle = 90
        End If
        If Pos.Y > CenterPos.Y Then
            GetAngle = 270
        End If
    End If
    GetAngle = Abs(GetAngle Mod 360)
End Function

Private Sub Timer2_Timer()
'//Clouds, not fully working so removed
    'NewPart 20, -40, 100, -40, -40, 0, vbWhite, 6, 0, 200, 0
    'NewPart 23, -50, 90, -40, -40, 0, vbWhite, 6, 0, 200, 0
End Sub

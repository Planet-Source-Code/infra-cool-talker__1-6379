VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cool Talker"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command15 
      Caption         =   "Code"
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Mixed Text"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Binary"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Timer Timer9 
      Interval        =   100
      Left            =   3960
      Top             =   0
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Stop All Animation"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Timer Timer8 
      Interval        =   500
      Left            =   3480
      Top             =   0
   End
   Begin VB.Timer Timer7 
      Interval        =   1000
      Left            =   3000
      Top             =   0
   End
   Begin VB.Timer Timer6 
      Interval        =   500
      Left            =   2520
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   2040
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   1560
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   0
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "About"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Copy"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Twisted"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Scrambled"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Echo"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Doubled"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Spaced"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Screwed"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Backwards"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alt. Caps"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Maintxt 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Line LineMove 
      Index           =   12
      Visible         =   0   'False
      X1              =   4680
      X2              =   4680
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4680
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3240
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Line LineMove 
      Index           =   9
      Visible         =   0   'False
      X1              =   4320
      X2              =   4320
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line LineMove 
      Index           =   8
      Visible         =   0   'False
      X1              =   4200
      X2              =   4200
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line LineMove 
      Index           =   7
      Visible         =   0   'False
      X1              =   4080
      X2              =   4080
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line LineMove 
      Index           =   6
      Visible         =   0   'False
      X1              =   3960
      X2              =   3960
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line LineMove 
      Index           =   5
      Visible         =   0   'False
      X1              =   3840
      X2              =   3840
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line LineMove 
      Index           =   4
      Visible         =   0   'False
      X1              =   3720
      X2              =   3720
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line LineMove 
      Index           =   3
      Visible         =   0   'False
      X1              =   3600
      X2              =   3600
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line LineMove 
      Index           =   2
      Visible         =   0   'False
      X1              =   3480
      X2              =   3480
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line LineMove 
      Index           =   11
      Visible         =   0   'False
      X1              =   4560
      X2              =   4560
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line LineMove 
      Index           =   10
      Visible         =   0   'False
      X1              =   4440
      X2              =   4440
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line LineMove 
      Index           =   1
      Visible         =   0   'False
      X1              =   3360
      X2              =   3360
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line LineMove 
      Index           =   0
      X1              =   3240
      X2              =   3240
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Shape LineHide4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   15
      Left            =   4080
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape LineHide3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   15
      Left            =   2760
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape LineHide2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   15
      Left            =   1440
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape LineHide 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   15
      Left            =   240
      Top             =   1920
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4680
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurLinePos As Integer
Dim ClearFor As Boolean
Dim OnFirst As Boolean

Function AltCaps(Text As String)
'This will make the caps in text go on and off for each letter, like this:  cOoL
On Error GoTo error
Dim i As Integer
Dim s As String
s = ""
For i = 1 To Len(Text$)
  keyval = Asc(Mid$(Text$, i, 1))
  If (keyval >= 96 And keyval < 96 + 26) Or (keyval >= 64 And keyval < 64 + 26) Then
    If (i And 1) = 1 Then
      If keyval < 96 Then
        s = s + Chr$(96 + keyval - 64)
      Else
        s = s + Chr$(keyval)
      End If
    Else
      If keyval >= 96 Then
        s = s + Chr$(64 + keyval - 96)
      Else
        s = s + Chr$(keyval)
      End If
    End If
  Else
    s = s + Chr$(keyval)
  End If
Next i
Text$ = s
AltCaps = Text$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function BackwardsText(Text As String)
'This will make text go backwards, like this:  looC (Cool)
On Error GoTo error
For i% = 1 To Len(Text$)
stringy$ = Mid$(Text$, i%, 1)
final$ = stringy$ + final$
Next i%
BackwardsText = final$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function EliteType(Text As String)
'This will change characters to make them "elite", example:  ©00|_
On Error GoTo error
Dim s(52) As String
s(0) = "æ"
s(1) = "å"
s(2) = "b"
s(3) = "<"
s(4) = "c|"
s(5) = "ë"
s(6) = "f"
s(7) = "9"
s(8) = "h"
s(9) = "ï"
s(10) = "j"
s(11) = "|<"
s(12) = "|_"
s(13) = "/x\"
s(14) = "|\|"
s(15) = "0"
s(16) = "p"
s(17) = "q"
s(18) = "r"
s(19) = "_/¯"
s(20) = "-|-"
s(21) = "µ"
s(22) = "\/"
s(23) = "\/\/"
s(24) = "×"
s(25) = "ÿ"
s(26) = "¯/_"
s(27) = "Ä"
s(28) = "ß"
s(29) = "©"
s(30) = "|}"
s(31) = "È"
s(32) = "F"
s(33) = "G"
s(34) = "|-|"
s(35) = "I"
s(36) = "J"
s(37) = "]<"
s(38) = "]_"
s(39) = "/\/\"
s(40) = "|\|"
s(41) = "{}"
s(42) = "P"
s(43) = "¶"
s(44) = "|2"
s(45) = "§"
s(46) = "¯|¯"
s(47) = "|_|"
s(48) = "\/"
s(49) = "\x/"
s(50) = "><"
s(51) = "¥"
s(52) = "¯/_"
Text$ = ReplaceC(Text$, "ae", s(0))
Text$ = ReplaceC(Text$, "a", s(1))
Text$ = ReplaceC(Text$, "b", s(2))
Text$ = ReplaceC(Text$, "c", s(3))
Text$ = ReplaceC(Text$, "d", s(4))
Text$ = ReplaceC(Text$, "e", s(5))
Text$ = ReplaceC(Text$, "f", s(6))
Text$ = ReplaceC(Text$, "g", s(7))
Text$ = ReplaceC(Text$, "h", s(8))
Text$ = ReplaceC(Text$, "i", s(9))
Text$ = ReplaceC(Text$, "j", s(10))
Text$ = ReplaceC(Text$, "k", s(11))
Text$ = ReplaceC(Text$, "l", s(12))
Text$ = ReplaceC(Text$, "m", s(13))
Text$ = ReplaceC(Text$, "n", s(14))
Text$ = ReplaceC(Text$, "o", s(15))
Text$ = ReplaceC(Text$, "p", s(16))
Text$ = ReplaceC(Text$, "q", s(17))
Text$ = ReplaceC(Text$, "r", s(18))
Text$ = ReplaceC(Text$, "s", s(19))
Text$ = ReplaceC(Text$, "t", s(20))
Text$ = ReplaceC(Text$, "u", s(21))
Text$ = ReplaceC(Text$, "v", s(22))
Text$ = ReplaceC(Text$, "w", s(23))
Text$ = ReplaceC(Text$, "x", s(24))
Text$ = ReplaceC(Text$, "y", s(25))
Text$ = ReplaceC(Text$, "z", s(26))
Text$ = ReplaceC(Text$, "A", s(27))
Text$ = ReplaceC(Text$, "B", s(28))
Text$ = ReplaceC(Text$, "C", s(29))
Text$ = ReplaceC(Text$, "D", s(30))
Text$ = ReplaceC(Text$, "E", s(31))
Text$ = ReplaceC(Text$, "F", s(32))
Text$ = ReplaceC(Text$, "G", s(33))
Text$ = ReplaceC(Text$, "H", s(34))
Text$ = ReplaceC(Text$, "I", s(35))
Text$ = ReplaceC(Text$, "J", s(36))
Text$ = ReplaceC(Text$, "K", s(37))
Text$ = ReplaceC(Text$, "L", s(38))
Text$ = ReplaceC(Text$, "M", s(39))
Text$ = ReplaceC(Text$, "N", s(40))
Text$ = ReplaceC(Text$, "O", s(41))
Text$ = ReplaceC(Text$, "P", s(42))
Text$ = ReplaceC(Text$, "Q", s(43))
Text$ = ReplaceC(Text$, "R", s(44))
Text$ = ReplaceC(Text$, "S", s(45))
Text$ = ReplaceC(Text$, "T", s(46))
Text$ = ReplaceC(Text$, "U", s(47))
Text$ = ReplaceC(Text$, "V", s(48))
Text$ = ReplaceC(Text$, "W", s(49))
Text$ = ReplaceC(Text$, "X", s(50))
Text$ = ReplaceC(Text$, "Y", s(51))
Text$ = ReplaceC(Text$, "Z", s(52))
EliteType = Text$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function SpaceCharacters(Text As String, AmountOfSpaces As Integer)
'This will put a space between every character in the text, like this:  C o o l
On Error GoTo error
Dim i As Long
Dim SpaceStr As String
If AmountOfSpaces > 100 Then
AmountOfSpaces = 100
ElseIf AmountOfSpaces < 1 Then
AmountOfSpaces = 1
End If
For i = 1 To AmountOfSpaces
SpaceStr$ = SpaceStr$ + " "
Next i
Dim endstr As String
For i = 1 To Len(Text$)
endstr$ = endstr$ & Mid$(Text$, i, 1) & SpaceStr$
Next i
endstr$ = Mid$(endstr$, 1, Len(endstr$) - 1)
SpaceCharacters = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function DoubleCharacters(Text As String, AmountOfExtras As Integer)
'This will double every character in the text, like this:  CCooooll
On Error GoTo error
Dim i As Long
Dim i2 As Long
Dim endstr As String
If AmountOfExtras > 100 Then
AmountOfExtras = 100
ElseIf AmountOfExtras < 1 Then
AmountOfExtras = 1
End If
For i = 1 To Len(Text$)
  For i2 = 1 To AmountOfExtras
  endstr$ = endstr$ & Mid$(Text$, i, 1)
  Next i2
Next i
DoubleCharacters = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function EchoText(Text As String, Reverse As Boolean)
'This will "echo" the text, like this:  Cool ool ol l
On Error GoTo error
Dim i As Long
Dim endstr As String
For i = 1 To Len(Text$)
  If Reverse = True Then
  endstr$ = Mid$(Text$, i, Len(Text$) - (i - 1)) & " " & endstr$
  Else
  endstr$ = endstr$ & Mid$(Text$, i, Len(Text$) - (i - 1)) & " "
  End If
Next i
endstr$ = Mid$(endstr$, 1, Len(endstr$) - 1)
EchoText = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Scramble(Text As String, Key As Integer)
'This will scramble text up, example:  oCol
On Error GoTo error
Dim RndNum As Long
Dim i As Long
Dim endstr As String
Dim ListN(10000) As Long
Dim CurPos As Long
Randomize Key
CurPos = 0
Text$ = Mid$(Text$, 1, 10000)
Start:
RndNum = Int((Len(Text$) - 1 + 1) * Rnd + 1)
For i = 0 To CurPos
  If RndNum = ListN(i) Then
  GoTo Start
  End If
Next i
ListN(CurPos) = RndNum
CurPos = CurPos + 1
If Not CurPos = Len(Text$) Then
GoTo Start
End If
For i = 0 To CurPos - 1
endstr$ = endstr$ & Mid$(Text$, ListN(i), 1)
Next i
Scramble = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function TwistText(Text As String)
'This will "twist" text, it is kind of like scramble, example:  oClo
Dim CurPos As Long
Dim endstr As String
CurPos = 1
Start:
endstr$ = endstr$ & Mid$(Text$, CurPos + 1, 1) & Mid$(Text$, CurPos, 1)
CurPos = CurPos + 2
If Len(Text$) > CurPos Then
GoTo Start
ElseIf Len(Text$) = CurPos Then
endstr$ = endstr$ & Mid$(Text$, Len(Text$), 1)
End If
TwistText = endstr$
End Function

Private Sub Command1_Click()
Maintxt.Text = AltCaps(Maintxt.Text)
End Sub

Private Sub Command10_Click()
MsgBox "This program created by InfraRed, because I was bored.", 48, "About"
End Sub

Private Sub Command11_Click()
Unload Me
End
End Sub

Private Sub Command12_Click()
If Command12.Caption = "Stop All Animation" Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Command1.Caption = "Alt. Caps"
Command2.Caption = "Backwards"
Command3.Caption = "Screwed"
Command4.Caption = "Spaced"
Command5.Caption = "Doubled"
Command6.Caption = "Echo"
Command7.Caption = "Scrambled"
Command8.Caption = "Twisted"
Command9.Caption = "Copy"
Command10.FontBold = False
Command11.Caption = "Exit"
LineMove(CurLinePos).Visible = False
Label1.ForeColor = &H0&
Line2.BorderColor = &H0&
LineHide.Visible = False
LineHide2.Visible = False
LineHide3.Visible = False
LineHide4.Visible = False
Command12.Caption = "Start All Animation"
Else
  If OnFirst = True Then
  Timer1.Enabled = True
  Else
  Timer2.Enabled = True
  End If
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = True
Timer9.Enabled = True
LineHide.Visible = True
LineHide2.Visible = True
LineHide3.Visible = True
LineHide4.Visible = True
Command12.Caption = "Stop All Animation"
End If
End Sub

Private Sub Command13_Click()
Dim endstr As String, xtra As String, i As Integer, i2 As Integer
endstr$ = ""
For i = 1 To Len(Maintxt.Text)
xtra$ = Bin(Asc(Mid$(Maintxt.Text, i, 1)))
  For i2 = 1 To 8 - Len(xtra$)
  xtra$ = "0" & xtra$
  Next i2
Do While Len(xtra$) > 8
xtra$ = Mid$(xtra$, 2, Len(xtra$) - 1)
Loop
endstr$ = endstr$ & xtra$
Next i
Maintxt.Text = endstr$
End Sub

Private Sub Command14_Click()
Maintxt.Text = Mix(Maintxt.Text)
End Sub

Private Sub Command15_Click()
Maintxt.Text = Code(Maintxt.Text)
End Sub

Private Sub Command2_Click()
Maintxt.Text = BackwardsText(Maintxt.Text)
End Sub

Private Sub Command3_Click()
Maintxt.Text = EliteType(Maintxt.Text)
End Sub

Private Sub Command4_Click()
Maintxt.Text = SpaceCharacters(Maintxt.Text, 1)
End Sub

Private Sub Command5_Click()
Maintxt.Text = DoubleCharacters(Maintxt.Text, 2)
End Sub

Private Sub Command6_Click()
Maintxt.Text = EchoText(Maintxt.Text, False)
End Sub

Private Sub Command7_Click()
Maintxt.Text = Scramble(Maintxt.Text, 20)
End Sub

Private Sub Command8_Click()
Maintxt.Text = TwistText(Maintxt.Text)
End Sub

Private Sub Command9_Click()
Clipboard.Clear
Clipboard.SetText Maintxt.Text
End Sub

Private Sub Form_Load()
CurLinePos = 0
ClearFor = True
Label1.ForeColor = &H0&
OnFirst = True
End Sub

Private Sub Image1_Click()
Maintxt.Text = ""
End Sub

Private Sub Timer1_Timer()
If CurLinePos = 12 Then
OnFirst = False
Timer2.Enabled = True
Timer1.Enabled = False
Else
LineMove(CurLinePos).Visible = False
CurLinePos = CurLinePos + 1
LineMove(CurLinePos).Visible = True
End If
End Sub

Private Sub Timer2_Timer()
If CurLinePos = 0 Then
OnFirst = True
Timer1.Enabled = True
Timer2.Enabled = False
Else
LineMove(CurLinePos).Visible = False
CurLinePos = CurLinePos - 1
LineMove(CurLinePos).Visible = True
End If
End Sub

Private Sub Timer3_Timer()
If LineHide.Left > Form1.Width Then
LineHide.Left = 0 - LineHide.Width
Else
LineHide.Left = LineHide.Left + 20
End If
If LineHide2.Left > Form1.Width Then
LineHide2.Left = 0 - LineHide2.Width
Else
LineHide2.Left = LineHide2.Left + 20
End If
If LineHide3.Left > Form1.Width Then
LineHide3.Left = 0 - LineHide3.Width
Else
LineHide3.Left = LineHide3.Left + 20
End If
If LineHide4.Left > Form1.Width Then
LineHide4.Left = 0 - LineHide4.Width
Else
LineHide4.Left = LineHide4.Left + 20
End If
End Sub

Private Function ReplaceC(MainStr As String, OldStr As String, NewStr As String) As String
'For Section 12 (Code/Decode):  Replaces one string with another
On Error GoTo error
ReplaceC = ""
Dim NewStrString As String
Dim i As Integer
For i = 1 To Len(MainStr)
  If Mid(MainStr, i, Len(OldStr)) = OldStr Then
  NewStrString = NewStrString & NewStr
  i = i + Len(OldStr) - 1
  Else
  NewStrString = NewStrString & Mid(MainStr, i, 1)
  End If
Next i
ReplaceC = NewStrString
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Private Sub Timer4_Timer()
If ClearFor = True Then
If Label1.ForeColor = &H0& Then
Label1.ForeColor = &H222222
ElseIf Label1.ForeColor = &H222222 Then
Label1.ForeColor = &H444444
ElseIf Label1.ForeColor = &H444444 Then
Label1.ForeColor = &H666666
ElseIf Label1.ForeColor = &H666666 Then
Label1.ForeColor = &H888888
ElseIf Label1.ForeColor = &H888888 Then
Label1.ForeColor = &HAAAAAA
ElseIf Label1.ForeColor = &HAAAAAA Then
Label1.ForeColor = &HC0C0C0
ElseIf Label1.ForeColor = &HC0C0C0 Then
ClearFor = False
End If
Else
If Label1.ForeColor = &HC0C0C0 Then
Label1.ForeColor = &HAAAAAA
ElseIf Label1.ForeColor = &HAAAAAA Then
Label1.ForeColor = &H888888
ElseIf Label1.ForeColor = &H888888 Then
Label1.ForeColor = &H666666
ElseIf Label1.ForeColor = &H666666 Then
Label1.ForeColor = &H444444
ElseIf Label1.ForeColor = &H444444 Then
Label1.ForeColor = &H222222
ElseIf Label1.ForeColor = &H222222 Then
Label1.ForeColor = &H0&
ElseIf Label1.ForeColor = &H0& Then
ClearFor = True
End If
End If
End Sub

Private Sub Timer5_Timer()
Randomize Timer
If Command1.Caption = "Alt. Caps" Then
Command1.Caption = AltCaps("Alt. Caps")
Else
Command1.Caption = "Alt. Caps"
End If
If Command2.Caption = "Backwards" Then
Command2.Caption = BackwardsText("Backwards")
Else
Command2.Caption = "Backwards"
End If
If Command3.Caption = "Screwed" Then
Command3.Caption = EliteType("Screwed")
Else
Command3.Caption = "Screwed"
End If
If Command4.Caption = "Spaced" Then
Command4.Caption = SpaceCharacters("Spaced", 1)
Else
Command4.Caption = "Spaced"
End If
If Command5.Caption = "Doubled" Then
Command5.Caption = DoubleCharacters("Doubled", 2)
Else
Command5.Caption = "Doubled"
End If
If Command6.Caption = "Echo" Then
Command6.Caption = EchoText("Echo", False)
Else
Command6.Caption = "Echo"
End If
If Command7.Caption = "Scrambled" Then
Command7.Caption = Scramble("Scrambled", 20)
Else
Command7.Caption = "Scrambled"
End If
If Command13.Caption = "Binary" Then
Command13.Caption = Mid$(Bin(Int((255 - 1 + 1) * Rnd + 1)), 9, 8)
Else
Command13.Caption = "Binary"
End If
If Command14.Caption = "Mixed Text" Then
Command14.Caption = Mix("Mixed Text")
Else
Command14.Caption = "Mixed Text"
End If
If Command15.Caption = "Code" Then
Command15.Caption = Code("Code")
Else
Command15.Caption = "Code"
End If
If Command8.Caption = "Twisted" Then
Command8.Caption = TwistText("Twisted")
Timer5.Interval = 500
Else
Command8.Caption = "Twisted"
Timer5.Interval = 1000
End If
End Sub

Private Sub Timer6_Timer()
If Command9.Caption = "Copy" Then
Command9.Caption = "C   "
Timer6.Interval = 500
ElseIf Command9.Caption = "C   " Then
Command9.Caption = "Co  "
ElseIf Command9.Caption = "Co  " Then
Command9.Caption = "Cop "
ElseIf Command9.Caption = "Cop " Then
Command9.Caption = "Copy"
Timer6.Interval = 1000
End If
End Sub

Private Sub Timer7_Timer()
If Command11.Caption = "Exit" Then
Command11.Caption = ""
Timer7.Interval = 100
Else
Command11.Caption = "Exit"
Timer7.Interval = 1000
End If
End Sub

Private Sub Timer8_Timer()
If Command10.FontBold = False Then
Command10.FontBold = True
Else
Command10.FontBold = False
End If
End Sub

Private Sub Timer9_Timer()
If Line2.BorderColor = &HFF& Then
Line2.BorderColor = &H33FF&
ElseIf Line2.BorderColor = &H33FF& Then
Line2.BorderColor = &H66FF&
ElseIf Line2.BorderColor = &H66FF& Then
Line2.BorderColor = &H99FF&
ElseIf Line2.BorderColor = &H99FF& Then
Line2.BorderColor = &HCCFF&
ElseIf Line2.BorderColor = &HCCFF& Then
Line2.BorderColor = &HFFFF&
ElseIf Line2.BorderColor = &HFFFF& Then
Line2.BorderColor = &HFFCC&
ElseIf Line2.BorderColor = &HFFCC& Then
Line2.BorderColor = &HFF99&
ElseIf Line2.BorderColor = &HFF99& Then
Line2.BorderColor = &HFF66&
ElseIf Line2.BorderColor = &HFF66& Then
Line2.BorderColor = &HFF33&
ElseIf Line2.BorderColor = &HFF33& Then
Line2.BorderColor = &HFF00&
ElseIf Line2.BorderColor = &HFF00& Then
Line2.BorderColor = &H33CC00
ElseIf Line2.BorderColor = &H33CC00 Then
Line2.BorderColor = &H669900
ElseIf Line2.BorderColor = &H669900 Then
Line2.BorderColor = &H996600
ElseIf Line2.BorderColor = &H996600 Then
Line2.BorderColor = &HCC3300
ElseIf Line2.BorderColor = &HCC3300 Then
Line2.BorderColor = &HFF0000
ElseIf Line2.BorderColor = &HFF0000 Then
Line2.BorderColor = &HFF0033
ElseIf Line2.BorderColor = &HFF0033 Then
Line2.BorderColor = &HFF0066
ElseIf Line2.BorderColor = &HFF0066 Then
Line2.BorderColor = &HFF0099
ElseIf Line2.BorderColor = &HFF0099 Then
Line2.BorderColor = &HFF00CC
ElseIf Line2.BorderColor = &HFF00CC Then
Line2.BorderColor = &HFF00FF
ElseIf Line2.BorderColor = &HFF00FF Then
Line2.BorderColor = &HCC00CC
ElseIf Line2.BorderColor = &HCC00CC Then
Line2.BorderColor = &H990099
ElseIf Line2.BorderColor = &H990099 Then
Line2.BorderColor = &H660066
ElseIf Line2.BorderColor = &H660066 Then
Line2.BorderColor = &H330033
ElseIf Line2.BorderColor = &H330033 Then
Line2.BorderColor = &H0&
ElseIf Line2.BorderColor = &H0& Then
Line2.BorderColor = &H333333
ElseIf Line2.BorderColor = &H333333 Then
Line2.BorderColor = &H666666
ElseIf Line2.BorderColor = &H666666 Then
Line2.BorderColor = &H999999
ElseIf Line2.BorderColor = &H999999 Then
Line2.BorderColor = &HCCCCCC
ElseIf Line2.BorderColor = &HCCCCCC Then
Line2.BorderColor = &HFFFFFF
ElseIf Line2.BorderColor = &HFFFFFF Then
Line2.BorderColor = &HCCCCFF
ElseIf Line2.BorderColor = &HCCCCFF Then
Line2.BorderColor = &H9999FF
ElseIf Line2.BorderColor = &H9999FF Then
Line2.BorderColor = &H6666FF
ElseIf Line2.BorderColor = &H6666FF Then
Line2.BorderColor = &H3333FF
ElseIf Line2.BorderColor = &H3333FF Then
Line2.BorderColor = &HFF&
Else
Line2.BorderColor = &HFF&
End If
End Sub

Function Bin(ByVal intDec As Integer) As String


    Dim strDummy As String
    Dim intLess As Integer
    Dim lngVal As Long
    Dim strTrue As String
    Dim strFalse As String
    lngVal = intDec
    intLess = 16384


    If Abs(lngVal) = lngVal Then
        strTrue = "1"
        strFalse = "0"
        strDummy = "0"
    Else
        strTrue = "0"
        strFalse = "1"
        strDummy = "1"
        lngVal = Abs(lngVal) - 1
    End If



    Do


        If lngVal - intLess >= 0 Then
            strDummy = strDummy & strTrue
            lngVal = lngVal - intLess
        Else
            strDummy = strDummy & strFalse
        End If

        intLess = intLess / 2
    Loop While intLess >= 1

    Bin = strDummy
End Function

Function Mix(Text As String) As String
Dim TT(64000) As String, i As Integer, ParseChr As String, LTP As Integer, Pos As Integer, endstr As String, DC As Boolean
ParseChr = " "
Pos = 0
LTP = 0
DS:
DC = False
For i = 1 To Len(Text)
  If Mid$(Text, i, 2) = "  " Then
  DC = True
  Text = Mid$(Text, 1, i) & Mid$(Text, i + 2, Len(Text) - (i + 1))
  End If
Next i
    If DC = True Then GoTo DS
DS2:
        If Mid$(Text, Len(Text), 1) = " " Then
        Text = Mid$(Text, 1, Len(Text) - 1)
        GoTo DS2
        End If
For i = 1 To Len(Text)
  If Mid$(Text, i, Len(ParseChr)) = " " Then
  TT(Pos) = Mid$(Text, LTP + Len(ParseChr), i - (LTP + Len(ParseChr)))
  LTP = i
  Pos = Pos + 1
  End If
Next i
TT(Pos) = Mid$(Text, LTP + Len(ParseChr), (Len(Text) + 1) - (LTP + Len(ParseChr)))
Pos = Pos + 1
  If Pos = 0 Then
  Pos = 1
  TT(0) = Text
  End If
Randomize Timer
endstr = ""
For i = 0 To Pos - 1
endstr = endstr & Scramble(TT(i), Int((10000 - 1 + 1) * Rnd + 1))
  If i < Pos - 1 Then endstr = endstr & " "
Next i
Mix = endstr
End Function

Function Code(Text As String)
Randomize Timer
Dim i As Integer, endstr As String, RndNum As Integer
endstr = ""
For i = 1 To Len(Text)
RndNum = Int((255 - 1 + 1) * Rnd + 1)
Check1:
  If RndNum < Asc("a") Or RndNum > Asc("z") Then
    If RndNum < Asc("A") Or RndNum > Asc("Z") Then
    RndNum = Int((255 - 1 + 1) * Rnd + 1)
    GoTo Check1
    End If
  End If
endstr = endstr & Chr(RndNum) & Mid$(Text, i, 1)
RndNum = Int((255 - 1 + 1) * Rnd + 1)
Check2:
  If RndNum < Asc("a") Or RndNum > Asc("z") Then
    If RndNum < Asc("A") Or RndNum > Asc("Z") Then
    RndNum = Int((255 - 1 + 1) * Rnd + 1)
    GoTo Check2
    End If
  End If
endstr = endstr & Chr(RndNum)
  If i > 1 Then endstr = " " & endstr
Next i
Check3:
  If Mid$(endstr, 1, 1) = " " Then
  endstr = Mid$(endstr, 2, Len(endstr) - 1)
  GoTo Check3
  End If
Code = endstr
End Function

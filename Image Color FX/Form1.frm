VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Cool COLOR Effects! Simple, Easy, Nice!!!"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   6360
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Render"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4800
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   240
      ScaleHeight     =   900
      ScaleWidth      =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4560
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4500
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   6060
      Begin VB.Line Line1 
         DrawMode        =   6  'Mask Pen Not
         X1              =   -120
         X2              =   -120
         Y1              =   -480
         Y2              =   4560
      End
   End
   Begin VB.Label DC 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form1.frx":D605
      Height          =   1815
      Left            =   6360
      TabIndex        =   4
      Top             =   2880
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, Temp As Long, Effect, R, G, B
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Function Red() As Integer
Red = Temp Mod &H100
End Function
Private Function Green() As Integer
Green = (Temp \ &H100) Mod &H100
End Function
Private Function Blue() As Integer
Blue = (Temp \ &H10000) Mod &H100
End Function

Private Sub Command1_Click()
Me.Enabled = False
For X = 0 To Picture2.Width + 30 Step 15
For Y = 0 To Picture2.Height Step 15
On Error Resume Next
Temp = Picture1.Point(X, Y)
Call UpdateColor
Picture2.Line (X, Y)-(X + 30, Y + 500), RGB(R, G, B)
Next
Picture2.Refresh
Next
Me.Enabled = True
End Sub

Private Sub Form_Load()
Picture2.Height = Picture1.Height
Picture2.Width = Picture1.Width
Effect = 1
For i = 1 To 22
List1.AddItem "Color Shift #" & i
Next
End Sub
Sub UpdateColor()
Select Case Effect
Case 1
R = Blue
G = Red
B = Green
Case 2
R = Green
G = Blue
B = Red
Case 3
R = Blue
G = Green
B = Red
Case 4
R = Red + Green / 2
G = Green + Blue / 2
B = Blue + Red / 2
Case 5
R = Green + Blue / 2
G = Blue + Red / 2
B = Red + Green / 2
Case 6
R = Blue + Red / 2
G = Green + Blue / 2
B = Red + Green / 2
Case 7
R = Blue + Red / 2
G = Red + Green / 2
B = Green + Blue / 2
Case 8
R = Blue + Green / 2
G = Red + Blue / 2
B = Green + Red / 2
Case 9
R = (Red * 2)
G = (Green * 2)
B = (Blue * 2)
Case 10
R = Red / 1.83798
G = Green / 1.564362
B = Blue / 1.2344
Case 11
R = Red / 2 + Green / 2
G = Green / 2 + Blue / 2
B = Blue / 2 + Red / 2
Case 12
R = Green / 2 + Blue / 2
G = Blue / 2 + Red / 2
B = Red / 2 + Green / 2
Case 13
R = Blue / 2 + Red / 2
G = Red / 2 + Green / 2
B = Green / 2 + Blue / 2
Case 14
R = Red / 2 + Blue / 4
G = Green / 2 + Green / 2
B = Blue / 2 + Red / 4
Case 15
R = 255 - (Blue + Green) / 2
G = 255 - (Red + Blue) / 2
B = 255 - (Green + Red) / 2
Case 16
R = (Red + Green + Blue) / 3
G = (Red + Green + Blue) / 3
B = (Red + Green + Blue) / 3
Case 17
R = (Red + Green + Blue) / 3
G = 0
B = 0
Case 18
R = 0
G = 0
B = ((Red + Green + Blue) / 3) * 1.65164546
Case 19
R = 0
G = ((Red + Green + Blue) / 3) * 1.65164546
B = 0
Case 20
R = 0
G = ((Red + Green + Blue) / 3) * 1.65164546
B = ((Red + Green + Blue) / 3) * 1.65164546
Case 21
R = ((Red + Green + Blue) / 3) * 1.65164546
G = 0
B = ((Red + Green + Blue) / 3) * 1.65164546
Case 22
R = Red
G = Green
B = Blue
End Select
End Sub

Private Sub List1_Click()
Effect = List1.ListIndex + 1
Command1.Enabled = True
End Sub

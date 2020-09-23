VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   LinkTopic       =   "Form2"
   ScaleHeight     =   5310
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   3600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   720
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H0080FFFF&
      Height          =   1335
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GlobalImageIndex
Dim GlobalImageX

Dim PreviousX

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Me.Width = Screen.Width * Screen.TwipsPerPixelX
Me.Height = Screen.Height * Screen.TwipsPerPixelY
Me.top = 0
Me.left = 0
Image1(0).Width = Me.Width / 12
Image1(0).Height = Me.Height / 12
Shape1(0).Width = Me.Width / 12 + 40
Shape1(0).Height = Me.Height / 12 + 40

'Image1(0).Height = Image1(0).Width

'fill first line
'For i = 1 To 9
'Load image1(i)
'image1(i).left = image1(i - 1).left + image1(i - 1).Width
'image1(i).Visible = True
'Next i
plus = 0
For j = 0 To 9
toppos = (Image1(0).Height * j) + 50
leftpos = 50
For i = 1 To 10
Load Image1(i + plus)
Image1(i + plus).top = toppos + a
Image1(i + plus).left = leftpos
Image1(i + plus).Visible = True
'leftpos = leftpos + Image1(0).Width + 50

Load Shape1(i + plus)
Shape1(i + plus).top = toppos + a - 20
Shape1(i + plus).left = leftpos - 20
Shape1(i + plus).Visible = True

leftpos = leftpos + Image1(0).Width + 50

Next i


a = a + 50
plus = plus + 10
Next j

MCCaptureMouseCursorIntoArea Me, Image1(1)

End Sub

Private Sub Image1_Click(Index As Integer)
ClipCursor ByVal 0&
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If X = PreviousX Then
a = a + 1
PreviousX = X
End If

If a = 5 Then
Beep
a = 0
End If

End Sub


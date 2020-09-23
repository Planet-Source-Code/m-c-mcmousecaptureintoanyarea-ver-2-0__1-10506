VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "McMouseCaptureHandler Demo Form"
   ClientHeight    =   6120
   ClientLeft      =   105
   ClientTop       =   675
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      Height          =   1020
      IntegralHeight  =   0   'False
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "This one tests nested control problem solution!"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   4440
      TabIndex        =   14
      Top             =   4080
      Width           =   855
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   6600
      TabIndex        =   13
      Top             =   3720
      Width           =   975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   6360
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   855
      Left            =   1680
      TabIndex        =   11
      Top             =   2760
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   735
      Left            =   5400
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   3600
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5520
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   4440
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   735
      Left            =   3000
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   735
      Left            =   5760
      TabIndex        =   6
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4560
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   3360
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   3900
      IntegralHeight  =   0   'False
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   2880
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   840
      Width           =   1575
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FFFF&
         Height          =   855
         Left            =   240
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   1455
      Left            =   5040
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select control or form area in list box - mouse will be captured in that control, to release mouse, click in that control !!!!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   15
      Left            =   4560
      Top             =   960
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1620
      Left            =   1560
      Picture         =   "Form1.frx":16A4B
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub Check1_Click()
ClipCursor ByVal 0&
End Sub
Private Sub Combo1_GotFocus()
ClipCursor ByVal 0&
End Sub

Private Sub Command1_Click()
ClipCursor ByVal 0&
End Sub



Private Sub Command2_Click()
MCCaptureMouseCursorIntoNestedArea Me, Picture1, Shape3
End Sub

Private Sub Command3_Click()
End
End Sub


Private Sub Dir1_Click()
ClipCursor ByVal 0&
End Sub


Private Sub Drive1_GotFocus()
ClipCursor ByVal 0&
End Sub
Private Sub File1_GotFocus()
ClipCursor ByVal 0&
End Sub

Private Sub Form_Activate()
List1.SetFocus
End Sub

Private Sub Form_Click()
ClipCursor ByVal 0& 'this is there to release capture from shapes
End Sub

Private Sub Form_Load()
Dim ControlOnForm
For Each ControlOnForm In Me
        If ControlOnForm.Name = "List1" Or ControlOnForm.Name = "command2" Or ControlOnForm.Name = "Shape3" Then GoTo skip
        If TypeOf ControlOnForm Is Menu Then GoTo skip
        List1.AddItem ControlOnForm.Name
skip:
Next

List3.AddItem "CompleteFormBorderExcluded"
List3.AddItem "CompleteForm"
List3.AddItem "FormClientArea"
List3.AddItem "CaptionBar"
List3.AddItem "MenuBar"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ClipCursor ByVal 0&
End Sub

Private Sub Form_Resize()
ClipCursor ByVal 0&
End Sub

Private Sub Frame1_Click()
ClipCursor ByVal 0&
End Sub

Private Sub HScroll1_GotFocus()
ClipCursor ByVal 0&
End Sub

Private Sub Image1_Click()
ClipCursor ByVal 0&
End Sub

Private Sub Image1_DblClick()
MCCaptureMouseCursorIntoArea Me, Image1
End Sub

Private Sub Label1_Click()
ClipCursor ByVal 0&
End Sub

Private Sub List1_Click()
Dim SelectedControl As Control
Dim SelectedControl1 As Control
For Each SelectedControl In Me
If SelectedControl.Name = List1.List(List1.ListIndex) Then
Set SelectedControl1 = SelectedControl
End If
Next
MCCaptureMouseCursorIntoArea Me, SelectedControl1
End Sub

Private Sub List2_GotFocus()
ClipCursor ByVal 0&
End Sub

Private Sub List3_Click()
'Dim SelectedArea As String
'Dim SelectedArea1 As String
''For Each SelectedArea In Me
'If SelectedArea.Name = List1.List(List1.ListIndex) Then
'Set SelectedArea1 = SelectedArea
'End If
'Next
'MCCaptureMouseCursorIntoArea Me, SelectedArea1
C = List1.List(List1.ListIndex)
MCCaptureMouseCursorIntoSpecialArea Me, List3.List(List3.ListIndex)
End Sub

Private Sub Option1_Click()
ClipCursor ByVal 0&
End Sub

Private Sub Picture1_Click()
ClipCursor ByVal 0&
End Sub

Private Sub Text1_Click()
ClipCursor ByVal 0&
End Sub

Private Sub VScroll1_GotFocus()
ClipCursor ByVal 0&
End Sub

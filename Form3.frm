VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   120
      Top             =   1800
   End
   Begin VB.Menu M 
      Caption         =   "M"
      Visible         =   0   'False
      Begin VB.Menu C 
         Caption         =   "C"
      End
      Begin VB.Menu D 
         Caption         =   "D"
         Begin VB.Menu E 
            Caption         =   "E"
            Begin VB.Menu F 
               Caption         =   "F"
            End
         End
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long







Private Sub Form_Click()
'Timer1.Enabled = True

'show menu
PopupMenu M

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Beep
End Sub

Private Sub Timer1_Timer()
''E = FindWindow("#32768", vbNullString)
Print E
PopUpMenuStuff (E)
Timer1.Enabled = False
End Sub

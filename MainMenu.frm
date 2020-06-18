VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   Picture         =   "MainMenu.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image3 
      Height          =   705
      Left            =   8760
      Picture         =   "MainMenu.frx":4BCD2
      Top             =   120
      Width           =   3075
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   4440
      Picture         =   "MainMenu.frx":4C0F3
      Top             =   1200
      Width           =   7230
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   120
      Picture         =   "MainMenu.frx":4EFA8
      Top             =   120
      Width           =   6615
   End
   Begin VB.Image imgstory 
      Height          =   450
      Left            =   8640
      MouseIcon       =   "MainMenu.frx":51AB5
      MousePointer    =   99  'Custom
      Picture         =   "MainMenu.frx":5221F
      Top             =   5280
      Width           =   2250
   End
   Begin VB.Image imgstart 
      Height          =   450
      Left            =   8040
      MouseIcon       =   "MainMenu.frx":523B3
      MousePointer    =   99  'Custom
      Picture         =   "MainMenu.frx":52B1D
      Top             =   4320
      Width           =   3150
   End
   Begin VB.Image imgexit 
      Height          =   360
      Left            =   9240
      MouseIcon       =   "MainMenu.frx":52D0F
      MousePointer    =   99  'Custom
      Picture         =   "MainMenu.frx":53479
      Top             =   6240
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Private Sub cmdstart_Click()
frm1.Visible = True
Form1.Visible = False
End Sub

Private Sub cmdstory_Click()
Form2.Visible = True
Form1.Hide
frm1.Hide
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub cmdconfig_Click()

End Sub

Private Sub Form_Load()
Dim wid As Long
Dim hgt As Long

    ' Get the screen's size including the task bar area.
    wid = GetSystemMetrics(SM_CXSCREEN)
    hgt = GetSystemMetrics(SM_CYSCREEN)

    ' Make the form this size.
    Move 0, 0, _
        ScaleX(wid, vbPixels, vbTwips), _
        ScaleY(hgt, vbPixels, vbTwips)
Call sndPlaySound(ByVal "Sound\A_RTHUND1.wav", SND_ASYNC)
End Sub
Private Sub imgexit_Click()
End
End Sub

Private Sub imgstart_Click()
frm1.Visible = True
Form1.Visible = False
End Sub

Private Sub imgstory_Click()
Form2.Visible = True
Form1.Hide
End Sub

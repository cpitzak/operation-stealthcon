VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11895
   LinkTopic       =   "Form3"
   ScaleHeight     =   8880
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   3180
      Left            =   5280
      Picture         =   "form3.frx":0000
      Top             =   3960
      Width           =   1185
   End
   Begin VB.Image imgre 
      Height          =   3240
      Left            =   1440
      MouseIcon       =   "form3.frx":4482
      MousePointer    =   99  'Custom
      Picture         =   "form3.frx":4BEC
      Top             =   4800
      Width           =   3675
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   6480
      MouseIcon       =   "form3.frx":5083
      MousePointer    =   99  'Custom
      Picture         =   "form3.frx":57ED
      Top             =   3000
      Width           =   3675
   End
   Begin VB.Image imgmain 
      Height          =   645
      Left            =   1680
      MouseIcon       =   "form3.frx":5A74
      MousePointer    =   99  'Custom
      Picture         =   "form3.frx":61DE
      Top             =   3000
      Width           =   3675
   End
   Begin VB.Image imgquit 
      Height          =   645
      Left            =   7680
      MouseIcon       =   "form3.frx":640F
      MousePointer    =   99  'Custom
      Picture         =   "form3.frx":6B79
      Top             =   6360
      Width           =   3675
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SCREENSAVERRUNNING = 97

' Clip mouse Declarations
Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function ClipCursorByNum Lib "user32" Alias "ClipCursor" (ByVal num As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Sub cmdquit_Click()
' Free the mouse
    ClipCursorByNum 0&
End
End Sub
Private Sub Form_Activate()
Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
Dim window As RECT
' Restrict the mouse to this window.
    GetWindowRect hwnd, window
    ClipCursor window
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
End Sub

Private Sub Image1_Click()
frm1.Hide
Form3.Hide
Form2.Visible = True
' Free the mouse
    ClipCursorByNum 0&
End Sub

Private Sub imgmain_Click()
Form1.Visible = True
Form3.Hide
frm1.Hide
' Free the mouse
    ClipCursorByNum 0&
End Sub

Private Sub imgquit_Click()
' Free the mouse
    ClipCursorByNum 0&
End
End Sub

Private Sub imgre_Click()
' Free the mouse
    ClipCursorByNum 0&
Form3.Hide
frm1.Visible = True
End Sub

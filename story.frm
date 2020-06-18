VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11835
   LinkTopic       =   "Form2"
   Picture         =   "story.frx":0000
   ScaleHeight     =   8880
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OLE OLE1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "Package"
      Height          =   855
      Left            =   9840
      OleObjectBlob   =   "story.frx":18EB5
      SourceDoc       =   "A:\Text\README.txt"
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label lblback 
      BackStyle       =   0  'Transparent
      Caption         =   "< Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      MouseIcon       =   "story.frx":1A8CD
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   8040
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Sub Label1_Click()

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

Private Sub lblback_Click()
Form1.Visible = True
Form2.Hide
End Sub

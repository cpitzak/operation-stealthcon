VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Assassins"
   ClientHeight    =   8940
   ClientLeft      =   105
   ClientTop       =   -3780
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pbhealth 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar pbcharger 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer rechargertim 
      Interval        =   200
      Left            =   4680
      Top             =   8040
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4080
      Top             =   8040
   End
   Begin VB.Image imgGen 
      Height          =   765
      Left            =   11040
      Picture         =   "Play.frx":0000
      Top             =   120
      Width           =   765
   End
   Begin VB.Image img2gen 
      Height          =   765
      Left            =   11040
      Picture         =   "Play.frx":0DEB
      Top             =   120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Shape Ojbar 
      BorderColor     =   &H000080FF&
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   8280
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Shape Ojbar 
      BorderColor     =   &H000080FF&
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   5775
      Index           =   2
      Left            =   10320
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Ojbar 
      BorderColor     =   &H000080FF&
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   2175
      Index           =   1
      Left            =   8040
      Top             =   5760
      Width           =   255
   End
   Begin VB.Shape Ojbar 
      BorderColor     =   &H000080FF&
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   2175
      Index           =   0
      Left            =   5640
      Top             =   5760
      Width           =   255
   End
   Begin VB.Shape Box 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   5400
      Top             =   5520
      Width           =   735
   End
   Begin VB.Shape Box 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   7920
      Top             =   5520
      Width           =   375
   End
   Begin VB.Shape Box 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Shape Box 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   0
      Top             =   3960
      Width           =   975
   End
   Begin VB.Shape Box 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   8040
      Top             =   5280
      Width           =   255
   End
   Begin VB.Shape Box 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   2775
      Index           =   5
      Left            =   5640
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Box 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   960
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Box 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   5775
      Index           =   3
      Left            =   3240
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Box 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   0
      Top             =   6120
      Width           =   975
   End
   Begin VB.Shape Box 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   960
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape Box 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   8040
      Top             =   2760
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   11880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   7920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00404040&
      X1              =   11880
      X2              =   11880
      Y1              =   0
      Y2              =   7920
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   11880
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      Height          =   6015
      Index           =   7
      Left            =   960
      Top             =   1920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      Height          =   2175
      Index           =   6
      Left            =   3240
      Top             =   5760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      Height          =   255
      Index           =   5
      Left            =   3480
      Top             =   2760
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      Height          =   2535
      Index           =   4
      Left            =   8040
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      Height          =   255
      Index           =   3
      Left            =   10560
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      Height          =   255
      Index           =   2
      Left            =   1200
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      Height          =   255
      Index           =   1
      Left            =   3480
      Top             =   840
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      Height          =   255
      Index           =   0
      Left            =   3480
      Top             =   5520
      Width           =   7095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      Height          =   135
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgattack 
      Height          =   735
      Index           =   10
      Left            =   11040
      Picture         =   "Play.frx":19BC
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgattack 
      Height          =   735
      Index           =   9
      Left            =   6120
      Picture         =   "Play.frx":212F
      Top             =   6480
      Width           =   615
   End
   Begin VB.Image imgattack 
      Height          =   735
      Index           =   8
      Left            =   9360
      Picture         =   "Play.frx":28A2
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgattack 
      Height          =   735
      Index           =   7
      Left            =   9600
      Picture         =   "Play.frx":3015
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgattack 
      Height          =   735
      Index           =   6
      Left            =   7680
      Picture         =   "Play.frx":3788
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgattack 
      Height          =   735
      Index           =   5
      Left            =   10680
      Picture         =   "Play.frx":3EFB
      Top             =   6840
      Width           =   615
   End
   Begin VB.Image imgattack 
      Height          =   735
      Index           =   4
      Left            =   120
      Picture         =   "Play.frx":466E
      Top             =   4800
      Width           =   615
   End
   Begin VB.Image imgattack 
      Height          =   735
      Index           =   3
      Left            =   6960
      Picture         =   "Play.frx":4DE1
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image imgattack 
      Height          =   735
      Index           =   2
      Left            =   0
      Picture         =   "Play.frx":5554
      Top             =   7080
      Width           =   615
   End
   Begin VB.Image imgattack 
      Height          =   735
      Index           =   1
      Left            =   120
      Picture         =   "Play.frx":5CC7
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Health"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   8520
      Width           =   735
   End
   Begin VB.Image imgFguy 
      Height          =   1320
      Left            =   2040
      Picture         =   "Play.frx":643A
      Top             =   240
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgBguy 
      Height          =   1320
      Left            =   2040
      Picture         =   "Play.frx":6FEA
      Top             =   240
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lblmenu 
      BackStyle       =   0  'Transparent
      Caption         =   "Abort Mission"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   10920
      MouseIcon       =   "Play.frx":7B52
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   8040
      Width           =   975
   End
   Begin VB.Image imgattack 
      Height          =   735
      Index           =   0
      Left            =   5400
      Picture         =   "Play.frx":82BC
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Charger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   8520
      Width           =   855
   End
   Begin VB.Image imgLguy 
      Height          =   1335
      Left            =   1680
      Picture         =   "Play.frx":8A2F
      Top             =   240
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgRguy 
      Height          =   1335
      Left            =   1680
      Picture         =   "Play.frx":95B7
      Top             =   240
      Width           =   840
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
'----------------------------------------------
'Clint Pitzak (Programer, Founder, Director)
'Daniel Ko    (Touch up Art, Design Leader)
'Sima Ahmad   (Flow chart, text files)
'----------------------------------------------
Dim lleft
Dim rright
Dim uup
Dim down
Dim sstop
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Dim shoot
Dim x
Dim shotflyR
Dim shotflyL
Dim shotflyF
Dim shotflyB
'API call to hide mouse cursor/pointer
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Dim W As Long
Dim health
Private Sub Command1_Click()
Form1.Visible = True
frm1.Visible = False
End Sub

Private Sub Form_Activate()
W = ShowCursor(False)
End Sub

Private Sub Form_Click()
'If click then show mouse
W = ShowCursor(True)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'-----------------------------
'KEY CODE FOR PLAYER MOVEMENTS
'-----------------------------
'UP
Select Case KeyCode
Case Is = 38
uup = True
lleft = False
rright = False
down = False
shoot = False

'RIGHT
Case Is = 39
rright = True
lleft = False
uup = False
down = False
shoot = False

'DOWN
Case Is = 40
down = True
rright = False
lleft = False
uup = False
shoot = False

'LEFT
Case Is = 37
lleft = True
rright = False
uup = False
down = False
shoot = False

'shoot
Case Is = 32
shoot = True
lleft = False
rright = False
uup = False
down = False
End Select



'------------------------
'SHOWING PICTURE OF PLAYER
'------------------------
'RIGHT
If rright = True Then
imgRguy.Visible = True
imgLguy.Visible = False
imgBguy.Visible = False
imgFguy.Visible = False
End If

'LEFT
If lleft = True Then
imgLguy.Visible = True
imgRguy.Visible = False
imgBguy.Visible = False
imgFguy.Visible = False
End If

'back
If uup = True Then
imgLguy.Visible = False
imgRguy.Visible = False
imgBguy.Visible = True
imgFguy.Visible = False
End If

'front
If down = True Then
imgFguy.Visible = True
imgLguy.Visible = False
imgRguy.Visible = False
imgBguy.Visible = False
End If
If pbhealth = 0 Then
MsgBox "Sorry you died, Try again! Bye: )"
End
End If

Dim b As Integer
For b = 0 To 10
'attack bounces off you
If Val(imgLguy.Top + imgLguy.Height) >= imgattack(b).Top Then
If imgLguy.Top < Val(imgattack(b).Top + imgattack(b).Height) Then
If Val(imgLguy.Left + imgLguy.Width) >= imgattack(b).Left Then
If imgLguy.Left < Val(imgattack(b).Left + imgattack(b).Width) Then

If imgattack(b).Left > imgLguy.Left Then imgattack(b).Left = imgattack(b).Left + 233
If imgattack(b).Top > imgLguy.Top Then imgattack(b).Top = imgattack(b).Top + 233
If imgattack(b).Left < imgLguy.Left Then imgattack(b).Left = imgattack(b).Left - 233
If imgattack(b).Top < imgLguy.Top Then imgattack(b).Top = imgattack(b).Top - 233

End If
End If
End If
End If
Next
'Dim y As Integer
'For y = 0 To 3
If shoot = True Then
Shape2.Visible = True
'If Shape2(3).Visible = False And Shape2(0).Visible = True And Shape2(1).Visible = True And Shape2(2).Visible = True Then
'Shape2(3).Visible = True
'End If
'If Shape2(2).Visible = False And Shape2(0).Visible = True And Shape2(1).Visible = True Then
'Shape2(2).Visible = True
'End If
'If Shape2(1).Visible = False And Shape2(0).Visible = True Then
'Shape2(1).Visible = True
'End If
'If Shape2(0).Visible = False Then
'Shape2(0).Visible = True
'End If
Call sndPlaySound(ByVal "Sound\Blaster_3.wav", SND_ASYNC)
End If


'----------------------
'TRAVEL SPEED FOR PLAYER
'----------------------
'LEFT VIEW
If rright = True Then
imgLguy.Left = imgLguy.Left + 200
imgRguy.Left = imgRguy.Left + 200
imgBguy.Left = imgBguy.Left + 200
imgFguy.Left = imgFguy.Left + 200

End If

If uup = True Then
imgLguy.Top = imgLguy.Top - 200
imgRguy.Top = imgRguy.Top - 200
imgBguy.Top = imgBguy.Top - 200
imgFguy.Top = imgFguy.Top - 200
End If

If down = True Then
imgLguy.Top = imgLguy.Top + 200
imgRguy.Top = imgRguy.Top + 200
imgBguy.Top = imgBguy.Top + 200
imgFguy.Top = imgFguy.Top + 200
End If

If lleft = True Then
imgLguy.Left = imgLguy.Left - 200
imgRguy.Left = imgRguy.Left - 200
imgBguy.Left = imgBguy.Left - 200
imgFguy.Left = imgFguy.Left - 200
End If

'-------------------
'Shot position Left
'-------------------

If shoot = True Then
If imgLguy.Visible = True Then
Shape2.Left = imgLguy.Left
Shape2.Top = imgLguy.Top
End If
End If
'Direction bullet fly's, Left<
If imgLguy.Visible = True And shoot = True Then
shotflyL = True
shotflyR = False
shotflyF = False
shotflyB = False
End If
'--------------------
'Shot position Right
'--------------------
If shoot = True Then
If imgRguy.Visible = True Then
Shape2.Left = imgRguy.Left + imgRguy.Width
Shape2.Top = imgRguy.Top
End If
End If
'Direction bullet fly's, Right>
If imgRguy.Visible = True And shoot = True Then
shotflyR = True
shotflyL = False
shotflyF = False
shotflyB = False
End If
'-------------------
'Shot position Back
'-------------------
If shoot = True Then
If imgBguy.Visible = True Then
Shape2.Left = imgBguy.Left + imgBguy.Width
Shape2.Top = imgBguy.Top
End If
End If
'Direction bullet fly's, Left<
If imgBguy.Visible = True And shoot = True Then
shotflyB = True
shotflyL = False
shotflyR = False
shotflyF = False
End If
'-------------------
'Shot position Front
'-------------------
If shoot = True Then
If imgFguy.Visible = True Then
Shape2.Left = imgFguy.Left
Shape2.Top = imgFguy.Top
End If
End If
'Direction bullet fly's, Left<
If imgFguy.Visible = True And shoot = True Then
shotflyF = True
shotflyB = False
shotflyL = False
shotflyR = False
End If
'when charger 0 no bullet
If pbcharger.Value < 100 And shoot = True Then
Shape2.Visible = False
Call sndPlaySound(ByVal "Sound\C_eqps.wav", SND_ASYNC)
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
End If

'--------------------------------------
'Progressbar goes down when your shoot
'--------------------------------------
If pbcharger.Value > 0 Then
If shoot = True Then pbcharger.Value = pbcharger.Value - 100
End If
'-----------------------------------------------
'   |      |
'---|------|---- THIS IS THE AREA MAKE UP FOR
'   | BOX  |     BOX EXECUTION CODE
'---|------|----
'   |      |
'-----------------------------------------------
Dim i As Integer
For i = 0 To 10

If Val(imgLguy.Top + imgLguy.Height) >= Box(i).Top Then
If imgLguy.Top < Val(Box(i).Top + Box(i).Height) Then
If Val(imgLguy.Left + imgLguy.Width) >= Box(i).Left Then
If imgLguy.Left < Val(Box(i).Left + Box(i).Width) Then

'guy bounce
If rright = True Then
imgLguy.Left = imgLguy.Left - 200
imgRguy.Left = imgRguy.Left - 200
imgBguy.Left = imgBguy.Left - 200
imgFguy.Left = imgFguy.Left - 200
End If
'----
If lleft = True Then
imgLguy.Left = imgLguy.Left + 200
imgRguy.Left = imgRguy.Left + 200
imgBguy.Left = imgBguy.Left + 200
imgFguy.Left = imgFguy.Left + 200
End If
'----
If down = True Then
imgLguy.Top = imgLguy.Top - 200
imgRguy.Top = imgRguy.Top - 200
imgBguy.Top = imgBguy.Top - 200
imgFguy.Top = imgFguy.Top - 200
End If
'----
If uup = True Then
imgLguy.Top = imgLguy.Top + 200
imgRguy.Top = imgRguy.Top + 200
imgBguy.Top = imgBguy.Top + 200
imgFguy.Top = imgFguy.Top + 200
End If
End If
End If
End If
End If
Next
'repells guy from orange bar--
Dim z As Integer
For z = 0 To 3
If Val(imgLguy.Top + imgLguy.Height) >= Ojbar(z).Top Then
If imgLguy.Top < Val(Ojbar(z).Top + Ojbar(z).Height) Then
If Val(imgLguy.Left + imgLguy.Width) >= Ojbar(z).Left Then
If imgLguy.Left < Val(Ojbar(z).Left + Ojbar(z).Width) Then

'guy bounce
If rright = True Then
imgLguy.Left = imgLguy.Left - 200
imgRguy.Left = imgRguy.Left - 200
imgBguy.Left = imgBguy.Left - 200
imgFguy.Left = imgFguy.Left - 200
End If
'----
If lleft = True Then
imgLguy.Left = imgLguy.Left + 200
imgRguy.Left = imgRguy.Left + 200
imgBguy.Left = imgBguy.Left + 200
imgFguy.Left = imgFguy.Left + 200
End If
'----
If down = True Then
imgLguy.Top = imgLguy.Top - 200
imgRguy.Top = imgRguy.Top - 200
imgBguy.Top = imgBguy.Top - 200
imgFguy.Top = imgFguy.Top - 200
End If
'----
If uup = True Then
imgLguy.Top = imgLguy.Top + 200
imgRguy.Top = imgRguy.Top + 200
imgBguy.Top = imgBguy.Top + 200
imgFguy.Top = imgFguy.Top + 200
End If
End If
End If
End If
End If
Next
End Sub

Private Sub Image13_Click()

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
pbcharger = 100
health = 100
pbhealth = 100
W = ShowCursor(False)
End Sub

Private Sub mnugameexit_Click()
End
End Sub

Private Sub lblend_Click()
End
End Sub

Private Sub Form_Terminate()
W = ShowCursor(True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
W = ShowCursor(True)
End Sub

Private Sub lblmenu_Click()
Call sndPlaySound(ByVal "Sound\na_abrt.wav", SND_ASYNC)
Form3.Visible = True
End Sub

Private Sub rechargertim_Timer()
If pbcharger.Value < 100 Then
x = x + 1
If x = 1 Then
pbcharger.Value = pbcharger.Value + 100
x = 0
End If
End If
'-------------------
'Bullet fly's Right
'-------------------
If shotflyR = True Then
Shape2.Left = Shape2.Left + 333
End If
'------------------
'Bullet fly's Left
'------------------
If shotflyL = True Then
Shape2.Left = Shape2.Left - 333
End If
'-------------------
'Bullet fly's Front
'-------------------
If shotflyF = True Then
Shape2.Top = Shape2.Top + 333
End If
'-------------------
'Bullet fly's Back
'-------------------
If shotflyB = True Then
Shape2.Top = Shape2.Top - 333
End If
Dim b As Integer
For b = 0 To 10
'Dies by attacks touch, health bar minuses-----
If Val(imgattack(b).Top + imgattack(b).Height) >= imgLguy.Top Then
If imgattack(b).Top < Val(imgLguy.Top + imgLguy.Height) Then
If Val(imgattack(b).Left + imgattack(b).Width) >= imgLguy.Left Then
If imgattack(b).Left < Val(imgLguy.Left + imgLguy.Width) Then
health = health - 15
If health > 1 Then health = health - 1 Else W = ShowCursor(True): MsgBox "Game Over! Try again:) ", vbExclamation, "You Lose : (": W = ShowCursor(True): End
pbhealth.Value = health
End If
End If
End If
End If
Next
Dim i As Integer
For i = 0 To 10
'Wall stops bullet
If Val(Shape2.Top + Shape2.Height) >= Box(i).Top Then
If Shape2.Top < Val(Box(i).Top + Box(i).Height) Then
If Val(Shape2.Left + Shape2.Width) >= Box(i).Left Then
If Shape2.Left < Val(Box(i).Left + Box(i).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
Shape2.Visible = False
End If
End If
End If
End If
Dim z As Integer
For z = 0 To 3
'Orange bar stops bullet--
If Val(Shape2.Top + Shape2.Height) >= Ojbar(z).Top Then
If Shape2.Top < Val(Ojbar(z).Top + Ojbar(z).Height) Then
If Val(Shape2.Left + Shape2.Width) >= Ojbar(z).Left Then
If Shape2.Left < Val(Ojbar(z).Left + Ojbar(z).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
Shape2.Visible = False
End If
End If
End If
End If
Dim j As Integer
For j = 0 To 7
'Blue Barrier stop bullet--
If Val(Shape2.Top + Shape2.Height) >= Shape1(j).Top Then
If Shape2.Top < Val(Shape1(j).Top + Shape1(j).Height) Then
If Val(Shape2.Left + Shape2.Width) >= Shape1(j).Left Then
If Shape2.Left < Val(Shape1(j).Left + Shape1(j).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
Shape2.Visible = False
End If
End If
End If
End If
Next
Next
'----------
'attack repelled by bullet--from here down same just different attack images 0-10--
'0-------------------------------
If Val(Shape2.Top + Shape2.Height) >= imgattack(0).Top Then
If Shape2.Top < Val(imgattack(0).Top + imgattack(0).Height) Then
If Val(Shape2.Left + Shape2.Width) >= imgattack(0).Left Then
If Shape2.Left < Val(imgattack(0).Left + imgattack(0).Width) Then
'-------------
'stops Bullet
'-------------
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
'guy fly's away---
If imgattack(0).Left > imgLguy.Left Then imgattack(0).Left = imgattack(0).Left + 9999
If imgattack(0).Top > imgLguy.Top Then imgattack(0).Top = imgattack(0).Top + 9999
If imgattack(0).Left < imgLguy.Left Then imgattack(0).Left = imgattack(0).Left - 9999
If imgattack(0).Top < imgLguy.Top Then imgattack(0).Top = imgattack(0).Top - 9999
Shape2.Visible = False
End If
End If
End If
End If
'1---------------------
If Val(Shape2.Top + Shape2.Height) >= imgattack(1).Top Then
If Shape2.Top < Val(imgattack(1).Top + imgattack(1).Height) Then
If Val(Shape2.Left + Shape2.Width) >= imgattack(1).Left Then
If Shape2.Left < Val(imgattack(1).Left + imgattack(1).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
'guy flys away---
If imgattack(1).Left > imgLguy.Left Then imgattack(1).Left = imgattack(1).Left + 9999
If imgattack(1).Top > imgLguy.Top Then imgattack(1).Top = imgattack(1).Top + 9999
If imgattack(1).Left < imgLguy.Left Then imgattack(1).Left = imgattack(1).Left - 9999
If imgattack(1).Top < imgLguy.Top Then imgattack(1).Top = imgattack(1).Top - 9999
Shape2.Visible = False
End If
End If
End If
End If
'2------------------------
If Val(Shape2.Top + Shape2.Height) >= imgattack(2).Top Then
If Shape2.Top < Val(imgattack(2).Top + imgattack(2).Height) Then
If Val(Shape2.Left + Shape2.Width) >= imgattack(2).Left Then
If Shape2.Left < Val(imgattack(2).Left + imgattack(2).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
'guy flys away---
If imgattack(2).Left > imgLguy.Left Then imgattack(2).Left = imgattack(2).Left + 9999
If imgattack(2).Top > imgLguy.Top Then imgattack(2).Top = imgattack(2).Top + 9999
If imgattack(2).Left < imgLguy.Left Then imgattack(2).Left = imgattack(2).Left - 9999
If imgattack(2).Top < imgLguy.Top Then imgattack(1).Top = imgattack(2).Top - 9999
Shape2.Visible = False
End If
End If
End If
End If
'3-------------------------
If Val(Shape2.Top + Shape2.Height) >= imgattack(3).Top Then
If Shape2.Top < Val(imgattack(3).Top + imgattack(3).Height) Then
If Val(Shape2.Left + Shape2.Width) >= imgattack(3).Left Then
If Shape2.Left < Val(imgattack(3).Left + imgattack(3).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
'guy flys away---
If imgattack(3).Left > imgLguy.Left Then imgattack(3).Left = imgattack(3).Left + 9999
If imgattack(3).Top > imgLguy.Top Then imgattack(3).Top = imgattack(3).Top + 9999
If imgattack(3).Left < imgLguy.Left Then imgattack(3).Left = imgattack(3).Left - 9999
If imgattack(3).Top < imgLguy.Top Then imgattack(3).Top = imgattack(3).Top - 9999
Shape2.Visible = False
End If
End If
End If
End If
'4------------------------------
If Val(Shape2.Top + Shape2.Height) >= imgattack(4).Top Then
If Shape2.Top < Val(imgattack(4).Top + imgattack(4).Height) Then
If Val(Shape2.Left + Shape2.Width) >= imgattack(4).Left Then
If Shape2.Left < Val(imgattack(4).Left + imgattack(4).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
'guy flys away---
If imgattack(4).Left > imgLguy.Left Then imgattack(4).Left = imgattack(4).Left + 9999
If imgattack(4).Top > imgLguy.Top Then imgattack(4).Top = imgattack(4).Top + 9999
If imgattack(4).Left < imgLguy.Left Then imgattack(4).Left = imgattack(4).Left - 9999
If imgattack(4).Top < imgLguy.Top Then imgattack(4).Top = imgattack(4).Top - 9999
Shape2.Visible = False
End If
End If
End If
End If
'5-------------------
If Val(Shape2.Top + Shape2.Height) >= imgattack(5).Top Then
If Shape2.Top < Val(imgattack(5).Top + imgattack(5).Height) Then
If Val(Shape2.Left + Shape2.Width) >= imgattack(5).Left Then
If Shape2.Left < Val(imgattack(5).Left + imgattack(5).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
'guy flys away---
If imgattack(5).Left > imgLguy.Left Then imgattack(5).Left = imgattack(5).Left + 9999
If imgattack(5).Top > imgLguy.Top Then imgattack(5).Top = imgattack(5).Top + 9999
If imgattack(5).Left < imgLguy.Left Then imgattack(5).Left = imgattack(5).Left - 9999
If imgattack(5).Top < imgLguy.Top Then imgattack(5).Top = imgattack(5).Top - 9999
Shape2.Visible = False
End If
End If
End If
End If
'6--------------------
If Val(Shape2.Top + Shape2.Height) >= imgattack(6).Top Then
If Shape2.Top < Val(imgattack(6).Top + imgattack(6).Height) Then
If Val(Shape2.Left + Shape2.Width) >= imgattack(6).Left Then
If Shape2.Left < Val(imgattack(6).Left + imgattack(6).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
'guy flys away---
If imgattack(6).Left > imgLguy.Left Then imgattack(6).Left = imgattack(6).Left + 9999
If imgattack(6).Top > imgLguy.Top Then imgattack(6).Top = imgattack(6).Top + 9999
If imgattack(6).Left < imgLguy.Left Then imgattack(6).Left = imgattack(6).Left - 9999
If imgattack(6).Top < imgLguy.Top Then imgattack(6).Top = imgattack(6).Top - 9999
Shape2.Visible = False
End If
End If
End If
End If
'7---------------------
If Val(Shape2.Top + Shape2.Height) >= imgattack(7).Top Then
If Shape2.Top < Val(imgattack(7).Top + imgattack(7).Height) Then
If Val(Shape2.Left + Shape2.Width) >= imgattack(7).Left Then
If Shape2.Left < Val(imgattack(7).Left + imgattack(7).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
'guy flys away---
If imgattack(7).Left > imgLguy.Left Then imgattack(7).Left = imgattack(7).Left + 9999
If imgattack(7).Top > imgLguy.Top Then imgattack(7).Top = imgattack(7).Top + 9999
If imgattack(7).Left < imgLguy.Left Then imgattack(7).Left = imgattack(7).Left - 9999
If imgattack(7).Top < imgLguy.Top Then imgattack(7).Top = imgattack(7).Top - 9999
Shape2.Visible = False
End If
End If
End If
End If
'8-------------------------
If Val(Shape2.Top + Shape2.Height) >= imgattack(8).Top Then
If Shape2.Top < Val(imgattack(8).Top + imgattack(8).Height) Then
If Val(Shape2.Left + Shape2.Width) >= imgattack(8).Left Then
If Shape2.Left < Val(imgattack(8).Left + imgattack(8).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
'guy flys away-----
If imgattack(8).Left > imgLguy.Left Then imgattack(8).Left = imgattack(8).Left + 9999
If imgattack(8).Top > imgLguy.Top Then imgattack(8).Top = imgattack(8).Top + 9999
If imgattack(8).Left < imgLguy.Left Then imgattack(8).Left = imgattack(8).Left - 9999
If imgattack(8).Top < imgLguy.Top Then imgattack(8).Top = imgattack(8).Top - 9999
Shape2.Visible = False
End If
End If
End If
End If
'9----------------
If Val(Shape2.Top + Shape2.Height) >= imgattack(9).Top Then
If Shape2.Top < Val(imgattack(9).Top + imgattack(9).Height) Then
If Val(Shape2.Left + Shape2.Width) >= imgattack(9).Left Then
If Shape2.Left < Val(imgattack(9).Left + imgattack(9).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
'guy flys away---
If imgattack(9).Left > imgLguy.Left Then imgattack(9).Left = imgattack(9).Left + 9999
If imgattack(9).Top > imgLguy.Top Then imgattack(9).Top = imgattack(9).Top + 9999
If imgattack(9).Left < imgLguy.Left Then imgattack(9).Left = imgattack(9).Left - 9999
If imgattack(9).Top < imgLguy.Top Then imgattack(9).Top = imgattack(9).Top - 9999
Shape2.Visible = False
End If
End If
End If
End If
'10--------------
If Val(Shape2.Top + Shape2.Height) >= imgattack(10).Top Then
If Shape2.Top < Val(imgattack(10).Top + imgattack(10).Height) Then
If Val(Shape2.Left + Shape2.Width) >= imgattack(10).Left Then
If Shape2.Left < Val(imgattack(10).Left + imgattack(10).Width) Then
shotflyR = False
shotflyL = False
shotflyF = False
shotflyB = False
'guy flys away---
If imgattack(10).Left > imgLguy.Left Then imgattack(10).Left = imgattack(10).Left + 9999
If imgattack(10).Top > imgLguy.Top Then imgattack(10).Top = imgattack(10).Top + 9999
If imgattack(10).Left < imgLguy.Left Then imgattack(10).Left = imgattack(10).Left - 9999
If imgattack(10).Top < imgLguy.Top Then imgattack(10).Top = imgattack(10).Top - 9999
Shape2.Visible = False
End If
End If
End If
End If
Next
'--End of bullet repelling attacker--
If Val(Shape2.Top + Shape2.Height) >= imgGen.Top Then
If Shape2.Top < Val(imgGen.Top + imgGen.Height) Then
If Val(Shape2.Left + Shape2.Width) >= imgGen.Left Then
If Shape2.Left < Val(imgGen.Left + imgGen.Width) Then
imgGen.Visible = False
If imgGen.Visible = False Then img2gen.Visible = True
If img2gen.Visible = True Then W = ShowCursor(True): MsgBox "Congradulations! You have destroyed then Alien's Generator, You Win:) ", vbExclamation, "Winner : )": W = ShowCursor(True): End
End If
End If
End If
End If
End Sub

Private Sub Timer1_Timer()

'--------------------------------------
'-BOUNDARY CODE- for right view of PLAYER
'--------------------------------------
'LEFT
If imgLguy.Left < Line2.X2 Then
imgLguy.Left = Line2.X2
End If

'TOP
If imgLguy.Top < Line1.Y2 Then
imgLguy.Top = Line1.Y2
End If

'RIGHT
If imgLguy.Left + imgLguy.Width > Line5.X2 Then
imgLguy.Left = Line5.X2 - imgLguy.Width
End If

'BOTTOM
If imgLguy.Top + imgLguy.Height > Line6.Y2 Then
imgLguy.Top = Line6.Y2 - imgLguy.Height
End If
'-------------------------------------
'-BOUNDARY CODE- for left view of PLAYER
'-------------------------------------
'LEFT
If imgRguy.Left < Line2.X2 Then
imgRguy.Left = Line2.X2
End If

'TOP
If imgRguy.Top < Line1.Y2 Then
imgRguy.Top = Line1.Y2
End If

'RIGHT
If imgRguy.Left + imgRguy.Width > Line5.X2 Then
imgRguy.Left = Line5.X2 - imgRguy.Width
End If

'BOTTOM
If imgRguy.Top + imgRguy.Height > Line6.Y2 Then
imgRguy.Top = Line6.Y2 - imgRguy.Height
End If
'-----------------------------------------
'FRONT Guy
'-----------------------------------------
'LEFT
If imgFguy.Left < Line2.X2 Then
imgFguy.Left = Line2.X2
End If

'TOP
If imgFguy.Top < Line1.Y2 Then
imgFguy.Top = Line1.Y2
End If

'RIGHT
If imgFguy.Left + imgFguy.Width > Line5.X2 Then
imgFguy.Left = Line5.X2 - imgFguy.Width
End If

'BOTTOM
If imgFguy.Top + imgFguy.Height > Line6.Y2 Then
imgFguy.Top = Line6.Y2 - imgFguy.Height
End If

'-----------------------------------------
'BACK of Guy
'-----------------------------------------
'LEFT
If imgBguy.Left < Line2.X2 Then
imgBguy.Left = Line2.X2
End If

'TOP
If imgBguy.Top < Line1.Y2 Then
imgBguy.Top = Line1.Y2
End If

'RIGHT
If imgBguy.Left + imgBguy.Width > Line5.X2 Then
imgBguy.Left = Line5.X2 - imgBguy.Width
End If

'BOTTOM
If imgBguy.Top + imgBguy.Height > Line6.Y2 Then
imgBguy.Top = Line6.Y2 - imgBguy.Height
End If


Dim b As Integer
For b = 0 To 10
'BAD GUY ATTACK----
If imgattack(b).Left > imgLguy.Left Then imgattack(b).Left = imgattack(b).Left - 33
If imgattack(b).Top > imgLguy.Top Then imgattack(b).Top = imgattack(b).Top - 33
If imgattack(b).Left < imgLguy.Left Then imgattack(b).Left = imgattack(b).Left + 33
If imgattack(b).Top < imgLguy.Top Then imgattack(b).Top = imgattack(b).Top + 33

'PREVENTS ATTACKER FROM going through walls--
Dim i As Integer
For i = 0 To 10

If Val(imgattack(b).Top + imgattack(b).Height) >= Box(i).Top Then
If imgattack(b).Top < Val(Box(i).Top + Box(i).Height) Then
If Val(imgattack(b).Left + imgattack(b).Width) >= Box(i).Left Then
If imgattack(b).Left < Val(Box(i).Left + Box(i).Width) Then

If imgattack(b).Left > imgLguy.Left Then imgattack(b).Left = imgattack(b).Left + 33
If imgattack(b).Top > imgLguy.Top Then imgattack(b).Top = imgattack(b).Top + 33
If imgattack(b).Left < imgLguy.Left Then imgattack(b).Left = imgattack(b).Left - 33
If imgattack(b).Top < imgLguy.Top Then imgattack(b).Top = imgattack(b).Top - 33

End If
End If
End If
End If
Next
Next
End Sub


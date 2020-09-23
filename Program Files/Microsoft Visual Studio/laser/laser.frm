VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Quick Sketch"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   DrawWidth       =   25
   LinkTopic       =   "Form1"
   Picture         =   "laser.frx":0000
   ScaleHeight     =   6435
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Prog 
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   5640
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   5535
      Left            =   5280
      ScaleHeight     =   5475
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   0
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SKETCH"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   5880
      Width           =   1695
   End
   Begin VB.PictureBox pb 
      BackColor       =   &H80000009&
      Height          =   5535
      Left            =   0
      Picture         =   "laser.frx":D2C2
      ScaleHeight     =   5475
      ScaleWidth      =   5235
      TabIndex        =   3
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim GLine As String
Dim PicX As Single
Dim PicY As Single
Dim PC As Double

Prog.Value = 0
Prog.Max = pb.Height + 30

For PicY = 0 To pb.Height Step Screen.TwipsPerPixelY
Prog.Value = Prog.Value + Screen.TwipsPerPixelY
x = PicX
y = PicY

For PicX = 0 To pb.Width Step Screen.TwipsPerPixelX
PC = pb.Point(PicX, PicY) 'get pixle color

x = PicX
If x = 5895 Or x = 0 Or x = 15 Then
Picture2.PSet (x, y), PC
GoTo skip:
End If

If y = PicY Then
y = Picture2.Height
x = Picture2.Width \ 2
Else
y = PicY
x = PicX
End If
    
Picture2.Line -(x, y), PC
skip:
Picture2.Line -(PicX, PicY), PC
Next
Next
End Sub


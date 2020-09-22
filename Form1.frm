VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Start Button Animation"
   ClientHeight    =   3495
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      Height          =   255
      Left            =   3360
      TabIndex        =   36
      Top             =   0
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Style 4"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   51
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1635
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   123
      TabIndex        =   50
      Top             =   480
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   11
      Left            =   2640
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   49
      Top             =   2520
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   10
      Left            =   1800
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   48
      Top             =   2520
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   9
      Left            =   960
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   47
      Top             =   2520
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   8
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   46
      Top             =   2520
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   7
      Left            =   2640
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   45
      Top             =   2280
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   6
      Left            =   1800
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   44
      Top             =   2280
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   5
      Left            =   960
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   43
      Top             =   2280
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   4
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   42
      Top             =   2280
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   3
      Left            =   2640
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   41
      Top             =   2040
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   2
      Left            =   1800
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   40
      Top             =   2040
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   960
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   39
      Top             =   2040
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Style 3"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   38
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   37
      Top             =   2040
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Style 2"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   35
      Top             =   0
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Style 1"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   18
      Left            =   2280
      Picture         =   "Form1.frx":0FAB
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   33
      Top             =   4560
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   17
      Left            =   1560
      Picture         =   "Form1.frx":18ED
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   32
      Top             =   4560
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   16
      Left            =   840
      Picture         =   "Form1.frx":222F
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   31
      Top             =   4560
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   15
      Left            =   120
      Picture         =   "Form1.frx":2B71
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   30
      Top             =   4560
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   14
      Left            =   3000
      Picture         =   "Form1.frx":34B3
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   29
      Top             =   4320
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   13
      Left            =   2280
      Picture         =   "Form1.frx":3DF5
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   28
      Top             =   4320
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   12
      Left            =   1560
      Picture         =   "Form1.frx":4737
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   27
      Top             =   4320
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   11
      Left            =   840
      Picture         =   "Form1.frx":5079
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   26
      Top             =   4320
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   10
      Left            =   120
      Picture         =   "Form1.frx":59BB
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   25
      Top             =   4320
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   9
      Left            =   3000
      Picture         =   "Form1.frx":62FD
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   8
      Left            =   2280
      Picture         =   "Form1.frx":6C3F
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   23
      Top             =   4080
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   7
      Left            =   1560
      Picture         =   "Form1.frx":7581
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   22
      Top             =   4080
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   6
      Left            =   840
      Picture         =   "Form1.frx":7EC3
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   5
      Left            =   120
      Picture         =   "Form1.frx":8805
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   4
      Left            =   3000
      Picture         =   "Form1.frx":9147
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   3
      Left            =   2280
      Picture         =   "Form1.frx":9A89
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   2
      Left            =   1560
      Picture         =   "Form1.frx":A3CB
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   840
      Picture         =   "Form1.frx":AD0D
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":B64F
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   3960
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   14
      Left            =   3000
      Picture         =   "Form1.frx":BF91
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   13
      Left            =   2280
      Picture         =   "Form1.frx":C8D3
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   12
      Left            =   1560
      Picture         =   "Form1.frx":D215
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   12
      Top             =   3360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   11
      Left            =   840
      Picture         =   "Form1.frx":DB57
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   10
      Left            =   120
      Picture         =   "Form1.frx":E499
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   9
      Left            =   3000
      Picture         =   "Form1.frx":EDDB
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   8
      Left            =   2280
      Picture         =   "Form1.frx":F71D
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   7
      Left            =   1560
      Picture         =   "Form1.frx":1005F
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   6
      Left            =   840
      Picture         =   "Form1.frx":109A1
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   5
      Left            =   120
      Picture         =   "Form1.frx":112E3
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   4
      Left            =   3000
      Picture         =   "Form1.frx":11C25
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   3
      Left            =   2280
      Picture         =   "Form1.frx":12567
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   2
      Left            =   1560
      Picture         =   "Form1.frx":12EA9
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   840
      Picture         =   "Form1.frx":137EB
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":1412D
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lngTrayWnd As Long, lngStartButton As Long, SBhDc As Long, dCount As Byte, s As Single, n As Long, i As Integer, ii As Integer, c As Integer

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load(): Me.Hide

'Find StartButton's DC
lngTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
lngStartButton = FindChildByClass(lngTrayWnd, "Button")
SBhDc = GetDC(lngStartButton)

For c = 0 To 11
For i = 0 To Picture3(c).ScaleWidth: DoEvents
For ii = 0 To Picture3(c).ScaleHeight
n = Rnd * 255
Picture3(c).PSet (i, ii), RGB(n, n, n)
Next
Next
Next

'Start Tray Icon
Init_Tray Me, "Start Button Animation"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Msg = X / Screen.TwipsPerPixelX
Select Case Msg
Case WM_LBUTTONUP
PopupMenu mnuMain
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload_Tray
Call BitBlt(SBhDc, 3, 3, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight, Picture1(0).hdc, 0, 0, SRCCOPY)
End Sub

Private Sub mnuExit_Click()
Unload_Tray
Unload Me
End Sub

Private Sub mnuShow_Click()
Me.Show
End Sub

Private Sub Option1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
dCount = 0
End Sub

Private Sub Timer1_Timer(): On Error Resume Next
Timer1.Interval = 100
If Option1(0).Value Then
If s = 1 Then
dCount = dCount + s
If dCount = 15 Then s = -1
Call BitBlt(SBhDc, 3, 3, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight, Picture1(dCount).hdc, 0, 0, SRCCOPY)
Else
dCount = dCount + s
If dCount = 0 Then s = 1
BitBlt SBhDc, 3, 3, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight, Picture1(dCount).hdc, 0, 0, SRCCOPY
End If
ElseIf Option1(1).Value Then
dCount = dCount + 1
If dCount = 19 Then dCount = 0
Call BitBlt(SBhDc, 3, 3, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight, Picture2(dCount).hdc, 0, 0, SRCCOPY)
ElseIf Option1(2).Value Then
Call BitBlt(SBhDc, 3, 3, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight, Picture3(Int(Rnd * 11)).hdc, 0, 0, SRCCOPY)
Else
Timer1.Interval = 500
Call BitBlt(SBhDc, 3, 3, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight, Picture4.hdc, RndRange(0, Picture4.ScaleHeight - Picture1(0).ScaleWidth), RndRange(0, Picture4.ScaleHeight - Picture1(0).ScaleHeight), SRCCOPY)
End If

End Sub
Public Function RndRange(ByVal intMin As Integer, ByVal intMax As Integer)
RndRange = Int(Rnd * (intMax - intMin + 1)) + intMin
End Function

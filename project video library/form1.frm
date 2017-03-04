VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   12885
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   8999
      Left            =   12120
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   60
      Left            =   7560
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   11640
      Top             =   1920
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   6240
      TabIndex        =   6
      Top             =   1080
      Width           =   4815
      Begin VB.PictureBox Picture12 
         Height          =   1815
         Left            =   2760
         Picture         =   "form1.frx":0000
         ScaleHeight     =   1755
         ScaleWidth      =   1755
         TabIndex        =   12
         Top             =   4920
         Width           =   1815
      End
      Begin VB.PictureBox Picture11 
         Height          =   1695
         Left            =   360
         Picture         =   "form1.frx":17DC
         ScaleHeight     =   1635
         ScaleWidth      =   1395
         TabIndex        =   11
         Top             =   5160
         Width           =   1455
      End
      Begin VB.PictureBox Picture10 
         Height          =   1455
         Left            =   2640
         Picture         =   "form1.frx":280A
         ScaleHeight     =   1395
         ScaleWidth      =   1875
         TabIndex        =   10
         Top             =   2880
         Width           =   1935
      End
      Begin VB.PictureBox Picture9 
         Height          =   1815
         Left            =   360
         Picture         =   "form1.frx":3878
         ScaleHeight     =   1755
         ScaleWidth      =   1635
         TabIndex        =   9
         Top             =   2760
         Width           =   1695
      End
      Begin VB.PictureBox Picture8 
         Height          =   1935
         Left            =   2640
         Picture         =   "form1.frx":4D92
         ScaleHeight     =   1875
         ScaleWidth      =   1515
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.PictureBox Picture7 
         Height          =   1455
         Left            =   360
         Picture         =   "form1.frx":5E6F
         ScaleHeight     =   1395
         ScaleWidth      =   1875
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   5415
      Begin VB.PictureBox Picture4 
         Height          =   2055
         Left            =   2520
         Picture         =   "form1.frx":70D2
         ScaleHeight     =   1995
         ScaleWidth      =   1995
         TabIndex        =   27
         Top             =   2880
         Width           =   2055
      End
      Begin VB.PictureBox Picture6 
         Height          =   1695
         Left            =   2520
         Picture         =   "form1.frx":939C
         ScaleHeight     =   1635
         ScaleWidth      =   1875
         TabIndex        =   5
         Top             =   5640
         Width           =   1935
      End
      Begin VB.PictureBox Picture5 
         Height          =   1695
         Left            =   120
         Picture         =   "form1.frx":A4B9
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   4
         Top             =   5520
         Width           =   1815
      End
      Begin VB.PictureBox Picture3 
         Height          =   1815
         Left            =   240
         Picture         =   "form1.frx":BA3F
         ScaleHeight     =   1755
         ScaleWidth      =   1875
         TabIndex        =   3
         Top             =   2640
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         Height          =   1815
         Left            =   2640
         Picture         =   "form1.frx":D164
         ScaleHeight     =   1755
         ScaleWidth      =   2115
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Height          =   1815
         Left            =   240
         Picture         =   "form1.frx":E645
         ScaleHeight     =   1755
         ScaleWidth      =   2115
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   480
      TabIndex        =   13
      Top             =   1080
      Width           =   10695
      Begin VB.PictureBox Picture24 
         Height          =   1335
         Left            =   7800
         Picture         =   "form1.frx":F67A
         ScaleHeight     =   1275
         ScaleWidth      =   2115
         TabIndex        =   26
         Top             =   4920
         Width           =   2175
      End
      Begin VB.PictureBox Picture23 
         Height          =   1575
         Left            =   4920
         Picture         =   "form1.frx":10593
         ScaleHeight     =   1515
         ScaleWidth      =   2115
         TabIndex        =   24
         Top             =   4800
         Width           =   2175
      End
      Begin VB.PictureBox Picture22 
         Height          =   1575
         Left            =   2400
         Picture         =   "form1.frx":11B2C
         ScaleHeight     =   1515
         ScaleWidth      =   1995
         TabIndex        =   23
         Top             =   5040
         Width           =   2055
      End
      Begin VB.PictureBox Picture21 
         Height          =   1455
         Left            =   360
         Picture         =   "form1.frx":12EEF
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   22
         Top             =   5160
         Width           =   1455
      End
      Begin VB.PictureBox Picture20 
         Height          =   1815
         Left            =   7440
         Picture         =   "form1.frx":13FDB
         ScaleHeight     =   1755
         ScaleWidth      =   2115
         TabIndex        =   21
         Top             =   2640
         Width           =   2175
      End
      Begin VB.PictureBox Picture19 
         Height          =   1455
         Left            =   5160
         Picture         =   "form1.frx":157B7
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   20
         Top             =   2880
         Width           =   1455
      End
      Begin VB.PictureBox Picture18 
         Height          =   1575
         Left            =   2160
         Picture         =   "form1.frx":1683F
         ScaleHeight     =   1515
         ScaleWidth      =   2355
         TabIndex        =   19
         Top             =   2760
         Width           =   2415
      End
      Begin VB.PictureBox Picture17 
         Height          =   1695
         Left            =   240
         Picture         =   "form1.frx":17F83
         ScaleHeight     =   1635
         ScaleWidth      =   1395
         TabIndex        =   18
         Top             =   2640
         Width           =   1455
      End
      Begin VB.PictureBox Picture16 
         Height          =   1815
         Left            =   8160
         Picture         =   "form1.frx":189F3
         ScaleHeight     =   1755
         ScaleWidth      =   1395
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.PictureBox Picture15 
         Height          =   1935
         Left            =   5400
         Picture         =   "form1.frx":199C7
         ScaleHeight     =   1875
         ScaleWidth      =   1995
         TabIndex        =   16
         Top             =   360
         Width           =   2055
      End
      Begin VB.PictureBox Picture14 
         Height          =   1695
         Left            =   2880
         Picture         =   "form1.frx":1ADE7
         ScaleHeight     =   1635
         ScaleWidth      =   1995
         TabIndex        =   15
         Top             =   480
         Width           =   2055
      End
      Begin VB.PictureBox Picture13 
         Height          =   1695
         Left            =   240
         Picture         =   "form1.frx":1C50C
         ScaleHeight     =   1635
         ScaleWidth      =   2235
         TabIndex        =   14
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      Caption         =   "4 FOX  VIDEO LIBRARY"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      TabIndex        =   25
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim flag As Variant

Private Sub Form_Load()
Form2.WindowsMediaPlayer1.URL = "C:\Users\Admin\Documents\proj.wav"
 


End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Timer1_Timer()

If (flag = 0) Then

Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
End If

If (flag = 1) Then
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
End If

If (flag = 2) Then
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
End If
flag = (flag + 1) Mod 3


End Sub

Private Sub Timer2_Timer()
Label1.BackColor = RGB(256 * Rnd, 256 * Rnd, 0)
End Sub

Private Sub Timer3_Timer()
Form1.Visible = False
frmLogin.Visible = True
End Sub

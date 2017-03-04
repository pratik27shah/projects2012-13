VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000001&
   Caption         =   "Form3"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15060
   LinkTopic       =   "Form3"
   ScaleHeight     =   9600
   ScaleWidth      =   15060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   11
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   10080
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9720
      TabIndex        =   2
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View all"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      TabIndex        =   10
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   10320
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   720
      Top             =   2400
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      TabIndex        =   1
      Top             =   5280
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      TabIndex        =   0
      Top             =   3720
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   1080
      Picture         =   "form3.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   10755
      TabIndex        =   3
      Top             =   6480
      Width           =   10815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   10320
      TabIndex        =   4
      Top             =   5400
      Width           =   495
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   480
         TabIndex        =   7
         Text            =   "Text5"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   600
         Width           =   150
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   120
         Width           =   150
      End
   End
   Begin VB.Label Label2 
      Caption         =   "CD_Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "ID_no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   7125
      Index           =   0
      Left            =   360
      Picture         =   "form3.frx":138B2
      Top             =   0
      Width           =   9555
   End
   Begin VB.Image Image1 
      Height          =   6930
      Index           =   1
      Left            =   240
      Picture         =   "form3.frx":1BBA3
      Top             =   0
      Width           =   9690
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim f As Variant

Private Sub Command2_Click()
DataReport1.Show

End Sub

Private Sub Command3_Click()

If f = 1 Then
cn.Execute "delete from studd where CD_ID='" & UCase$(Text1.Text) & "'"


rs.MoveFirst

MsgBox ("Record  deleted")
rs.Update
Else
MsgBox ("Enter CD code")
End If
f = 0
End Sub

Private Sub Picture1_Click()
Picture1.Left = Picture1.Left - 10

End Sub

Private Sub text1_dblclick()
Dim s As Variant

Dim s1 As Variant, s2 As Variant
s = Len(Text2.Text)
s1 = Mid$(Text2.Text, 1, 2)
s2 = Right$(Text2.Text, 1)
Text1.Text = UCase$(s1 + Str(s) + s2)

f = 1
End Sub

Private Sub Command1_Click()

If f = 1 Then
cn.Execute "insert into studd values('" & Text1.Text & "' ,'" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "')"


rs.MoveFirst

MsgBox ("Record  inserted")
rs.Update
Else
MsgBox ("Enter CD code")
End If
f = 0
End Sub

Private Sub Form_Load()
Text3.Text = "0"
Text4.Text = "0"
Text5.Text = "0"

cn.Open "tdsn", "system", "pratik" 'connectivity
rs.Open "select * from studd", cn, adOpenStatic, adLockOptimistic
End Sub

Private Sub Text2_GotFocus()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Timer1_Timer()
If (Picture1.Left > 1000 And Picture1.Left < 4999) Then
Picture1.Left = Picture1.Left - 100
Else
'If (Picture1.Left > 5000) Then

Picture1.Left = Picture1.Left + 2500
'End If
End If
End Sub

Private Sub Timer2_Timer()
Image1(0).Visible = True
Timer2.Enabled = False


End Sub

Private Sub Timer3_Timer()
Image1(0).Visible = False
Timer2.Enabled = True
End Sub

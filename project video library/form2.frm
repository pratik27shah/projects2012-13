VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   Caption         =   "Form2"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
   LinkTopic       =   "Form2"
   ScaleHeight     =   9000
   ScaleWidth      =   15270
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Index           =   10
      Left            =   12360
      Picture         =   "form2.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   44
      Top             =   5160
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Index           =   9
      Left            =   12120
      Picture         =   "form2.frx":0EEE
      ScaleHeight     =   1275
      ScaleWidth      =   1755
      TabIndex        =   43
      Top             =   3000
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Index           =   8
      Left            =   12240
      Picture         =   "form2.frx":1C9F
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   42
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Index           =   7
      Left            =   10560
      Picture         =   "form2.frx":2757
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   41
      Top             =   7440
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Index           =   3
      Left            =   8040
      Picture         =   "form2.frx":33F1
      ScaleHeight     =   1155
      ScaleWidth      =   1875
      TabIndex        =   37
      Top             =   7440
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   1695
      Index           =   2
      Left            =   8040
      Picture         =   "form2.frx":456A
      ScaleHeight     =   1635
      ScaleWidth      =   1395
      TabIndex        =   36
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Timer Timer6 
      Interval        =   8500
      Left            =   0
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Index           =   9
      Left            =   12120
      Picture         =   "form2.frx":54E3
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   32
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Index           =   8
      Left            =   12480
      Picture         =   "form2.frx":6181
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   31
      Top             =   840
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Index           =   7
      Left            =   10560
      Picture         =   "form2.frx":6D8F
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   30
      Top             =   7680
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Index           =   3
      Left            =   8040
      Picture         =   "form2.frx":79F5
      ScaleHeight     =   1035
      ScaleWidth      =   1875
      TabIndex        =   26
      Top             =   7560
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Index           =   2
      Left            =   8040
      Picture         =   "form2.frx":845F
      ScaleHeight     =   1635
      ScaleWidth      =   1395
      TabIndex        =   25
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Timer Timer5 
      Interval        =   1800
      Left            =   7680
      Top             =   4200
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   7680
      Top             =   4560
   End
   Begin VB.Timer Timer3 
      Interval        =   300
      Left            =   360
      Top             =   7440
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   6960
      Top             =   1080
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   480
      TabIndex        =   22
      Top             =   6240
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "1 year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   21
      Top             =   5640
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "6 months"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   5640
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "2 months"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   19
      Top             =   5640
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1095
      Left            =   2760
      TabIndex        =   18
      Top             =   3240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1931
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"form2.frx":90F9
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H000000FF&
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ADD CD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   15
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Already Exsisting Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   14
      Top             =   8280
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generate ID"
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
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   1920
      Top             =   360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   10
      Top             =   6960
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Text            =   "+91"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   7
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1560
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Text            =   "Text4"
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   17
      Top             =   6720
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Index           =   4
      Left            =   10200
      Picture         =   "form2.frx":917B
      ScaleHeight     =   1515
      ScaleWidth      =   1635
      TabIndex        =   38
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Index           =   4
      Left            =   10200
      Picture         =   "form2.frx":9EA0
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   27
      Top             =   840
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Index           =   5
      Left            =   9960
      Picture         =   "form2.frx":A96B
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   39
      Top             =   3000
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Index           =   6
      Left            =   10200
      Picture         =   "form2.frx":B3BC
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   40
      Top             =   5400
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Index           =   10
      Left            =   12360
      Picture         =   "form2.frx":BF91
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   33
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Index           =   5
      Left            =   10320
      Picture         =   "form2.frx":CE88
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   28
      Top             =   3000
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Index           =   6
      Left            =   10200
      Picture         =   "form2.frx":D940
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   29
      Top             =   5400
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Index           =   11
      Left            =   12480
      Picture         =   "form2.frx":E7DD
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   45
      Top             =   7320
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Index           =   11
      Left            =   12480
      Picture         =   "form2.frx":F6D4
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   34
      Top             =   7320
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Index           =   1
      Left            =   8040
      Picture         =   "form2.frx":1054C
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   35
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Index           =   0
      Left            =   8040
      Picture         =   "form2.frx":10F86
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   23
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   1695
      Index           =   0
      Left            =   8040
      Picture         =   "form2.frx":11C95
      ScaleHeight     =   1635
      ScaleWidth      =   1395
      TabIndex        =   46
      Top             =   720
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Index           =   1
      Left            =   8040
      Picture         =   "form2.frx":126E6
      ScaleHeight     =   1635
      ScaleWidth      =   1395
      TabIndex        =   24
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "             REGISTRATIONS"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer2 
      Height          =   135
      Left            =   3840
      TabIndex        =   48
      Top             =   240
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   238
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   135
      Left            =   3000
      TabIndex        =   47
      Top             =   360
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   238
   End
   Begin VB.Label Label4 
      Caption         =   "Deposit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Mobile Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Customer Id:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   6855
      Left            =   0
      Picture         =   "form2.frx":13408
      Top             =   240
      Width           =   15645
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim i As Integer

Dim ch As Integer

Dim rs As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim n As Integer




'table name vb1
Private Sub Command1_Click()
'Print ch
If (ch > 3) Then

rs.MoveFirst

cn.Execute "insert into vb1 values('" & Text1.Text & "' ,'" & Text2.Text & "','" & Text5.Text & "','" & RichTextBox1.Text & "','" & Text3.Text & "','" & Text7.Text & "')"
MsgBox ("Record  inserted")
rs.Update
Else
MsgBox ("Values Missing")
End If

End Sub


Private Sub Command3_Click()
ch = ch + 1

Dim s1 As Variant
s1 = Left(Text2.Text, 3)
Dim mob As Variant
mob = Right(Text5.Text, 4) + Mid$(Text5.Text, 2, 2)
Text1.Text = s1 + mob
Timer2.Enabled = False
End Sub


Private Sub Command2_Click()
End
End Sub


Private Sub Command5_Click()
Form3.Visible = True

End Sub

Private Sub Command4_Click()
Form4.Visible = True
Form4.Enabled = True
Load Form4
End Sub

Private Sub Command6_Click()
Dim roll As Variant
Dim n As Integer
roll = Text1.Text
rs.MoveFirst

n = 1
Dim flag As Integer
flag = 0


Do Until n = rs.RecordCount
If InStr(LCase$(rs(0)), LCase$(Text1.Text)) Then    'delete  record
flag = 1
n = rs.RecordCount
cn.Execute "delete from vb1 where CUS_ID='" & LCase$(Text1.Text) & "'"
MsgBox ("Record Deleted")
Else
n = n + 1
rs.MoveNext   'move to next record
flag = 0
End If
Loop
rs.Update
'rsc.Update
rs.Update
If flag = 0 Then
MsgBox ("Not Found  to delete")
End If
End Sub


Private Sub Command7_Click()

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

Private Sub Form_Load()

ch = 0

Timer2.Enabled = True
Timer3.Enabled = True
Text4.Visible = False
Text4.Text = 0


cn.Open "tdsn", "system", "pratik" 'connectivity
rs.Open "select * from vb1", cn, adOpenStatic, adLockOptimistic
End Sub


Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
ch = ch + 1
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or (KeyAscii = 8) Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or (KeyAscii = 32) Then
Else
    MsgBox ("Enter valid address....")
     
End If
End Sub





Private Sub Text2_GotFocus()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text5.Text = ""
ch = ch + 1

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or (KeyAscii = 8) Then
Else
    MsgBox ("Enter valid name....")
     KeyAscii = 0
End If
End Sub

Private Sub Text3_Gotfocus()
ch = ch + 1

If (Option1.Value = True) Then
Text3.Text = 2000
Text7.Text = "2 months"
End If
If (Option2.Value = True) Then
Text7.Text = "6 months"
Text3.Text = 3000
End If
If (Option3.Value = True) Then
Text7.Text = "1 year"

Text3.Text = 5000
End If


End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
ch = ch + 1
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or (KeyAscii = 8) Then
Else
    MsgBox ("Enter valid mobile number...")
     KeyAscii = 0
End If
End Sub



Private Sub Timer1_Timer()
Label1.BackColor = RGB(256 * Rnd, 256 * Rnd, 256 * Rnd)
End Sub

Private Sub Timer2_Timer()
Command3.BackColor = RGB(256 * Rnd, 256 * Rnd, 256)
End Sub

Private Sub Timer3_Timer()
Label2.BackColor = RGB(0, 256 * Rnd, 256 * Rnd)
Label3.BackColor = RGB(0, 256 * Rnd, 256 * Rnd)
Label4.BackColor = RGB(256 * Rnd, 0, 256 * Rnd)
Label6.BackColor = RGB(256 * Rnd, 256 * Rnd, 256 * Rnd)
Label7.BackColor = RGB(0, 0, 256 * Rnd)

If (ch > 4) Then
Command1.BackColor = RGB(256 * Rnd, 256 * Rnd, 0)
End If

End Sub

Private Sub Timer4_Timer()
For i = 0 To 11 Step 1
Picture2(i).Visible = False
Next i
End Sub

Private Sub Timer5_Timer()
For i = 0 To 11 Step 1
Picture2(i).Visible = True
Next i
End Sub

Private Sub Timer6_Timer()
Form2.BackColor = RGB(256 * Rnd, 230 * Rnd, 0)
End Sub

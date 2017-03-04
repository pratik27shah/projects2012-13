VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   15045
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer6 
      Interval        =   1300
      Left            =   14400
      Top             =   1200
   End
   Begin VB.Timer Timer5 
      Interval        =   2500
      Left            =   14280
      Top             =   480
   End
   Begin VB.Timer Timer4 
      Interval        =   400
      Left            =   3360
      Top             =   3360
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Avalability"
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
      Left            =   4560
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.Timer Timer3 
      Interval        =   1600
      Left            =   5880
      Top             =   9600
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   12120
      Top             =   2280
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Index           =   9
      Left            =   12000
      Picture         =   "form4.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   1635
      TabIndex        =   36
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Index           =   8
      Left            =   12240
      Picture         =   "form4.frx":0F79
      ScaleHeight     =   1995
      ScaleWidth      =   1515
      TabIndex        =   34
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Index           =   6
      Left            =   1320
      Picture         =   "form4.frx":2274
      ScaleHeight     =   2115
      ScaleWidth      =   1275
      TabIndex        =   31
      Top             =   0
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Index           =   5
      Left            =   0
      Picture         =   "form4.frx":356F
      ScaleHeight     =   2115
      ScaleWidth      =   1275
      TabIndex        =   30
      Top             =   0
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Index           =   4
      Left            =   5760
      Picture         =   "form4.frx":4A89
      ScaleHeight     =   2115
      ScaleWidth      =   1395
      TabIndex        =   29
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Index           =   3
      Left            =   4320
      Picture         =   "form4.frx":5E16
      ScaleHeight     =   2115
      ScaleWidth      =   1395
      TabIndex        =   28
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Index           =   1
      Left            =   7320
      Picture         =   "form4.frx":6BB2
      ScaleHeight     =   2115
      ScaleWidth      =   1395
      TabIndex        =   26
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Index           =   5
      Left            =   5760
      Picture         =   "form4.frx":7CE9
      ScaleHeight     =   2115
      ScaleWidth      =   1515
      TabIndex        =   23
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Index           =   6
      Left            =   7320
      Picture         =   "form4.frx":90B5
      ScaleHeight     =   2115
      ScaleWidth      =   1515
      TabIndex        =   22
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Index           =   3
      Left            =   4320
      Picture         =   "form4.frx":A442
      ScaleHeight     =   2115
      ScaleWidth      =   1515
      TabIndex        =   21
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Index           =   2
      Left            =   -240
      Picture         =   "form4.frx":B73D
      ScaleHeight     =   2115
      ScaleWidth      =   1515
      TabIndex        =   20
      Top             =   0
      Width           =   1575
      Begin VB.PictureBox Picture2 
         Height          =   2175
         Index           =   0
         Left            =   240
         ScaleHeight     =   2115
         ScaleWidth      =   1275
         TabIndex        =   25
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Index           =   1
      Left            =   1320
      Picture         =   "form4.frx":C4D9
      ScaleHeight     =   2115
      ScaleWidth      =   1515
      TabIndex        =   19
      Top             =   0
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   120
      Top             =   2160
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C000C0&
      Caption         =   "OUT CDs"
      Height          =   495
      Left            =   8160
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8760
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "RETURN"
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
      Left            =   1440
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   13
      Text            =   "Mobile Number"
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C000C0&
      Caption         =   "Taken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8520
      UseMaskColor    =   -1  'True
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C000C0&
      Caption         =   "RETURN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8520
      UseMaskColor    =   -1  'True
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d.MMMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   10
      Top             =   7440
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d.MMMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   9
      Top             =   6480
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   6
      Top             =   5640
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Top             =   4920
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      Top             =   2280
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Index           =   2
      Left            =   2640
      Picture         =   "form4.frx":D9F3
      ScaleHeight     =   2115
      ScaleWidth      =   1635
      TabIndex        =   27
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Index           =   0
      Left            =   2880
      Picture         =   "form4.frx":EA0E
      ScaleHeight     =   2115
      ScaleWidth      =   1515
      TabIndex        =   18
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Index           =   7
      Left            =   8880
      Picture         =   "form4.frx":FB45
      ScaleHeight     =   2115
      ScaleWidth      =   1395
      TabIndex        =   32
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Index           =   8
      Left            =   10320
      Picture         =   "form4.frx":10A72
      ScaleHeight     =   1995
      ScaleWidth      =   1635
      TabIndex        =   35
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Index           =   7
      Left            =   10440
      Picture         =   "form4.frx":119D5
      ScaleHeight     =   2115
      ScaleWidth      =   1515
      TabIndex        =   33
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Index           =   4
      Left            =   8880
      Picture         =   "form4.frx":12902
      ScaleHeight     =   1995
      ScaleWidth      =   1515
      TabIndex        =   24
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CD_ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   17
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "OUT Date:"
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
      TabIndex        =   8
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "IN Date:"
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
      Left            =   720
      TabIndex        =   7
      Top             =   7440
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Custmor ID:"
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
      Left            =   720
      TabIndex        =   4
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Custmor Name:"
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
      Left            =   720
      TabIndex        =   3
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CD Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   10365
      Left            =   0
      Picture         =   "form4.frx":13554
      Top             =   120
      Width           =   14505
   End
   Begin VB.Image Image1 
      Height          =   12135
      Left            =   -1320
      Picture         =   "form4.frx":29585
      Top             =   -1920
      Width           =   16290
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim cn As New ADODB.Connection

Dim rs1 As New ADODB.Recordset
Dim cn1 As New ADODB.Connection


Dim rsc As New ADODB.Recordset
Dim cnc As New ADODB.Connection


Dim ch As Integer, ch2 As Integer



'studd2 out cd table

Private Sub Command1_Click()
Dim roll As Variant
Dim n As Integer
roll = Text1.Text
rs1.MoveFirst

n = 0
Dim flag As Integer
flag = 0


Do Until n = rs1.RecordCount
If InStr(LCase$(rs1(1)), LCase$(Text1.Text)) Then    'search record
Label2.Caption = "found"
n = rs1.RecordCount
Dim s As Variant, s1 As Variant, s2 As Variant
s = Len(Text1.Text)
s1 = Mid$(Text1.Text, 1, 2)
s2 = Right$(Text1.Text, 1)
Text8.Text = s1 + Str(s) + s2
ch2 = 1

flag = 1
Else
n = n + 1
rs1.MoveNext   'move to next record
flag = 0
End If
Loop

If flag = 0 Then
Label2.Caption = " Not found sorry"
End If
If (Text1.Text = " ") Then
ch = 0
Else
ch = 1
End If
End Sub




Private Sub Command2_Click()
Dim flag As Integer, n As Integer
n = 0
flag = 0
If (ch2 = 1 And Text3.Text <> " " And Text2.Text <> " ") Then

rsc.MoveFirst


Do Until n = rsc.RecordCount
If InStr(LCase$(rsc(1)), LCase$(Text1.Text)) Then    'search record
Label2.Caption = "found"
n = rsc.RecordCount
flag = 1
Else
n = n + 1
rsc.MoveNext   'move to next record
flag = 0
End If
Loop


If flag = 0 Then
Label2.Caption = " Not found sorry"


Else
Text3.Text = "0"
Text4.Text = "0"
Text5.Text = "0"


cn1.Execute "insert into studd  values ('" & Text8.Text & "' ,'" & Text1.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "')"
cnc.Execute "delete from cstudd  where CD_ID= '" & UCase$(Text8.Text) & "'"
'.Update
MsgBox ("Value updated")

End If
Else
MsgBox ("Incorrect values")
End If

Text1.Text = " "
Text4.Text = ""
Text5.Text = ""
Text2.Text = " "



Unload Form1
Unload Form2
Unload Form3
Unload Form4
Form5.Visible = True
Form5.Enabled = True



End Sub

Private Sub Command4_Click()
Label5.Caption = "take date"

Label6.Caption = "Return date"
Dim n As Integer
Dim flag As Integer
n = 0
rsc.MoveFirst


Do Until n = rsc.RecordCount
If InStr(UCase$(rsc(1)), UCase$(Text1.Text)) Then   'search record for out cd
Label2.Caption = "found"
Text8.Text = rsc(0)
n = rsc.RecordCount
Text3.Text = rsc(4)
Text4.Text = rsc(3)
Text5.Text = rsc(2)
ch2 = 1

flag = 1
Else
n = n + 1
rsc.MoveNext   'move to next record
flag = 0
End If
Loop

If flag = 0 Then
Label2.Caption = "Not found"
ch2 = 0
MsgBox ("Check if CD/DVD in store")

End If



End Sub

Private Sub Command5_Click()
DataReport2.Show

End Sub




Private Sub Text3_Gotfocus()
Text4.Text = "00"
Text5.Text = "00"
End Sub


'RETURN

Private Sub Command3_Click()
Timer4.Enabled = False


Dim n As Integer
If (ch = 1 And Text3.Text <> " " And Text2.Text <> " ") Then
rs1.MoveFirst

n = 1
Dim flag As Integer
flag = 0


Do Until n = rs.RecordCount
If InStr(UCase$(rs(0)), UCase$(Text3.Text)) Then    'search record
Label2.Caption = "found"
n = rs.RecordCount
flag = 1
Else
n = n + 1
rs.MoveNext   'move to next record
flag = 0
End If
Loop


If flag = 0 Then
Label2.Caption = " Not found sorry"


Else
rs1.MoveFirst


cn1.Execute "update  studd result set Cus_id= '" & Text3.Text & "'  where Movie_name='" & Text1.Text & "'"
cn1.Execute "update  studd result set Outdate= '" & Text4.Text & "'  where Movie_name='" & Text1.Text & "'"
cn1.Execute "update  studd result set Indate= '" & Text5.Text & "'  where Movie_name='" & Text1.Text & "'"

cnc.Execute "insert into cstudd  values ('" & UCase$(Text8.Text) & "' ,'" & Text1.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text3.Text & "')"
cn1.Execute "delete from studd  where CD_ID= '" & UCase$(Text8.Text) & "'"

MsgBox ("Record  updated")
rsc.Update
rs1.Update
Text2.Text = rs(1)
Print rs.RecordCount


End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text8.Text = ""

Else
MsgBox ("Values Are not correct")
End If
ch = 0

Form4.Visible = False
Form4.Enabled = False
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Form5.Visible = True


End Sub

Private Sub Form_Load() 'vb1 customer
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text8.Text = ""


cn.Open "tdsn", "system", "pratik"
rs.Open "select * from vb1", cn, adOpenStatic, adLockOptimistic

'studd cd/dvd
cn1.Open "tdsn", "system", "pratik"
rs1.Open "select * from studd", cn1, adOpenStatic, adLockOptimistic

'insert/delete cstudd
cnc.Open "tdsn", "system", "pratik"
rsc.Open "select * from cstudd", cn1, adOpenStatic, adLockOptimistic



End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
Label2.Caption = ""
End Sub

Private Sub Text3_Click()
Dim s1 As Variant
s1 = Left(Text2.Text, 3)
Dim mob As Variant
mob = Right(Text7.Text, 4) + Mid$(Text7.Text, 2, 2)
Text3.Text = s1 + mob
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
For i = 0 To 9 Step 1
Picture2(i).Visible = False
Next i
Timer2.Enabled = True



End Sub

Private Sub Timer2_Timer()
Dim i As Integer
For i = 0 To 9 Step 1
Picture2(i).Visible = True
Next i
Timer2.Enabled = False

End Sub

Private Sub Timer3_Timer()
Form4.BackColor = RGB(0, 256 * Rnd, 256 * Rnd)

End Sub

Private Sub Timer4_Timer()
Command3.BackColor = RGB(256 * Rnd, 0, 256 * Rnd)
Command2.BackColor = RGB(0, 256 * Rnd, 230 * Rnd)



End Sub

Private Sub Timer5_Timer()
Image1.Visible = True
Timer6.Enabled = True

End Sub

Private Sub Timer6_Timer()
Image1.Visible = False
Timer6.Enabled = False

End Sub

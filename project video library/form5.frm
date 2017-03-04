VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000007&
   Caption         =   "Form5"
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14415
   LinkTopic       =   "Form5"
   ScaleHeight     =   10170
   ScaleWidth      =   14415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "THANK YOU"
      DisabledPicture =   "form5.frx":0000
      DownPicture     =   "form5.frx":0F79
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2520
      MaskColor       =   &H000080FF&
      OLEDropMode     =   1  'Manual
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   7335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   12360
      Left            =   120
      Picture         =   "form5.frx":1EF2
      ScaleHeight     =   12300
      ScaleWidth      =   15915
      TabIndex        =   1
      Top             =   120
      Width           =   15975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

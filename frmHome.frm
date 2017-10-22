VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00400040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Home"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12225
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Image Image7 
      Height          =   630
      Left            =   11160
      Picture         =   "frmHome.frx":0000
      Top             =   240
      Width           =   630
   End
   Begin VB.Image Image6 
      Height          =   3135
      Left            =   480
      Picture         =   "frmHome.frx":6762
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Image Image5 
      Height          =   3135
      Left            =   4560
      Picture         =   "frmHome.frx":168DF
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Image Image4 
      Height          =   3135
      Left            =   8640
      Picture         =   "frmHome.frx":27572
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Image Image3 
      Height          =   3135
      Left            =   8640
      Picture         =   "frmHome.frx":3818B
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   3135
      Left            =   4560
      Picture         =   "frmHome.frx":49347
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   480
      Picture         =   "frmHome.frx":5A71A
      Top             =   1200
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    LoginCond = 1
    Login_load
    Form4.Show
    Form2.Hide
End Sub

Private Sub Image1_Click()
    Form1.Show
    Form2.Hide
End Sub

Private Sub Image2_Click()
    LoginCond = 0
    Login_load
    Form4.Show
    Form2.Hide
End Sub

Private Sub Image3_Click()
    LoginCond = 2
    Login_load
    Form4.Show
    Form2.Hide
End Sub

Private Sub Image4_Click()
    LoginCond = 5
    Login_load
    Form4.Show
    Form2.Hide
End Sub

Private Sub Image5_Click()
    LoginCond = 4
    Login_load
    Form4.Show
    Form2.Hide
End Sub

Private Sub Image6_Click()
    LoginCond = 3
    Login_load
    Form4.Show
    Form2.Hide
End Sub

Private Sub Image7_Click()
    Form2.Hide
    End
End Sub

VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Account List (Admin View)"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15285
   LinkTopic       =   "Form5"
   ScaleHeight     =   8115
   ScaleWidth      =   15285
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstBalance 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   6330
      ItemData        =   "frmAccList.frx":0000
      Left            =   2520
      List            =   "frmAccList.frx":0002
      TabIndex        =   16
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ListBox lstAccType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   6330
      ItemData        =   "frmAccList.frx":0004
      Left            =   1080
      List            =   "frmAccList.frx":0006
      TabIndex        =   15
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ListBox lstPhone 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   6330
      ItemData        =   "frmAccList.frx":0008
      Left            =   13440
      List            =   "frmAccList.frx":000A
      TabIndex        =   14
      Top             =   1800
      Width           =   1815
   End
   Begin VB.ListBox lstAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   6330
      ItemData        =   "frmAccList.frx":000C
      Left            =   9600
      List            =   "frmAccList.frx":000E
      TabIndex        =   13
      Top             =   1800
      Width           =   3855
   End
   Begin VB.ListBox lstGender 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   6330
      ItemData        =   "frmAccList.frx":0010
      Left            =   8400
      List            =   "frmAccList.frx":0012
      TabIndex        =   12
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox lstAge 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   6330
      ItemData        =   "frmAccList.frx":0014
      Left            =   7440
      List            =   "frmAccList.frx":0016
      TabIndex        =   11
      Top             =   1800
      Width           =   975
   End
   Begin VB.ListBox lstName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   6330
      ItemData        =   "frmAccList.frx":0018
      Left            =   4440
      List            =   "frmAccList.frx":001A
      TabIndex        =   10
      Top             =   1800
      Width           =   3015
   End
   Begin VB.ListBox lstAccNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   6330
      ItemData        =   "frmAccList.frx":001C
      Left            =   0
      List            =   "frmAccList.frx":001E
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.OptionButton optSavings 
      BackColor       =   &H00400040&
      Caption         =   "Savings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5040
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.OptionButton optCurrent 
      BackColor       =   &H00400040&
      Caption         =   "Current"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.OptionButton optAll 
      BackColor       =   &H00400040&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtListName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7440
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox txtListAccountNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Data dataSBI 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "Z:\ThirdB VB projects\batch 3\BaNkInG\project.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tab"
      Top             =   0
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15255
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00808080&
         Caption         =   "Close"
         Height          =   375
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Image btnHome 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   14040
         Picture         =   "frmAccList.frx":0020
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblAccountNoNotFound 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "A/c not Found"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   8040
         TabIndex        =   26
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name   :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   6360
         TabIndex        =   8
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "A/c Type    :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   600
         TabIndex        =   7
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Account No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   3000
      TabIndex        =   24
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   11280
      TabIndex        =   23
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   14160
      TabIndex        =   22
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   8760
      TabIndex        =   21
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Type"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   1440
      TabIndex        =   20
      Top             =   1560
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   7800
      TabIndex        =   19
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   5760
      TabIndex        =   18
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "A/c No"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   240
      TabIndex        =   17
      Top             =   1560
      Width           =   540
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnHome_Click()
    Unload Me
    Form2.Show
End Sub
Private Sub Form_Activate()
    lblAccountNoNotFound.Caption = ""
    txtListAccountNo.SetFocus
    dataSBI.Refresh
    dataSBI.Recordset.MoveFirst
    Do While Not dataSBI.Recordset.EOF
        Form5.lstName.AddItem (dataSBI.Recordset.Fields(0))
        Form5.lstAge.AddItem (dataSBI.Recordset.Fields(1))
        Form5.lstAddress.AddItem (dataSBI.Recordset.Fields(2))
        Form5.lstPhone.AddItem (dataSBI.Recordset.Fields(3))
        Form5.lstGender.AddItem (dataSBI.Recordset.Fields(4))
        Form5.lstAccNo.AddItem (dataSBI.Recordset.Fields(6))
        Form5.lstBalance.AddItem (dataSBI.Recordset.Fields(7))
        Form5.lstAccType.AddItem (dataSBI.Recordset.Fields(8))
        dataSBI.Recordset.MoveNext
    Loop
End Sub

Private Sub Form_Load()
    dataSBI.DatabaseName = App.Path & "\SBI Database.mdb"
    dataSBI.RecordSource = "SBI"
    dataSBI.Visible = False
    
    lblAccountNoNotFound.Caption = ""
End Sub


Private Sub optAll_Click()
txtListAccountNo.Text = ""
    lstName.Clear
    lstAge.Clear
    lstAddress.Clear
    lstPhone.Clear
    lstGender.Clear
    lstAccNo.Clear
    lstBalance.Clear
    lstAccType.Clear

dataSBI.Recordset.MoveFirst
Do While Not dataSBI.Recordset.EOF
 If optAll.Value = True Then
 
    lstName.AddItem (dataSBI.Recordset.Fields(0))
    lstAge.AddItem (dataSBI.Recordset.Fields(1))
    lstAddress.AddItem (dataSBI.Recordset.Fields(2))
    lstPhone.AddItem (dataSBI.Recordset.Fields(3))
    lstGender.AddItem (dataSBI.Recordset.Fields(4))
    lstAccNo.AddItem (dataSBI.Recordset.Fields(6))
    lstBalance.AddItem (dataSBI.Recordset.Fields(7))
    lstAccType.AddItem (dataSBI.Recordset.Fields(8))
    IAN = 1
  End If
dataSBI.Recordset.MoveNext
Loop

End Sub

Private Sub optCurrent_Click()
txtListAccountNo.Text = ""
    lstName.Clear
    lstAge.Clear
    lstAddress.Clear
    lstPhone.Clear
    lstGender.Clear
    lstAccNo.Clear
    lstBalance.Clear
    lstAccType.Clear

dataSBI.Recordset.MoveFirst
Do While Not dataSBI.Recordset.EOF
 If optCurrent.Value = True Then
    If dataSBI.Recordset.Fields(8) = "Current" Then
    
    lstName.AddItem (dataSBI.Recordset.Fields(0))
    lstAge.AddItem (dataSBI.Recordset.Fields(1))
    lstAddress.AddItem (dataSBI.Recordset.Fields(2))
    lstPhone.AddItem (dataSBI.Recordset.Fields(3))
    lstGender.AddItem (dataSBI.Recordset.Fields(4))
    lstAccNo.AddItem (dataSBI.Recordset.Fields(6))
    lstBalance.AddItem (dataSBI.Recordset.Fields(7))
    lstAccType.AddItem (dataSBI.Recordset.Fields(8))
    IAN = 1
  End If
  End If
dataSBI.Recordset.MoveNext
Loop

End Sub

Private Sub optSavings_Click()
txtListAccountNo.Text = ""
    lstName.Clear
    lstAge.Clear
    lstAddress.Clear
    lstPhone.Clear
    lstGender.Clear
    lstAccNo.Clear
    lstBalance.Clear
    lstAccType.Clear

dataSBI.Recordset.MoveFirst
Do While Not dataSBI.Recordset.EOF
 If optSavings.Value = True Then
    If dataSBI.Recordset.Fields(8) = "Savings" Then
    
    lstName.AddItem (dataSBI.Recordset.Fields(0))
    lstAge.AddItem (dataSBI.Recordset.Fields(1))
    lstAddress.AddItem (dataSBI.Recordset.Fields(2))
    lstPhone.AddItem (dataSBI.Recordset.Fields(3))
    lstGender.AddItem (dataSBI.Recordset.Fields(4))
    lstAccNo.AddItem (dataSBI.Recordset.Fields(6))
    lstBalance.AddItem (dataSBI.Recordset.Fields(7))
    lstAccType.AddItem (dataSBI.Recordset.Fields(8))
    IAN = 1
  End If
  End If
dataSBI.Recordset.MoveNext
Loop

End Sub

Private Sub txtListAccountNo_Change()
    lblAccountNoNotFound.Caption = ""
    lstName.Clear
    lstAge.Clear
    lstAddress.Clear
    lstPhone.Clear
    lstGender.Clear
    lstAccNo.Clear
    lstBalance.Clear
    lstAccType.Clear

Dim IAN As Integer  'Initial A/c No
Dim LAN As Integer  'Length of A/c No
Dim MAN As String   'Mid of A/c No
Dim TAN As String   'Trim of A/c No

IAN = 0
LAN = Len(txtListAccountNo.Text)
TAN = Trim(txtListAccountNo.Text)
dataSBI.Recordset.MoveFirst
Do While Not dataSBI.Recordset.EOF
MAN = Mid(dataSBI.Recordset.Fields(6), 1, LAN)
If MAN = TAN Then
 
    lstName.AddItem (dataSBI.Recordset.Fields(0))
    lstAge.AddItem (dataSBI.Recordset.Fields(1))
    lstAddress.AddItem (dataSBI.Recordset.Fields(2))
    lstPhone.AddItem (dataSBI.Recordset.Fields(3))
    lstGender.AddItem (dataSBI.Recordset.Fields(4))
    lstAccNo.AddItem (dataSBI.Recordset.Fields(6))
    lstBalance.AddItem (dataSBI.Recordset.Fields(7))
    lstAccType.AddItem (dataSBI.Recordset.Fields(8))
    IAN = 1

End If
dataSBI.Recordset.MoveNext
Loop

If IAN <> 1 Then
 lblAccountNoNotFound.Caption = "A/c not Found"
End If

End Sub

Private Sub txtListName_Change()
    'lblNameNotFound.Caption = ""
    lstName.Clear
    lstAge.Clear
    lstAddress.Clear
    lstPhone.Clear
    lstGender.Clear
    lstAccNo.Clear
    lstBalance.Clear
    lstAccType.Clear

Dim INa As Integer   'Initial Name
Dim LN As Integer  'Length of Name
Dim MN As String   'Mid of Name
Dim TN As String   'Trim of Name

INa = 0
LN = Len(txtListName.Text)
TN = Trim(txtListName.Text)
dataSBI.Recordset.MoveFirst
Do While Not dataSBI.Recordset.EOF
MN = Mid(dataSBI.Recordset.Fields(0), 1, LN)
If MN = TN Then
 
    lstName.AddItem (dataSBI.Recordset.Fields(0))
    lstAge.AddItem (dataSBI.Recordset.Fields(1))
    lstAddress.AddItem (dataSBI.Recordset.Fields(2))
    lstPhone.AddItem (dataSBI.Recordset.Fields(3))
    lstGender.AddItem (dataSBI.Recordset.Fields(4))
    lstAccNo.AddItem (dataSBI.Recordset.Fields(6))
    lstBalance.AddItem (dataSBI.Recordset.Fields(7))
    lstAccType.AddItem (dataSBI.Recordset.Fields(8))
    IAN = 1

End If
dataSBI.Recordset.MoveNext
Loop

End Sub


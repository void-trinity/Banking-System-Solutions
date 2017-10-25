VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H00400040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Account"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12225
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   5160
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtPhoneNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   720
         TabIndex        =   10
         Top             =   6360
         Width           =   3255
      End
      Begin VB.TextBox txtAcType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   720
         TabIndex        =   9
         Top             =   5280
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   720
         TabIndex        =   4
         Top             =   2280
         Width           =   5000
      End
      Begin VB.TextBox txtBalance 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   720
         TabIndex        =   3
         Top             =   3240
         Width           =   3255
      End
      Begin VB.TextBox txtAcNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   720
         TabIndex        =   2
         Top             =   4320
         Width           =   3255
      End
      Begin VB.Image Image2 
         Height          =   630
         Left            =   3960
         Picture         =   "frmDeleteAcc.frx":0000
         Top             =   7200
         Width           =   1875
      End
      Begin VB.Image Image1 
         Height          =   630
         Left            =   1200
         Picture         =   "frmDeleteAcc.frx":7467
         Top             =   7200
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   12
         Top             =   6120
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   11
         Top             =   5040
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "THE FOLLOWING ACCOUNT WILL BE DELETED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   720
         TabIndex        =   8
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   7
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   6
         Top             =   3000
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   5
         Top             =   4080
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   5160
      TabIndex        =   13
      Top             =   0
      Width           =   7095
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "THE ACCOUNT WAS DELETED. PRESS HOME BUTTON TO RETURN TO MAIN SCREEN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1455
         Left            =   840
         TabIndex        =   14
         Top             =   1560
         Width           =   5655
      End
   End
   Begin VB.Image btnHome 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   360
      Picture         =   "frmDeleteAcc.frx":ED6D
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   1005
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer

Private Sub btnHome_Click()
    Unload Me
    Form2.Show
End Sub

Private Sub Form_Load()
    dataSBI.DatabaseName = App.Path & "\SBI Database.mdb"
    dataSBI.RecordSource = "SBI"
    dataSBI.Visible = False
End Sub
Private Sub Form_Activate()
    Dim X As Integer
    X = 1
    Frame1.Visible = True
    Frame2.Visible = False
    dataSBI.Refresh
    dataSBI.Recordset.MoveFirst
    txtName.Enabled = False
    txtBalance.Enabled = False
    txtAcNo.Enabled = False
    txtAcType.Enabled = False
    txtPhoneNo.Enabled = False
    btnHome.Visible = False
    Do While (Not dataSBI.Recordset.EOF) And X = 1
        If dataSBI.Recordset.Fields(6) = AcNo Then
            txtName.Text = dataSBI.Recordset.Fields(0)
            txtAcType.Text = dataSBI.Recordset.Fields(8)
            txtPhoneNo.Text = dataSBI.Recordset.Fields(3)
            txtBalance.Text = dataSBI.Recordset.Fields(7)
            txtAcNo.Text = dataSBI.Recordset.Fields(6)
            X = 0
        End If
        dataSBI.Recordset.MoveNext
    Loop
End Sub
Private Sub Image1_Click()
    dataSBI.Recordset.MoveFirst
    X = 1
    dataSBI.Refresh
    Do While (Not dataSBI.Recordset.EOF) And X = 1
      
        If dataSBI.Recordset.Fields(6) = AcNo Then
            dataSBI.Recordset.Delete
            X = 0
        End If
    dataSBI.Recordset.MoveNext
    Loop
    Frame1.Visible = False
    Frame2.Visible = True
    btnHome.Visible = True
End Sub

Private Sub Image2_Click()
    Unload Me
    Form2.Show
End Sub

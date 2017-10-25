VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00400040&
   Caption         =   "View Details"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12075
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   8115
   ScaleWidth      =   12075
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data dataSBI 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "Z:\ThirdB VB projects\batch 3\BaNkInG\project.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Frame frmView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   5040
      TabIndex        =   0
      Top             =   -120
      Width           =   7095
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
         TabIndex        =   17
         Top             =   7320
         Width           =   3255
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
         TabIndex        =   15
         Top             =   6480
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
         TabIndex        =   14
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox txtGender 
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
         TabIndex        =   13
         Top             =   2160
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
         TabIndex        =   5
         Top             =   480
         Width           =   5000
      End
      Begin VB.TextBox txtAge 
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
         TabIndex        =   4
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtAddress 
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
         Height          =   800
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2880
         Width           =   5000
      End
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
         TabIndex        =   2
         Top             =   5640
         Width           =   3255
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   4800
         Width           =   5000
      End
      Begin VB.Label lblWarningView 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Valid Details"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   7800
         Width           =   4335
      End
      Begin VB.Image Image2 
         Height          =   630
         Left            =   6120
         Picture         =   "frmView.frx":0000
         Top             =   7320
         Width           =   630
      End
      Begin VB.Image Image1 
         Height          =   630
         Left            =   6120
         Picture         =   "frmView.frx":7C18
         Top             =   7320
         Width           =   630
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
         TabIndex        =   18
         Top             =   7080
         Width           =   1455
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
         TabIndex        =   16
         Top             =   6240
         Width           =   705
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
         TabIndex        =   12
         Top             =   3840
         Width           =   810
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
         TabIndex        =   11
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
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
         TabIndex        =   10
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         TabIndex        =   9
         Top             =   2640
         Width           =   720
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
         TabIndex        =   8
         Top             =   5400
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Top             =   4560
         Width           =   855
      End
   End
   Begin VB.Image btnHome 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   360
      Picture         =   "frmView.frx":F1D1
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1005
   End
   Begin VB.Label lblHead 
      BackStyle       =   0  'Transparent
      Caption         =   "View Details"
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
      Left            =   720
      TabIndex        =   19
      Top             =   1200
      Width           =   3495
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnHome_Click()
    Form2.Show
    Unload Me
End Sub

Private Sub Form_Activate()
    lblWarningView.Visible = False
    dataSBI.Refresh
    dataSBI.Recordset.MoveFirst
    txtName.Enabled = False
    txtAge.Enabled = False
    txtGender.Enabled = False
    txtAddress.Enabled = False
    txtAcType.Enabled = False
    txtPassword.Enabled = False
    txtPhoneNo.Enabled = False
    txtBalance.Enabled = False
    txtAcNo.Enabled = False
    Image1.Visible = True
    Image2.Visible = False
    btnHome.Visible = True
    Dim X As Integer
    X = 1
    Do While (Not dataSBI.Recordset.EOF) Or (X = 1)

        If AcNo = dataSBI.Recordset.Fields(6) Then
            txtName.Text = dataSBI.Recordset.Fields(0)
            txtAge.Text = dataSBI.Recordset.Fields(1)
            txtGender.Text = dataSBI.Recordset.Fields(4)
            txtAddress.Text = dataSBI.Recordset.Fields(2)
            txtAcType.Text = dataSBI.Recordset.Fields(8)
            txtPassword.Text = dataSBI.Recordset.Fields(5)
            txtPhoneNo.Text = dataSBI.Recordset.Fields(3)
            txtBalance.Text = dataSBI.Recordset.Fields(7)
            txtAcNo.Text = dataSBI.Recordset.Fields(6)
            X = 0
        End If
    dataSBI.Recordset.MoveNext
    Loop
End Sub

Private Sub Form_Load()
    dataSBI.DatabaseName = App.Path & "\SBI Database.mdb"
    dataSBI.RecordSource = "SBI"
    dataSBI.Visible = False
End Sub

Private Sub Image1_Click()
    txtName.Enabled = True
    txtAge.Enabled = True
    txtGender.Enabled = False
    txtAddress.Enabled = True
    txtAcType.Enabled = False
    txtPassword.Enabled = False
    txtPhoneNo.Enabled = True
    txtBalance.Enabled = False
    txtAcNo.Enabled = False
    Image1.Visible = False
    Image2.Visible = True
    btnHome.Visible = False
    dataSBI.Refresh
End Sub

Private Sub Image2_Click()
    If Len(txtName.Text) > 0 And IsNumeric(txtAge.Text) And Len(txtAddress.Text) > 0 And IsNumeric(txtPhoneNo.Text) And Len(txtPhoneNo.Text) = 10 Then
        X = 1
        dataSBI.Refresh
        dataSBI.Recordset.MoveFirst
        Image1.Visible = True
        Image2.Visible = False
        Do While (Not dataSBI.Recordset.EOF) Or (X = 1)
            If AcNo = dataSBI.Recordset.Fields(6) Then
                dataSBI.Recordset.Edit
                dataSBI.Recordset.Fields(0) = txtName.Text
                dataSBI.Recordset.Fields(1) = txtAge.Text
                dataSBI.Recordset.Fields(2) = txtAddress.Text
                dataSBI.Recordset.Fields(3) = txtPhoneNo.Text
                dataSBI.Recordset.Update
                X = 0
            End If
            dataSBI.Recordset.MoveNext
        Loop
        txtName.Enabled = False
        txtAge.Enabled = False
        txtGender.Enabled = False
        txtAddress.Enabled = False
        txtAcType.Enabled = False
        txtPassword.Enabled = False
        txtPhoneNo.Enabled = False
        txtBalance.Enabled = False
        txtAcNo.Enabled = False
        Image1.Visible = True
        Image2.Visible = False
        btnHome.Visible = True
        lblWarningView.Visible = False
    Else
        lblWarningView.Visible = True
    End If
End Sub

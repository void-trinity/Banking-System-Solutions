VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Account"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12225
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
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
   Begin VB.Frame frameAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   5160
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   29
         Top             =   6960
         Width           =   2775
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   720
         TabIndex        =   16
         Top             =   4680
         Width           =   2695
         Begin VB.OptionButton optSavings 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Savings"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   18
            Top             =   120
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optCurrent 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Current"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   720
         TabIndex        =   13
         Top             =   2400
         Width           =   2655
         Begin VB.OptionButton optFemale 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Female"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   15
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton optMale 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Male"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   120
            Value           =   -1  'True
            Width           =   1095
         End
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
         TabIndex        =   5
         Top             =   5520
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
         TabIndex        =   4
         Top             =   6360
         Width           =   3255
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
         Top             =   3360
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
         TabIndex        =   2
         Top             =   1440
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
         TabIndex        =   1
         Top             =   480
         Width           =   5000
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "All the fields are required. Fields maybe empty or invalid data is entered"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1080
         TabIndex        =   31
         Top             =   7680
         Width           =   5295
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
         TabIndex        =   12
         Top             =   5280
         Width           =   855
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
         TabIndex        =   11
         Top             =   2040
         Width           =   630
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
         TabIndex        =   10
         Top             =   6120
         Width           =   1305
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
         Top             =   3120
         Width           =   720
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
         TabIndex        =   8
         Top             =   1200
         Width           =   345
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
         Top             =   240
         Width           =   510
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
         TabIndex        =   6
         Top             =   4320
         Width           =   810
      End
   End
   Begin VB.Frame frameSuccess 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   5160
      TabIndex        =   35
      Top             =   0
      Width           =   7095
      Begin VB.Label lblSuceessMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Created Successfully. Press the home button to return to home page"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1935
         Left            =   600
         TabIndex        =   36
         Top             =   1560
         Width           =   6255
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   8295
      Left            =   5160
      TabIndex        =   19
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   28
         Top             =   6840
         Width           =   2535
      End
      Begin VB.TextBox txtOpenBal 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   23
         Top             =   5880
         Width           =   4500
      End
      Begin VB.TextBox txtAccType 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   22
         Top             =   4440
         Width           =   4500
      End
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   21
         Top             =   2880
         Width           =   4500
      End
      Begin VB.TextBox txtAccNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   20
         Top             =   1680
         Width           =   4500
      End
      Begin VB.Label lblAmtWarning 
         Caption         =   "Entered amount is invalid. Please Enter Valid amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   7560
         Width           =   3975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Successfully Created.               Enter the Opening Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   735
         Left            =   720
         TabIndex        =   33
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Please note the Account No and Password for future references"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         TabIndex        =   30
         Top             =   3600
         Width           =   4815
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Opening Balance (min 1000)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   5520
         Width           =   2655
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Account Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   26
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblNewAccountNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account No"
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
         TabIndex        =   24
         Top             =   1320
         Width           =   1020
      End
   End
   Begin VB.Image btnHome 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   240
      Picture         =   "frmCreateAcc.frx":0000
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1005
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Create Account"
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
      Left            =   360
      TabIndex        =   32
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnHome_Click()
    Unload Me
    Form2.Show
End Sub

Private Sub cmdSubmit_Click()
    If (Len(txtName.Text) > 0 And IsNumeric(txtAge.Text) And Len(txtAddress.Text) > 0 And Len(txtPassword.Text) > 0 And IsNumeric(txtPhoneNo.Text) And Len(txtPhoneNo.Text) = 10) Then
        Dim a As Double
        dataSBI.Recordset.AddNew
        dataSBI.Recordset.Fields(0) = txtName.Text
        dataSBI.Recordset.Fields(1) = txtAge.Text
        dataSBI.Recordset.Fields(2) = txtAddress.Text
        dataSBI.Recordset.Fields(3) = txtPhoneNo.Text
        dataSBI.Recordset.Fields(5) = txtPassword.Text
        
        If optMale.Value = True Then
         dataSBI.Recordset.Fields(4) = "Male"
        Else
         dataSBI.Recordset.Fields(4) = "Female"
        End If
        
        If optCurrent.Value = True Then
            dataSBI.Recordset.Fields(8) = "Current"
        Else
            dataSBI.Recordset.Fields(8) = "Savings"
        End If
        
        dataSBI.Recordset.Update
         
        dataSBI.Recordset.MoveLast
        dataSBI.Recordset.MovePrevious
        a = dataSBI.Recordset.Fields(6)
        dataSBI.Recordset.MoveNext
        dataSBI.Recordset.Edit
        dataSBI.Recordset.Fields(6) = a + 1
        dataSBI.Recordset.Update
        
        Frame4.Visible = True
        txtAccNo.Text = a + 1
        txtPass.Text = txtPassword.Text
        txtAccType = dataSBI.Recordset.Fields(8)
        frameAdd.Visible = False
        txtOpenBal.SetFocus
        lblAmtWarning.Visible = False
        btnHome.Visible = False
    Else
        lblWarning.Visible = True
    End If
End Sub

Private Sub cmdAdd_Click()
    If (IsNumeric(txtOpenBal.Text) And Val(txtOpenBal.Text) >= 1000) Then
        dataSBI.Recordset.MoveLast
        dataSBI.Recordset.Edit
        dataSBI.Recordset.Fields(7) = txtOpenBal.Text
        dataSBI.Recordset.Update
        dataSBI.Refresh
        Frame4.Visible = False
        frameSuccess.Visible = True
        btnHome.Visible = True
    Else
        lblAmtWarning.Visible = True
    End If
End Sub

Private Sub Form_Activate()
    txtName.SetFocus
End Sub

Private Sub Form_Load()
    dataSBI.DatabaseName = App.Path & "\SBI Database.mdb"
    dataSBI.RecordSource = "SBI"
    dataSBI.Visible = False
    Frame4.Visible = False
    lblWarning.Visible = False
    txtName.Text = ""
    txtAge.Text = ""
    txtAddress.Text = ""
    txtPhoneNo.Text = ""
    txtPassword.Text = ""
    frameSuccess.Visible = False
End Sub
Private Sub Form_Deactivate()
    frameAdd.Visible = True
    Frame4.Visible = False
    frameSuccess.Visible = False
    lblWarning.Visible = False
    txtName.Text = ""
    txtAge.Text = ""
    txtAddress.Text = ""
    txtPhoneNo.Text = ""
    txtPassword.Text = ""
    txtOpenBal.Text = ""
End Sub

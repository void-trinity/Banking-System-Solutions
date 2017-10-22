VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00400040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12225
   FillColor       =   &H00400040&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data dataSBI 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\EDUCATION\BE CSE AC09UCS099\Sem-5\SE\Output\batch 3\BaNkInG\project.mdb"
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
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   5160
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   6
         Top             =   5640
         Width           =   1815
      End
      Begin VB.TextBox txtPassword 
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
         TabIndex        =   3
         Top             =   3000
         Width           =   4500
      End
      Begin VB.TextBox txtAccNo 
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
         TabIndex        =   2
         Top             =   1560
         Width           =   4500
      End
      Begin VB.Label lblLoginWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   6600
         Width           =   3135
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Image btnHome 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   120
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   1005
   End
   Begin VB.Label lblHeader 
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
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnHome_Click()
    Form2.Show
    Form4.Hide
End Sub

Private Sub Command1_Click()
    If txtAccNo.Text = "" And txtPassword.Text = "" Then
    lblLoginWarning.Caption = "Enter the Acc No and Password"
    txtAccNo.SetFocus
    
    ElseIf txtAccNo.Text = "" Then
        lblLoginWarning.Caption = "Enter the Acc No"
        txtAccNo.SetFocus
        
    ElseIf txtPassword.Text = "" Then
        lblLoginWarning.Caption = "Enter the Password"
        txtPassword.SetFocus
        
    ElseIf txtAccNo.Text = "admin" And txtPassword.Text = "admin" And LoginCond = 1 Then
            Form5.Show
            Login_load
            
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
    Else

        Dim X As Integer
        Dim Login As String
        Dim Password As String
        
        X = 0
        
        Login = Trim(txtAccNo.Text)         'Account No
        Password = Trim(txtPassword.Text)   'Password
        
        dataSBI.Recordset.MoveFirst
            
        Do While Not dataSBI.Recordset.EOF
        
        If Login = dataSBI.Recordset.Fields(6) And Password = dataSBI.Recordset.Fields(5) Then
            X = 1
            If LoginCond = 0 Then
                Form3.Show
                Login_load
                
            ElseIf LoginCond = 2 Then
                Form6.Show
                Login_load
            
            ElseIf LoginCond = 3 Then
                Form7.Show
                Login_load
            
            ElseIf LoginCond = 4 Then
                Form8.Show
                Login_load
            
            ElseIf LoginCond = 5 Then
                Form9.Show
                Login_load
            End If
        End If
        dataSBI.Recordset.MoveNext
        Loop
        
        If X <> 1 Then
         lblLoginWarning.Caption = "Invalid Acc.No and Password,Try Again"
         txtAccNo.Text = ""
         txtPassword.Text = ""
         txtAccNo.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    dataSBI.Refresh
    txtAccNo.SetFocus
    If LoginCond = 0 Then
        lblHeader.Caption = "Delete Account"
    ElseIf LoginCond = 2 Then
        lblHeader.Caption = "View Details"
    ElseIf LoginCond = 3 Then
        lblHeader.Caption = "Deposit"
    ElseIf LoginCond = 4 Then
        lblHeader.Caption = "Withdraw"
    ElseIf LoginCond = 5 Then
        lblHeader.Caption = "Transfer Money"
    ElseIf LoginCond = 1 Then
        lblHeader.Caption = "Admin Login"
    End If
    
End Sub

Private Sub Form_Load()
    dataSBI.DatabaseName = App.Path & "\SBI Database.mdb"
    dataSBI.RecordSource = "SBI"
    dataSBI.Visible = False
End Sub

VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00400040&
   Caption         =   "Transfer"
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
   LinkTopic       =   "Form9"
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
      Left            =   5040
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtAmtTrans 
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
         TabIndex        =   9
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox txtCurr 
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
         Top             =   2520
         Width           =   2655
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Transfer"
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
         TabIndex        =   3
         Top             =   5880
         Width           =   2775
      End
      Begin VB.TextBox txtAcNo 
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
         TabIndex        =   2
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtDestAccNo 
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
         TabIndex        =   1
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Top             =   4560
         Width           =   60
      End
      Begin VB.Label lblWarningAmt 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Valid Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   6600
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Balance"
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
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source Account Number"
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
         Top             =   1080
         Width           =   2130
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination Account Number"
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
         Top             =   3360
         Width           =   2490
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   5040
      TabIndex        =   11
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtUpdatedAmt 
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
         TabIndex        =   12
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Money Successfully Transferred"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1575
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Updated Amount in Source"
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
         TabIndex        =   13
         Top             =   3240
         Width           =   2340
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer"
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
      Left            =   1200
      TabIndex        =   15
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Image btnHome 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   240
      Picture         =   "frmTransfer.frx":0000
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1005
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim amt As Long
Dim X As Integer
Dim max As Long
Dim DestAcNo As Long
Dim j As Long
Private Sub btnHome_Click()
    Unload Me
    Form2.Show
End Sub

Private Sub cmdSubmit_Click()
    If IsNumeric(txtAmtTrans.Text) And Val(txtAmtTrans.Text) <= max And Val(txtAmtTrans.Text) > 0 Then
        X = 1
        dataSBI.Refresh
        dataSBI.Recordset.MoveFirst
        DestAcNo = Val(txtDestAccNo.Text)
        Do While ((Not dataSBI.Recordset.EOF) And X = 1)
            If dataSBI.Recordset.Fields(6) = Val(txtDestAccNo.Text) Then
                dataSBI.Recordset.Edit
                j = dataSBI.Recordset.Fields(7)
                dataSBI.Recordset.Fields(7) = j + Val(txtAmtTrans.Text)
                dataSBI.Recordset.Update
                X = 0
            End If
            dataSBI.Recordset.MoveNext
        Loop
        dataSBI.Refresh
        dataSBI.Recordset.MoveFirst
        If X = 0 Then
            X = 1
            Dim i As Long
            i = amt - Val(txtAmtTrans.Text)
            Do While (Not dataSBI.Recordset.EOF) And X = 1
                If dataSBI.Recordset.Fields(6) = AcNo Then
                    dataSBI.Recordset.Edit
                    dataSBI.Recordset.Fields(7) = i
                    dataSBI.Recordset.Update
                    X = 0
                End If
                dataSBI.Recordset.MoveNext
            Loop
            Frame1.Visible = False
            Frame2.Visible = True
            txtUpdatedAmt.Text = i
            txtUpdatedAmt.Enabled = False
        Else
            lblWarningAmt.Caption = "Enter valid destination account number"
            lblWarningAmt.Visible = True
        End If
    Else
        lblWarningAmt.Visible = True
        lblWarningAmt.Caption = "Enter valid amount"
    End If
End Sub

Private Sub Form_Activate()
    Frame1.Visible = True
    Frame2.Visible = False
    lblWarningAmt.Visible = False
    dataSBI.Refresh
    dataSBI.Recordset.MoveFirst
    X = 1
    Do While (Not dataSBI.Recordset.EOF) And X = 1
        If dataSBI.Recordset.Fields(6) = AcNo Then
            txtAcNo.Text = AcNo
            txtCurr = dataSBI.Recordset.Fields(7)
            amt = Val(txtCurr.Text)
            max = IIf(amt / 2 > 50000, 50000, amt / 2)
            Label4.Caption = "Amount to be Transferred (max " & max & ")"
            X = 0
        End If
        dataSBI.Recordset.MoveNext
    Loop
    txtAmtTrans.Text = ""
    txtDestAccNo.SetFocus
    txtAcNo.Enabled = False
    txtCurr.Enabled = False
End Sub

Private Sub Form_Load()
    dataSBI.DatabaseName = App.Path & "\SBI Database.mdb"
    dataSBI.RecordSource = "SBI"
    dataSBI.Visible = False
End Sub



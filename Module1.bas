Attribute VB_Name = "Module1"
Public MaxAmount As Long
Public LoginCond As Long
Public AcNo As Long
Public Sub Login_load()
    Form4.txtAccNo.Text = ""
    Form4.txtPassword.Text = ""
    Form4.lblLoginWarning.Caption = ""
    Form4.lblHeader.Caption = ""
End Sub

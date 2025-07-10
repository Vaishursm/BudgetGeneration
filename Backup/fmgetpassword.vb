Public Class fmgetpassword
    Dim oForm As New Form
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        mPassword = Me.txtPassword.Text
        Me.Close()
        frmProjectDetails.lblMessage.Text = "Click " & Chr(34) & "Save and Proceed " & Chr(34) & "to continue"
        frmProjectDetails.lblMessage.Visible = True
        frmProjectDetails.Refresh()
        'oForm = New frmOptions()
        ' oForm.ShowDialog()
        'oForm = Nothing
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Me.txtPassword.Text = ""
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub fmgetpassword_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Label2.Text = "Remember the password with exact case of letters. " & vbNewLine & "You will need the same later when you open the prject file"
    End Sub
End Class
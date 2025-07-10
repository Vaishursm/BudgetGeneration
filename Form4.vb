Public Class Form4

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblMsgbeforeClose.Text = "You have opted to close the application without saving any updated data." & vbNewLine & _
        "Are you sure ?"
    End Sub

    Private Sub btnYes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnYes.Click
        answer = vbYes
        Me.Close()
    End Sub

    Private Sub btnNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNo.Click
        answer = vbNo
        Me.Close()
    End Sub
End Class
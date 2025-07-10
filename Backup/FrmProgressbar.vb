Public Class frmProgressbar

    Private Sub ProgressBar1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProgressBar1.Click

    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    <STAThread()> _
Shared Sub Main()
        Application.Run(New frmProgressbar)
    End Sub 'Main
End Class
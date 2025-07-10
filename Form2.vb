Public Class Form2A

    'Private Sub Form2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
    '    Dim workbook1 As Microsoft.Office.Interop.Excel.Workbook
    '    workbook1 = xlApp.Workbooks.Open("D:\Users\RSM\Desktop\Book1.xls")
    '    'workbook1.LoadFromFile("D:\Users\RSM\Desktop")
    '    Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet = workbook1.Worksheets("Sheet3")
    '    'add worksheets and name them
    '    'MsgBox(worksheet.Name)
    '    Dim j As Integer
    '    'For j = 1 To workbook1.Worksheets.Count
    '    '    worksheet = workbook1.Worksheets(j)
    '    '    MsgBox(worksheet.Name)
    '    'Next
    '    worksheet.Copy(Before:=workbook1.Worksheets("Sheet3"))
    '    Dim inti As Integer = worksheet.Index
    '    'MsgBox(worksheet.Name)
    '    inti = inti - 1
    '    worksheet = workbook1.Worksheets(inti)
    '    'MsgBox(worksheet.Index & "....." & worksheet.Name)
    '    worksheet.Name = "NewSheet"
    '    worksheet.Visible = True
    '    'worksheet = workbook1.Worksheets.Add(, workbook1.Worksheets("Sheet3"))
    '    'worksheet.Name = "NewSheet"
    '    ' workbook1.Worksheets.Add()
    '    'copy worksheet to the new added worksheets

    '    'workbook1.Worksheets(2).CopyFrom(workbook1.Worksheets(1))
    '    workbook1.Save()
    '    workbook1.Close()
    '    xlApp.Quit()
    '    xlApp = Nothing
    '    MsgBox("Done")
    '    'System.Diagnostics.Process.Start("D:\Users\RSM\Desktop")
    'End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblMsgBeforeSave.Text = "You have started saving the details in  the workbook." & vbNewLine & _
              " After saving you can only close the application. " & _
              vbNewLine & "Saving the details in the workbook will take few minutes." & vbNewLine & _
              "Are you sure to start saving?"
    End Sub
    'Sub s()

    '    VP = Me.Top + 20
    '    HP = Me.Left + 2
    '    HP = 10
    '    Dim pnlTitle As New Panel
    '    Dim x As New TextBox
    '    pnlTitle.Enabled = True
    '    pnlTitle.BringToFront()
    '    x.Text = "sample"

    '    pnlTitle.BorderStyle = BorderStyle.FixedSingle
    '    pnlTitle.Controls.Add(x)





    '    pnlTitle.Left = HP
    '    pnlTitle.Top = VP
    '    pnlTitle.Width = 4000 'Me.tbcBudgetHeads.Width - 5
    '    pnlTitle.Height = 3000
    '    pnlTitle.Location = New System.Drawing.Point(10, 50)
    '    pnlTitle.Name = "pnlTitle"
    '    Dim L1 As New Label, L2 As New Label
    '    L1.Text = "Category"
    '    L1.Font = New Font("vedana", 8)
    '    L1.Left = HP
    '    L1.Top = VP
    '    L1.TextAlign = ContentAlignment.MiddleCenter
    '    L1.Width = 150
    '    L1.Height = 100
    '    L1.BorderStyle = BorderStyle.FixedSingle
    '    pnlTitle.Controls.Add(L1)
    '    HP = HP + L1.Width + 10

    '    L2.Text = "Equip Name"
    '    L2.Font = New Font("vedana", 8)
    '    L2.Left = HP
    '    L2.Top = VP
    '    L2.TextAlign = ContentAlignment.MiddleCenter
    '    L2.Width = 110
    '    L2.Height = 30
    '    L2.BorderStyle = BorderStyle.FixedSingle
    '    pnlTitle.Controls.Add(L2)
    '    HP = HP + L2.Width + 10
    '    pnlTitle.Visible = True
    '    Me.Controls.Add(pnlTitle)

    '    ' Me.Show()


    'End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        answer = vbYes
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        answer = vbNo
        Me.Close()
    End Sub
End Class
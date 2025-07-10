Imports System.Data
Imports System.Data.OleDb
Public Class AddedItems_new
    Public strConnection As String
    Public Binding As BindingSource = New BindingSource()

    Private mAdapter As OleDb.OleDbDataAdapter
    Private mDataset As New DataSet

    Public moledbConnection As OleDbConnection
    Dim strStatement As String
    Dim moledbCommand As OleDbCommand
    Dim mOledbDataAdapter As OleDbDataAdapter
    Private Sub AddedItems_new_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim mcategory As String, addString As String
        mcategory = SelectedCategory
        Me.lblTitle.Text = mcategory
        Dim intI As Integer, items As Integer = CategoryTextBoxes.Count
        If mcategory <> "Lighting" Then
            For intI = 0 To items - 1
                addString = ""
                Dim j As Integer = 1
                If Checkboxes(intI).Checked Then
                    'Dim lvs As New ListViewItem.ListViewSubItem
                    Dim lvi As New ListViewItem
                    lvi.SubItems.Add(intI)
                    lvi.SubItems.Add(EquipNameTextBoxes(intI).Text)
                    lvi.SubItems.Add(CapacityTextBoxes(intI).Text)
                    lvi.SubItems.Add(MakeModelTextBoxes(intI).Text)
                    lvi.SubItems.Add(QtyTextBoxes(intI).Text)
                    lvi.SubItems.Add(HrsPermonthTextBoxes(intI).Text)
                    lvi.SubItems.Add(MobdatePickers(intI).Text)
                    lvi.SubItems.Add(DemobDatePickers(intI).Text)
                    lvi.SubItems.Add(DepPercComboboxes(intI).Text)
                    ListView1.Items.Add(lvi)

                    ' & Space(30 - Len(EquipNameTextBoxes(intI).Text)))
                    'lvi.
                    'lvs.Text = EquipNameTextBoxes(intI).Text & Space(30 - Len(EquipNameTextBoxes(intI).Text))
                    'lvi.SubItems.Add(lvs)
                    ' ListView1.Items.Add(lvi)
                    ListView1.View = View.Details


                    ListView1.Refresh()
                End If
            Next
        Else
            items = EquipNameTextBoxes.Count
            For intI = 0 To items - 1
                addString = ""
                Dim j As Integer = 1
                If Checkboxes(intI).Checked Then
                    'Dim lvs As New ListViewItem.ListViewSubItem
                    Dim lvi As New ListViewItem
                    lvi.SubItems.Add(intI)
                    lvi.SubItems.Add(EquipNameTextBoxes(intI).Text)
                    lvi.SubItems.Add(CapacityTextBoxes(intI).Text)
                    lvi.SubItems.Add(MakeModelTextBoxes(intI).Text)
                    lvi.SubItems.Add(MobdatePickers(intI).Text)
                    lvi.SubItems.Add(DemobDatePickers(intI).Text)
                    lvi.SubItems.Add(QtyTextBoxes(intI).Text)
                    lvi.SubItems.Add(PowerPerUnitTextBoxes(intI).Text)
                    lvi.SubItems.Add(ConnectLoadTextBoxes(intI).Text)
                    lvi.SubItems.Add(UtilityFactorTextBoxes(intI).Text)
                    ListView1.Items.Add(lvi)

                    ListView1.View = View.Details
                    ListView1.Refresh()
                End If
            Next
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class
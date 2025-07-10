Imports System.Data
Imports System.Data.OleDb
Public Class AddedItems
    'Public xlApp As New Microsoft.Office.Interop.Excel.Application
    'Public xlWorkbook As Microsoft.Office.Interop.Excel.Workbook
    'Public xlWorksheet As Microsoft.Office.Interop.Excel.Worksheet
    'Public xlRange As Microsoft.Office.Interop.Excel.Range
    Public strConnection As String
    Public Binding As BindingSource = New BindingSource()

    Private mAdapter As OleDb.OleDbDataAdapter
    Private mDataset As New DataSet

    Public moledbConnection As OleDbConnection
    Dim strStatement As String
    Dim moledbCommand As OleDbCommand
    Dim mOledbDataAdapter As OleDbDataAdapter
    'Dim mReader As OleDbDataReader
    'Dim mDataSet As DataSet
    Private Sub AddedItems_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        mcategory = SelectedCategory
        Tablename = GetTablename(mcategory)
        Dim strSelect As String = "select * from " & Tablename
        mAdapter = New OleDb.OleDbDataAdapter(strSelect, moledbConnection1)
        mDataset = New DataSet
        mAdapter.Fill(mDataset, mcategory)
        If mDataset.Tables(mcategory).Rows.Count = 0 Then
            MsgBox("NO RECORDS TO DISPLAY")
            Me.btnClose_Click(Me, e)
            Exit Sub
        End If
        Me.DataGridView1.DataSource = mDataset.Tables(mcategory)
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        If SelectedCategory = "External Others" Then
            EditOperationSheet = "External Others"
        ElseIf SelectedCategory = "Hired Equipments" Or SelectedCategory = "Conveyance_Hired" Then
            EditOperationSheet = "external Hire"
        ElseIf SelectedCategory = "Minor Equipments" Then
            EditOperationSheet = "Minor Eqpts"
        Else
            EditOperationSheet = SelectedCategory
        End If
        mAdapter = Nothing
        mDataset.Clear()
        mDataset = Nothing
        Me.Close()
        frmOptions.Show()
        frmOptions.Focus()
        If EditOrDelete = "Edit" Then
            If Not (SelectedCategory = "Conveyance_Hired" Or SelectedCategory = "HiredEquipments" Or SelectedCategory = "External Others") Then
                MsgBox("To update the Representation value, Select the  Model in the selection combo control and press Tab")
            End If
        End If
        'frmOptions.cmbMajorEquipModel_SelectedIndexChanged(frmOptions, e)
    End Sub

    Private Sub btnMajorConvItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMajorConvItems.Click
        mcategory = SelectedCategory
        EditOperationSheet = "Conveyance"
        Tablename = "MajorConveyanceEquips"
        Dim strSelect As String = "select * from " & Tablename
        mAdapter = New OleDb.OleDbDataAdapter(strSelect, moledbConnection1)
        mDataset = New DataSet
        mAdapter.Fill(mDataset, mcategory)
        If mDataset.Tables(mcategory).Rows.Count = 0 Then
            MsgBox("NO RECORDS TO DISPLAY")
            Exit Sub
        End If
        Me.DataGridView1.DataSource = mDataset.Tables(mcategory)
    End Sub

    Private Sub btnMajorConcreteItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMajorConcreteItems.Click
        mcategory = SelectedCategory
        EditOperationSheet = "Concreting"
        Tablename = "MajorConcreteEquips"
        Dim strSelect As String = "select * from " & Tablename
        mAdapter = New OleDb.OleDbDataAdapter(strSelect, moledbConnection1)
        mDataset = New DataSet
        mAdapter.Fill(mDataset, mcategory)
        If mDataset.Tables(mcategory).Rows.Count = 0 Then
            MsgBox("NO RECORDS TO DISPLAY")
            Exit Sub
        End If
        Me.DataGridView1.DataSource = mDataset.Tables(mcategory)
    End Sub

    Private Sub btnMajorCraneItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMajorCraneItems.Click
        mcategory = SelectedCategory
        EditOperationSheet = "Cranes"
        Tablename = "MajorCraneEquips"
        Dim strSelect As String = "select * from " & Tablename
        mAdapter = New OleDb.OleDbDataAdapter(strSelect, moledbConnection1)
        mDataset = New DataSet
        mAdapter.Fill(mDataset, mcategory)
        If mDataset.Tables(mcategory).Rows.Count = 0 Then
            MsgBox("NO RECORDS TO DISPLAY")
            Exit Sub
        End If
        Me.DataGridView1.DataSource = mDataset.Tables(mcategory)
    End Sub

    Private Sub btnMajorMHItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMajorMHItems.Click
        mcategory = SelectedCategory
        EditOperationSheet = "Material Handling"
        Tablename = "MajorMaterialhandlingEquips"
        Dim strSelect As String = "select * from " & Tablename
        mAdapter = New OleDb.OleDbDataAdapter(strSelect, moledbConnection1)
        mDataset = New DataSet
        mAdapter.Fill(mDataset, mcategory)
        If mDataset.Tables(mcategory).Rows.Count = 0 Then
            MsgBox("NO RECORDS TO DISPLAY")
            Exit Sub
        End If
        Me.DataGridView1.DataSource = mDataset.Tables(mcategory)
    End Sub

    Private Sub btnMajorNCItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMajorNCItems.Click
        mcategory = SelectedCategory
        EditOperationSheet = "Non Concreting"
        Tablename = "MajornonconcreteEquips"
        Dim strSelect As String = "select * from " & Tablename
        mAdapter = New OleDb.OleDbDataAdapter(strSelect, moledbConnection1)
        mDataset = New DataSet
        mAdapter.Fill(mDataset, mcategory)
        If mDataset.Tables(mcategory).Rows.Count = 0 Then
            MsgBox("NO RECORDS TO DISPLAY")
            Exit Sub
        End If
        Me.DataGridView1.DataSource = mDataset.Tables(mcategory)
    End Sub

    Private Sub btnMajorDGItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMajorDGItems.Click
        mcategory = SelectedCategory
        EditOperationSheet = "DG Sets"
        Tablename = "MajorDGSetsEquips"
        Dim strSelect As String = "select * from " & Tablename
        mAdapter = New OleDb.OleDbDataAdapter(strSelect, moledbConnection1)
        mDataset = New DataSet
        mAdapter.Fill(mDataset, mcategory)
        If mDataset.Tables(mcategory).Rows.Count = 0 Then
            MsgBox("NO RECORDS TO DISPLAY")
            Exit Sub
        End If
        Me.DataGridView1.DataSource = mDataset.Tables(mcategory)
    End Sub

    Private Sub btnHiredItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHiredItems.Click
        mcategory = SelectedCategory
        EditOperationSheet = "external Hire"
        Tablename = "HiredEquips"
        Dim strSelect As String = "select * from " & Tablename
        mAdapter = New OleDb.OleDbDataAdapter(strSelect, moledbConnection1)
        mDataset = New DataSet
        mAdapter.Fill(mDataset, mcategory)
        If mDataset.Tables(mcategory).Rows.Count = 0 Then
            MsgBox("NO RECORDS TO DISPLAY")
            Exit Sub
        End If
        Me.DataGridView1.DataSource = mDataset.Tables(mcategory)
    End Sub

    Private Sub btnMinorItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMinorItems.Click
        mcategory = SelectedCategory
        EditOperationSheet = "Minor Eqpts"
        Tablename = "MinorEquips"
        Dim strSelect As String = "select * from " & Tablename
        mAdapter = New OleDb.OleDbDataAdapter(strSelect, moledbConnection1)
        mDataset = New DataSet
        mAdapter.Fill(mDataset, mcategory)
        If mDataset.Tables(mcategory).Rows.Count = 0 Then
            MsgBox("NO RECORDS TO DISPLAY")
            Exit Sub
        End If

        Me.DataGridView1.DataSource = mDataset.Tables(mcategory)
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim row As DataGridViewRow
        Dim xldescription, xlcapacity, xlmakemodel As String
        Dim xlmobdate, xldemobdate As Date
        Dim xlqty As Integer
        row = Me.DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex)
        Dim sheetno, x As Integer, cols As Integer
        EditOrDelete = "Edit"
        mDescription = row.Cells("Description").Value
        mCapacity = row.Cells("Capacity").Value
        mMake = row.Cells("Make").Value
        mModel = row.Cells("Model").Value
        mMobDate = row.Cells("MobDate").Value
        mDeMobDate = row.Cells("DemobDate").Value
        mQty = row.Cells("Qty").Value
        mMonths = System.Math.Round((mDeMobDate.Date - mMobDate.Date).Days / 30, 0)
        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With
        Dim intI As Integer = 2
        If SelectedCategory = "External Others" Then
            EditOperationSheet = "External Others"
        ElseIf SelectedCategory = "Hired Equipments" Or SelectedCategory = "Conveyance_Hired" Then
            EditOperationSheet = "external Hire"
        ElseIf SelectedCategory = "Minor Equipments" Then
            EditOperationSheet = "Minor Eqpts"
        Else
            EditOperationSheet = SelectedCategory
        End If


        xlWorksheet = xlWorkbook.Sheets.Item(EditOperationSheet)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "SlNo")
        xlRange = xlRange.Offset(1, 0)

        While (True)
            xlRange = xlRange.Offset(0, 1)
            'MsgBox(xlRange.Value)
            xldescription = xlRange.Value
            xlRange = xlRange.Offset(0, 2)
            'MsgBox(xlRange.Value)
            xlcapacity = xlRange.Value
            xlRange = xlRange.Offset(0, 1)
            'MsgBox(xlRange.Value)
            xlmakemodel = xlRange.Value
            xlRange = xlRange.Offset(0, 1)
            'MsgBox(xlRange.Value)
            xlqty = xlRange.Value
            xlRange = xlRange.Offset(0, 1)
            'MsgBox(xlRange.Value)
            xlmobdate = xlRange.Value
            xlRange = xlRange.Offset(0, 1)
            'MsgBox(xlRange.Value)
            xldemobdate = xlRange.Value
            Dim xx As String = mMake & vbNewLine & "/" & mModel
            If (xldescription = mDescription And xlcapacity = mCapacity And _
                xlmakemodel = xx And xlqty = mQty And _
                xlmobdate = mMobDate And xldemobdate = mDeMobDate) Then
                xlRange.Select()
                If RecordsInserted(sheetno) = 1 Then
                    If Category_Shortname = "Concrete_" Then
                        cols = 34
                    Else
                        cols = 33
                    End If
                    xlRange = xlWorksheet.Range(Category_Shortname & "Slno").Offset(1, 0)
                    For intI = 1 To cols
                        'MsgBox(xlRange.Address & "......." & xlRange.Formula)
                        If Not xlRange.HasFormula Then
                            xlRange.Value = ""
                        End If
                        xlRange = xlRange.Offset(0, 1)
                    Next
                    RecordsInserted(sheetno) = RecordsInserted(sheetno) - 1
                    xlRange = xlRange.Offset(1, -cols)
                    xlRange = xlRange.Offset(0, 1)
                    If xlRange.Value = "Total" Then
                        xlRange.Select()
                        MsgBox(xlRange.Application.ActiveCell.Address)
                        xlRange.Application.ActiveCell.EntireRow.Delete()
                    End If
                    xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
                    xlRange.Value = RecordsInserted(sheetno)
                    xlApp.CalculateBeforeSave = True
                    xlWorkbook.Save()
                    xlWorkbook.Close()
                    xlWorkbook = Nothing
                    With xlApp
                        If xlWorkbook Is Nothing Then
                            xlWorkbook = .Workbooks.Open(xlFilename)
                        End If
                        xlWorksheet = xlWorkbook.Sheets.Item(EditOperationSheet)
                        xlWorksheet.Activate()
                        getCategoryShortname(xlWorksheet)
                    End With
                    Exit While
                Else
                    'MsgBox(xlRange.Address)
                    'MsgBox(xlApp.ActiveCell.Address & "....." & xlApp.ActiveCell.Value)
                    xlRange.Application.ActiveCell.EntireRow.Delete()
                    xlApp.CalculateBeforeSave = True
                    xlWorkbook.Save()
                    xlWorkbook.Close()
                    xlWorkbook = Nothing
                    With xlApp
                        If xlWorkbook Is Nothing Then
                            xlWorkbook = .Workbooks.Open(xlFilename)
                        End If
                    End With
                    'MsgBox(xlWorkbook.Name)
                    xlWorksheet = xlWorkbook.Sheets.Item(EditOperationSheet)
                    sheetno = getSheetNo(xlWorksheet.Name)
                    RecordsInserted(sheetno) = RecordsInserted(sheetno) - 1

                    xlRange = xlWorksheet.Range(Category_Shortname & "Slno").Offset(1, 0)
                    xlRange = xlRange.Offset(0, 1)
                    If xlRange.Value = "Total" Then
                        xlRange.Select()
                        MsgBox(xlRange.Application.ActiveCell.Address)
                        xlRange.Application.ActiveCell.EntireRow.Delete()
                    End If

                    xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
                    xlRange.Value = RecordsInserted(sheetno)
                    xlApp.CalculateBeforeSave = True
                    xlWorkbook.Save()
                    xlWorkbook.Close()
                    xlWorkbook = Nothing
                    With xlApp
                        If xlWorkbook Is Nothing Then
                            xlWorkbook = .Workbooks.Open(xlFilename)
                        End If
                        xlWorksheet = xlWorkbook.Sheets.Item(EditOperationSheet)
                        xlWorksheet.Activate()
                        getCategoryShortname(xlWorksheet)
                    End With
                    'xlRange = xlWorksheet.Range(Category_Shortname & "SlNo")
                    'xlRange = xlRange.Offset(1, 0)
                    ''MsgBox(xlRange.Value)
                    'xlRange = xlRange.Offset(0, 1)
                    'While (True)
                    '    If xlRange.Value = "Total" Then
                    '        xlRange.Select()
                    '        xlApp.ActiveCell.EntireRow.Delete()
                    '        xlWorkbook.Save()
                    '        xlWorkbook.Close()
                    '        xlWorkbook = Nothing
                    '        With xlApp
                    '            If xlWorkbook Is Nothing Then
                    '                xlWorkbook = .Workbooks.Open(xlFilename)
                    '            End If
                    '        End With
                    '        'MsgBox(xlWorkbook.Name)
                    '        xlWorksheet = xlWorkbook.Sheets.Item(EditOperationSheet)
                    '        Exit While
                    '    Else
                    '        xlRange = xlRange.Offset(1, 0)
                    '    End If
                    'End While
                    x = 1
                    xlRange = xlWorksheet.Range(Category_Shortname & "SlNo")
                    'MsgBox(xlRange.Value)
                    xlRange = xlRange.Offset(1, 0)
                    'MsgBox(xlRange.Value)
                    While (True)
                        If Not Len(Trim(xlRange.Value)) = 0 Then
                            xlRange.Value = x
                            x = x + 1
                            xlRange = xlRange.Offset(1, 0)
                        Else
                            Exit While
                        End If
                    End While
                    'xlWorkbook.Save()
                    Exit While
                End If
            Else
                xlRange = xlRange.Offset(1, -7)
            End If
        End While

        Dim strDeleteStatement As String
        strDeleteStatement = "DELETE FROM " & Tablename & " WHERE "
        strDeleteStatement = strDeleteStatement & "Description ='" & mDescription & "' and Capacity ='" & mCapacity & _
            "' and Make = '" & mMake & "' and model = '" & mModel & "' and  MobDate = #" & mMobDate & _
            "# and DeMobDate = #" & mDeMobDate & "# and Qty = " & mQty
        Try
            If (moledbConnection1.State.ToString().Equals("Closed")) Then
                moledbConnection1.Open()
            End If
            moledbCommand = New OleDbCommand
            moledbCommand.CommandType = CommandType.Text
            moledbCommand.CommandText = strDeleteStatement
            moledbCommand.Connection = moledbConnection1
            moledbCommand.ExecuteNonQuery()
            DataGridView1.DataSource = Nothing
            mDataset.Tables(mcategory).Clear()
            DataGridView1.Rows.Clear()
            Dim strSelect As String = "select * from " & Tablename
            Dim mAdapter1 As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(strSelect, moledbConnection1)
            mAdapter1.Fill(mDataset, mcategory)
            DataGridView1.DataSource = mDataset.Tables(mcategory)
            DataGridView1.Refresh()
            ''MsgBox(Me.txtProjectdescription.Text & "Record inserted in list of projects.")
        Catch ex As Exception
            MsgBox(e.ToString())
        Finally
            moledbCommand = Nothing
            'moledbconnection3.Close()
        End Try
        Me.btnDelete.Enabled = False
        Me.btnEdit.Enabled = False
        If EditOperationSheet = "Concreting" Or EditOperationSheet = "Cranes" Or EditOperationSheet = "Material Handling" Or _
            EditOperationSheet = "Non Concreting" Or EditOperationSheet = "DG Sets" Or EditOperationSheet = "Conveyance" Or _
            EditOperationSheet = "Major Others" Then
            frmOptions.cmbMajorEqipCategory.Text = SelectedCategory
            frmOptions.cmbMajorEquipName.Text = mDescription
            frmOptions.cmbMajorEquipCapacity.Text = mCapacity
            frmOptions.cmbMajorEquipModel.Text = mModel
            frmOptions.cmbMajorEquipMake.Text = mMake
            frmOptions.txtMajorEquipQty.Text = mQty
            frmOptions.dpMobDate.Text = mMobDate.Date
            frmOptions.dpDemobDate.Text = mDeMobDate.Date
            frmOptions.txtMonths.Text = 0
            frmOptions.cmbDrive.SelectedIndex = 0
            frmOptions.txtRepvalue.Text = 0
            frmOptions.cmbDepPercentage.SelectedIndex = 0
            frmOptions.txtDepreciation.Text = 0
            frmOptions.txtPropHrs.Text = 0
            frmOptions.txtFuelPerHr.Text = 0
            frmOptions.txtPowerPerHr.Text = 0
            frmOptions.txtOprCostPerMCPerMonth.Text = 0
            frmOptions.txtConsumablesPerMonth.Text = 0
            frmOptions.cmbMinorequip_Shifts.SelectedIndex = 0
            frmOptions.tbclEqipsEntry.SelectTab(0)
            frmOptions.txtMaintPercPerMC_PerMonth.Text = 0
        ElseIf EditOperationSheet = "external Hire" Or EditOperationSheet = "External Others" Then
            frmOptions.cmbExtEquip_Category.Text = SelectedCategory
            frmOptions.cmbExtEquip_Name.Text = mDescription
            frmOptions.cmbExtEquip_Capacity.Text = mCapacity
            frmOptions.cmbExtEquip_Model.Text = mModel
            frmOptions.cmbExtEquip_Make.Text = mMake
            frmOptions.cmbExtEquip_Drive.SelectedIndex = 0
            frmOptions.txtExtEquip_Qty.Text = mQty
            frmOptions.dpExtEquip_MobDate.Text = mMobDate.Date
            frmOptions.dpExtEquip_DemobDate.Text = mDeMobDate.Date
            frmOptions.txtExtEquip_Months.Text = 0
            frmOptions.txtExtEquip_RepValue.Text = 0
            frmOptions.cmbExtEquip_DepPerc.SelectedIndex = 0
            frmOptions.txtExtEquip_Depreciation.Text = 0
            frmOptions.txtExtEquip_Hire_Charges.Text = 0
            frmOptions.txtExtEquip_PropHours.Text = 0
            frmOptions.txtExtEquip_FuelPerHr.Text = 0
            frmOptions.txtExtEquip_PowerPerhr.Text = 0
            frmOptions.txtExtEquip_OprCostPerMCPerMonth.Text = 0
            frmOptions.txtExtEquip_ConxumablesPerMonth.Text = 0
            frmOptions.txtExtEquip_OperatingCost.Text = 0
            frmOptions.cmbExtEquip_Shifts.SelectedIndex = 0
            frmOptions.txtMaintPercPerMC_PerMonth.Text = 0
            frmOptions.tbclEqipsEntry.SelectTab(2)
        ElseIf EditOperationSheet = "Minor Eqpts" Then
            frmOptions.cmbMinoEquip_Category.Text = SelectedCategory
            frmOptions.cmbMinorEquip_Name.Text = mDescription
            frmOptions.cmbMinorEquip_capacity.Text = mCapacity
            frmOptions.cmbMinorEquip_Model.Text = mModel
            frmOptions.cmbMinorEquip_Make.Text = mMake
            frmOptions.cmbMinorEquip_Drive.SelectedIndex = 0
            frmOptions.txtMinorEquip_Qty.Text = mQty
            frmOptions.dpMinorEquip_MobDate.Text = mMobDate.Date
            frmOptions.dpMinorequip_Demobdate.Text = mDeMobDate.Date
            frmOptions.txtMinorequip_Months.Text = 0
            frmOptions.txtMinorequip_NewEquipCost.Text = 0
            frmOptions.cmbMinorequip_DepPerc.Text = 0
            frmOptions.txtMinorequip_Depreciation.Text = 0
            frmOptions.txtMinorEquip_HireCharges.Text = 0
            frmOptions.txtMinorequip_PropHours.Text = 0
            frmOptions.txtMinorequip_FuelPerhr.Text = 0
            frmOptions.txtMinorequip_PowerPerHr.Text = 0
            frmOptions.txtMinorequip_OprCostPerMCPerMonth.Text = 0
            frmOptions.txtMinorequip_ConxumablesPerMonth.Text = 0
            frmOptions.txtMinorequip_OperatingCost.Text = 0
            frmOptions.cmbMinorequip_Shifts.SelectedIndex = 0
            frmOptions.txtMinorequip_ConsumablesPercPerMCPerMonth.Text = 0
            frmOptions.tbclEqipsEntry.SelectTab(1)
        End If
        'frmOptions.Refresh()
        BlankForm = False
        Me.btnClose_Click(Me, e)

    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim row As DataGridViewRow, rowno As Integer
        If NomoreDelete Then Exit Sub
        row = Me.DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex)
        rowno = DataGridView1.CurrentCell.RowIndex
        Dim sheetno, x As Integer, cols As Integer, intI As Integer
        EditOrDelete = "Delete"
        mDescription = row.Cells("Description").Value
        mCapacity = row.Cells("Capacity").Value
        mMake = row.Cells("Make").Value
        mModel = row.Cells("Model").Value
        mMobDate = row.Cells("MobDate").Value
        mDeMobDate = row.Cells("DemobDate").Value
        mQty = row.Cells("Qty").Value
        mMonths = System.Math.Round((mDeMobDate.Date - mMobDate.Date).Days / 30, 0)
        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With
        intI = 2
        If SelectedCategory = "External Others" Then
            EditOperationSheet = "External Others"
        ElseIf SelectedCategory = "Hired Equipments" Or SelectedCategory = "Conveyance_Hired" Then
            EditOperationSheet = "external Hire"
        ElseIf SelectedCategory = "Minor Equipments" Then
            EditOperationSheet = "Minor Eqpts"
        Else
            EditOperationSheet = SelectedCategory
        End If
        xlWorksheet = xlWorkbook.Sheets.Item(EditOperationSheet)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "SlNo")
        xlRange = xlRange.Offset(1, 0)
        sheetno = getSheetNo(xlWorksheet.Name)
        'xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        If RecordsInserted(sheetno) = 1 Then
            If Category_Shortname = "Concrete_" Then
                cols = 34
            Else
                cols = 33
            End If
            xlRange = xlWorksheet.Range(Category_Shortname & "Slno").Offset(1, 0)
            For intI = 1 To cols
                'MsgBox(xlRange.Address & "......." & xlRange.Formula)
                If Not xlRange.HasFormula Then
                    xlRange.Value = ""
                End If
                xlRange = xlRange.Offset(0, 1)
            Next
            RecordsInserted(sheetno) = RecordsInserted(sheetno) - 1
            xlRange = xlRange.Offset(1, -cols)
            xlRange = xlRange.Offset(0, 1)
            If xlRange.Value = "Total" Then
                xlRange.Select()
                'MsgBox(xlRange.Application.ActiveCell.Address)
                xlRange.Application.ActiveCell.EntireRow.Delete()
            End If
            xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
            xlRange.Value = RecordsInserted(sheetno)
            xlApp.CalculateBeforeSave = True
            xlWorkbook.Save()
            xlWorkbook.Close()
            xlWorkbook = Nothing
            With xlApp
                If xlWorkbook Is Nothing Then
                    xlWorkbook = .Workbooks.Open(xlFilename)
                End If
            End With

        Else
            'If rowno = 0 Then
            '    MsgBox("First entry in the grid can be only deleted last")
            '    Exit Sub
            'End If
            xlRange = xlWorksheet.Range(Category_Shortname & "Slno").Offset(1, 0)
            While (True)
                Dim xldescription, xlcapacity, xlmakemodel As String
                Dim xlmobdate, xldemobdate As Date
                Dim xlqty As Integer
                xlRange = xlRange.Offset(0, 1)
                'MsgBox(xlRange.Value)
                xldescription = xlRange.Value
                xlRange = xlRange.Offset(0, 2)
                'MsgBox(xlRange.Value)
                xlcapacity = xlRange.Value
                xlRange = xlRange.Offset(0, 1)
                'MsgBox(xlRange.Value)
                xlmakemodel = xlRange.Value
                xlRange = xlRange.Offset(0, 1)
                'MsgBox(xlRange.Value)
                xlqty = xlRange.Value
                xlRange = xlRange.Offset(0, 1)
                'MsgBox(xlRange.Value)
                xlmobdate = xlRange.Value
                xlRange = xlRange.Offset(0, 1)
                'MsgBox(xlRange.Value)
                xldemobdate = xlRange.Value
                Dim xx As String = mMake & vbNewLine & "/" & mModel
                If (xldescription = mDescription And xlcapacity = mCapacity And _
                    xlmakemodel = xx And xlqty = mQty And _
                    xlmobdate = mMobDate And xldemobdate = mDeMobDate) Then
                    xlRange.Select()
                    'MsgBox(xlRange.Address)
                    'MsgBox(xlApp.ActiveCell.Address & "....." & xlApp.ActiveCell.Value)
                    xlRange.Application.ActiveCell.EntireRow.Delete()
                    xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
                    'sheetno = getSheetNo(xlWorksheet.Name)
                    'RecordsInserted(sheetno) = RecordsInserted(sheetno)
                    xlApp.CalculateBeforeSave = True
                    xlWorkbook.Save()
                    xlWorkbook.Close()
                    xlWorkbook = Nothing
                    With xlApp
                        If xlWorkbook Is Nothing Then
                            xlWorkbook = .Workbooks.Open(xlFilename)
                        End If
                    End With
                    'MsgBox(xlWorkbook.Name)
                    xlWorksheet = xlWorkbook.Sheets.Item(EditOperationSheet)
                    sheetno = getSheetNo(xlWorksheet.Name)
                    RecordsInserted(sheetno) = RecordsInserted(sheetno) - 1
                    xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
                    xlRange.Value = RecordsInserted(sheetno)
                    If xlRange.Value = 0 Then
                        xlRange = xlWorksheet.Range(Category_Shortname & "Equip_Name")
                        xlRange = xlRange.Offset(1, 0).Select
                        xlRange.Application.ActiveCell.EntireRow.Delete()
                    End If

                    x = 1
                    xlRange = xlWorksheet.Range(Category_Shortname & "SlNo")
                    'MsgBox(xlRange.Value)
                    xlRange = xlRange.Offset(1, 0)
                    'MsgBox(xlRange.Value)
                    While (True)
                        If Not Len(Trim(xlRange.Value)) = 0 Then
                            xlRange.Value = x
                            x = x + 1
                            xlRange = xlRange.Offset(1, 0)
                        Else
                            Exit While
                        End If
                    End While
                    'xlWorkbook.Save()
                    Exit While
                Else
                    xlRange = xlRange.Offset(1, -7)
                End If
            End While
        End If
        Dim strDeleteStatement As String
        strDeleteStatement = "DELETE FROM " & Tablename & " WHERE "
        strDeleteStatement = strDeleteStatement & "Description ='" & mDescription & "' and Capacity ='" & mCapacity & _
            "' and Make = '" & mMake & "' and model = '" & mModel & "' and  MobDate = #" & mMobDate & _
            "# and DeMobDate = #" & mDeMobDate & "# and Qty = " & mQty
        Try
            If (moledbConnection1.State.ToString().Equals("Closed")) Then
                moledbConnection1.Open()
            End If
            moledbCommand = New OleDbCommand
            moledbCommand.CommandType = CommandType.Text
            moledbCommand.CommandText = strDeleteStatement
            moledbCommand.Connection = moledbConnection1
            moledbCommand.ExecuteNonQuery()
            DataGridView1.DataSource = Nothing
            mDataset.Tables(mcategory).Clear()
            DataGridView1.Rows.Clear()
            Dim strSelect As String = "select * from " & Tablename
            Dim mAdapter1 As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(strSelect, moledbConnection1)
            mAdapter1.Fill(mDataset, mcategory)
            DataGridView1.DataSource = mDataset.Tables(mcategory)
            DataGridView1.Refresh()
            ''MsgBox(Me.txtProjectdescription.Text & "Record inserted in list of projects.")
        Catch ex As Exception
            MsgBox(e.ToString())
        Finally
            moledbCommand = Nothing
            'moledbconnection3.Close()
        End Try
        Me.btnDelete.Enabled = False
        Me.btnEdit.Enabled = False
        MsgBox("Entry for " & vbNewLine & "Equpment Name: " & mDescription & vbNewLine & _
            "Capacity: " & mCapacity & vbNewLine & "Make: " & mMake & vbNewLine & _
            "Model: " & mModel & vbNewLine & "Mobilization Date: " & mMobDate & vbNewLine & _
            "Demobilization Date" & mDeMobDate & " Deleted")
        If EditOperationSheet = "Concreting" Or EditOperationSheet = "Cranes" Or EditOperationSheet = "Material Handling" Or _
        EditOperationSheet = "Non Concreting" Or EditOperationSheet = "DG Sets" Or EditOperationSheet = "Conveyance" Or _
        EditOperationSheet = "Major Others" Then
            frmOptions.cmbMajorEqipCategory.Text = ""
            frmOptions.cmbMajorEquipName.Text = ""
            frmOptions.cmbMajorEquipCapacity.Text = ""
            frmOptions.cmbMajorEquipModel.Text = ""
            frmOptions.cmbMajorEquipMake.Text = ""
            frmOptions.txtMajorEquipQty.Text = 0
            frmOptions.dpMobDate.Text = ""
            frmOptions.dpDemobDate.Text = ""
            frmOptions.txtMonths.Text = 0
            frmOptions.cmbDrive.SelectedIndex = 0
            frmOptions.txtRepvalue.Text = 0
            frmOptions.cmbDepPercentage.SelectedIndex = 0
            frmOptions.txtDepreciation.Text = 0
            frmOptions.txtPropHrs.Text = 0
            frmOptions.txtFuelPerHr.Text = 0
            frmOptions.txtPowerPerHr.Text = 0
            frmOptions.txtOprCostPerMCPerMonth.Text = 0
            frmOptions.txtConsumablesPerMonth.Text = 0
            frmOptions.cmbMinorequip_Shifts.SelectedIndex = 0
            frmOptions.txtMaintPercPerMC_PerMonth.Text = 0
            frmOptions.tbclEqipsEntry.SelectTab(0)
        ElseIf EditOperationSheet = "external Hire" Or EditOperationSheet = "External Others" Then
            frmOptions.cmbExtEquip_Category.Text = ""
            frmOptions.cmbExtEquip_Name.Text = ""
            frmOptions.cmbExtEquip_Capacity.Text = ""
            frmOptions.cmbExtEquip_Model.Text = ""
            frmOptions.cmbExtEquip_Make.Text = ""
            frmOptions.txtExtEquip_Qty.Text = 0
            frmOptions.dpExtEquip_MobDate.Text = ""
            frmOptions.dpExtEquip_DemobDate.Text = ""
            frmOptions.cmbExtEquip_Drive.SelectedIndex = 0
            frmOptions.txtExtEquip_Months.Text = 0
            frmOptions.txtExtEquip_RepValue.Text = 0
            frmOptions.cmbExtEquip_DepPerc.SelectedIndex = 0
            frmOptions.txtExtEquip_Depreciation.Text = 0
            frmOptions.txtExtEquip_Hire_Charges.Text = 0
            frmOptions.txtExtEquip_PropHours.Text = 0
            frmOptions.txtExtEquip_FuelPerHr.Text = 0
            frmOptions.txtExtEquip_PowerPerhr.Text = 0
            frmOptions.txtExtEquip_OprCostPerMCPerMonth.Text = 0
            frmOptions.txtExtEquip_ConxumablesPerMonth.Text = 0
            frmOptions.txtExtEquip_OperatingCost.Text = 0
            frmOptions.cmbExtEquip_Shifts.SelectedIndex = 0
            frmOptions.txtMaintPercPerMC_PerMonth.Text = 0
            frmOptions.tbclEqipsEntry.SelectTab(2)
        ElseIf EditOperationSheet = "Minor Eqpts" Then
            frmOptions.cmbMinoEquip_Category.Text = ""
            frmOptions.cmbMinorEquip_Name.Text = ""
            frmOptions.cmbMinorEquip_capacity.Text = ""
            frmOptions.cmbMinorEquip_Model.Text = ""
            frmOptions.cmbMinorEquip_Make.Text = ""
            frmOptions.txtMinorEquip_Qty.Text = 0
            frmOptions.dpMinorEquip_MobDate.Text = ""
            frmOptions.dpMinorequip_Demobdate.Text = ""
            frmOptions.cmbMinorEquip_Drive.SelectedIndex = 0
            frmOptions.txtMinorequip_Months.Text = 0
            frmOptions.txtMinorequip_NewEquipCost.Text = 0
            frmOptions.cmbMinorequip_DepPerc.Text = 0
            frmOptions.txtMinorequip_Depreciation.Text = 0
            frmOptions.txtMinorEquip_HireCharges.Text = 0
            frmOptions.txtMinorequip_PropHours.Text = 0
            frmOptions.txtMinorequip_FuelPerhr.Text = 0
            frmOptions.txtMinorequip_PowerPerHr.Text = 0
            frmOptions.txtMinorequip_OprCostPerMCPerMonth.Text = 0
            frmOptions.txtMinorequip_ConxumablesPerMonth.Text = 0
            frmOptions.txtMinorequip_OperatingCost.Text = 0
            frmOptions.cmbMinorequip_Shifts.SelectedIndex = 0
            frmOptions.txtMinorequip_ConsumablesPercPerMCPerMonth.Text = 0
            frmOptions.tbclEqipsEntry.SelectTab(1)
        End If
        'frmOptions.Refresh()
        BlankForm = False
        'If rowno = 0 Then NomoreDelete = True
    End Sub
End Class

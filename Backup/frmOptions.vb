
Imports System.Data
Imports System.Data.OleDb

Public Class frmOptions
    Public xlApp As New Microsoft.Office.Interop.Excel.Application
    Public strConnection As String

    Public moledbConnection As OleDbConnection
    Dim strStatement As String
    Dim moledbCommand As OleDbCommand
    Dim mOledbDataAdapter As OleDbDataAdapter
    Dim mReader As OleDbDataReader
    Dim mDataSet As DataSet
    Dim RowCount As Integer = 0
    Dim RangeCols As Integer
    Dim oForm As Form

    Private Sub cmbMajorEqipCategory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMajorEqipCategory.SelectedIndexChanged
        TemplatePath = (My.Application.Info.DirectoryPath)
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
        moledbConnection = New OleDbConnection(strConnection)
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strStatement = "Select distinct Equipmentname from MajorEquipments where Categoryname = '" & Me.cmbMajorEqipCategory.Text & "'"
        moledbCommand = New OleDbCommand(strStatement, moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbMajorEquipName.Items.Clear()
        If mReader.HasRows Then
            mReader.Read()
            Do
                Me.cmbMajorEquipName.Items.Add(mReader("EquipmentName"))
            Loop While mReader.Read()
        End If
        moledbCommand = Nothing
        mReader.Close()
        Me.btnViewItemsMajor.Text = "Click here to view items already added  from " & Me.cmbMajorEqipCategory.Text & " category"
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.ClearFields()
    End Sub

    Private Sub frmOptions_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If BlankForm Then
            Exit Sub
        Else
            BlankForm = True
            Me.btnQuit.Enabled = True
            Me.btnClose.Enabled = False
            'Me.cmbMinoEquip_Category.Text = SelectedCategory
            If EditOperationSheet = "Concreting" Or EditOperationSheet = "Cranes" Or EditOperationSheet = "Material Handling" Or _
            EditOperationSheet = "Non Concreting" Or EditOperationSheet = "DG Sets" Or EditOperationSheet = "Conveyance" Or _
            EditOperationSheet = "Major Others" Then
                Me.cmbMajorEqipCategory.Text = SelectedCategory
                Me.cmbMajorEquipModel_SelectedIndexChanged(Me, e)
            ElseIf EditOperationSheet = "external Hire" Or EditOperationSheet = "External Others" Then
                Me.cmbExtEquip_Category.Text = SelectedCategory
                Me.cmbExtEquip_Model_SelectedIndexChanged(Me, e)
            ElseIf EditOperationSheet = "Minor Eqpts" Then
                Me.cmbMinoEquip_Category.Text = SelectedCategory
                Me.cmbMinorEquip_Model_SelectedIndexChanged(Me, e)
            End If
        End If
    End Sub

    Private Sub frmOptions_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Leave

    End Sub

    Private Sub frmOptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim intI As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
        If moledbConnection Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        Me.btnClose.Enabled = False
        Me.btnQuit.Enabled = True
        moledbCommand = New OleDbCommand("Select distinct categoryname from MajorEquipments", moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbMajorEqipCategory.Items.Clear()
        If mReader.HasRows Then
            mReader.Read()
            Do
                cmbMajorEqipCategory.Items.Add(mReader("Categoryname"))
            Loop While mReader.Read()
        End If
        moledbCommand = New OleDbCommand("Select distinct categoryname from HiredEquipments", moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbExtEquip_Category.Items.Clear()
        If mReader.HasRows Then
            mReader.Read()
            Do
                cmbExtEquip_Category.Items.Add(mReader("Categoryname"))
            Loop While mReader.Read()
        End If
        Me.cmbMinoEquip_Category.Items.Clear()
        Me.cmbMinoEquip_Category.Items.Add("Minor Equipments")
        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With
        SheetsCount = 12
        If Startup Then
            For intI = 1 To SheetsCount
                xlWorksheet = xlWorkbook.Sheets.Item(intI)
                Sheetnames(intI) = (xlWorksheet.Name)
                SheetIndices(intI) = intI
                getCategoryShortname(xlWorksheet)
                RecordsInserted(intI) = xlWorksheet.Range(Category_Shortname & "RecordsTotal").Value
            Next
            Startup = False
            Me.btnClose.Enabled = False
        End If
        Me.cmbMajorEqipCategory.SelectedIndex = 0
        moledbCommand = Nothing
        mReader.Close()
        Me.dpMobDate.Value = mStartDate.Date
        Me.dpDemobDate.Value = mEndDate.Date
        Me.dpMinorEquip_MobDate.Value = mStartDate.Date
        Me.dpMinorequip_Demobdate.Value = mEndDate.Date
        Me.dpExtEquip_MobDate.Value = mStartDate.Date
        Me.dpExtEquip_DemobDate.Value = mEndDate.Date
        FormLoaded = True
        Me.tbclEqipsEntry.SelectTab(Currenttab)
        If Currenttab = 0 Then
            Me.cmbMajorEqipCategory.Text = SelectedCategory
        ElseIf Currenttab = 1 Then
            Me.cmbMinoEquip_Category.Text = SelectedCategory
        ElseIf Currenttab = 2 Then
            Me.cmbExtEquip_Category.Text = SelectedCategory
        End If
    End Sub

    Private Sub cmbMajorEquipName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMajorEquipName.SelectedIndexChanged
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        Dim strSql As String
        strSql = "categoryname = '" & Me.cmbMajorEqipCategory.Text & "' And EquipmentName = '" & Me.cmbMajorEquipName.Text & "'"
        strStatement = "Select distinct Capacity From MajorEquipments where " & strSql
        moledbCommand = New OleDbCommand(strStatement, moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbMajorEquipCapacity.Items.Clear()

        If mReader.HasRows Then
            mReader.Read()
            Do
                Me.cmbMajorEquipCapacity.Items.Add(mReader("Capacity"))
            Loop While mReader.Read()
        End If
        moledbCommand = Nothing
        mReader.Close()
    End Sub

    Private Sub txtPowerPerHr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPowerPerHr.TextChanged

    End Sub

    Private Sub txtMajorEquipQty_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMajorEquipQty.Validating
        If Not IsNumeric(Me.txtMajorEquipQty.Text) Then
            MsgBox("Enter Numeric value for the Quantity")
            Me.txtMajorEquipQty.Text = ""
            e.Cancel = True
        Else
            Me.txtMajorEquipQty.Text = Int(Val(Me.txtMajorEquipQty.Text))
            Me.txtHireCharges.Text = Val(Me.txtDepreciation.Text) * Val(Me.txtMajorEquipQty.Text) * Val(Me.txtMonths.Text)
        End If
    End Sub

    Private Sub dpMobDate_MouseCaptureChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dpMobDate.MouseCaptureChanged

    End Sub

    Private Sub dpMobDate_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dpMobDate.Validated
        If (Me.dpDemobDate.Value.Date > Me.dpMobDate.Value.Date) Then
            Me.txtMonths.Text = System.Math.Round((Me.dpDemobDate.Value.Date - Me.dpMobDate.Value.Date).Days / 30, 0)
        Else
            Me.txtMonths.Text = 0
        End If
    End Sub

    Private Sub dpMobDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dpMobDate.Validating
        If Me.dpMobDate.Value.Date < mStartDate.Date Then
            MsgBox("Mobilisation Date cannot be past date. Please select again")
            e.Cancel = True
            Exit Sub
        End If
        If Me.dpMobDate.Value.Date < mStartDate Then
            MsgBox("Mobilization date cannot be less than than the project start date")
            Me.dpMobDate.Focus()
            e.Cancel = True
        End If
    End Sub

    Private Sub dpDemobDate_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dpDemobDate.Validated
        If (Me.dpDemobDate.Value > Me.dpMobDate.Value) Then
            Me.txtMonths.Text = System.Math.Round((Me.dpDemobDate.Value.Date - Me.dpMobDate.Value.Date).Days / 30, 0)
        Else
            Me.txtMonths.Text = 0
        End If
    End Sub

    Private Sub dpDemobDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dpDemobDate.Validating
        If Me.dpDemobDate.Value < mStartDate.Date Then
            MsgBox("De-Mobilisation Date cannot be less than Project Start date. Please select again")
            Me.dpDemobDate.Value = mEndDate.Date
            e.Cancel = True
            Exit Sub
        End If
        If Me.dpDemobDate.Value.Date > mEndDate Then
            MsgBox("De-mobilization date cannot be beyond project end date")
            Me.dpDemobDate.Focus()
            e.Cancel = True
        End If

        If Me.dpDemobDate.Value <= Me.dpMobDate.Value Then
            MsgBox("De-mobilisation Date must be greater than Mobilisation Date")
            Me.dpDemobDate.Value = mEndDate.Date
            e.Cancel = True
        End If
    End Sub

    Private Sub txtMonths_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMonths.Enter
        If (Me.dpDemobDate.Value > Me.dpMobDate.Value) Then
            Me.txtMonths.Text = System.Math.Round((Me.dpDemobDate.Value.Date - Me.dpMobDate.Value.Date).Days / 30, 0)
        Else
            Me.txtMonths.Text = 0
        End If
    End Sub

    Private Sub txtMonths_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMonths.TextChanged

    End Sub

    Private Sub cmbMajorEquipMake_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMajorEquipMake.SelectedIndexChanged
        Dim strSql As String
        strSql = "categoryname = '" & Me.cmbMajorEqipCategory.Text & "' And EquipmentName = '" & Me.cmbMajorEquipName.Text & "' and Capacity = '" & _
           Me.cmbMajorEquipCapacity.Text & "' and Make = '" & Me.cmbMajorEquipMake.Text & "'"
        strStatement = "Select distinct Model From MajorEquipments where " & strSql
        moledbCommand = New OleDbCommand(strStatement, moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbMajorEquipModel.Items.Clear()

        If mReader.HasRows Then
            mReader.Read()
            Do
                Me.cmbMajorEquipModel.Items.Add(mReader("Model"))
            Loop While mReader.Read()
        End If
        moledbCommand = Nothing
        mReader.Close()
    End Sub

    Private Sub cmbDrive_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDrive.SelectedIndexChanged, cmbShifts.SelectedIndexChanged
    End Sub

    Private Sub cmbDepPercentage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDepPercentage.SelectedIndexChanged
    End Sub

    Private Sub txtFuelPerHr_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFuelPerHr.Validating
        If Not IsNumeric(Me.txtFuelPerHr.Text) Then
            MsgBox("Fuel per hour should be numeric")
            e.Cancel = True
        End If
    End Sub

    Private Sub btnComputevalues_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnComputevalues.Click
        Dim AllValid As Boolean
        AllValid = True
        Dim msgstring As String
        msgstring = "The following were not entered/selected" & vbNewLine
        If (Me.cmbMajorEqipCategory.Text = "") Then
            msgstring = msgstring & "*** Equipment catefory not selected"
            AllValid = False
        End If
        If UCase(Me.cmbMajorEqipCategory.Text) = UCase("Concrete") And (Me.txtConcreteQty.Text = "" Or _
            Val(Me.txtConcreteQty.Text) = 0) Then
            msgstring = msgstring & "*** Concrete Qunatity is not entered of is zero"
            AllValid = False
        End If
        If (Me.cmbMajorEquipName.Text = "") Then
            msgstring = msgstring & "*** Equipment Name not selected" & vbNewLine
            AllValid = False
        End If
        If (Me.cmbDrive.Text = "") Then
            msgstring = msgstring & "*** Equipment drive not selected" & vbNewLine
            AllValid = False
        End If
        If (Me.cmbMajorEquipCapacity.Text = "") Then
            msgstring = msgstring & "*** Equipment capacity not selected" & vbNewLine
            AllValid = False
        End If
        If (Me.cmbMajorEquipMake.Text = "") Then
            msgstring = msgstring & "*** Equipment Make not selected" & vbNewLine
            AllValid = False
        End If
        If (Me.cmbMajorEquipModel.Text = "") Then
            msgstring = msgstring & "*** Equipment model not selected" & vbNewLine
            AllValid = False
        End If
        If (Me.txtMajorEquipQty.Text = "" Or Len(Trim(Me.txtMajorEquipQty.Text)) = 0) Then
            msgstring = msgstring & "*** Equipment quantity  not entered" & vbNewLine
            AllValid = False
        End If
        If (Me.dpMobDate.Value.ToString() = "") Then
            msgstring = msgstring & "*** Equipment mobilisation date not selected" & vbNewLine
            AllValid = False
        End If
        If (dpDemobDate.Value.ToString() = "") Then
            msgstring = msgstring & "*** Equipment demobilisation date  not selected" & vbNewLine
            AllValid = False
        End If
        If Trim(Me.cmbDrive.Text) = "Fuel(Diesel)" Then
            If (Me.txtFuelPerHr.Text = "") Then
                msgstring = msgstring & "*** Fuel consumption per Hour missing"
                AllValid = False
            End If
            If (Val(Me.txtFuelCostPerLtr.Text) = 0 Or Me.txtFuelCostPerLtr.Text = "") Then
                msgstring = msgstring & "*** Fuel cost per Ltr not entered or is zero" & vbNewLine
                AllValid = False
            End If
        End If
        If Trim(Me.cmbDrive.Text) = "Electrical()" Then
            If (Me.txtPowerPerHr.Text = "") Then
                msgstring = msgstring & "*** Power Consumption Per Hr" & vbNewLine
                AllValid = False
            End If
            If (Val(Me.txtPowerCostPerUnit.Text) = 0 Or Me.txtPowerCostPerUnit.Text = "") Then
                msgstring = msgstring & "*** Power Cost per Unit is not entered or is zero"
                AllValid = False
            End If
        End If
        If (Me.cmbShifts.Text = "") Then
            Me.cmbShifts.SelectedIndex = 0
        End If
        If (Me.txtMaintPercPerMC_PerMonth.Text = "" Or Val(txtMaintPercPerMC_PerMonth.Text) = 0) Then
            msgstring = msgstring & "*** Maintenance Percentage not entered."
            AllValid = False
        End If
        If Not AllValid Then
            MsgBox(msgstring)
        Else
            'mConcreteQty = Val(Me.txtConcreteQty.Text)
            If UCase(Me.cmbMajorEqipCategory.Text) = UCase("Conveyance") Or UCase(Me.cmbMajorEquipName.Text) = UCase("Tipper") Or _
                  UCase(Me.cmbMajorEquipName.Text) = UCase("Truck") Then
                Me.txtFuelPerunitpermonth.Text = Val(Me.txtPropHrs.Text) / Val(Me.txtFuelPerHr.Text)
            Else
                Me.txtFuelPerunitpermonth.Text = Val(Me.txtFuelPerHr.Text) * Val(Me.txtPropHrs.Text)
            End If
            Me.txtFuelCostPerMonth.Text = Val(txtFuelPerunitpermonth.Text) * Val(Me.txtMajorEquipQty.Text) * Val(Me.txtFuelCostPerLtr.Text)
            Me.txtFuelCostProject.Text = Val(txtFuelCostPerMonth.Text) * Val(Me.txtMonths.Text)
            Me.txtPowerCostPerMonth.Text = Val(Me.txtPowerPerHr.Text) * Val(Me.txtPropHrs.Text) * Val(Me.txtMajorEquipQty.Text) * Val(Me.txtPowerCostPerUnit.Text)
            Me.txtPowerCostProject.Text = Val(txtPowerCostPerMonth.Text) * Val(Me.txtMonths.Text)
            Me.txtOprCostPerMonth.Text = Val(Me.txtOprCostPerMCPerMonth.Text) * Val(Me.txtMajorEquipQty.Text) * Val(Me.cmbShifts.Text)
            Me.txtoprCostProject.Text = Val(Me.txtOprCostPerMonth.Text) * Val(Me.txtMonths.Text)
            Me.txtConsumablesPerMonth.Text = Val(Me.txtMaintCostPerMC_PerMonth.Text) * Val(Me.txtMajorEquipQty.Text)
            Me.txtConsumblesProject.Text = Val(txtConsumablesPerMonth.Text) * Val(Me.txtMonths.Text)
            Me.txtOperatingCost_MajorEquips.Text = Val(Me.txtHireCharges.Text) + Val(Me.txtFuelCostProject.Text) + Val(Me.txtPowerCostProject.Text) + _
                  Val(Me.txtoprCostProject.Text) + Val(Me.txtConsumblesProject.Text)
        End If
        Me.Button1.Enabled = True
    End Sub

    Private Sub txtFuelPerunitpermonth_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFuelPerunitpermonth.TextChanged

    End Sub

    Private Sub cmbMajorEquipCapacity_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMajorEquipCapacity.SelectedIndexChanged
        Dim strSql As String
        strSql = "categoryname = '" & Me.cmbMajorEqipCategory.Text & "' And EquipmentName = '" & Me.cmbMajorEquipName.Text & "' and Capacity = '" & _
           Me.cmbMajorEquipCapacity.Text & "'"
        strStatement = "Select distinct Make From MajorEquipments where " & strSql
        moledbCommand = New OleDbCommand(strStatement, moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbMajorEquipMake.Items.Clear()

        If mReader.HasRows Then
            mReader.Read()
            Do
                Me.cmbMajorEquipMake.Items.Add(mReader("Make"))
            Loop While mReader.Read()
        End If
        moledbCommand = Nothing
        mReader.Close()
    End Sub

    Private Sub cmbMajorEquipModel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMajorEquipModel.SelectedIndexChanged
        Dim strSql As String
        strSql = "categoryname = '" & Me.cmbMajorEqipCategory.Text & "' And EquipmentName = '" & Me.cmbMajorEquipName.Text & "' and " & _
           "Make ='" & Me.cmbMajorEquipMake.Text & "' and Model ='" & Me.cmbMajorEquipModel.Text & "' and Capacity ='" & Me.cmbMajorEquipCapacity.Text & "'"
        strStatement = "Select * from MajorEquipments where " & strSql
        If moledbConnection Is Nothing Then
            strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        'create a data adapter
        mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
        'create a dataset
        mDataSet = New DataSet()
        'fill the dataset using data adapter
        mOledbDataAdapter.Fill(mDataSet, "MajorEquipments")
        Dim Machine As DataRow
        For Each Machine In mDataSet.Tables("MajorEquipments").Rows
            Me.txtRepvalue.Text = Machine("Repvalue").ToString()
            If Val(Machine("Depreciation_Fixed").ToString()) > 0 Then
                Me.txtDepreciation.Text = Machine("Depreciation_Fixed").ToString()
                Me.cmbDepPercentage.Text = "Fixed"
                Me.cmbDepPercentage.Enabled = False
            Else
                Me.txtDepreciation.Text = 0
                Me.cmbDepPercentage.SelectedIndex = 0
                Me.cmbDepPercentage.Enabled = True
            End If
            Me.txtPropHrs.Text = Machine("Hrs_PerMonth").ToString()
            Me.txtFuelPerHr.Text = Machine("Fuel_PerHour").ToString()
            Me.txtPowerPerHr.Text = Machine("Power_PerHour").ToString()
            Me.txtOprCostPerMCPerMonth.Text = Machine("OperatorCost_PerMonth").ToString()
            Me.txtMaintCostPerMC_PerMonth.Text = Machine("MaintCost_PerMonth").ToString()
            mMinMaintperc = Val(Machine("MinMaintCostperc").ToString())
            mMaxMaintPerc = Val(Machine("MaxMaintCostperc").ToString)
            Me.txtMaintPercPerMC_PerMonth.Text = Machine("DefaultMaintCostPerc")
            Me.Label123.Text = "(Between " & mMinMaintperc & " and " & mMaxMaintPerc & ")"
            Me.txtMaintCostPerMC_PerMonth.Text = Val(Me.txtRepvalue.Text) * Val(Me.txtMaintPercPerMC_PerMonth.Text) / 100
            If UCase(Me.cmbMajorEqipCategory.Text) = UCase("Conveyance") Or UCase(Me.cmbMajorEquipName.Text) = UCase("Tipper") Or _
                  UCase(Me.cmbMajorEquipName.Text) = UCase("Truck") Then
                Me.lblMajorHrsPerMonth.Text = "Usage Kilometers per Month"
            Else
                Me.lblMajorHrsPerMonth.Text = "Proposed Hrs. per Month"
            End If

            If (Me.txtFuelPerHr.Text = 0 Or Len(Trim(txtFuelPerHr.Text)) = 0) Then
                Me.cmbDrive.Text = "Electrical"
            Else
                Me.cmbDrive.Text = "Fuel (Diesel)"
            End If
            If Me.cmbDepPercentage.Text <> "0" Then
                Me.txtDepreciation.Text = Val(Me.txtRepvalue.Text) * Val(Me.cmbDepPercentage.Text) / 100
                Me.txtHireCharges.Text = Val(Me.txtDepreciation.Text) * Val(Me.txtMajorEquipQty.Text) * Val(Me.txtMonths.Text)
                Me.txtDepreciation.Enabled = False
            End If
            mOledbDataAdapter = Nothing
        Next
        'End If
    End Sub

    Private Sub txtConsumblesProject_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtConsumblesProject.TextChanged

    End Sub

    Private Sub txtFuelCostPerMonth_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFuelCostPerMonth.TextChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim intI As Integer, mcategory As String
        Dim InsertCommand As String
        Dim moleDBInsertCommand As OleDbCommand

        Dim currentsheetname1, currentsheetname2 As String
        Me.btnQuit.Enabled = True
        Me.btnClose.Enabled = False
        If Len(Trim(xlFilename)) = 0 Then
            MsgBox("Filename not specified to save. cannot Save data. Start the application again.")
            Application.Exit()
        End If
        If Me.dpMobDate.Value.Date > Me.dpDemobDate.Value.Date Then
            MsgBox("Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
              "And mobilisation and demobilisation dates must fall within Projet Start and End dates")
            Me.dpDemobDate.Value = mEndDate.Date
            Me.dpMobDate.Focus()
            Exit Sub
        End If

        If (Val(Me.txtMajorEquipQty.Text) = 0 Or Len(Trim(Me.txtMajorEquipQty.Text)) = 0) Then
            MsgBox("Quantity is missing")
            Me.txtMajorEquipQty.Text = 0
            Me.txtMajorEquipQty.Focus()
            Exit Sub
        End If

        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With

        SheetsCount = 12   '= xlWorkbook.Sheets.Count
        If Startup Then
            For intI = 1 To SheetsCount
                xlWorksheet = xlWorkbook.Sheets.Item(intI)
                Sheetnames(intI) = (xlWorksheet.Name)
                SheetIndices(intI) = intI
                getCategoryShortname(xlWorksheet)
                'MsgBox(xlWorksheet.Name)
                RecordsInserted(intI) = xlWorksheet.Range(Category_Shortname & "RecordsTotal").Value
            Next
            Startup = False
            Me.btnClose.Enabled = False
        End If
        currentsheetname2 = Me.cmbMajorEqipCategory.Text

        For intI = 0 To SheetsCount - 1
            currentsheetname1 = Sheetnames(intI + 1)
            If UCase(Trim(currentsheetname1)) = UCase(Trim(currentsheetname2)) Then
                PrevSheetname = xlWorksheet.Name
                xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
                SheetNo = intI + 1
                Exit For
            End If
        Next

        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        Me.txtConcreteQty.Text = mConcreteQty
        If Category_Shortname = "Concrete_" Then
            xlRange = xlWorksheet.Range(Category_Shortname & "ConcreteQty")
            xlRange.Value = mConcreteQty
        End If
        xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
        xlRange.Value = mMainTitle1
        xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
        xlRange.Value = mMainTitle2
        xlRange = xlWorksheet.Range(Category_Shortname & "Client")
        xlRange.Value = mClient
        xlRange = xlWorksheet.Range(Category_Shortname & "Location")
        xlRange.Value = mLocation
        xlRange = xlWorksheet.Range(Category_Shortname & "StartDate")
        xlRange.Value = mStartDate.Date.ToString()
        xlRange.NumberFormat = "dd-mmm-yyyy"
        xlRange = xlWorksheet.Range(Category_Shortname & "EndDate")
        xlRange.Value = mEndDate.Date.ToString()
        xlRange.NumberFormat = "dd-mmm-yyyy"
        xlRange = xlWorksheet.Range(Category_Shortname & "ProjectValue")
        xlRange.Value = mProjectvalue

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
        xlRange.Value = "Fuel Cost for all mc/month @Rs. " & Me.txtFuelCostPerLtr.Text & " per Lt"
        xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
        xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
        'End If
        RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1
        If Trim(xlWorksheet.Name) = "Concreting" Or Trim(xlWorksheet.Name) = "Cranes" Or _
            Trim(xlWorksheet.Name) = "Material Handling" Or Trim(xlWorksheet.Name) = "Non Concreting" Or _
            Trim(xlWorksheet.Name) = "DG Sets" Or Trim(xlWorksheet.Name) = "Conveyance" _
              Or Trim(xlWorksheet.Name) = "Major Others" Then
            xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo), 0)
            If xlRange.Offset(1, 1).Value = "Total" Then
                xlWorksheet.Range(xlRange.Offset(1, 1), xlRange.Offset(0, 35)).ClearContents()
            End If
            xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo), 0)
            xlRange.Value = RecordsInserted(SheetNo)
            xlRange = xlRange.Offset(0, 1)
            xlRange.Value = Me.cmbMajorEquipName.Text
            xlRange = xlRange.Offset(0, 1)
            xlRange.Value = Me.cmbDrive.Text
            xlRange = xlRange.Offset(0, 1)
            xlRange.Value = Me.cmbMajorEquipCapacity.Text
            xlRange = xlRange.Offset(0, 1)
            xlRange.Value = Me.cmbMajorEquipMake.Text & vbNewLine & "/" & Me.cmbMajorEquipModel.Text
            xlRange = xlRange.Offset(0, 1)
            xlRange.Value = IIf(Len(Trim(Me.txtMajorEquipQty.Text)) = 0, 0, Me.txtMajorEquipQty.Text)
            xlRange = xlRange.Offset(0, 1)
            xlRange.Value = Me.dpMobDate.Value.Date.ToString()
            xlRange.NumberFormat = "dd-mmm-yyyy"
            xlRange = xlRange.Offset(0, 1)
            xlRange.Value = Me.dpDemobDate.Value.Date.ToString()
            xlRange.NumberFormat = "dd-mmm-yyyy"
            xlRange = xlRange.Offset(0, 2)
            xlRange.Value = IIf(Len(Trim(Me.txtRepvalue.Text)) = 0, 0, Me.txtRepvalue.Text)
            xlRange = xlRange.Offset(0, 1)
            xlRange.Value = IIf(Len(Trim(Me.cmbDepPercentage.Text)) = 0, 0, Me.cmbDepPercentage.Text)
            xlRange = xlRange.Offset(0, 3)
            xlRange.Value = IIf(Len(Trim(Me.txtPropHrs.Text)) = 0, 0, Me.txtPropHrs.Text)
            xlRange = xlRange.Offset(0, 1)
            xlRange.Value = IIf(Len(Trim(Me.txtFuelPerHr.Text)) = 0, 0, Me.txtFuelPerHr.Text)
            xlRange = xlRange.Offset(0, 4)
            xlRange.Value = IIf(Len(Trim(Me.txtFuelCostPerLtr.Text)) = 0, 0, Me.txtFuelCostPerLtr.Text)
            xlRange = xlRange.Offset(0, 3)
            xlRange.Value = IIf(Len(Trim(Me.txtPowerPerHr.Text)) = 0, 0, Me.txtPowerPerHr.Text)
            xlRange = xlRange.Offset(0, 1)
            xlRange.Value = IIf(Len(Trim(Me.txtPowerCostPerUnit.Text)) = 0, 0, Me.txtPowerCostPerUnit.Text)
            xlRange = xlRange.Offset(0, 3)
            xlRange.Value = IIf(Len(Trim(Me.txtOprCostPerMCPerMonth.Text)) = 0, 0, Me.txtOprCostPerMCPerMonth.Text)
            xlRange = xlRange.Offset(0, 1)
            xlRange.Value = IIf(Len(Trim(Me.cmbShifts.Text)) = 0, 1, Me.cmbShifts.Text)
            xlRange.NumberFormat = "#.0#"
            xlRange = xlRange.Offset(0, 3)
            xlRange.Value = IIf(Len(Trim(Me.txtConsumablesPerMonth.Text)) = 0, 0, Val(Me.txtConsumablesPerMonth.Text)) '* Val(Me.txtMajorEquipQty.Text))
            If UCase(Category_Shortname) = UCase("Concrete_") Then
                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = IIf(mConcreteQty = 0, 0, mConcreteQty)
            End If
            If Trim(xlWorksheet.Name) <> Trim(PrevSheetname) Then
                PrevSheetname = xlWorksheet.Name
            End If
            If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
            MsgBox("Record No. " & RecordsInserted(SheetNo) & " added in " & xlWorksheet.Name)
        End If
        'End
        InsertCommand = ""
        mcategory = Me.cmbMajorEqipCategory.Text
        InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
        InsertCommand = InsertCommand & "'" & Me.cmbMajorEquipName.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.cmbMajorEquipCapacity.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.cmbMajorEquipMake.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.cmbMajorEquipModel.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.dpMobDate.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.dpDemobDate.Text & "',"
        InsertCommand = InsertCommand & Me.txtMajorEquipQty.Text & ")"

        Try
            If (moledbConnection1.State.ToString().Equals("Closed")) Then
                moledbConnection1.Open()
            End If
            moleDBInsertCommand = New OleDbCommand
            moleDBInsertCommand.CommandType = CommandType.Text
            moleDBInsertCommand.CommandText = InsertCommand
            moleDBInsertCommand.Connection = moledbConnection1
            moleDBInsertCommand.ExecuteNonQuery()
            'MsgBox(Me.txtProjectdescription.Text & "Record inserted in list of projects.")
        Catch ex As Exception
            MsgBox(e.ToString())
        Finally
            moleDBInsertCommand = Nothing
        End Try
        xlApp.CalculateBeforeSave = True
        xlWorkbook.Save()
        ClearFields()
        Me.Button1.Enabled = False
        Me.txtFuelPerunitpermonth.Text = ""
        Me.txtFuelCostPerMonth.Text = ""
        Me.txtFuelCostProject.Text = ""
        Me.txtPowerCostPerMonth.Text = ""
        Me.txtPowerCostProject.Text = ""
        Me.txtOprCostPerMonth.Text = ""
        Me.txtoprCostProject.Text = ""
        Me.txtConsumablesPerMonth.Text = ""
        Me.txtConsumblesProject.Text = ""
        Me.txtConsumblesProject.Text = ""
        EditOrDelete = ""
    End Sub
    Private Sub ClearFields()
        Me.cmbMajorEquipName.Text = ""
        Me.cmbMajorEquipName.Text = ""
        Me.cmbMajorEquipCapacity.Text = ""
        Me.cmbMajorEquipModel.Text = ""
        Me.cmbMajorEquipMake.Text = ""
        Me.txtMajorEquipQty.Text = 0
        Me.dpMobDate.Text = mStartDate.Date
        Me.dpDemobDate.Text = mEndDate.Date
        Me.txtMonths.Text = 0
        Me.cmbDrive.SelectedIndex = 0
        Me.txtRepvalue.Text = 0
        Me.cmbDepPercentage.SelectedIndex = 0
        Me.txtDepreciation.Text = 0
        Me.txtPropHrs.Text = 0
        Me.txtFuelPerHr.Text = 0
        Me.txtPowerPerHr.Text = 0
        Me.txtOprCostPerMCPerMonth.Text = 0
        Me.txtConsumablesPerMonth.Text = 0
        Me.cmbMinorequip_Shifts.SelectedIndex = 0
        Me.txtMaintPercPerMC_PerMonth.Text = 0
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.tbclEqipsEntry.SelectTab(1)
        Currenttab = 1
    End Sub
    Private Sub FillFormulas(ByVal sheetno As Integer, ByVal record As Integer)

        xlWorksheet = xlWorkbook.Sheets(sheetno)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "Months").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)

        xlRange = xlWorksheet.Range(Category_Shortname & "Depreciation").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)

        If Category_Shortname = "Ext_" Or Category_Shortname = "ExtOthers_" Then
            xlRange = xlWorksheet.Range(Category_Shortname & "Hire_Charges_PerMonth").Offset(1, 0)
            xlRange.Copy()
            xlRange = xlRange.Offset(record - 1, 0)
            xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)
        End If

        xlRange = xlWorksheet.Range(Category_Shortname & "Hire_Charges").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)


        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Per_Mc").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        If Category_Shortname = "Ext_" And Me.cmbExtEquip_Category.Text = "Conveyance" Then
            xlRange.Formula = "=RC[-2]/RC[-1]"
        ElseIf Category_Shortname = "Conv_" Then
            xlRange.Formula = "=RC[-2]/RC[-1]"
        ElseIf Category_Shortname = "MH_" And (Me.cmbMajorEquipName.Text = "Tipper" Or Me.cmbMajorEquipName.Text = "Truck") Then
            xlRange.Formula = "=RC[-2]/RC[-1]"
        Else
            xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)
        End If

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Total").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_For_Project").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)


        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_Project").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)

        xlRange = xlWorksheet.Range(Category_Shortname & "Power_Per_Month").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)

        xlRange = xlWorksheet.Range(Category_Shortname & "Power_Cost_Project").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)

        xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)

        xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Project").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)

        xlRange = xlWorksheet.Range(Category_Shortname & "Consumables_Project").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)
        xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Month").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)

        xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Project").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)

        If UCase(Trim(xlWorksheet.Name)) = UCase("Minor Eqpts") Then
            xlRange = xlWorksheet.Range(Category_Shortname & "PurchCost_Project").Offset(1, 0)
            xlRange.Copy()
            xlRange = xlRange.Offset(record - 1, 0)
            xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)
        End If
    End Sub
    Private Sub GetTotals()
        Dim FormulaString As String
        Dim intI As Integer

        For SheetNo = 2 To 8
            With xlApp
                If xlWorkbook Is Nothing Then
                    xlWorkbook = .Workbooks.Open(xlFilename)
                End If
            End With
            xlWorksheet = xlWorkbook.Sheets(SheetNo)
            xlWorksheet.Select()
            If RecordsInserted(SheetNo) > 1 Then
                xlRange = xlWorksheet.Range(Category_Shortname & "Equip_Name")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = "Total"
                xlRange = xlWorksheet.Range(Category_Shortname & "Depreciation")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString

                xlRange = xlWorksheet.Range(Category_Shortname & "Hire_Charges")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString

                xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Per_Mc")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString

                xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Total")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString

                xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_For_Project")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString

                xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_Project")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange = xlWorksheet.Range(Category_Shortname & "Power_Per_Month")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange = xlWorksheet.Range(Category_Shortname & "Power_Cost_Project")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Project")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange = xlWorksheet.Range(Category_Shortname & "Consumables_Project")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString

                xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Month")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString

                xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Project")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString

                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(1, 0)
                x1 = xlRange.Address
                xlRange.Select()
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight)
                x2 = xlRange.Address
                xlRange = xlWorksheet.Range(x1, x2)
                RangeCols = xlRange.Columns.Count() - 1
                With xlRange
                    .BorderAround(, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                    .Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .Interior.ColorIndex = 2
                    .Font.Bold = False
                    '.Interior.Pattern = 1
                End With
                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(1, 0)
                x1 = xlRange.Address
                xlRange.Select()
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange = xlRange.Offset(1, 0)
                xlRange = xlRange.Offset(0, RangeCols)
                x2 = xlRange.Address
                xlRange = xlWorksheet.Range(x1, x2)
                With xlRange
                    .BorderAround(, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                    .Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                End With

                xlRange = xlWorksheet.Range(Category_Shortname & "Slno")
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange = xlRange.Offset(1, 1)
                x1 = xlRange.Address
                xlRange = xlRange.Offset(0, RangeCols - 1)
                x2 = xlRange.Address
                xlRange = xlWorksheet.Range(x1, x2)
                With xlRange
                    .Font.Size = 12
                    .Font.Bold = True
                End With
            End If
        Next
    End Sub
    Private Sub btnQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuit.Click
        Me.lblmessage.Visible = True
        Me.Refresh()
        GetTotals()
        GetTotals_MinorEquipments()
        GetTotals_ExternalEquipments()

        Dim intI As Integer
        For intI = 2 To 12
            xlWorksheet = xlWorkbook.Sheets(intI)
            getCategoryShortname(xlWorksheet)
            xlWorksheet.Range(Category_Shortname & "RecordsTotal").Value = RecordsInserted(intI)
        Next

        SetRangeNamesForTotals()

        If Not moledbConnection.State.ToString().Equals("Closed") Then
            moledbConnection.Close()
            moledbConnection = Nothing
        End If
        If Not xlWorkbook Is Nothing Then
            xlWorksheet = xlWorkbook.Sheets.Item(1)
            xlWorksheet.Select()
            xlApp.CalculateBeforeSave = True
            xlWorkbook.Save()
            xlWorkbook.Close()
            xlWorkbook = Nothing
        End If
        Me.Close()
        xlApp.Quit()
        xlApp = Nothing
        System.GC.Collect()
        frmProjectDetails.Show()
    End Sub
    Private Sub SetRangeNamesForTotals()
        Dim intI As Integer, CatShortnames(12) As String, Rangenamestring As String   ', CellAddress As String, Addstring As String
        Sheetnames(0) = "None"
        For intI = 1 To 12
            xlWorksheet = xlWorkbook.Sheets(intI)
            Sheetnames(intI) = xlWorksheet.Name
            getCategoryShortname(xlWorksheet)
            CatShortnames(intI) = Category_Shortname
        Next

        For intI = 2 To 11
            xlWorksheet = xlWorkbook.Sheets(intI)
            xlWorksheet.Select()
            With xlWorksheet
                If CatShortnames(intI) = "Ext_" Then
                    Continue For
                End If
                Rangenamestring = CatShortnames(intI) & "HireChargesTotal"
                xlWorkbook.Names.Item(Rangenamestring).Delete()
                xlRange = .Range(CatShortnames(intI) & "Hire_Charges")
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange.Name = Rangenamestring

                If CatShortnames(intI) = "Ext_" Or CatShortnames(intI) = "ExtOthers" Then
                    Rangenamestring = CatShortnames(intI) & "HireChargesPerMonthTotal"
                    xlWorkbook.Names.Item(Rangenamestring).Delete()
                    xlRange = .Range(CatShortnames(intI) & "Hire_Charges_PerMonth")
                    xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                    xlRange.Name = Rangenamestring
                End If

                Rangenamestring = CatShortnames(intI) & "HireChargesTotal"
                xlWorkbook.Names.Item(Rangenamestring).Delete()
                xlRange = .Range(CatShortnames(intI) & "Hire_Charges")
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange.Name = Rangenamestring

                Rangenamestring = CatShortnames(intI) & "FuelPerMonthPerMCTotal"
                xlWorkbook.Names.Item(Rangenamestring).Delete()
                xlRange = .Range(CatShortnames(intI) & "Fuel_Per_Month_Per_Mc")
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange.Name = Rangenamestring

                Rangenamestring = CatShortnames(intI) & "FuelPerMonthTotal"
                xlWorkbook.Names.Item(Rangenamestring).Delete()
                xlRange = .Range(CatShortnames(intI) & "Fuel_Per_Month_Total")
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange.Name = Rangenamestring

                Rangenamestring = CatShortnames(intI) & "FuelForProjectTotal"
                xlWorkbook.Names.Item(Rangenamestring).Delete()
                xlRange = .Range(CatShortnames(intI) & "Fuel_For_Project")
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange.Name = Rangenamestring

                Rangenamestring = CatShortnames(intI) & "FuelCostPerMonthTotal"
                xlWorkbook.Names.Item(Rangenamestring).Delete()
                xlRange = .Range(CatShortnames(intI) & "Fuel_Cost_per_Month")
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange.Name = Rangenamestring

                Rangenamestring = CatShortnames(intI) & "FuelCostProjectTotal"
                xlWorkbook.Names.Item(Rangenamestring).Delete()
                xlRange = .Range(CatShortnames(intI) & "Fuel_Cost_Project")
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange.Name = Rangenamestring

                Rangenamestring = CatShortnames(intI) & "PowerPerMonthTotal"
                xlWorkbook.Names.Item(Rangenamestring).Delete()
                xlRange = .Range(CatShortnames(intI) & "Power_Per_Month")
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange.Name = Rangenamestring

                Rangenamestring = CatShortnames(intI) & "PowerCostProjectTotal"
                xlWorkbook.Names.Item(Rangenamestring).Delete()
                xlRange = .Range(CatShortnames(intI) & "Power_Cost_Project")
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange.Name = Rangenamestring

                Rangenamestring = CatShortnames(intI) & "OprCostProjectTotal"
                xlWorkbook.Names.Item(Rangenamestring).Delete()
                xlRange = .Range(CatShortnames(intI) & "OprCost_Project")
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange.Name = Rangenamestring

                Rangenamestring = CatShortnames(intI) & "ConsumablesProjectTotal"
                xlWorkbook.Names.Item(Rangenamestring).Delete()
                xlRange = .Range(CatShortnames(intI) & "Consumables_Project")
                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange.Name = Rangenamestring

                If CatShortnames(intI) = "Min_" Then
                    Rangenamestring = CatShortnames(intI) & "PurchaseCostProjectTotal"
                    xlWorkbook.Names.Item(Rangenamestring).Delete()
                    xlRange = .Range(CatShortnames(intI) & "PurchCost_Project")
                    xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                    xlRange.Name = Rangenamestring
                End If
            End With
        Next
    End Sub
    Private Sub GetTotals_MinorEquipments()
        Dim FormulaString As String
        Dim intI As Integer, currentsheetname1 As String, currentsheetname2 As String
        currentsheetname2 = "Minor Eqpts"
        xlWorksheet = xlWorkbook.Sheets(currentsheetname2)
        SheetNo = getSheetNo(currentsheetname2)
        xlWorksheet.Select()
        If RecordsInserted(SheetNo) > 1 Then
            For intI = 2 To RecordsInserted(SheetNo)
                FillFormulas(SheetNo, intI)
            Next intI
            Category_Shortname = "Min_"
            xlRange = xlWorksheet.Range(Category_Shortname & "Equip_Name")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            xlRange.Value = "Total"
            xlRange = xlWorksheet.Range(Category_Shortname & "Depreciation")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "Hire_Charges")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Total")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_For_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
            xlRange = xlWorksheet.Range(Category_Shortname & "Power_Per_Month")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
            xlRange = xlWorksheet.Range(Category_Shortname & "Power_Cost_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
            xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
            xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "Consumables_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Month")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
            xlRange = xlWorksheet.Range(Category_Shortname & "PurchCost_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString


            xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(1, 0)
            x1 = xlRange.Address
            xlRange.Select()
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight)
            x2 = xlRange.Address
            xlRange = xlWorksheet.Range(x1, x2)
            RangeCols = xlRange.Columns.Count() - 1
            With xlRange
                .BorderAround(, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                .Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .Interior.ColorIndex = 2
                .Font.Bold = False
                '.Interior.Pattern = 1
            End With
            xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(1, 0)
            x1 = xlRange.Address
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
            xlRange = xlRange.Offset(1, 0)
            xlRange = xlRange.Offset(0, RangeCols)
            x2 = xlRange.Address
            xlRange = xlWorksheet.Range(x1, x2)
            With xlRange
                .BorderAround(, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                .Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            End With

            xlRange = xlWorksheet.Range(Category_Shortname & "Slno")
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
            xlRange = xlRange.Offset(1, 1)
            x1 = xlRange.Address
            xlRange = xlRange.Offset(0, RangeCols - 1)
            x2 = xlRange.Address
            xlRange = xlWorksheet.Range(x1, x2)
            With xlRange
                .Font.Size = 12
                .Font.Bold = True
            End With

            xlRange = xlWorksheet.Range(Category_Shortname & "Equip_Name")
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
            x1 = xlRange.Address
            xlRange = xlRange.Offset(-1, 0)
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight)
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
            x2 = xlRange.Address
            xlRange = xlWorksheet.Range(x1, x2)
            With xlRange
                .Font.Size = 12
                .Font.Bold = True
            End With
        End If
    End Sub
    Private Sub GetTotals_ExternalEquipments()
        Dim FormulaString As String
        Dim intI As Integer, currentsheetname1 As String, currentsheetname2 As String
        currentsheetname2 = "external Hire"
        For intI = 0 To SheetsCount - 1
            currentsheetname1 = Sheetnames(intI + 1)
            If UCase(Trim(currentsheetname1)) = UCase(Trim(currentsheetname2)) Then
                PrevSheetname = xlWorksheet.Name
                xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
                SheetNo = intI + 1
                Exit For
            End If
        Next
        xlWorksheet = xlWorkbook.Sheets(SheetNo)
        xlWorksheet.Select()
        If RecordsInserted(SheetNo) > 1 Then
            For intI = 2 To RecordsInserted(SheetNo)
                FillFormulas(SheetNo, intI)
            Next intI
            xlRange = xlWorksheet.Range(Category_Shortname & "Equip_Name")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            xlRange.Value = "Total"
            xlRange = xlWorksheet.Range(Category_Shortname & "Depreciation")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "Hire_Charges_PerMonth")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "Hire_Charges")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Per_Mc")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Total")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_For_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
            xlRange = xlWorksheet.Range(Category_Shortname & "Power_Per_Month")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
            xlRange = xlWorksheet.Range(Category_Shortname & "Power_Cost_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
            xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
            xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
            xlRange = xlWorksheet.Range(Category_Shortname & "Consumables_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Month")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Project")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString

            xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(1, 0)
            x1 = xlRange.Address
            'xlRange.Select()
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight)
            x2 = xlRange.Address
            xlRange = xlWorksheet.Range(x1, x2)
            Dim RangeCols As Integer = xlRange.Columns.Count() - 1
            With xlRange
                .BorderAround(, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                .Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            End With
            xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(1, 0)
            x1 = xlRange.Address
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
            xlRange = xlRange.Offset(1, 0)
            xlRange = xlRange.Offset(0, RangeCols)
            x2 = xlRange.Address
            xlRange = xlWorksheet.Range(x1, x2)
            With xlRange
                .BorderAround(, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                .Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            End With

            xlRange = xlWorksheet.Range(Category_Shortname & "Equip_Name")
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
            x1 = xlRange.Address
            xlRange = xlRange.Offset(-1, 0)
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight)
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
            x2 = xlRange.Address
            xlRange = xlWorksheet.Range(x1, x2)
            With xlRange
                .Font.Size = 12
                .Font.Bold = True
            End With
        End If
    End Sub

    Private Sub cmbMinorEquip_Model_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMinorEquip_Model.SelectedIndexChanged
        Dim strSql As String
        strSql = "categoryname = '" & Me.cmbMinoEquip_Category.Text & "' And EquipmentName = '" & Me.cmbMinorEquip_Name.Text & "' and Capacity = '" & _
           Me.cmbMinorEquip_capacity.Text & "' and Make = '" & Me.cmbMinorEquip_Make.Text & "' And  Model = '" & Me.cmbMinorEquip_Model.Text & "'"
        strStatement = "Select * from MinorEquipments where " & strSql
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        'create a data adapter
        mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
        'create a dataset
        mDataSet = New DataSet()
        'fill the dataset using data adapter
        mOledbDataAdapter.Fill(mDataSet, "MinorEquipments")
        Dim Machine As DataRow
        For Each Machine In mDataSet.Tables("MinorEquipments").Rows
            Me.txtMinorequip_NewEquipCost.Text = Machine("CostofnewEquipment").ToString()
            If Val(Machine("Depreciation_Fixed").ToString()) > 0 Then
                Me.txtMinorequip_Depreciation.Text = Machine("Depreciation_Fixed").ToString()
                Me.cmbMinorequip_DepPerc.Text = "Fixed"
                Me.cmbMinorequip_DepPerc.Enabled = False
            Else
                Me.txtMinorequip_Depreciation.Text = 0
                Me.cmbMinorequip_DepPerc.SelectedIndex = 0
                Me.cmbMinorequip_DepPerc.Enabled = True
            End If
            Me.txtMinorequip_PropHours.Text = Machine("Hrs_PerMonth").ToString()
            Me.txtMinorequip_FuelPerhr.Text = Machine("Fuel_PerHour").ToString()
            Me.txtMinorequip_PowerPerHr.Text = Machine("Power_PerHour").ToString()
            Me.txtMinorequip_OprCostPerMCPerMonth.Text = Machine("OperatorCost_PerMonth").ToString()
            Me.txtMinorequip_ConsumablesPerMCPerMonth.Text = Machine("minMaintCostPerc")
            Me.Label124.Text = "(Between " & mMinMaintperc & " and " & mMaxMaintPerc & ")"
            Me.txtMinorequip_ConsumablesPerMCPerMonth.Text = Val(Me.txtMinorequip_NewEquipCost.Text) * Val(Me.txtMinorequip_ConsumablesPercPerMCPerMonth.Text) / 100            'Me.txtMinorequip_NewEquipCost.Enabled = False

            If (Val(Me.txtMinorequip_FuelPerhr.Text) = 0 Or Len(Trim(txtMinorequip_FuelPerhr.Text)) = 0) Then
                Me.cmbMinorEquip_Drive.Text = "Electrical"
            Else
                Me.cmbMinorEquip_Drive.Text = "Fuel (Diesel)"
            End If
            mOledbDataAdapter = Nothing
        Next
    End Sub

    Private Sub cmbMinoEquip_Category_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMinoEquip_Category.SelectedIndexChanged
        TemplatePath = (My.Application.Info.DirectoryPath)
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
        moledbConnection = New OleDbConnection(strConnection)
        ' MsgBox(moledbConnection.State.ToString())
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strStatement = "Select distinct Equipmentname from MinorEquipments where Categoryname = '" & Me.cmbMinoEquip_Category.Text & "'"
        moledbCommand = New OleDbCommand(strStatement, moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbMinorEquip_Name.Items.Clear()
        If mReader.HasRows Then
            mReader.Read()
            Do
                Me.cmbMinorEquip_Name.Items.Add(mReader("EquipmentName"))
            Loop While mReader.Read()
        End If
        moledbCommand = Nothing
        mReader.Close()
        Me.btnViewItemsMinor.Text = "Click here to view items already added  from " & Me.cmbMinoEquip_Category.Text & " category"
    End Sub

    Private Sub cmbMinorEquip_Name_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMinorEquip_Name.SelectedIndexChanged
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        Dim strSql As String
        strSql = "categoryname = '" & Me.cmbMinoEquip_Category.Text & "' And EquipmentName = '" & Me.cmbMinorEquip_Name.Text & "'"
        strStatement = "Select distinct Capacity From MinorEquipments where " & strSql
        moledbCommand = New OleDbCommand(strStatement, moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbMinorEquip_capacity.Items.Clear()
        If mReader.HasRows Then
            mReader.Read()
            Do
                Me.cmbMinorEquip_capacity.Items.Add(mReader("Capacity"))
            Loop While mReader.Read()
        End If
        moledbCommand = Nothing
        mReader.Close()
    End Sub

    Private Sub cmbMinorEquip_capacity_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMinorEquip_capacity.SelectedIndexChanged
        Dim strSql As String
        strSql = "categoryname = '" & Me.cmbMinoEquip_Category.Text & "' And EquipmentName = '" & Me.cmbMinorEquip_Name.Text & "' and Capacity = '" & _
           Me.cmbMinorEquip_capacity.Text & "'"
        strStatement = "Select distinct Make From MinorEquipments where " & strSql
        moledbCommand = New OleDbCommand(strStatement, moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbMinorEquip_Make.Items.Clear()

        If mReader.HasRows Then
            mReader.Read()
            Do
                Me.cmbMinorEquip_Make.Items.Add(mReader("Make"))
            Loop While mReader.Read()
        End If
        moledbCommand = Nothing
        mReader.Close()
    End Sub

    Private Sub cmbMinorEquip_Make_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMinorEquip_Make.SelectedIndexChanged
        Dim strSql As String
        strSql = "categoryname = '" & Me.cmbMinoEquip_Category.Text & "' And EquipmentName = '" & Me.cmbMinorEquip_Name.Text & "' and Capacity = '" & _
           Me.cmbMinorEquip_capacity.Text & "' and Make = '" & Me.cmbMinorEquip_Make.Text & "'"
        strStatement = "Select distinct Model From MinorEquipments where " & strSql
        moledbCommand = New OleDbCommand(strStatement, moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbMinorEquip_Model.Items.Clear()

        If mReader.HasRows Then
            mReader.Read()
            Do
                Me.cmbMinorEquip_Model.Items.Add(mReader("Model"))
            Loop While mReader.Read()
        End If
        moledbCommand = Nothing
        mReader.Close()
    End Sub

    Private Sub dpMinorEquip_MobDate_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dpMinorEquip_MobDate.Validated
        If (Me.dpMinorequip_Demobdate.Value > Me.dpMinorEquip_MobDate.Value) Then
            Me.txtMinorequip_Months.Text = System.Math.Round((Me.dpMinorequip_Demobdate.Value.Date - Me.dpMinorEquip_MobDate.Value.Date).Days / 30, 0)
        Else
            Me.txtMonths.Text = 0
        End If
    End Sub

    Private Sub dpMinorEquip_MobDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dpMinorEquip_MobDate.Validating
        If Me.dpMinorEquip_MobDate.Value.Date < mStartDate.Date Then
            MsgBox("Mobilisation Date cannot be past date. Please select again")
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub dpMinorequip_Demobdate_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dpMinorequip_Demobdate.Validated
        If (Me.dpMinorequip_Demobdate.Value > Me.dpMinorEquip_MobDate.Value) Then
            Me.txtMinorequip_Months.Text = System.Math.Round((Me.dpMinorequip_Demobdate.Value.Date - Me.dpMinorEquip_MobDate.Value.Date).Days / 30, 0)
        Else
            Me.txtMonths.Text = 0
        End If
    End Sub

    Private Sub dpMinorequip_Demobdate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dpMinorequip_Demobdate.Validating
        If Me.dpMinorequip_Demobdate.Value.Date <= Me.dpMinorEquip_MobDate.Value.Date Then
            MsgBox("De-mobilisation Date cannot be past date. Please select again")
            e.Cancel = True
        End If
        If Me.dpMinorequip_Demobdate.Value < mStartDate.Date Then
            MsgBox("De-Mobilisation Date cannot be less than Project Start date. Please select again")
            e.Cancel = True
        End If
        If Me.dpMinorequip_Demobdate.Value.Date > mEndDate Then
            MsgBox("De-mobilization date cannot be beyond project end date")
            Me.dpMinorequip_Demobdate.Value = mEndDate.Date
            e.Cancel = True
        End If
    End Sub

    Private Sub cmbMinorequip_DepPerc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMinorequip_DepPerc.SelectedIndexChanged
        Me.txtMinorequip_Depreciation.Text = Val(Me.txtMinorequip_NewEquipCost.Text) * Val(Me.cmbMinorequip_DepPerc.Text) / 100
        Me.txtMinorEquip_HireCharges.Text = Val(Me.txtMinorequip_Depreciation.Text) * Val(Me.txtMinorEquip_Qty.Text) * Val(Me.txtMinorequip_Months.Text)
    End Sub

    Private Sub btnMinorequip_ComputeValues_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMinorequip_ComputeValues.Click
        Dim AllValid As Boolean
        AllValid = True
        Dim msgstring As String
        msgstring = "The following were not entered/selected" & vbNewLine
        If (Me.cmbMinorEquip_capacity.Text = "") Then
            msgstring = msgstring & "Equipment Category not selected"
            AllValid = False
        ElseIf (Me.cmbMinorEquip_Name.Text = "") Then
            msgstring = msgstring & "Equipment Name not selected" & vbNewLine
            AllValid = False
        ElseIf (Me.cmbMinorEquip_Drive.Text = "") Then
            msgstring = msgstring & "Equipment drive not selected" & vbNewLine
            AllValid = False
        ElseIf (Me.cmbMinorEquip_capacity.Text) = "" Then
            msgstring = msgstring & "Equipment capacity not selected" & vbNewLine
            AllValid = False
        ElseIf (Me.cmbMinorEquip_Make.Text = "") Then
            msgstring = msgstring & "Equipment Make not selected" & vbNewLine
            AllValid = False
        ElseIf (Me.cmbMinorEquip_Model.Text = "") Then
            msgstring = msgstring & "Equipment Model not selected" & vbNewLine
            AllValid = False
        ElseIf (Me.txtMinorEquip_Qty.Text = "" Or Len(Trim(Me.txtMinorEquip_Qty.Text)) = 0) Then
            msgstring = msgstring & "Equipment quantity  not entered" & vbNewLine
            AllValid = False
        ElseIf (Me.dpMinorEquip_MobDate.Value.ToString() = "") Then
            msgstring = msgstring & "Equipment mobilisation date not selected" & vbNewLine
            AllValid = False
        ElseIf (dpMinorequip_Demobdate.Value.ToString() = "") Then
            msgstring = msgstring & "Equipment demobilisation date  not selected" & vbNewLine
            AllValid = False
        ElseIf Trim(Me.cmbMinorEquip_Drive.Text) = "Fuel(Diesel)" Then
            If (Me.txtMinorequip_FuelPerhr.Text = "") Then
                msgstring = msgstring & "*** Fuel consumption per Hour missing"
                AllValid = False
            End If
            If (Val(Me.txtMinorequip_FuelCostperLtr.Text) = 0 Or Me.txtMinorequip_FuelCostperLtr.Text = "") Then
                msgstring = msgstring & "*** Fuel cost per Ltr not entered or is zero" & vbNewLine
                AllValid = False
            End If
        ElseIf Trim(Me.cmbMinorEquip_Drive.Text) = "Electrical()" Then
            If (Me.txtMinorequip_PowerPerHr.Text = "") Then
                msgstring = msgstring & "*** Power Consumption Per Hr missing" & vbNewLine
                AllValid = False
            End If
            If (Val(Me.txtMinorequip_PowerCostPerUnit.Text) = 0 Or Me.txtMinorequip_PowerCostPerUnit.Text = "") Then
                msgstring = msgstring & "*** Power Cost per Unit is not entered or is zero"
                AllValid = False
            End If
        End If

        If Not AllValid Then
            MsgBox(msgstring)
        Else
            Me.txtMinorequip_FuelperMCPerMonth.Text = Val(Me.txtMinorequip_FuelPerhr.Text) * Val(Me.txtMinorequip_PropHours.Text)
            Me.txtMinorequip_FuelCostPerMonth.Text = Val(txtMinorequip_FuelperMCPerMonth.Text) * Val(Me.txtMinorEquip_Qty.Text) * Val(Me.txtMinorequip_FuelCostperLtr.Text)
            Me.txtMinorequip_FuelCostProject.Text = Val(txtMinorequip_FuelCostPerMonth.Text) * Val(Me.txtMinorequip_Months.Text)
            Me.txtMinorequip_PowerCostPerMonth.Text = Val(Me.txtMinorequip_PowerPerHr.Text) * Val(Me.txtMinorequip_PropHours.Text) * Val(Me.txtMinorEquip_Qty.Text) * Val(Me.txtMinorequip_PowerCostPerUnit.Text)
            Me.txtMinorequip_PowerCostProject.Text = Val(txtMinorequip_PowerCostPerMonth.Text) * Val(Me.txtMinorEquip_Qty.Text)
            Me.txtMinorequip_OprCostPerMonth.Text = Val(Me.txtMinorequip_OprCostPerMCPerMonth.Text) * Val(Me.txtMinorEquip_Qty.Text) * Val(Me.cmbMinorequip_Shifts.Text)
            Me.txtMinorequip_OprCostProject.Text = Val(txtMinorequip_OprCostPerMonth.Text) * Val(Me.txtMinorequip_Months.Text)
            Me.txtMinorequip_ConxumablesPerMonth.Text = Val(Me.txtMinorequip_ConsumablesPerMCPerMonth.Text) * Val(Me.txtMinorEquip_Qty.Text)
            Me.txtMinorequip_ConxumablesProject.Text = Val(txtMinorequip_ConxumablesPerMonth.Text) * Val(Me.txtMinorequip_Months.Text)
            Me.txtMinorequip_OperatingCost.Text = Val(Me.txtMinorEquip_HireCharges.Text) + Val(Me.txtMinorequip_FuelCostProject.Text) + Val(Me.txtPowerCostProject.Text) + _
                  Val(Me.txtoprCostProject.Text) + Val(Me.txtConsumblesProject.Text)
        End If
        Me.bbtnMinorequip_SaveEntry.Enabled = True
    End Sub

    Private Sub bbtnMinorequip_SaveEntry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bbtnMinorequip_SaveEntry.Click
        Dim intI As Integer
        Dim currentsheetname1, currentsheetname2 As String
        Dim moleDBinsertCommand As OleDbCommand

        Me.btnQuit.Enabled = True
        Me.btnClose.Enabled = False
        If Len(Trim(xlFilename)) = 0 Then
            MsgBox("Filename not specified to save. cannot Save data. Start the application again.")
            Exit Sub
        End If

        If Me.dpMinorEquip_MobDate.Value.Date > Me.dpMinorequip_Demobdate.Value.Date Then
            MsgBox("Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
              "And mobilisation and demobilisation dates must fall within Projet Start and End dates")
            Me.dpMinorequip_Demobdate.Value = mStartDate.Date
            Me.dpMinorEquip_MobDate.Focus()
            Exit Sub
        End If

        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With
        SheetsCount = 12     '= xlWorkbook.Sheets.Count

        If Startup Then
            For intI = 1 To SheetsCount
                xlWorksheet = xlWorkbook.Sheets.Item(intI)
                Sheetnames(intI) = (xlWorksheet.Name)
                SheetIndices(intI) = intI
                getCategoryShortname(xlWorksheet)
                RecordsInserted(intI) = xlWorksheet.Range(Category_Shortname & "RecordsTotal").Value
            Next
            Startup = False
            Me.btnClose.Enabled = False
        End If
        currentsheetname2 = "Minor Eqpts"

        For intI = 0 To SheetsCount - 1
            currentsheetname1 = Sheetnames(intI + 1)
            If UCase(Trim(currentsheetname1)) = UCase(Trim(currentsheetname2)) Then
                'PrevSheetname = xlWorksheet.Name
                xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
                SheetNo = intI + 1
                Exit For
            End If
        Next

        xlWorksheet.Activate()
        Category_Shortname = "Min_"
        xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
        xlRange.Value = mMainTitle1
        xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
        xlRange.Value = mMainTitle2
        xlRange = xlWorksheet.Range(Category_Shortname & "Client")
        xlRange.Value = mClient
        xlRange = xlWorksheet.Range(Category_Shortname & "Location")
        xlRange.Value = mLocation
        xlRange = xlWorksheet.Range(Category_Shortname & "StartDate")
        xlRange.Value = mStartDate.Date.ToString()
        xlRange = xlWorksheet.Range(Category_Shortname & "EndDate")
        xlRange.Value = mEndDate.Date.ToString()
        xlRange = xlWorksheet.Range(Category_Shortname & "ProjectValue")
        xlRange.Value = mProjectvalue
        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
        xlRange.Value = "Fuel Cost for all mc/month @Rs. " & Me.txtMinorequip_FuelCostperLtr.Text & " per Lt"
        xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
        xlRange.Value = "Opr Cost for all m/c"   ' with " & Me.cmbMinorequip_Shifts.Text & " per day"
        RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1

        xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo), 0)
        If xlRange.Offset(1, 1).Value = "Total" Then
            xlWorksheet.Range(xlRange.Offset(1, 1), xlRange.Offset(1, 35)).ClearContents()
        End If
        xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo), 0)

        xlRange.Value = RecordsInserted(SheetNo)
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.cmbMinorEquip_Name.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.cmbMinorEquip_Drive.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.cmbMinorEquip_capacity.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.cmbMinorEquip_Make.Text & vbNewLine & "/" & Me.cmbMinorEquip_Model.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = IIf(Len(Trim(Me.txtMinorEquip_Qty.Text)) = 0, 0, Me.txtMinorEquip_Qty.Text)
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.dpMinorEquip_MobDate.Value.Date.ToString()
        xlRange.NumberFormat = "dd-mmm-yyyy"
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.dpMinorequip_Demobdate.Value.Date.ToString()
        xlRange.NumberFormat = "dd-mmm-yyyy"
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = IIf(Len(Trim(Me.txtMinorequip_NewEquipCost.Text)) = 0, 0, Me.txtMinorequip_NewEquipCost.Text)
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = IIf(Len(Trim(Me.cmbMinorequip_DepPerc.Text)) = 0, 0, Me.cmbMinorequip_DepPerc.Text)
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = IIf(Len(Trim(Me.txtMinorequip_PropHours.Text)) = 0, 0, Me.txtMinorequip_PropHours.Text)
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = IIf(Len(Trim(Me.txtMinorequip_FuelPerhr.Text)) = 0, 0, Me.txtMinorequip_FuelPerhr.Text)
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = IIf(Len(Trim(Me.txtMinorequip_FuelCostperLtr.Text)) = 0, 0, Me.txtMinorequip_FuelCostperLtr.Text)
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = IIf(Len(Trim(Me.txtMinorequip_PowerPerHr.Text)) = 0, 0, Me.txtMinorequip_PowerPerHr.Text)
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = IIf(Len(Trim(Me.txtMinorequip_PowerCostPerUnit.Text)) = 0, 0, Me.txtMinorequip_PowerCostPerUnit.Text)
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = IIf(Len(Trim(Me.txtMinorequip_OprCostPerMCPerMonth.Text)) = 0, 0, Me.txtMinorequip_OprCostPerMCPerMonth.Text)
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = IIf(Len(Trim(Me.cmbMinorequip_Shifts.Text)) = 0, 1, Me.cmbMinorequip_Shifts.Text)
        xlRange.NumberFormat = "#.0#"
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = IIf(Len(Trim(Me.txtMinorequip_ConxumablesPerMonth.Text)) = 0, 0, Val(Me.txtMinorequip_ConxumablesPerMonth.Text))  ' * Val(Me.txtMinorEquip_Qty.Text))

        If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
        MsgBox("Record No. " & RecordsInserted(SheetNo) & " added in " & xlWorksheet.Name)
        'End If
        Dim InsertCommand As String = ""
        mcategory = cmbMinoEquip_Category.Text
        InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
        InsertCommand = InsertCommand & "'" & Me.cmbMinorEquip_Name.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.cmbMinorEquip_capacity.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.cmbMinorEquip_Make.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.cmbMinorEquip_Model.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.dpMinorEquip_MobDate.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.dpMinorequip_Demobdate.Text & "',"
        InsertCommand = InsertCommand & Me.txtMinorEquip_Qty.Text & ")"
        'InsertCommand = InsertCommand & "'" & Me.dpEndDDeate.Text & "',"

        Try
            moleDBInsertCommand = New OleDbCommand
            moleDBInsertCommand.CommandType = CommandType.Text
            moledbInsertCommand.CommandText = InsertCommand
            moleDBInsertCommand.Connection = moledbConnection1
            moledbInsertCommand.ExecuteNonQuery()
            'MsgBox(Me.txtProjectdescription.Text & "Record inserted in list of projects.")
        Catch ex As Exception
            MsgBox(e.ToString())
        Finally
            moleDBinsertCommand = Nothing
        End Try
        ClearFields_MinorEquipments()
        Me.bbtnMinorequip_SaveEntry.Enabled = False
        Me.txtMinorequip_FuelperMCPerMonth.Text = ""
        Me.txtMinorequip_FuelCostPerMonth.Text = ""
        Me.txtMinorequip_FuelCostProject.Text = ""
        Me.txtMinorequip_PowerCostPerMonth.Text = ""
        Me.txtMinorequip_PowerCostProject.Text = ""
        Me.txtMinorequip_OprCostPerMonth.Text = ""
        Me.txtMinorequip_OprCostProject.Text = ""
        Me.txtMinorequip_ConxumablesPerMonth.Text = ""
        Me.txtMinorequip_ConxumablesProject.Text = ""
        Me.txtMinorequip_OperatingCost.Text = ""
        EditOrDelete = ""
    End Sub
    Private Sub ClearFields_MinorEquipments()
        Me.cmbMinorEquip_Name.Text = ""
        Me.cmbMinorEquip_capacity.Text = ""
        Me.cmbMinorEquip_Model.Text = ""
        Me.cmbMinorEquip_Make.Text = ""
        Me.cmbMinorEquip_Drive.SelectedIndex = 0
        Me.txtMinorEquip_Qty.Text = 0
        Me.dpMinorEquip_MobDate.Text = mStartDate.Date
        Me.dpMinorequip_Demobdate.Text = mEndDate.Date
        Me.txtMinorequip_Months.Text = 0
        Me.txtMinorequip_NewEquipCost.Text = 0
        Me.cmbMinorequip_DepPerc.Text = 0
        Me.txtMinorequip_Depreciation.Text = 0
        Me.txtMinorEquip_HireCharges.Text = 0
        Me.txtMinorequip_PropHours.Text = 0
        Me.txtMinorequip_FuelPerhr.Text = 0
        Me.txtMinorequip_PowerPerHr.Text = 0
        Me.txtMinorequip_OprCostPerMCPerMonth.Text = 0
        Me.txtMinorequip_ConxumablesPerMonth.Text = 0
        Me.txtMinorequip_OperatingCost.Text = 0
        Me.cmbMinorequip_Shifts.SelectedIndex = 0
        Me.txtMinorequip_ConsumablesPercPerMCPerMonth.Text = 0
    End Sub

    Private Sub cmbExtEquip_Category_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbExtEquip_Category.SelectedIndexChanged
        TemplatePath = (My.Application.Info.DirectoryPath)
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
        moledbConnection = New OleDbConnection(strConnection)
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strStatement = "Select distinct Equipmentname from HiredEquipments where Categoryname = '" & Me.cmbExtEquip_Category.Text & "'"
        moledbCommand = New OleDbCommand(strStatement, moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbExtEquip_Name.Items.Clear()
        If mReader.HasRows Then
            mReader.Read()
            Do
                Me.cmbExtEquip_Name.Items.Add(mReader("EquipmentName"))
            Loop While mReader.Read()
        End If
        moledbCommand = Nothing
        mReader.Close()
        Me.btnHiredItemsView.Text = "Click here to view items already added  from " & Me.cmbExtEquip_Category.Text & " category"
    End Sub

    Private Sub cmbExtEquip_Name_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbExtEquip_Name.SelectedIndexChanged
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        Dim strSql As String
        strSql = "Categoryname = '" & Me.cmbExtEquip_Category.Text & "' And EquipmentName = '" & Me.cmbExtEquip_Name.Text & "'"
        strStatement = "Select distinct Capacity From HiredEquipments where " & strSql
        moledbCommand = New OleDbCommand(strStatement, moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbExtEquip_Capacity.Items.Clear()
        If mReader.HasRows Then
            mReader.Read()
            Do
                Me.cmbExtEquip_Capacity.Items.Add(mReader("Capacity"))
            Loop While mReader.Read()
        End If
        moledbCommand = Nothing
        mReader.Close()
    End Sub

    Private Sub cmbExtEquip_Capacity_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbExtEquip_Capacity.SelectedIndexChanged
        Dim strSql As String
        strSql = "Categoryname = '" & Me.cmbExtEquip_Category.Text & "' And EquipmentName = '" & Me.cmbExtEquip_Name.Text & "' and Capacity = '" & _
           Me.cmbExtEquip_Capacity.Text & "'"
        strStatement = "Select distinct Make From HiredEquipments where " & strSql
        moledbCommand = New OleDbCommand(strStatement, moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbExtEquip_Make.Items.Clear()

        If mReader.HasRows Then
            mReader.Read()
            Do
                Me.cmbExtEquip_Make.Items.Add(mReader("Make"))
            Loop While mReader.Read()
        End If
        moledbCommand = Nothing
        mReader.Close()
    End Sub

    Private Sub cmbExtEquip_Model_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbExtEquip_Model.SelectedIndexChanged
        Dim strSql As String
        strSql = "Categoryname = '" & Me.cmbExtEquip_Category.Text & "' And EquipmentName = '" & Me.cmbExtEquip_Name.Text & "' and Capacity = '" & _
           Me.cmbExtEquip_Capacity.Text & "' and Make = '" & Me.cmbExtEquip_Make.Text & "' And  Model = '" & Me.cmbExtEquip_Model.Text & "'"
        strStatement = "Select * from HiredEquipments where " & strSql
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
        If moledbConnection Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        'create a data adapter
        mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
        'create a dataset
        mDataSet = New DataSet()
        'fill the dataset using data adapter
        mOledbDataAdapter.Fill(mDataSet, "HiredEquipments")
        Dim Machine As DataRow
        For Each Machine In mDataSet.Tables("HiredEquipments").Rows
            Me.txtExtEquip_RepValue.Text = Machine("RepValue").ToString()
            If Val(Machine("Depreciation_Fixed").ToString()) > 0 Then
                Me.txtExtEquip_Depreciation.Text = Machine("Depreciation_Fixed").ToString()
                Me.cmbExtEquip_DepPerc.Text = "Fixed"
                Me.cmbExtEquip_DepPerc.Enabled = False
            Else
                Me.txtExtEquip_Depreciation.Text = 0
                Me.cmbExtEquip_DepPerc.SelectedIndex = 0
                Me.cmbExtEquip_DepPerc.Enabled = True
            End If
            Me.txtExtEquip_PropHours.Text = Machine("Hrs_PerMonth").ToString()
            Me.txtExtEquip_FuelPerHr.Text = Machine("Fuel_PerHour").ToString()
            Me.txtExtEquip_PowerPerhr.Text = Machine("Power_PerHour").ToString()
            Me.txtExtEquip_OprCostPerMCPerMonth.Text = Machine("OperatorCost_PerMonth").ToString()
            Me.txtExtEquip_ConsumablesPerMCPerMonth.Text = Machine("MaintCost_PerMonth").ToString()
            If (Val(Me.txtExtEquip_FuelPerHr.Text) = 0 Or Len(Trim(txtExtEquip_FuelPerHr.Text)) = 0) Then
                Me.cmbExtEquip_Drive.Text = "Electrical"
            Else
                Me.cmbExtEquip_Drive.Text = "Fuel (Diesel)"
            End If
            mOledbDataAdapter = Nothing
        Next
    End Sub

    Private Sub cmbExtEquip_Make_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbExtEquip_Make.SelectedIndexChanged
        Dim strSql As String
        strSql = "Categoryname = '" & Me.cmbExtEquip_Category.Text & "' And EquipmentName = '" & Me.cmbExtEquip_Name.Text & "' and Capacity = '" & _
           Me.cmbExtEquip_Capacity.Text & "' and Make = '" & Me.cmbExtEquip_Make.Text & "'"
        strStatement = "Select distinct Model From HiredEquipments where " & strSql
        moledbCommand = New OleDbCommand(strStatement, moledbConnection)
        mReader = moledbCommand.ExecuteReader()
        Me.cmbExtEquip_Model.Items.Clear()

        If mReader.HasRows Then
            mReader.Read()
            Do
                Me.cmbExtEquip_Model.Items.Add(mReader("Model"))
            Loop While mReader.Read()
        End If
        moledbCommand = Nothing
        mReader.Close()
    End Sub

    Private Sub txtMinorEquip_Qty_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMinorEquip_Qty.Validating
        If Not IsNumeric(Me.txtMinorEquip_Qty.Text) Then
            MsgBox("Enter Numeric value for the Quantity")
            Me.txtMinorEquip_Qty.Text = ""
            e.Cancel = True
        Else
            Me.txtMinorEquip_Qty.Text = Int(Val(Me.txtMinorEquip_Qty.Text))
            Me.txtMinorEquip_HireCharges.Text = Val(Me.txtMinorequip_Depreciation.Text) * Val(Me.txtMinorEquip_Qty.Text) * Val(Me.txtMinorequip_Months.Text)
        End If
    End Sub

    Private Sub txtExtEquip_Qty_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtExtEquip_Qty.Validating
        If Not IsNumeric(Me.txtExtEquip_Qty.Text) Then
            MsgBox("Enter Numeric value for the Quantity")
            Me.txtExtEquip_Qty.Text = ""
            e.Cancel = True
        Else
            Me.txtExtEquip_Qty.Text = Int(Val(Me.txtExtEquip_Qty.Text))
        End If
    End Sub

    Private Sub dpExtEquip_MobDate_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dpExtEquip_MobDate.Validated
        If (Me.dpExtEquip_DemobDate.Value > Me.dpExtEquip_MobDate.Value) Then
            Me.txtExtEquip_Months.Text = System.Math.Round((Me.dpExtEquip_DemobDate.Value.Date - Me.dpExtEquip_MobDate.Value.Date).Days / 30, 0)
        Else
            Me.txtExtEquip_Months.Text = 0
        End If
    End Sub

    Private Sub dpExtEquip_MobDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dpExtEquip_MobDate.Validating
        If Me.dpExtEquip_MobDate.Value.Date < mStartDate.Date Then
            MsgBox("Mobilisation Date cannot be past date. Please select again")
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub dpExtEquip_MobDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dpExtEquip_MobDate.ValueChanged

    End Sub

    Private Sub dpExtEquip_DemobDate_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dpExtEquip_DemobDate.Validated
        If (Me.dpExtEquip_DemobDate.Value > Me.dpExtEquip_MobDate.Value) Then
            Me.txtExtEquip_Months.Text = System.Math.Round((Me.dpExtEquip_DemobDate.Value.Date - Me.dpExtEquip_MobDate.Value.Date).Days / 30, 0)
        Else
            Me.txtExtEquip_Months.Text = 0
        End If
    End Sub

    Private Sub dpExtEquip_DemobDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dpExtEquip_DemobDate.Validating
        If Me.dpExtEquip_DemobDate.Value.Date <= dpExtEquip_MobDate.Value.Date Then
            MsgBox("Mobilisation Date cannot be past date. Please select again")
            Me.dpExtEquip_DemobDate.Value = mEndDate.Date
            e.Cancel = True
            Exit Sub
        End If
        If Me.dpExtEquip_DemobDate.Value < mStartDate.Date Then
            MsgBox("De-Mobilisation Date cannot be less than Project Start date. Please select again")
            Me.dpExtEquip_DemobDate.Value = mEndDate.Date
            e.Cancel = True
            Exit Sub
        End If
        If Me.dpExtEquip_DemobDate.Value.Date > mEndDate Then
            MsgBox("De-mobilization date cannot be beyond project end date")
            Me.dpExtEquip_DemobDate.Value = mEndDate.Date
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub txtMinorequip_Months_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMinorequip_Months.Validated
        If Val(Me.txtMinorequip_Months.Text) <= 0 Then
            txtMinorequip_Months.Text = 1
        End If
    End Sub

    Private Sub txtMinorequip_Months_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMinorequip_Months.Validating
        If Not IsNumeric(txtMinorequip_Months.Text) Then
            MsgBox("Months should be entered as numeric")
            Me.txtMinorequip_Months.Focus()
            e.Cancel = True
        End If
    End Sub
    Private Sub txtExtEquip_Months_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExtEquip_Months.Validated
        If Val(txtExtEquip_Months.Text) <= 0 Then
            txtExtEquip_Months.Text = 1
        End If
    End Sub

    Private Sub txtExtEquip_Months_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtExtEquip_Months.Validating
        If Not IsNumeric(txtExtEquip_Months.Text) Then
            MsgBox("Months should be entered as numeric")
            Me.txtExtEquip_Months.Focus()
            e.Cancel = True
        End If
    End Sub

    Private Sub cmbExtEquip_DepPerc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbExtEquip_DepPerc.SelectedIndexChanged
        '    Dim mDepPerc As Single = 0
        '    Me.txtExtEquip_Depreciation.Text = Val(Me.txtExtEquip_RepValue.Text) * mDepPerc / 100
        '    Me.txtExtEquip_Hire_Charges.Text = Val(Me.txtExtEquip_Depreciation.Text) * Val(Me.txtExtEquip_Qty.Text) * Val(Me.txtExtEquip_Months.Text)
    End Sub

    Private Sub btnExtEquip_ComputeValues_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExtEquip_ComputeValues.Click
        Dim AllValid As Boolean
        AllValid = True
        Dim msgstring As String
        msgstring = "The following are missing. Cannot Proceed now." & vbNewLine
        If (Me.cmbExtEquip_Capacity.Text = "") Then
            msgstring = msgstring & "*** Equipment Category"
            AllValid = False
        End If
        If (Me.cmbExtEquip_Name.Text = "") Then
            msgstring = msgstring & "*** Equipment Name" & vbNewLine
            AllValid = False
        End If
        If (Me.cmbExtEquip_Drive.Text = "") Then
            msgstring = msgstring & "*** Equipment drive" & vbNewLine
            AllValid = False
        End If
        If (Me.cmbExtEquip_Capacity.Text) = "" Then
            msgstring = msgstring & "*** Equipment capacity" & vbNewLine
            AllValid = False
        End If
        If (Me.cmbExtEquip_Make.Text = "") Then
            msgstring = msgstring & "*** Equipment Make " & vbNewLine
            AllValid = False
        End If
        If (Me.cmbExtEquip_Model.Text = "") Then
            msgstring = msgstring & "*** Equipment Model" & vbNewLine
            AllValid = False
        End If
        If (Me.txtExtEquip_Qty.Text = "" Or Len(Trim(Me.txtExtEquip_Qty.Text)) = 0) Then
            msgstring = msgstring & "Equipment quantity  not entered" & vbNewLine
            AllValid = False
        End If
        If (Me.dpExtEquip_MobDate.Value.ToString() = "") Then
            msgstring = msgstring & "*** Equipment mobilisation date" & vbNewLine
            AllValid = False
        End If
        If (dpExtEquip_DemobDate.Value.ToString() = "") Then
            msgstring = msgstring & "*** Equipment demobilisation" & vbNewLine
            AllValid = False
        End If
        If Trim(Me.cmbExtEquip_Drive.Text) = "Fuel(Diesel)" Then
            If (Me.txtExtEquip_FuelPerHr.Text = "" Or Val(Me.txtExtEquip_FuelPerHr.Text) = 0) Then
                msgstring = msgstring & "*** Fuel consumption per Hour missing"
                AllValid = False
            End If
            If (Val(Me.txtExtEquip_FuelCostperLtr.Text) = 0 Or Me.txtExtEquip_FuelCostperLtr.Text = "") Then
                msgstring = msgstring & "*** Fuel cost per Ltr not entered or is zero" & vbNewLine
                AllValid = False
            End If
        End If
        If Trim(Me.cmbExtEquip_Drive.Text) = "Electrical()" Then
            If (Me.txtExtEquip_PowerPerhr.Text = "" Or Val(Me.txtExtEquip_PowerPerhr.Text) = 0) Then
                msgstring = msgstring & "*** Power Consumption Per Hr missing" & vbNewLine
                AllValid = False
            End If
            If (Val(Me.txtExtEquip_PowerCostPerUnit.Text) = 0 Or Me.txtExtEquip_PowerCostPerUnit.Text = "") Then
                msgstring = msgstring & "*** Power Cost per Unit is not entered or is zero"
                AllValid = False
            End If
        End If
        If Not AllValid Then
            MsgBox(msgstring)
        Else
            Me.txtExtEquip_FuelperMCPerMonth.Text = IIf(Me.cmbExtEquip_Category.Text = "Conveyance_Hired", Val(Me.txtExtEquip_PropHours.Text) / Val(Me.txtExtEquip_FuelPerHr.Text), Val(Me.txtExtEquip_PropHours.Text) * Val(Me.txtExtEquip_FuelPerHr.Text))
            Me.txtExtEquip_FuelCostPerMonth.Text = Val(txtExtEquip_FuelperMCPerMonth.Text) * Val(Me.txtExtEquip_Qty.Text) * Val(Me.txtExtEquip_FuelCostperLtr.Text)
            Me.txtExtEquip_FuelCostProject.Text = Val(txtExtEquip_FuelCostPerMonth.Text) * Val(Me.txtExtEquip_Months.Text)
            Me.txtExtEquip_PowerCostPerMonth.Text = Val(Me.txtExtEquip_PowerPerhr.Text) * Val(Me.txtExtEquip_PropHours.Text) * Val(Me.txtExtEquip_Qty.Text) * Val(Me.txtExtEquip_PowerCostPerUnit.Text)
            Me.txtExtEquip_PowerCostProject.Text = Val(txtExtEquip_PowerCostPerMonth.Text) * Val(Me.txtExtEquip_Months.Text)
            Me.txtExtEquip_OprCostPerMonth.Text = Val(Me.txtExtEquip_OprCostPerMCPerMonth.Text) * Val(Me.txtExtEquip_Qty.Text) * Val(Me.cmbExtEquip_Shifts.Text)
            Me.txtExtEquip_OprCostProject.Text = Val(txtExtEquip_OprCostPerMonth.Text) * Val(Me.txtExtEquip_Months.Text)
            Me.txtExtEquip_ConxumablesPerMonth.Text = Val(Me.txtExtEquip_ConsumablesPerMCPerMonth.Text) * Val(Me.txtExtEquip_Qty.Text)
            Me.txtExtEquip_ConxumablesProject.Text = Val(txtExtEquip_ConxumablesPerMonth.Text) * Val(Me.txtExtEquip_Months.Text)
            Me.txtExtEquip_OperatingCost.Text = Val(Me.txtExtEquip_Hire_Charges.Text) + Val(Me.txtExtEquip_FuelCostProject.Text) + Val(Me.txtExtEquip_PowerCostProject.Text) + _
                  Val(Me.txtExtEquip_OprCostProject.Text) + Val(Me.txtExtEquip_ConxumablesProject.Text)
        End If
        Me.btnExtEquip_SaveEntry.Enabled = True
    End Sub

    Private Sub btnExtEquip_SaveEntry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExtEquip_SaveEntry.Click
        Dim intI As Integer
        Dim currentsheetname1, currentsheetname2 As String
        Dim moleDBInsertCommand As OleDbCommand
        Me.btnQuit.Enabled = True
        Me.btnClose.Enabled = False
        If Len(Trim(xlFilename)) = 0 Then
            MsgBox("Filename not specified to save. cannot Save data. Start the application again.")
            Application.Exit()
        End If
        If Me.dpExtEquip_DemobDate.Value.Date < Me.dpExtEquip_MobDate.Value.Date Then
            MsgBox("Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
              "And mobilisation and demobilisation dates must fall within Projet Start and End dates")
            Me.dpExtEquip_DemobDate.Value = mEndDate.Date
            Me.dpExtEquip_DemobDate.Focus()
            Exit Sub
        End If

        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With
        SheetsCount = 12     '= xlWorkbook.Sheets.Count
        If Startup Then
            For intI = 1 To SheetsCount
                xlWorksheet = xlWorkbook.Sheets.Item(intI)
                Sheetnames(intI) = (xlWorksheet.Name)
                SheetIndices(intI) = intI
                getCategoryShortname(xlWorksheet)
                RecordsInserted(intI) = xlWorksheet.Range(Category_Shortname & "RecordsTotal").Value
            Next
            Startup = False
            Me.btnClose.Enabled = False
        End If
        If Me.cmbExtEquip_Category.Text = "External Others" Then
            currentsheetname2 = "External Others"
        Else
            currentsheetname2 = "external Hire"
        End If

        For intI = 0 To SheetsCount - 1
            currentsheetname1 = Sheetnames(intI + 1)
            If UCase(Trim(currentsheetname1)) = UCase(Trim(currentsheetname2)) Then
                xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
                SheetNo = intI + 1
                Exit For
            End If
        Next

        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)    '= "Ext_"

        xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
        xlRange.Value = mMainTitle1
        xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
        xlRange.Value = mMainTitle2
        'xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
        'xlRange.Value = mMainTitle3
        xlRange = xlWorksheet.Range(Category_Shortname & "Client")
        xlRange.Value = mClient
        xlRange = xlWorksheet.Range(Category_Shortname & "Location")
        xlRange.Value = mLocation
        xlRange = xlWorksheet.Range(Category_Shortname & "StartDate")
        xlRange.Value = mStartDate.Date.ToString()
        xlRange = xlWorksheet.Range(Category_Shortname & "EndDate")
        xlRange.Value = mEndDate.Date.ToString()
        xlRange = xlWorksheet.Range(Category_Shortname & "ProjectValue")
        xlRange.Value = mProjectvalue
        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
        xlRange.Value = "Fuel Cost for all mc/month @Rs. " & Me.txtExtEquip_FuelCostperLtr.Text & " per Lt"
        xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
        xlRange.Value = "Opr Cost for all m/c with " & Me.cmbExtEquip_Shifts.Text & " per day"
        'End If
        RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1

        xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo), 0)
        xlRange.Value = RecordsInserted(SheetNo)
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.cmbExtEquip_Name.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.cmbExtEquip_Drive.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.cmbExtEquip_Capacity.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.cmbExtEquip_Make.Text & vbNewLine & "/" & Me.cmbExtEquip_Model.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = IIf(Len(Trim(Me.txtExtEquip_Qty.Text)) = 0, 0, Me.txtExtEquip_Qty.Text)
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.dpExtEquip_MobDate.Text
        xlRange.NumberFormat = "dd-mmm-yyyy"
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.dpExtEquip_DemobDate.Text
        xlRange.NumberFormat = "dd-mmm-yyyy"
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = IIf(Len(Trim(Me.txtExtEquip_RepValue.Text)) = 0, 0, Me.txtExtEquip_RepValue.Text)
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = IIf((Me.cmbExtEquip_DepPerc.Text = "Nil" Or Len(Trim(Me.cmbExtEquip_DepPerc.Text)) = 0), 0, Me.cmbExtEquip_DepPerc.Text)
        xlRange = xlRange.Offset(0, 1)
        If IsNumeric(Me.cmbExtEquip_DepPerc.Text) Then
            xlRange.Value = Val(Me.txtExtEquip_RepValue) * Val(Me.cmbExtEquip_DepPerc.Text)
        Else
            xlRange.Value = Me.txtExtEquip_Depreciation.Text
        End If

        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = IIf(Len(Trim(Me.txtExtEquip_PropHours.Text)) = 0, 0, Me.txtExtEquip_PropHours.Text)
        xlRange = xlRange.Offset(0, 1)

        xlRange.Value = IIf(Len(Trim(Me.txtExtEquip_FuelPerHr.Text)) = 0, 0, Me.txtExtEquip_FuelPerHr.Text)
        xlRange = xlRange.Offset(0, 4)

        xlRange.Value = IIf(Len(Trim(Me.txtExtEquip_FuelCostperLtr.Text)) = 0, 0, Me.txtExtEquip_FuelCostperLtr.Text)

        xlRange = xlRange.Offset(0, 3)

        xlRange.Value = IIf(Len(Trim(Me.txtExtEquip_PowerPerhr.Text)) = 0, 0, Me.txtExtEquip_PowerPerhr.Text)
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = IIf(Len(Trim(Me.txtExtEquip_PowerCostPerUnit.Text)) = 0, 0, Me.txtExtEquip_PowerCostPerUnit.Text)
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = IIf(Len(Trim(Me.txtExtEquip_OprCostPerMCPerMonth.Text)) = 0, 0, Me.txtExtEquip_OprCostPerMCPerMonth.Text)
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = IIf(Len(Trim(Me.cmbExtEquip_Shifts.Text)) = 0, 1, Val(Me.cmbExtEquip_Shifts.Text))
        xlRange.NumberFormat = "0.0#"
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = IIf(Len(Trim(Me.txtExtEquip_ConxumablesPerMonth.Text)) = 0, 0, Val(Me.txtExtEquip_ConxumablesPerMonth.Text))  ' * Val(Me.txtExtEquip_Qty.Text))

        If (UCase(Me.cmbExtEquip_Category.Text) = UCase("Conveyance_Hired") And UCase(Me.cmbExtEquip_Name.Text) = UCase("Bus")) Then
            xlRange = xlWorksheet.Range("Ext_Bus_HirechargesTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_Hire_Charges.Text) * Val(Me.txtExtEquip_Qty.Text) * Val(Me.txtExtEquip_Months.Text)
            xlRange = xlWorksheet.Range("Ext_Bus_FuelPerMonthTotal")
            xlRange.Value = xlRange.Value + Val(txtExtEquip_FuelperMCPerMonth.Text) * Val(Me.txtExtEquip_Qty.Text)
            xlRange = xlWorksheet.Range("Ext_Bus_FuelProjectTotal")
            xlRange.Value = xlRange.Value + Val(txtExtEquip_FuelperMCPerMonth.Text) * Val(Me.txtExtEquip_Qty.Text) * Val(Me.txtExtEquip_Months.Text)

            xlRange = xlWorksheet.Range("Ext_Bus_FuelCostperMonthTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_FuelCostPerMonth.Text)
            xlRange = xlWorksheet.Range("Ext_Bus_FuelCostProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_FuelCostProject.Text)
            xlRange = xlWorksheet.Range("Ext_Bus_PowerCostperMonthTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_PowerCostPerMonth.Text)
            xlRange = xlWorksheet.Range("Ext_Bus_PowerCostProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_PowerCostProject.Text)
            xlRange = xlWorksheet.Range("Ext_Bus_OperatorCostProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_OprCostProject.Text)
            xlRange = xlWorksheet.Range("Ext_Bus_consumablesProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_ConxumablesProject.Text)
        ElseIf (UCase(Me.cmbExtEquip_Category.Text) = UCase("Conveyance_Hired") And UCase(Me.cmbExtEquip_Name.Text) = UCase("Car")) Then
            xlRange = xlWorksheet.Range("Ext_Car_HirechargesTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_Hire_Charges.Text) * Val(Me.txtExtEquip_Qty.Text) * Val(Me.txtExtEquip_Months.Text)
            xlRange = xlWorksheet.Range("Ext_Car_FuelPerMonthTotal")
            xlRange.Value = xlRange.Value + Val(txtExtEquip_FuelperMCPerMonth.Text) * Val(Me.txtExtEquip_Qty.Text)
            xlRange = xlWorksheet.Range("Ext_Car_FuelProjectTotal")
            xlRange.Value = xlRange.Value + Val(txtExtEquip_FuelperMCPerMonth.Text) * Val(Me.txtExtEquip_Qty.Text) * Val(Me.txtExtEquip_Months.Text)

            xlRange = xlWorksheet.Range("Ext_Car_FuelCostperMonthTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_FuelCostPerMonth.Text)
            xlRange = xlWorksheet.Range("Ext_Car_FuelCostProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_FuelCostProject.Text)
            xlRange = xlWorksheet.Range("Ext_Car_PowerCostperMonthTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_PowerCostPerMonth.Text)
            xlRange = xlWorksheet.Range("Ext_Car_PowerCostProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_PowerCostProject.Text)
            xlRange = xlWorksheet.Range("Ext_Car_OperatoCostProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_OprCostProject.Text)
            xlRange = xlWorksheet.Range("Ext_Car_consumablesProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_ConxumablesProject.Text)
        ElseIf (UCase(Me.cmbExtEquip_Category.Text) = UCase("Conveyance_Hired") And UCase(Me.cmbExtEquip_Name.Text) = UCase("Sumo")) Then
            xlRange = xlWorksheet.Range("Ext_Sumo_HirechargesTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_Hire_Charges.Text) * Val(Me.txtExtEquip_Qty.Text) * Val(Me.txtExtEquip_Months.Text)

            xlRange = xlWorksheet.Range("Ext_Sumo_FuelPerMonthTotal")
            xlRange.Value = xlRange.Value + Val(txtExtEquip_FuelperMCPerMonth.Text) * Val(Me.txtExtEquip_Qty.Text)
            xlRange = xlWorksheet.Range("Ext_Sumo_FuelProjectTotal")
            xlRange.Value = xlRange.Value + Val(txtExtEquip_FuelperMCPerMonth.Text) * Val(Me.txtExtEquip_Qty.Text) * Val(Me.txtExtEquip_Months.Text)


            xlRange = xlWorksheet.Range("Ext_Sumo_FuelCostperMonthTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_FuelCostPerMonth.Text)
            xlRange = xlWorksheet.Range("Ext_Sumo_FuelCostProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_FuelCostProject.Text)
            xlRange = xlWorksheet.Range("Ext_Sumo_PowerCostperMonthTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_PowerCostPerMonth.Text)
            xlRange = xlWorksheet.Range("Ext_Sumo_PowerCostProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_PowerCostProject.Text)
            xlRange = xlWorksheet.Range("Ext_Sumo_OperatoCostProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_OprCostProject.Text)
            xlRange = xlWorksheet.Range("Ext_Sumo_consumablesProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_ConxumablesProject.Text)
        ElseIf (UCase(Me.cmbExtEquip_Category.Text) = UCase("Hired Equipments")) Then
            xlRange = xlWorksheet.Range("Ext_Equpments_HirechargesTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_Hire_Charges.Text) * Val(Me.txtExtEquip_Qty.Text) * Val(Me.txtExtEquip_Months.Text)

            xlRange = xlWorksheet.Range("Ext_Equpments_FuelPerMonthTotal")
            xlRange.Value = xlRange.Value + Val(txtExtEquip_FuelperMCPerMonth.Text) * Val(Me.txtExtEquip_Qty.Text)
            xlRange = xlWorksheet.Range("Ext_Equpments_FuelProjectTotal")
            xlRange.Value = xlRange.Value + Val(txtExtEquip_FuelperMCPerMonth.Text) * Val(Me.txtExtEquip_Qty.Text) * Val(Me.txtExtEquip_Months.Text)

            xlRange = xlWorksheet.Range("Ext_Equpments_FuelCostperMonthTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_FuelCostPerMonth.Text)
            xlRange = xlWorksheet.Range("Ext_Equpments_FuelCostProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_FuelCostProject.Text)
            xlRange = xlWorksheet.Range("Ext_Equpments_PowerCostperMonthTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_PowerCostPerMonth.Text)
            xlRange = xlWorksheet.Range("Ext_Equpments_PowerCostProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_PowerCostProject.Text)
            xlRange = xlWorksheet.Range("Ext_Equpments_OperatoCostProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_OprCostProject.Text)
            xlRange = xlWorksheet.Range("Ext_Equpments_consumablesProjectTotal")
            xlRange.Value = xlRange.Value + Val(Me.txtExtEquip_ConxumablesProject.Text)
        End If
        If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
        MsgBox("Record No. " & RecordsInserted(SheetNo) & " added in " & xlWorksheet.Name)
        Dim InsertCommand As String = ""
        mcategory = cmbExtEquip_Category.Text
        InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
        InsertCommand = InsertCommand & "'" & Me.cmbExtEquip_Name.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.cmbExtEquip_Capacity.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.cmbExtEquip_Make.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.cmbExtEquip_Model.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.dpExtEquip_MobDate.Text & "',"
        InsertCommand = InsertCommand & "'" & Me.dpExtEquip_DemobDate.Text & "',"
        InsertCommand = InsertCommand & Me.txtExtEquip_Qty.Text & ")"
        'InsertCommand = InsertCommand & "'" & Me.dpEndDDeate.Text & "',"

        Try
            moleDBInsertCommand = New OleDbCommand
            moleDBInsertCommand.CommandType = CommandType.Text
            moleDBInsertCommand.CommandText = InsertCommand
            moleDBInsertCommand.Connection = moledbConnection1
            moleDBInsertCommand.ExecuteNonQuery()
            'MsgBox(Me.txtProjectdescription.Text & "Record inserted in list of projects.")
        Catch ex As Exception
            MsgBox(e.ToString())
        Finally
            'moledbConnection.Close()
            moleDBInsertCommand = Nothing
        End Try
        ClearFields_ExtEquipments()
        Me.btnExtEquip_SaveEntry.Enabled = False
        Me.txtExtEquip_FuelperMCPerMonth.Text = ""
        Me.txtExtEquip_FuelCostPerMonth.Text = ""
        Me.txtExtEquip_FuelCostProject.Text = ""
        Me.txtExtEquip_PowerCostPerMonth.Text = ""
        Me.txtMinorequip_PowerCostProject.Text = ""
        Me.txtExtEquip_OprCostPerMonth.Text = ""
        Me.txtExtEquip_OprCostProject.Text = ""
        Me.txtExtEquip_ConxumablesPerMonth.Text = ""
        Me.txtExtEquip_ConxumablesProject.Text = ""
        Me.txtExtEquip_OperatingCost.Text = ""
        EditOrDelete = ""
    End Sub
    Private Sub ClearFields_ExtEquipments()
        Me.cmbExtEquip_Name.Text = ""
        Me.cmbExtEquip_Capacity.Text = ""
        Me.cmbExtEquip_Model.Text = ""
        Me.cmbExtEquip_Make.Text = ""
        Me.cmbExtEquip_Drive.SelectedIndex = 0
        Me.txtExtEquip_Qty.Text = 0
        Me.dpExtEquip_MobDate.Text = mStartDate.Date
        Me.dpExtEquip_DemobDate.Text = mEndDate.Date
        Me.txtExtEquip_Months.Text = 0
        Me.txtExtEquip_RepValue.Text = 0
        Me.cmbExtEquip_DepPerc.SelectedIndex = 0
        Me.txtExtEquip_Depreciation.Text = 0
        Me.txtExtEquip_Hire_Charges.Text = 0
        Me.txtExtEquip_PropHours.Text = 0
        Me.txtExtEquip_FuelPerHr.Text = 0
        Me.txtExtEquip_PowerPerhr.Text = 0
        Me.txtExtEquip_OprCostPerMCPerMonth.Text = 0
        Me.txtExtEquip_ConxumablesPerMonth.Text = 0
        Me.txtExtEquip_OperatingCost.Text = 0
        Me.cmbExtEquip_Shifts.SelectedIndex = 0
        Me.txtMaintPercPerMC_PerMonth.Text = 0
    End Sub

    Private Sub btnMinorequip_SaveAndNextTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMinorequip_SaveAndNextTab.Click
        Me.tbclEqipsEntry.SelectTab(2)
        Currenttab = 2
    End Sub

    Private Sub btnExtEquip_SaveAndNextTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExtEquip_SaveAndNextTab.Click
        Me.tbclEqipsEntry.SelectTab(3)
        Currenttab = 3
    End Sub

    Private Sub txtMonths_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMonths.Validated
        If Val(txtMonths.Text) <= 0 Then
            txtMonths.Text = 1
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        If Not Startup = False Then
            If Not xlWorkbook Is Nothing Then
                xlWorkbook.Save()
                xlWorkbook.Close()
            End If
            xlWorkbook = Nothing
            xlApp.Quit()
            xlApp = Nothing
            System.GC.Collect()
            Me.Close()
            frmProjectDetails.Show()
        End If
    End Sub

    Private Sub btnMinorequip_ClearFields_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMinorequip_ClearFields.Click
        Me.ClearFields_MinorEquipments()
    End Sub

    Private Sub btnExtEquip_ClearFields_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExtEquip_ClearFields.Click
        Me.ClearFields_ExtEquipments()
    End Sub

    Private Sub txtDepreciation_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDepreciation.Validated
        Me.txtHireCharges.Text = Val(Me.txtDepreciation.Text) * Val(Me.txtMajorEquipQty.Text) * Val(Me.txtMonths.Text)
    End Sub

    Private Sub txtDepreciation_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDepreciation.Validating
        If Not IsNumeric(Me.txtDepreciation.Text) Then
            MsgBox("Deprecision should be a numeric value")
            Me.txtDepreciation.Focus()
            e.Cancel = True
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim mTotalExp As Double, mCablesExp As Double, mPanelExp As Double, mLightsExp As Double, mEBDeposit As Double
        Dim mMiscExp As Double, msgstring As String = ""
        Dim currentsheetname2 As String = "Electrical"
        Dim currentsheetname1 As String, intI As Integer
        Dim Allvalid As Boolean = True
        Me.btnQuit.Enabled = True
        Me.btnClose.Enabled = False
        Me.lblError.Text = ""
        If Val(Me.txtTotalElecExpensesPerc.Text) >= 100 Then
            msgstring = msgstring & "Total Electrical expenses percentage should be less than 100" & vbNewLine
            Me.txtTotalElecExpensesPerc.Focus()
            Allvalid = False
        ElseIf Val(Me.txtCablesPerc.Text) >= 100 Then
            msgstring = msgstring & "Cables expenses percentage should be less than 100" & vbNewLine
            Me.txtCablesPerc.Focus()
            Allvalid = False
        ElseIf Val(Me.txtPanelsperc.Text) >= 100 Then
            msgstring = msgstring & "Total Electrical expenses percentage should be less than 100" & vbNewLine
            Me.txtPanelsperc.Focus()
            Allvalid = False
        ElseIf Val(Me.txtEbDepositPerc.Text) >= 100 Then
            msgstring = msgstring & "EB Deposit and Other expenses percentage should be less than 100" & vbNewLine
            Me.txtEbDepositPerc.Focus()
            Allvalid = False
        ElseIf Val(Me.txtMiscPerc.Text) >= 100 Then
            msgstring = msgstring & "Miscellaneous expenses percentage should be less than 100" & vbNewLine
            Me.txtMiscPerc.Focus()
            Allvalid = False
        ElseIf Val(Me.txtLightsPerc.Text) >= 100 Then
            msgstring = msgstring & "Lights and Accessories expenses percentage should be less than 100" & vbNewLine
            Me.txtLightsPerc.Focus()
            Allvalid = False
        End If
        If Not Allvalid Then
            Me.lblError.Text = msgstring
            Exit Sub
        End If
        mTotalExp = Val(Me.txtTotalElecExpensesPerc.Text)
        mCablesExp = Val(Me.txtCablesPerc.Text)
        mPanelExp = Val(Me.txtPanelsperc.Text)
        mLightsExp = Val(Me.txtLightsPerc.Text)
        mEBDeposit = Val(Me.txtEbDepositPerc.Text)
        mMiscExp = Val(Me.txtMiscPerc.Text)

        If System.Math.Round((mCablesExp + mPanelExp + mLightsExp + mEBDeposit + mMiscExp), 2) <> System.Math.Round(mTotalExp, 2) Then
            Me.lblError.Text = "Percentage Break-up is not equal to the Total"
            Exit Sub
        End If

        If Len(Trim(xlFilename)) = 0 Then
            MsgBox("Filename not specified to save. cannot Save data. Start the application again.")
            Application.Exit()
        End If

        If xlWorkbook Is Nothing Then
            xlWorkbook = xlApp.Workbooks.Open(xlFilename)
        End If
        SheetsCount = 12    '= xlWorkbook.Sheets.Count

        If Startup Then
            For intI = 1 To SheetsCount
                xlWorksheet = xlWorkbook.Sheets.Item(intI)
                Sheetnames(intI) = (xlWorksheet.Name)
                SheetIndices(intI) = intI
                getCategoryShortname(xlWorksheet)
                RecordsInserted(intI) = xlWorksheet.Range(Category_Shortname & "RecordsTotal").Value
            Next
            Startup = False
        End If
        For intI = 0 To SheetsCount - 1
            currentsheetname1 = Sheetnames(intI + 1)
            If UCase(Trim(currentsheetname1)) = UCase(Trim(currentsheetname2)) Then
                xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
                SheetNo = intI + 1
                Exit For
            End If
        Next
        xlWorksheet = xlWorkbook.Sheets.Item("Electrical")
        xlWorksheet.Activate()
        Category_Shortname = "Elec_"
        xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
        xlRange.Value = mMainTitle2
        xlRange = xlWorksheet.Range(Category_Shortname & "Client")
        xlRange.Value = mClient
        xlRange = xlWorksheet.Range(Category_Shortname & "Location")
        xlRange.Value = mLocation
        xlRange = xlWorksheet.Range(Category_Shortname & "StartDate")
        xlRange.Value = mStartDate.Date
        xlRange = xlWorksheet.Range(Category_Shortname & "EndDate")
        xlRange.Value = mEndDate.Date
        xlRange = xlWorksheet.Range(Category_Shortname & "ProjectValue")
        xlRange.Value = mProjectvalue
        xlRange = xlWorksheet.Range("Elec_CablesExp")
        xlRange.Value = mProjectvalue * (Val(Me.txtCablesPerc.Text) / 100)
        xlRange = xlWorksheet.Range("Elec_PanelsExp")
        xlRange.Value = mProjectvalue * (Val(Me.txtPanelsperc.Text) / 100)
        xlRange = xlWorksheet.Range("Elec_LightsExp")
        xlRange.Value = mProjectvalue * (Val(Me.txtLightsPerc.Text) / 100)
        xlRange = xlWorksheet.Range("Elec_EBDeposit")
        xlRange.Value = mProjectvalue * (Val(Me.txtEbDepositPerc.Text) / 100)
        xlRange = xlWorksheet.Range("Elec_MiscExp")
        xlRange.Value = mProjectvalue * (Val(Me.txtMiscPerc.Text) / 100)
        xlRange = xlWorksheet.Range("Elec_reusablePanels")
        xlRange.Value = Me.txtReusablePanels.Text
        xlRange = xlWorksheet.Range("Elec_ReusableLights")
        xlRange.Value = Me.txtReusablesLights.Text
        MsgBox("Electrical Expense Budget details added in " & xlWorksheet.Name)
        Me.btnSave.Enabled = False
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Me.txtTotalElecExpensesPerc.Text = ""
        Me.txtCablesPerc.Text = ""
        Me.txtPanelsperc.Text = ""
        Me.txtLightsPerc.Text = ""
        Me.txtEbDepositPerc.Text = ""
        Me.txtMiscPerc.Text = ""
        Me.lblError.Text = ""
        Me.btnSave.Enabled = True
    End Sub

    Private Sub txtMaintCostPerMC_PerMonth_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMaintCostPerMC_PerMonth.Validated
        Me.txtMaintCostPerMC_PerMonth.Text = Val(Me.txtRepvalue.Text) * Val(Me.txtMaintPercPerMC_PerMonth.Text) / 100
    End Sub

    Private Sub txtMaintCostPerMC_PerMonth_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMaintCostPerMC_PerMonth.Validating
        If Not IsNumeric(Me.txtMaintCostPerMC_PerMonth.Text) Then
            MsgBox("Only numeric  values accepted")
            Me.txtMaintCostPerMC_PerMonth.Text = 0
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub txtMinorequip_ConsumablesPercPerMCPerMonth_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMinorequip_ConsumablesPercPerMCPerMonth.Validated
        Me.txtMinorequip_ConsumablesPerMCPerMonth.Text = Val(Me.txtMinorequip_NewEquipCost.Text) * Val(Me.txtMinorequip_ConsumablesPercPerMCPerMonth.Text) / 100
    End Sub

    Private Sub txtMinorequip_ConsumablesPercPerMCPerMonth_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMinorequip_ConsumablesPercPerMCPerMonth.Validating
        Dim Response As Integer
        If Not IsNumeric(txtMinorequip_ConsumablesPercPerMCPerMonth.Text) Then
            MsgBox("Type only numeric value for maintenance percentage ")
            Me.txtMinorequip_ConsumablesPercPerMCPerMonth.Text = 0
            e.Cancel = True
            Exit Sub
        End If
        If Val(Me.txtMinorequip_ConsumablesPercPerMCPerMonth.Text) < mMinMaintperc Then
            Response = MsgBox("The Maint cost perecentage for the equipment entered is less than the minimum prescribed. Do you want me to accept it?", MsgBoxStyle.Information + MsgBoxStyle.YesNo)
            If Response = vbNo Then
                e.Cancel = True
                Exit Sub
            End If
        ElseIf Val(Me.txtMinorequip_ConsumablesPercPerMCPerMonth.Text) > mMaxMaintPerc Then
            Response = MsgBox("The Maint cost perecentage for the equipment entered is more than the maximum prescribed. Do you want me to accept it?", MsgBoxStyle.Information + MsgBoxStyle.YesNo)
            If Response = vbNo Then
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtExtEquip_ConsumablesPercPerMCPerMonth_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExtEquip_ConsumablesPercPerMCPerMonth.Validated
        Me.txtExtEquip_ConsumablesPerMCPerMonth.Text = Val(Me.txtExtEquip_RepValue.Text) * Val(Me.txtExtEquip_ConsumablesPercPerMCPerMonth.Text) / 100
    End Sub

    Private Sub txtExtEquip_ConsumablesPercPerMCPerMonth_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtExtEquip_ConsumablesPercPerMCPerMonth.Validating
        Dim Response As Integer
        If Not IsNumeric(Me.txtExtEquip_ConsumablesPercPerMCPerMonth.Text) Then
            MsgBox("Type only numeric value for maintenance percentage ")
            Me.txtExtEquip_ConsumablesPercPerMCPerMonth.Text = 0
            e.Cancel = True
            Exit Sub
        End If
        If Val(Me.txtExtEquip_ConsumablesPercPerMCPerMonth.Text) < mMinMaintperc Then
            Response = MsgBox("The Maint cost perecentage for the equipment entered is less than the minumum prescribed. Do you want me to accept it?", MsgBoxStyle.Information + MsgBoxStyle.YesNo)
            If Response = vbNo Then
                e.Cancel = True
                Exit Sub
            End If
        ElseIf Val(Me.txtExtEquip_ConsumablesPercPerMCPerMonth.Text) > mMaxMaintPerc Then
            Response = MsgBox("The Maint cost perecentage for the equipment entered is more than the maximum prescribed. Do you want me to accept it?", MsgBoxStyle.Information + MsgBoxStyle.YesNo)
            If Response = vbNo Then
                e.Cancel = True
                Exit Sub
            End If
        End If

    End Sub

    Private Sub txtMaintPercPerMC_PerMonth_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMaintPercPerMC_PerMonth.Validated
        Me.txtMaintCostPerMC_PerMonth.Text = Val(Me.txtRepvalue.Text) * Val(Me.txtMaintPercPerMC_PerMonth.Text) / 100
    End Sub

    Private Sub txtMaintPercPerMC_PerMonth_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMaintPercPerMC_PerMonth.Validating
        Dim Response As Integer
        If Not IsNumeric(Me.txtMaintPercPerMC_PerMonth.Text) Then
            MsgBox("Type only numeric value for maintenance percentage ")
            Me.txtMaintPercPerMC_PerMonth.Text = 0
            e.Cancel = True
            Exit Sub
        End If
        If System.Math.Round(Val(Me.txtMaintPercPerMC_PerMonth.Text), 2) < System.Math.Round(mMinMaintperc, 2) Then
            Response = MsgBox("The Maint cost perecentage for the equipment entered is less than the minimum prescribed. Do you want me to accept it?", MsgBoxStyle.Information + MsgBoxStyle.YesNo)
            If Response = vbNo Then
                e.Cancel = True
                Exit Sub
            End If
        ElseIf System.Math.Round(Val(Me.txtMaintPercPerMC_PerMonth.Text), 2) > System.Math.Round(mMaxMaintPerc, 2) Then
            Response = MsgBox("The Maint cost perecentage for the equipment entered is more than the mximum prescribed. Do you want me to accept it?", MsgBoxStyle.Information + MsgBoxStyle.YesNo)
            If Response = vbNo Then
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub

    Private Sub cmbMinorequip_DepPerc_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbMinorequip_DepPerc.Validated
        Me.txtMinorequip_Depreciation.Text = Val(Me.txtMinorequip_NewEquipCost.Text) * Val(Me.cmbMinorequip_DepPerc.Text) / 100
        Me.txtMinorEquip_HireCharges.Text = Val(Me.txtMinorequip_Depreciation.Text) * Val(Me.txtMinorEquip_Qty.Text) * Val(Me.txtMinorequip_Months.Text)
    End Sub

    Private Sub cmbExtEquip_DepPerc_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbExtEquip_DepPerc.Validated
        Dim mDepPerc As Single = 0
        Me.txtExtEquip_Depreciation.Text = Val(Me.txtExtEquip_RepValue.Text) * mDepPerc / 100
    End Sub

    Private Sub cmbDepPercentage_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbDepPercentage.Validated
        If Me.cmbDepPercentage.Text <> "0" Then
            Me.txtDepreciation.Text = Val(Me.txtRepvalue.Text) * Val(Me.cmbDepPercentage.Text) / 100
            Me.txtHireCharges.Text = Val(Me.txtDepreciation.Text) * Val(Me.txtMajorEquipQty.Text) * Val(Me.txtMonths.Text)
            Me.txtDepreciation.Enabled = False
        Else
            Me.txtDepreciation.Enabled = True
        End If
    End Sub

    Private Sub txtExtEquip_Hire_Charges_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExtEquip_Hire_Charges.Validated
        Me.txtExtEquip_RepValue.Text = Me.txtExtEquip_Hire_Charges.Text
    End Sub
    Private Sub pgMajorEquips_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles pgMajorEquips.Enter
        Me.Button1.Enabled = False
    End Sub

    Private Sub pgMinorEquips_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles pgMinorEquips.Enter
        Me.bbtnMinorequip_SaveEntry.Enabled = False
    End Sub

    Private Sub pgHireEquips_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles pgHireEquips.Enter
        Me.btnExtEquip_SaveEntry.Enabled = False
    End Sub
    Private Sub txtTotalElecExpensesPerc_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTotalElecExpensesPerc.Validating
        If Val(Me.txtTotalElecExpensesPerc.Text) < 0 Or Val(Me.txtTotalElecExpensesPerc.Text) > 1.5 Then
            MsgBox("total Electrical Expenses Percentage should be within the range of 0 and 1.5")
            e.Cancel = True
        End If
    End Sub

    Private Sub chkNewOrOldEquip_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNewOrOldEquip.CheckedChanged
        If Me.chkNewOrOldEquip.Checked = False Then
            txtMinorequip_NewEquipCost.Text = 0
            txtMinorequip_NewEquipCost.Enabled = False
        Else
            txtMinorequip_NewEquipCost.Enabled = True
        End If
    End Sub

    Private Sub btnViewItemsMajor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewItemsMajor.Click
        Currenttab = 0
        If Not Len(Trim(Me.cmbMajorEqipCategory.Text)) = 0 Then
            SelectedCategory = Me.cmbMajorEqipCategory.Text
            Me.Hide()
            oForm = New AddedItems()
            oForm.ShowDialog()
            oForm = Nothing
        Else
            MsgBox("First Select the Category in Category combobox")
        End If
    End Sub

    Private Sub btnViewItemsMinor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewItemsMinor.Click
        Currenttab = 1
        If Not Len(Trim(Me.cmbMinoEquip_Category.Text)) = 0 Then
            SelectedCategory = Me.cmbMinoEquip_Category.Text
            Me.Hide()
            oForm = New AddedItems()
            oForm.ShowDialog()
            oForm = Nothing
        Else
            MsgBox("First Select the Category in Category combobox")
        End If
    End Sub

    Private Sub btnHiredItemsView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHiredItemsView.Click
        Currenttab = 2
        If Not Len(Trim(Me.cmbExtEquip_Category.Text)) = 0 Then
            SelectedCategory = Me.cmbExtEquip_Category.Text
            Me.Hide()
            oForm = New AddedItems()
            oForm.ShowDialog()
            oForm = Nothing
        Else
            MsgBox("First Select the Category in Category combobox")
        End If
    End Sub

    Private Sub TextBox2_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles qty2.TextChanged
        Me.Amount2.Text = Val(Me.qty2.Text) * Val(Me.Cost2.Text)
    End Sub

    Private Sub txtCost1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Amount1.Text = Val(Me.Qty1.Text) * Val(Me.Cost1.Text)
    End Sub

    Private Sub btnTCRExpSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTCRExpSave.Click
        Dim currentsheetname2 As String, currentsheetname1 As String, intI As Integer
        currentsheetname2 = "Tr crane related exp"
        Me.btnQuit.Enabled = True
        Me.btnClose.Enabled = False
        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With

        For intI = 0 To xlWorkbook.Sheets.Count - 1
            xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
            currentsheetname1 = xlWorksheet.Name
            If UCase(Trim(currentsheetname1)) = UCase(Trim(currentsheetname2)) Then
                SheetNo = intI + 1
                Exit For
            End If
        Next
        intI = 1

        xlRange = xlWorksheet.Range("TCRExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.Qty1.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.Cost1.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.Remarks1.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("TCRExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.qty2.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.Cost2.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.Remarks2.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("TCRExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.Qty3.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.Cost3.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.Remarks3.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("TCRExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.Qty4.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.Cost4.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.Remarks4.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("TCRExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.Qty5.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.Cost5.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.Remarks5.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("TCRExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.Qty6.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.Cost6.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.Remarks6.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("TCRExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.Qty7.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.Cost7.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.Remarks7.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("TCRExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.Qty8.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.Cost8.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.Remarks8.Text
        Label180.Text = "Tower Crane related Expenses  Saved. "
        Me.btnTCRExpSave.Enabled = False

    End Sub
    Private Sub txtQty1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Amount1.Text = Val(Me.Qty1.Text) * Val(Me.Cost1.Text)
    End Sub

    Private Sub Cost2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cost2.TextChanged
        Me.Amount2.Text = Val(Me.qty2.Text) * Val(Me.Cost2.Text)
    End Sub

    Private Sub Qty3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Qty3.TextChanged
        Me.Amount3.Text = Val(Me.Qty3.Text) * Val(Me.Cost3.Text)
    End Sub

    Private Sub Cost3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cost3.TextChanged
        Me.Amount3.Text = Val(Me.Qty3.Text) * Val(Me.Cost3.Text)
    End Sub

    Private Sub Qty4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Qty4.TextChanged
        Me.Amount4.Text = Val(Me.Qty4.Text) * Val(Me.Cost4.Text)
    End Sub

    Private Sub Cost4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cost4.TextChanged
        Me.Amount4.Text = Val(Me.Qty4.Text) * Val(Me.Cost4.Text)
    End Sub

    Private Sub Qty5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Qty5.TextChanged
        Me.Amount5.Text = Val(Me.Qty5.Text) * Val(Me.Cost5.Text)
    End Sub

    Private Sub Cost5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cost5.TextChanged
        Me.Amount5.Text = Val(Me.Qty5.Text) * Val(Me.Cost5.Text)
    End Sub

    Private Sub btnTCRExpCleaar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTCRExpCleaar.Click
        Me.Qty1.Text = ""
        Me.Cost1.Text = ""
        Me.Amount1.Text = ""
        Me.Remarks1.Text = ""

        Me.qty2.Text = ""
        Me.Cost2.Text = ""
        Me.Amount2.Text = ""
        Me.Remarks2.Text = ""

        Me.Qty3.Text = ""
        Me.Cost3.Text = ""
        Me.Amount3.Text = ""
        Me.Remarks3.Text = ""

        Me.Qty4.Text = ""
        Me.Cost4.Text = ""
        Me.Amount4.Text = ""
        Me.Remarks1.Text = ""

        Me.Qty5.Text = ""
        Me.Cost5.Text = ""
        Me.Amount5.Text = ""
        Me.Remarks5.Text = ""

    End Sub

    Private Sub btnBPlantExpSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBPlantExpSave.Click
        Dim currentsheetname2 As String, currentsheetname1 As String, intI As Integer
        currentsheetname2 = "Bplant related exp"
        Me.btnQuit.Enabled = True
        Me.btnClose.Enabled = False
        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With

        For intI = 0 To xlWorkbook.Sheets.Count - 1
            xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
            currentsheetname1 = xlWorksheet.Name
            If UCase(Trim(currentsheetname1)) = UCase(Trim(currentsheetname2)) Then
                SheetNo = intI + 1
                Exit For
            End If
        Next
        intI = 1

        xlRange = xlWorksheet.Range("BPlantExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.BPlantQty1.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.BPlantCost1.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.BPlantRemarks1.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("BPlantExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.BPlantQty2.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.BPlantCost2.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.BPlantRemarks2.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("BPlantExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.BPlantQty3.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.BPlantCost3.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.BPlantRemarks3.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("BPlantExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.BPlantQty4.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.BPlantCost4.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.BPlantRemarks4.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("BPlantExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.BPlantQty5.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.BPlantCost5.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.BPlantRemarks4.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("BPlantExp_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.BPlantQty6.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.BPlantCost6.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.BPlantRemarks4.Text

        Label179.Text = "BPlant expenses saved. "
        Me.btnBPlantExpSave.Enabled = False

    End Sub

    Private Sub BPlantQty1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BPlantQty1.TextChanged
        Me.BPlantAmount1.Text = Val(Me.BPlantQty1.Text) * Val(Me.BPlantCost1.Text)
    End Sub

    Private Sub BPlantCost1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BPlantCost1.TextChanged
        Me.BPlantAmount1.Text = Val(Me.BPlantQty1.Text) * Val(Me.BPlantCost1.Text)
    End Sub

    Private Sub BPlantExpHead2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BPlantQty2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BPlantQty2.TextChanged
        Me.BPlantAmount2.Text = Val(Me.BPlantQty2.Text) * Val(Me.BPlantCost2.Text)
    End Sub

    Private Sub BPlantCost2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BPlantCost2.TextChanged
        Me.BPlantAmount2.Text = Val(Me.BPlantQty2.Text) * Val(Me.BPlantCost2.Text)
    End Sub

    Private Sub BPlantQty3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BPlantQty3.TextChanged
        Me.BPlantAmount3.Text = Val(Me.BPlantQty3.Text) * Val(Me.BPlantCost3.Text)
    End Sub

    Private Sub BPlantCost3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BPlantCost3.TextChanged
        Me.BPlantAmount3.Text = Val(Me.BPlantQty3.Text) * Val(Me.BPlantCost3.Text)
    End Sub

    Private Sub BPlantQty4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BPlantQty4.TextChanged
        Me.BPlantAmount4.Text = Val(Me.BPlantQty4.Text) * Val(Me.BPlantCost4.Text)
    End Sub

    Private Sub BPlantCost4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BPlantCost4.TextChanged
        Me.BPlantAmount4.Text = Val(Me.BPlantQty4.Text) * Val(Me.BPlantCost4.Text)
    End Sub

    Private Sub btnBPlantExpClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBPlantExpClear.Click
        Me.BPlantQty1.Text = ""
        Me.BPlantCost1.Text = ""
        Me.BPlantAmount1.Text = ""
        Me.BPlantRemarks1.Text = ""

        Me.BPlantQty2.Text = ""
        Me.BPlantCost2.Text = ""
        Me.BPlantAmount2.Text = ""
        Me.BPlantRemarks2.Text = ""

        Me.BPlantQty3.Text = ""
        Me.BPlantCost3.Text = ""
        Me.BPlantAmount3.Text = ""
        Me.BPlantRemarks3.Text = ""

        Me.BPlantQty4.Text = ""
        Me.BPlantCost4.Text = ""
        Me.BPlantAmount4.Text = ""
        Me.BPlantRemarks4.Text = ""

        Label178.Text = "Pipeline cost details Saved. "
        Me.PipelineExpSave.Enabled = False


    End Sub

    Private Sub PipelineExpSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PipelineExpSave.Click
        Dim currentsheetname2 As String, currentsheetname1 As String, intI As Integer
        currentsheetname2 = "Extra pipeline"
        Me.btnQuit.Enabled = True
        Me.btnClose.Enabled = False
        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With

        For intI = 0 To xlWorkbook.Sheets.Count - 1
            xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
            currentsheetname1 = xlWorksheet.Name
            If UCase(Trim(currentsheetname1)) = UCase(Trim(currentsheetname2)) Then
                SheetNo = intI + 1
                Exit For
            End If
        Next
        intI = 1

        xlRange = xlWorksheet.Range("ExPipe_SlNo").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.PipelineQty1.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.PipelineCost1.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.PipelineRemarks1.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("ExPipe_SlNo").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.PipelineQty2.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.PipelineCost2.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.PipelineRemarks2.Text
        Label178.Text = "Pipeline related expenses saved. "
        Me.PipelineExpSave.Enabled = False
    End Sub

    Private Sub PipelineExpClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PipelineExpClear.Click
        Me.PipelineQty1.Text = ""
        Me.pipelineAmount1.Text = ""
        Me.PipelineRemarks1.Text = ""

        Me.PipelineQty1.Text = ""
        Me.pipelineAmount1.Text = ""
        Me.PipelineRemarks1.Text = ""


    End Sub

    Private Sub Esal_Nos_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Esal_Nos.TextChanged
        Me.EsalMonthTotal.Text = Val(Me.Esal_Nos.Text) * Val(Me.EsalAvgSalary.Text)
        Dim span As TimeSpan = mEndDate.Subtract(mStartDate)
        Dim ProjectDays As Integer = span.Days
        Me.EsalProjectTotal.Text = Val(Me.EsalMonthTotal.Text) * System.Math.Round(ProjectDays / 30, 0)
    End Sub

    Private Sub EsalAvgSalary_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EsalAvgSalary.TextChanged
        Me.EsalMonthTotal.Text = Val(Me.Esal_Nos.Text) * Val(Me.EsalAvgSalary.Text)
        Dim span As TimeSpan = mEndDate.Subtract(mStartDate)
        Dim ProjectDays As Integer = span.Days
        Me.EsalProjectTotal.Text = Val(Me.EsalMonthTotal.Text) * System.Math.Round(ProjectDays / 30, 0)
    End Sub

    Private Sub btnEsalSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEsalSave.Click
        Dim currentsheetname2 As String, currentsheetname1 As String, intI As Integer
        currentsheetname2 = "Elect salary"
        Me.btnQuit.Enabled = True
        Me.btnClose.Enabled = False
        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With

        For intI = 0 To xlWorkbook.Sheets.Count - 1
            xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
            currentsheetname1 = xlWorksheet.Name
            If UCase(Trim(currentsheetname1)) = UCase(Trim(currentsheetname2)) Then
                SheetNo = intI + 1
                Exit For
            End If
        Next
        intI = 1
        xlRange = xlWorksheet.Range("Esal_NoOFPeople")
        xlRange.Value = Me.Esal_Nos.Text
        xlRange = xlWorksheet.Range("Esal_AvgSalary")
        xlRange.Value = Me.EsalAvgSalary.Text
        Label172.Text = "Electricians cost details Saved. "
        Me.btnEsalSave.Enabled = False

    End Sub

    Private Sub btnEsalClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEsalClear.Click
        Me.Esal_Nos.Text = ""
        Me.EsalAvgSalary.Text = ""
    End Sub


    Private Sub btnMiscExpSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMiscExpSave.Click
        Dim currentsheetname2 As String, currentsheetname1 As String, intI As Integer
        currentsheetname2 = "Misc and Non ERP Purchases"
        Me.btnQuit.Enabled = True
        Me.btnClose.Enabled = False
        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With

        For intI = 0 To xlWorkbook.Sheets.Count - 1
            xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
            currentsheetname1 = xlWorksheet.Name
            If UCase(Trim(currentsheetname1)) = UCase(Trim(currentsheetname2)) Then
                SheetNo = intI + 1
                Exit For
            End If
        Next
        intI = 1

        xlRange = xlWorksheet.Range("Misc_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.Misc_Amount1.Text
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = Me.Misc_Remarks1.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("Misc_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.Misc_Amount2.Text
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = Me.Misc_Remarks2.Text


        intI = intI + 1
        xlRange = xlWorksheet.Range("Misc_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.Misc_Amount3.Text
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = Me.Misc_Remarks3.Text


        intI = intI + 1
        xlRange = xlWorksheet.Range("Misc_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.Misc_Amount4.Text
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = Me.Misc_Remarks4.Text


        intI = intI + 1
        xlRange = xlWorksheet.Range("Misc_Slno").Offset(intI, 0)
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.Misc_Amount5.Text
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = Me.Misc_Remarks5.Text

        Label170.Text = "Misc Expenses details Saved. "
        Me.btnMiscExpSave.Enabled = False

    End Sub

    Private Sub btnMiscExpClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMiscExpClear.Click
        Me.Misc_Amount1.Text = ""
        Me.Misc_Remarks1.Text = ""

        Me.Misc_Amount2.Text = ""
        Me.Misc_Remarks2.Text = ""

        Me.Misc_Amount3.Text = ""
        Me.Misc_Remarks3.Text = ""

        Me.Misc_Amount4.Text = ""
        Me.Misc_Remarks4.Text = ""

        Me.Misc_Amount5.Text = ""
        Me.Misc_Remarks5.Text = ""

    End Sub

    Private Sub btnPrevTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevTab.Click
        Me.tbclEqipsEntry.SelectTab(4)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.tbclEqipsEntry.SelectTab(5)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.tbclEqipsEntry.SelectTab(6)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.tbclEqipsEntry.SelectTab(7)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.tbclEqipsEntry.SelectTab(8)
    End Sub

    Private Sub PipelineQty1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PipelineQty1.TextChanged
        Me.pipelineAmount1.Text = Val(Me.PipelineQty1.Text) * Val(Me.PipelineCost1.Text)
    End Sub

    Private Sub PipelineCost1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PipelineCost1.TextChanged
        Me.pipelineAmount1.Text = Val(Me.PipelineQty1.Text) * Val(Me.PipelineCost1.Text)
    End Sub

    Private Sub PipelineQty2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PipelineQty2.TextChanged
        Me.PipelineAmount2.Text = Val(Me.PipelineQty2.Text) * Val(Me.PipelineCost2.Text)

    End Sub

    Private Sub PipelineCost2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PipelineCost2.TextChanged
        Me.PipelineAmount2.Text = Val(Me.PipelineQty2.Text) * Val(Me.PipelineCost2.Text)
    End Sub

    Private Sub EsalMonthTotal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EsalMonthTotal.TextChanged

    End Sub

    Private Sub tbclEqipsEntry_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbclEqipsEntry.SelectedIndexChanged
        If Me.tbclEqipsEntry.SelectedTab.Name = "MiscExp" Then
            Label170.Text = ""
            Me.btnMiscExpSave.Enabled = True
        ElseIf Me.tbclEqipsEntry.SelectedTab.Name = "ElectriciansCost" Then
            Label172.Text = ""
            Me.btnEsalSave.Enabled = True
        ElseIf Me.tbclEqipsEntry.SelectedTab.Name = "PipelineExp" Then
            Label178.Text = ""
            Me.PipelineExpSave.Enabled = True
        ElseIf Me.tbclEqipsEntry.SelectedTab.Name = "BPlantExp" Then
            Label179.Text = ""
            Me.btnBPlantExpSave.Enabled = True
        ElseIf Me.tbclEqipsEntry.SelectedTab.Name = "TowerCraneRelatedExp" Then
            Label180.Text = ""
            Me.btnTCRExpSave.Enabled = True
        ElseIf Me.tbclEqipsEntry.SelectedTab.Name = "StaffSalary" Then
            Label129.Text = ""
            Me.btnSalaryCostSave.Enabled = True
        End If
    End Sub

    Private Sub StaffSalry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StaffSalary.Click

    End Sub

    Private Sub TextBox14_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Forman_Cost.TextChanged
        Me.TotSalaryCost.Text = Val(Me.Manager_Cost.Text) + Val(Me.Eng_Cost.Text) + Val(Me.Sup_Cost.Text) + Val(Me.Forman_Cost.Text) + Val(Me.FormanElec_Cost.Text)
    End Sub

    Private Sub TextBox13_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Sup_Cost.TextChanged
        Me.TotSalaryCost.Text = Val(Me.Manager_Cost.Text) + Val(Me.Eng_Cost.Text) + Val(Me.Sup_Cost.Text) + Val(Me.Forman_Cost.Text) + Val(Me.FormanElec_Cost.Text)
    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Manager_Cost.TextChanged
        Me.TotSalaryCost.Text = Val(Me.Manager_Cost.Text) + Val(Me.Eng_Cost.Text) + Val(Me.Sup_Cost.Text) + Val(Me.Forman_Cost.Text) + Val(Me.FormanElec_Cost.Text)
    End Sub

    Private Sub Manager_Nos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Manager_Nos.Validated
        Me.Manager_Cost.Text = Val(Int(Me.Manager_Nos.Text)) * Val(Me.Manager_SalaryPM.Text) * Val(Me.Manager_Months.Text)
    End Sub

    Private Sub Manager_Nos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Manager_Nos.Validating
        If Not IsNumeric(Manager_Nos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub Eng_Nos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Eng_Nos.Validated
        Me.Eng_Cost.Text = Val(Int(Me.Eng_Nos.Text)) * Val(Me.Eng_SalaryPM.Text) * Val(Me.Sup_Months.Text)
    End Sub

    Private Sub Eng_Nos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Eng_Nos.Validating
        If Not IsNumeric(Eng_Nos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub Sup_Nos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Sup_Nos.Validated
        Me.Sup_Cost.Text = Val(Int(Me.Sup_Nos.Text)) * Val(Me.Sup_SalaryPM.Text) * Val(Me.Sup_Months.Text)
    End Sub

    Private Sub Sup_Nos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Sup_Nos.Validating
        If Not IsNumeric(Sup_Nos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub
    Private Sub Forman_Nos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Forman_Nos.Validated
        Me.Forman_Cost.Text = Val(Int(Me.Forman_Nos.Text)) * Val(Me.Foreman_Months.Text) * Val(Me.Foreman_Months.Text)
    End Sub

    Private Sub Forman_Nos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Forman_Nos.Validating
        If Not IsNumeric(Forman_Nos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub Manager_SalaryPM_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Manager_SalaryPM.Validated
        Me.Manager_Cost.Text = Val(Int(Me.Manager_Nos.Text)) * Val(Me.Manager_SalaryPM.Text) * Val(Me.Manager_Months.Text)
    End Sub

    Private Sub Manager_SalaryPM_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Manager_SalaryPM.Validating
        If Not IsNumeric(Manager_SalaryPM.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub Eng_SalaryPM_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Eng_SalaryPM.Validated
        Me.Eng_Cost.Text = Val(Int(Me.Eng_Nos.Text)) * Val(Me.Eng_SalaryPM.Text) * Val(Me.Eng_Months.Text)
    End Sub

    Private Sub Eng_SalaryPM_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Eng_SalaryPM.Validating
        If Not IsNumeric(Eng_SalaryPM.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub Sup_SalaryPM_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Sup_SalaryPM.Validated
        Me.Sup_Cost.Text = Val(Int(Me.Sup_Nos.Text)) * Val(Me.Sup_SalaryPM.Text) * Val(Me.Sup_Months.Text)
    End Sub

    Private Sub Sup_SalaryPM_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Sup_SalaryPM.Validating
        If Not IsNumeric(Sup_SalaryPM.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub Foreman_SalaryPM_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Foreman_SalaryPM.Validated
        Me.Forman_Cost.Text = Val(Int(Me.Forman_Nos.Text)) * Val(Me.Foreman_SalaryPM.Text) * Val(Me.Foreman_Months.Text)
    End Sub

    Private Sub Foreman_SalaryPM_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Foreman_SalaryPM.Validating
        If Not IsNumeric(Foreman_SalaryPM.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub Manager_Months_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Manager_Months.Validated
        Me.Manager_Cost.Text = Val(Int(Me.Manager_Nos.Text)) * Val(Me.Manager_SalaryPM.Text) * Val(Me.Manager_Months.Text)
    End Sub

    Private Sub Manager_Months_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Manager_Months.Validating
        If Not IsNumeric(Manager_Months.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub Eng_Months_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Eng_Months.Validated
        Me.Eng_Cost.Text = Val(Int(Me.Eng_Nos.Text)) * Val(Me.Eng_SalaryPM.Text) * Val(Me.Eng_Months.Text)
    End Sub

    Private Sub Eng_Months_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Eng_Months.Validating
        If Not IsNumeric(Eng_Months.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub Sup_Months_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Sup_Months.Validated
        Me.Sup_Cost.Text = Val(Int(Me.Sup_Nos.Text)) * Val(Me.Sup_SalaryPM.Text) * Val(Me.Sup_Months.Text)
    End Sub

    Private Sub Sup_Months_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Sup_Months.Validating
        If Not IsNumeric(Sup_Months.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub Foreman_Months_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Foreman_Months.Validated
        Me.Forman_Cost.Text = Val(Int(Me.Forman_Nos.Text)) * Val(Me.Foreman_SalaryPM.Text) * Val(Me.Foreman_Months.Text)
    End Sub

    Private Sub Foreman_Months_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Foreman_Months.Validating
        If Not IsNumeric(Foreman_Months.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub Eng_Cost_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Eng_Cost.TextChanged
        Me.TotSalaryCost.Text = Val(Me.Manager_Cost.Text) + Val(Me.Eng_Cost.Text) + Val(Me.Sup_Cost.Text) + Val(Me.Forman_Cost.Text) + Val(Me.FormanElec_Cost.Text)
    End Sub

    Private Sub FormanElec_Nos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles FormanElec_Nos.Validated
        Me.FormanElec_Cost.Text = Val(Me.FormanElec_Nos.Text) * Val(Me.FormanElec_SalaryPM.Text) * Val(Me.FormanElec_Months.Text)
    End Sub

    Private Sub FormanElec_Nos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles FormanElec_Nos.Validating
        If Not IsNumeric(FormanElec_Nos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub FormanElec_SalaryPM_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles FormanElec_SalaryPM.Validated
        Me.FormanElec_Cost.Text = Val(Me.FormanElec_Nos.Text) * Val(Me.FormanElec_SalaryPM.Text) * Val(Me.FormanElec_Months.Text)
    End Sub

    Private Sub FormanElec_SalaryPM_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles FormanElec_SalaryPM.Validating
        If Not IsNumeric(FormanElec_SalaryPM.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub FormanElec_Months_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles FormanElec_Months.Validated
        Me.FormanElec_Cost.Text = Val(Me.FormanElec_Nos.Text) * Val(Me.FormanElec_SalaryPM.Text) * Val(Me.FormanElec_Months.Text)
    End Sub

    Private Sub FormanElec_Months_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles FormanElec_Months.Validating
        If Not IsNumeric(FormanElec_Months.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub FormanElec_Cost_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles FormanElec_Cost.TextChanged
        Me.TotSalaryCost.Text = Val(Me.Manager_Cost.Text) + Val(Me.Eng_Cost.Text) + Val(Me.Sup_Cost.Text) + Val(Me.Forman_Cost.Text) + Val(Me.FormanElec_Cost.Text)
    End Sub

    Private Sub btnSalaryCostSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalaryCostSave.Click
        Dim currentsheetname2 As String, currentsheetname1 As String, intI As Integer
        currentsheetname2 = "Staff Salary"
        Me.btnQuit.Enabled = True
        Me.btnClose.Enabled = False
        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With

        For intI = 0 To xlWorkbook.Sheets.Count - 1
            xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
            currentsheetname1 = xlWorksheet.Name
            If UCase(Trim(currentsheetname1)) = UCase(Trim(currentsheetname2)) Then
                SheetNo = intI + 1
                Exit For
            End If
        Next
        intI = 1

        xlRange = xlWorksheet.Range("StaffSalary_ManagerNos")
        xlRange.Value = Me.Manager_Nos.Text
        xlRange = xlWorksheet.Range("StaffSalary_EngNos")
        xlRange.Value = Me.Eng_Nos.Text
        xlRange = xlWorksheet.Range("StaffSalary_SupNos")
        xlRange.Value = Me.Sup_Nos.Text
        xlRange = xlWorksheet.Range("StaffSalary_FormanNos")
        xlRange.Value = Me.Forman_Nos.Text
        xlRange = xlWorksheet.Range("StaffSalary_FormanElecNos")
        xlRange.Value = Me.FormanElec_Nos.Text
        xlRange = xlWorksheet.Range("StaffSalary_ManagerSalaryPM")
        xlRange.Value = Me.Manager_SalaryPM.Text
        xlRange = xlWorksheet.Range("StaffSalary_EngSalaryPM")
        xlRange.Value = Me.Eng_SalaryPM.Text
        xlRange = xlWorksheet.Range("StaffSalary_SupSalaryPM")
        xlRange.Value = Me.Sup_SalaryPM.Text
        xlRange = xlWorksheet.Range("StaffSalary_FormanSalaryPM")
        xlRange.Value = Me.Foreman_SalaryPM.Text
        xlRange = xlWorksheet.Range("StaffSalary_FormanElecSalaryPM")
        xlRange.Value = Me.FormanElec_SalaryPM.Text
        xlRange = xlWorksheet.Range("StaffSalary_ManagerMonths")
        xlRange.Value = Me.Manager_Months.Text
        xlRange = xlWorksheet.Range("StaffSalary_EngMonths")
        xlRange.Value = Me.Eng_Months.Text
        xlRange = xlWorksheet.Range("StaffSalary_SupMonths")
        xlRange.Value = Me.Sup_Months.Text
        xlRange = xlWorksheet.Range("StaffSalary_FormanMonths")
        xlRange.Value = Me.Foreman_Months.Text
        xlRange = xlWorksheet.Range("StaffSalary_FormanElecMonths")
        xlRange.Value = Me.FormanElec_Months.Text

        Label129.Text = "Salary Expense Saved. "
        Me.btnMiscExpSave.Enabled = False
    End Sub

    Private Sub btnSalaryCostClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalaryCostClear.Click
        Me.Manager_Nos.Text = ""
        Me.Manager_SalaryPM.Text = ""
        Me.Manager_Months.Text = ""
        Me.Eng_Nos.Text = ""
        Me.Eng_SalaryPM.Text = ""
        Me.Eng_Months.Text = ""
        Me.Sup_Nos.Text = ""
        Me.Sup_SalaryPM.Text = ""
        Me.Sup_Months.Text = ""
        Me.Forman_Nos.Text = ""
        Me.Foreman_SalaryPM.Text = ""
        Me.Foreman_Months.Text = ""
        Me.FormanElec_Nos.Text = ""
        Me.FormanElec_SalaryPM.Text = ""
        Me.FormanElec_Months.Text = ""
        Me.btnSalaryCostSave.Enabled = True
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Me.tbclEqipsEntry.SelectTab(0)
        Currenttab = 0
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Me.tbclEqipsEntry.SelectTab(1)
        Currenttab = 1
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Me.tbclEqipsEntry.SelectTab(2)
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Me.tbclEqipsEntry.SelectTab(3)
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Me.tbclEqipsEntry.SelectTab(4)
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Me.tbclEqipsEntry.SelectTab(5)
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Me.tbclEqipsEntry.SelectTab(6)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Me.tbclEqipsEntry.SelectTab(9)
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Me.tbclEqipsEntry.SelectTab(7)
    End Sub

    Private Sub TowerCraneRelatedExp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TowerCraneRelatedExp.Click

    End Sub

    Private Sub TextBox1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Qty6.TextChanged
        Me.Amount6.Text = Val(Me.Qty6.Text) * Val(Me.Cost6.Text)
    End Sub

    Private Sub TextBox2_TextChanged_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cost6.TextChanged
        Me.Amount6.Text = Val(Me.Qty6.Text) * Val(Me.Cost6.Text)
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Amount6.TextChanged

    End Sub

    Private Sub Qty7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Qty7.TextChanged
        Me.Amount7.Text = Val(Me.Qty7.Text) * Val(Me.Cost7.Text)
    End Sub

    Private Sub Cost7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cost7.TextChanged
        Me.Amount7.Text = Val(Me.Qty7.Text) * Val(Me.Cost7.Text)
    End Sub

    Private Sub Qty8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Qty8.TextChanged
        Me.Amount8.Text = Val(Me.Qty8.Text) * Val(Me.Cost8.Text)
    End Sub

    Private Sub Cost8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cost8.TextChanged
        Me.Amount8.Text = Val(Me.Qty8.Text) * Val(Me.Cost8.Text)
    End Sub

    Private Sub BPlantQty5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BPlantQty5.TextChanged
        Me.BPlantAmount5.Text = Val(Me.BPlantQty5.Text) * Val(Me.BPlantCost5.Text)
    End Sub

    Private Sub BPlantCost5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BPlantCost5.TextChanged
        Me.BPlantAmount5.Text = Val(Me.BPlantQty5.Text) * Val(Me.BPlantCost5.Text)
    End Sub

    Private Sub BPlantQty6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BPlantQty6.TextChanged
        Me.BPlantAmount6.Text = Val(Me.BPlantQty6.Text) * Val(Me.BPlantCost6.Text)
    End Sub

    Private Sub BPlantCost6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BPlantCost6.TextChanged
        Me.BPlantAmount6.Text = Val(Me.BPlantQty6.Text) * Val(Me.BPlantCost6.Text)
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub frmOptions_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        '    If BlankForm Then
        '        Exit Sub
        '    Else
        '        BlankForm = True
        '        Me.cmbMinoEquip_Category.Text = SelectedCategory
        '        If EditOperationSheet = "Concreting" Or EditOperationSheet = "Cranes" Or EditOperationSheet = "Material Handling" Or _
        '        EditOperationSheet = "Non Concreting" Or EditOperationSheet = "DG Sets" Or EditOperationSheet = "Conveyance" Or _
        '        EditOperationSheet = "Major Others" Then
        '            Me.cmbMajorEquipModel_SelectedIndexChanged(Me, e)
        '        ElseIf EditOperationSheet = "external Hire" Or EditOperationSheet = "External Others" Then
        '            Me.cmbExtEquip_Model_SelectedIndexChanged(Me, e)
        '        ElseIf EditOperationSheet = "Minor Eqpts" Then
        '            Me.cmbMinorEquip_Model_SelectedIndexChanged(Me, e)
        '        End If
        '    End If
    End Sub

End Class
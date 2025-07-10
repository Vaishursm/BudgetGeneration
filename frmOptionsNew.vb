Imports System.Data
Imports System.Data.OleDb

Public Class frmOptionsNew


    Public strConnection As String

    Public moledbConnection As OleDbConnection
    Dim strStatement As String
    Dim moledbCommand As OleDbCommand
    Dim mOledbDataAdapter As OleDbDataAdapter
    Dim mReader As OleDbDataReader
    Dim mDataSet As DataSet
    Dim RowCount As Integer = 0
    Dim RangeCols As Integer
    Dim cntItems As Integer = 0
    Dim txtMonths As Integer
    Dim HrsPerMonth As Integer
    Dim txtRepvalue As Long
    Dim txtDepreciation As Single
    Dim txtHirecharges As Long
    Dim txtFuelPerHr As Single
    Dim txtPowerperHr As Single
    Dim txtOprCostPerMCPerMonth As Single
    Dim txtMaintCostperMC_PerMonth As Single
    Dim txtFuelperUnitPerMonth As Single
    Dim txtFuelCostPerMonth As Integer
    Dim txtFuelCostProject As double
    Dim txtPowerCostperMonth As Double
    Dim txtPowerCostProject As Double
    Dim txtOprCostPerMonth As Long
    Dim txtOprCostProject As Long
    Dim txtConsumablesPerMonth As Single
    Dim txtConsumablesProject As Double
    Dim txtDrive As String
    Dim OperatingCost_MajorEquips As Double
    Dim txtMaintPercPerMC_PerMonth As Single
    Dim txtshifts As Single
    Dim mTabindex As Integer = 1
    Dim intI As Integer, mcategory As String

    Dim Concretesaved As Boolean = False
    Dim ConveyanceSaved As Boolean = False
    Dim CranesSaved As Boolean = False
    Dim dgsetssaved As Boolean = False
    Dim MHSaved As Boolean = False
    Dim NCSaved As Boolean = False
    Dim MajOthersSaved As Boolean = False
    Dim MinorEquipsSaved As Boolean = False
    Dim HiredEquipsSaved As Boolean = False
    Dim FixedExpSaved As Boolean
    Dim fixedBpExpSaved As Boolean
    Dim LightingSaved As Boolean
    Dim oForm As Form


    Private Sub frmOptionsNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label255.Text = Label255.Text & " " & mMainTitle2 & " - Budget Generation"
        VP = 0    'Me.Top
        HP = 0   'Me.Left + 10
        appPath = Application.StartupPath
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
        mDataSet = New DataSet()
        If moledbConnection Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If

        Me.btnClose.Enabled = True
        Me.btnQuit.Enabled = True
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        mcategory = "Concreting"
        mTabindex = 1
        Me.optMajConcrete.Checked = True
        Me.lblmessage.Visible = False

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
        End If
        If mWorkingMode = "New" Then
            CopyPowerReqTemplate()
        End If
        deleteOlddata("PowerGen Cost")
        moledbCommand = Nothing
        Panel2.Left = 1
        Panel2.Top = 20
        Panel2.Height = 30
        Panel2.Visible = True
        Panel4.Visible = False
        Panel5.Visible = False
        Panel6.Visible = False
        Panel7.Visible = False
        Me.tbcBdgetHeads.SelectTab(Currenttab)
    End Sub
    Private Sub LoadControlsInpage(ByVal mcategoryname As String)
        Dim intI As Integer
        Dim cnt As Integer
        Dim valCategory As String, valEquipsname As String, valCapacity As String, valMake As String, valModel As String
        Dim valMobDate As String, valDemobDate As String, valQty As Integer, valHPM As Single, valRepvalue As Long, valConcQty As Integer
        Dim valchkd As Integer


        If UCase(mcategoryname) = UCase("Concreting") Then
            HP = Me.Left + 1
            intI = 1
            HP = 1
            VP = 0
            CategoryTextBoxes.Clear()
            EquipNameTextBoxes.Clear()
            CapacityTextBoxes.Clear()
            MakeModelTextBoxes.Clear()
            QtyTextBoxes.Clear()
            HrsPermonthTextBoxes.Clear()
            RepValueTextBoxes.Clear()
            MaintPercTextBoxes.Clear()
            concreteqtyTextboxes.Clear()
            Checkboxes.Clear()
            MobdatePickers.Clear()
            DemobDatePickers.Clear()
            DepPercComboboxes.Clear()
            ShiftsComboboxes.Clear()
            AddButtons.Clear()

            Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
            mTabindex = 0
            If concEquipsNames.Length > 0 Then
                For cnt = 0 To concEquipsNames.Length - 1
                    valCategory = mcategoryname
                    valEquipsname = concEquipsNames(cnt)
                    valCapacity = concEquipsCapacity(cnt)
                    valMake = concEquipsMake(cnt)
                    valModel = concEquipsModel(cnt)
                    valMobDate = concEquipsMobDate(cnt)
                    valDemobDate = concEquipsDemobDate(cnt)
                    valQty = concEquipsQty(cnt)
                    valHPM = concEquipsHPM(cnt)
                    valRepvalue = concEquipsRepValue(cnt)
                    valConcQty = concEquipsConcQty(cnt)
                    valchkd = concEquipsChkd(cnt)
                    buildMajorcontrols(valCategory, valEquipsname, valCapacity, _
                        valMake, valModel, valMobDate, valDemobDate, valQty, valHPM, valRepvalue, valConcQty, valchkd, cnt)
                Next
            End If
        ElseIf UCase(mcategoryname) = UCase("Conveyance") Then
            HP = Me.Left + 1
            intI = 1
            HP = 1
            VP = 0
            'Me.Refresh()
            CategoryTextBoxes.Clear()
            EquipNameTextBoxes.Clear()
            CapacityTextBoxes.Clear()
            MakeModelTextBoxes.Clear()
            QtyTextBoxes.Clear()
            HrsPermonthTextBoxes.Clear()
            RepValueTextBoxes.Clear()
            MaintPercTextBoxes.Clear()
            concreteqtyTextboxes.Clear()
            Checkboxes.Clear()
            MobdatePickers.Clear()
            DemobDatePickers.Clear()
            DepPercComboboxes.Clear()
            ShiftsComboboxes.Clear()
            AddButtons.Clear()

            If convEquipsNames.Length > 0 Then
                For cnt = 0 To convEquipsNames.Length - 1
                    valCategory = mcategoryname
                    valEquipsname = convEquipsNames(cnt)
                    valCapacity = convEquipsCapacity(cnt)
                    valMake = convEquipsMake(cnt)
                    valModel = convEquipsModel(cnt)
                    valMobDate = convEquipsMobDate(cnt)
                    valDemobDate = convEquipsDemobDate(cnt)
                    valQty = convEquipsQty(cnt)
                    valHPM = convEquipsHPM(cnt)
                    valRepvalue = convEquipsRepValue(cnt)
                    valConcQty = convEquipsConcQty(cnt)
                    valchkd = convEquipsChkd(cnt)
                    buildMajorcontrols(valCategory, valEquipsname, valCapacity, _
                        valMake, valModel, valMobDate, valDemobDate, valQty, valHPM, valRepvalue, valConcQty, valchkd, cnt)
                Next
            End If
        ElseIf UCase(mcategoryname) = UCase("Cranes") Then
            Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
            HP = Me.Left + 1
            intI = 1
            HP = 1
            VP = 0
            CategoryTextBoxes.Clear()
            EquipNameTextBoxes.Clear()
            CapacityTextBoxes.Clear()
            MakeModelTextBoxes.Clear()
            QtyTextBoxes.Clear()
            HrsPermonthTextBoxes.Clear()
            RepValueTextBoxes.Clear()
            MaintPercTextBoxes.Clear()
            concreteqtyTextboxes.Clear()
            Checkboxes.Clear()
            MobdatePickers.Clear()
            DemobDatePickers.Clear()
            DepPercComboboxes.Clear()
            ShiftsComboboxes.Clear()
            AddButtons.Clear()

            If craneEquipsNames.Length > 0 Then
                For cnt = 0 To craneEquipsNames.Length - 1
                    valCategory = mcategoryname
                    valEquipsname = craneEquipsNames(cnt)
                    valCapacity = craneEquipsCapacity(cnt)
                    valMake = craneEquipsMake(cnt)
                    valModel = craneEquipsModel(cnt)
                    valMobDate = craneEquipsMobDate(cnt)
                    valDemobDate = craneEquipsDemobDate(cnt)
                    valQty = craneEquipsQty(cnt)
                    valHPM = craneEquipsHPM(cnt)
                    valRepvalue = craneEquipsRepValue(cnt)
                    valConcQty = craneEquipsConcQty(cnt)
                    valchkd = craneEquipsChkd(cnt)
                    buildMajorcontrols(valCategory, valEquipsname, valCapacity, _
                        valMake, valModel, valMobDate, valDemobDate, valQty, valHPM, valRepvalue, valConcQty, valchkd, cnt)
                Next
            End If
        ElseIf UCase(mcategoryname) = UCase("DG sets") Then
            Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
            HP = Me.Left + 1
            intI = 1
            HP = 1
            VP = 0
            CategoryTextBoxes.Clear()
            EquipNameTextBoxes.Clear()
            CapacityTextBoxes.Clear()
            MakeModelTextBoxes.Clear()
            QtyTextBoxes.Clear()
            HrsPermonthTextBoxes.Clear()
            RepValueTextBoxes.Clear()
            MaintPercTextBoxes.Clear()
            concreteqtyTextboxes.Clear()
            Checkboxes.Clear()
            MobdatePickers.Clear()
            DemobDatePickers.Clear()
            DepPercComboboxes.Clear()
            ShiftsComboboxes.Clear()
            AddButtons.Clear()

            If dgsetsEquipsNames.Length > 0 Then
                For cnt = 0 To dgsetsEquipsNames.Length - 1
                    valCategory = mcategoryname
                    valEquipsname = dgsetsEquipsNames(cnt)
                    valCapacity = dgsetsEquipsCapacity(cnt)
                    valMake = dgsetsEquipsMake(cnt)
                    valModel = dgsetsEquipsModel(cnt)
                    valMobDate = dgsetsEquipsMobDate(cnt)
                    valDemobDate = dgsetsEquipsDemobDate(cnt)
                    valQty = dgsetsEquipsQty(cnt)
                    valHPM = dgsetsEquipsHPM(cnt)
                    valRepvalue = dgsetsEquipsRepValue(cnt)
                    valConcQty = dgsetsEquipsConcQty(cnt)
                    valchkd = dgsetsEquipsChkd(cnt)
                    buildMajorcontrols(valCategory, valEquipsname, valCapacity, _
                        valMake, valModel, valMobDate, valDemobDate, valQty, valHPM, valRepvalue, valConcQty, valchkd, cnt)
                Next
            End If
        ElseIf UCase(mcategoryname) = UCase("Material Handling") Then
            Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
            HP = Me.Left + 1
            intI = 1
            HP = 1
            VP = 0
            'Me.Refresh()
            CategoryTextBoxes.Clear()
            EquipNameTextBoxes.Clear()
            CapacityTextBoxes.Clear()
            MakeModelTextBoxes.Clear()
            QtyTextBoxes.Clear()
            HrsPermonthTextBoxes.Clear()
            RepValueTextBoxes.Clear()
            MaintPercTextBoxes.Clear()
            concreteqtyTextboxes.Clear()
            Checkboxes.Clear()
            MobdatePickers.Clear()
            DemobDatePickers.Clear()
            DepPercComboboxes.Clear()
            ShiftsComboboxes.Clear()
            AddButtons.Clear()

            If MHEquipsNames.Length > 0 Then
                For cnt = 0 To MHEquipsNames.Length - 1
                    valCategory = mcategoryname
                    valEquipsname = MHEquipsNames(cnt)
                    valCapacity = MHEquipsCapacity(cnt)
                    valMake = MHEquipsMake(cnt)
                    valModel = MHEquipsModel(cnt)
                    valMobDate = MHEquipsMobDate(cnt)
                    valDemobDate = MHEquipsDemobDate(cnt)
                    valQty = MHEquipsQty(cnt)
                    valHPM = MHEquipsHPM(cnt)
                    valRepvalue = MHEquipsRepValue(cnt)
                    valConcQty = MHEquipsConcQty(cnt)
                    valchkd = MHEquipsChkd(cnt)
                    buildMajorcontrols(valCategory, valEquipsname, valCapacity, _
                        valMake, valModel, valMobDate, valDemobDate, valQty, valHPM, valRepvalue, valConcQty, valchkd, cnt)
                Next
            End If
        ElseIf UCase(mcategoryname) = UCase("Non Concreting") Then
            Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
            HP = Me.Left + 1
            intI = 1
            HP = 1
            VP = 0
            'Me.Refresh()
            CategoryTextBoxes.Clear()
            EquipNameTextBoxes.Clear()
            CapacityTextBoxes.Clear()
            MakeModelTextBoxes.Clear()
            QtyTextBoxes.Clear()
            HrsPermonthTextBoxes.Clear()
            RepValueTextBoxes.Clear()
            MaintPercTextBoxes.Clear()
            concreteqtyTextboxes.Clear()
            Checkboxes.Clear()
            MobdatePickers.Clear()
            DemobDatePickers.Clear()
            DepPercComboboxes.Clear()
            ShiftsComboboxes.Clear()
            AddButtons.Clear()

            If NCEquipsNames.Length > 0 Then
                For cnt = 0 To NCEquipsNames.Length - 1
                    valCategory = mcategoryname
                    valEquipsname = NCEquipsNames(cnt)
                    valCapacity = NCEquipsCapacity(cnt)
                    valMake = NCEquipsMake(cnt)
                    valModel = NCEquipsModel(cnt)
                    valMobDate = NCEquipsMobDate(cnt)
                    valDemobDate = NCEquipsDemobDate(cnt)
                    valQty = NCEquipsQty(cnt)
                    valHPM = NCEquipsHPM(cnt)
                    valRepvalue = NCEquipsRepValue(cnt)
                    valConcQty = NCEquipsConcQty(cnt)
                    valchkd = NCEquipsChkd(cnt)
                    buildMajorcontrols(valCategory, valEquipsname, valCapacity, _
                        valMake, valModel, valMobDate, valDemobDate, valQty, valHPM, valRepvalue, valConcQty, valchkd, cnt)
                Next
            End If
        ElseIf UCase(mcategoryname) = UCase("Major Others") Then
            Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
            HP = Me.Left + 1
            intI = 1
            HP = 1
            VP = 0
            CategoryTextBoxes.Clear()
            EquipNameTextBoxes.Clear()
            CapacityTextBoxes.Clear()
            MakeModelTextBoxes.Clear()
            QtyTextBoxes.Clear()
            HrsPermonthTextBoxes.Clear()
            RepValueTextBoxes.Clear()
            MaintPercTextBoxes.Clear()
            concreteqtyTextboxes.Clear()
            Checkboxes.Clear()
            MobdatePickers.Clear()
            DemobDatePickers.Clear()
            DepPercComboboxes.Clear()
            ShiftsComboboxes.Clear()
            AddButtons.Clear()

            If majOthersEquipsNames.Length > 0 Then
                For cnt = 0 To majOthersEquipsNames.Length - 1
                    valCategory = mcategoryname
                    valEquipsname = majOthersEquipsNames(cnt)
                    valCapacity = majOthersEquipsCapacity(cnt)
                    valMake = majOthersEquipsMake(cnt)
                    valModel = majOthersEquipsModel(cnt)
                    valMobDate = majOthersEquipsMobDate(cnt)
                    valDemobDate = majOthersEquipsDemobDate(cnt)
                    valQty = majOthersEquipsQty(cnt)
                    valHPM = majOthersEquipsHPM(cnt)
                    valRepvalue = majOthersEquipsRepValue(cnt)
                    valConcQty = majOthersEquipsConcQty(cnt)
                    valchkd = majOthersEquipsChkd(cnt)
                    buildMajorcontrols(valCategory, valEquipsname, valCapacity, _
                        valMake, valModel, valMobDate, valDemobDate, valQty, valHPM, valRepvalue, valConcQty, valchkd, cnt)
                Next
            End If
        End If

        Dim j As Integer
        For j = 0 To AddButtons.Count - 1
            If Checkboxes(j).Checked Then AddButtons(j).Enabled = True
        Next
        Me.Button1.Enabled = True
    End Sub
    Private Sub buildMajorcontrols(ByVal valCategory As String, ByVal valEquipname As String, ByVal valcapacity As String, _
          ByVal valmake As String, ByVal valmodel As String, ByVal valmobdate As String, ByVal valdemobdate As String, ByVal valqty As Integer, _
          ByVal valHPM As Single, ByVal valrepvalue As Long, ByVal valconcqty As Integer, ByVal valchkd As Integer, ByVal cnt As Integer)

        Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
        Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
        Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
        Dim txtHrsPerMonth As New TextBox
        Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox
        Dim txtRepValue As New TextBox, txtMaintPerc As New TextBox, chkSelected As New CheckBox
        Dim txtConcreteQty As New TextBox   ', cmbDrive As New ComboBox
        Dim btnAddExtra As New Button, chkval As Boolean
        Dim intI As Integer

        intI = cnt
        chkval = False
        chkSelected.Checked = valchkd

        chkSelected.Height = 30
        chkSelected.Width = 24
        chkSelected.Left = HP
        chkSelected.Top = VP - 3
        HP = HP + chkSelected.Width + 1
        chkSelected.Tag = intI
        chkSelected.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler chkSelected.CheckedChanged, AddressOf HandleCheckboxStatus
        Checkboxes.Add(intI, chkSelected)

        txtCategory.Text = valCategory
        txtCategory.WordWrap = True
        txtCategory.Height = 50
        txtCategory.Width = 65
        txtCategory.TextAlign = HorizontalAlignment.Center
        txtCategory.Left = HP
        txtCategory.Top = VP
        HP = HP + txtCategory.Width + 1
        txtCategory.Font = New Font("arial", 8)
        txtCategory.Tag = intI
        txtCategory.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCategory.Enabled = False
        CategoryTextBoxes.Add(intI, txtCategory)

        txtEquipname.Text = valEquipname
        ' txtEquipname.Multiline = True
        txtEquipname.WordWrap = True
        txtEquipname.Height = 60
        txtEquipname.Width = 150
        txtEquipname.TextAlign = HorizontalAlignment.Center
        txtEquipname.Left = HP
        txtEquipname.Top = VP
        HP = HP + txtEquipname.Width + 1
        txtEquipname.Font = New Font("arial", 8)
        txtEquipname.Tag = intI
        txtEquipname.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtEquipname.Enabled = False
        EquipNameTextBoxes.Add(intI, txtEquipname)

        txtCapacity.Text = valcapacity
        'txtCapacity.Multiline = True
        txtCapacity.WordWrap = True
        txtCapacity.Height = 30
        txtCapacity.Width = 70
        txtCapacity.TextAlign = HorizontalAlignment.Center
        txtCapacity.Left = HP
        txtCapacity.Top = VP
        HP = HP + txtCapacity.Width + 1
        txtCapacity.Font = New Font("arial", 8)
        txtCapacity.Tag = intI
        txtCapacity.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCapacity.Enabled = False
        CapacityTextBoxes.Add(intI, txtCapacity)

        txtMakeModel.Text = valmake & " / " & valmodel
        txtMakeModel.WordWrap = True
        txtMakeModel.Height = 50
        txtMakeModel.Width = 160
        txtMakeModel.TextAlign = HorizontalAlignment.Center
        txtMakeModel.Left = HP
        txtMakeModel.Top = VP
        HP = HP + txtMakeModel.Width + 1
        txtMakeModel.Font = New Font("arial", 8)
        txtMakeModel.Tag = intI
        txtMakeModel.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtMakeModel.Enabled = False
        MakeModelTextBoxes.Add(intI, txtMakeModel)

        dpMobDate.Value = valmobdate
        dpMobDate.Name = "MobDate"
        dpMobDate.Format = DateTimePickerFormat.Custom
        dpMobDate.CustomFormat = "dd-MMM-yyyy"
        dpMobDate.Height = 30
        dpMobDate.Width = 100
        dpMobDate.Left = HP
        dpMobDate.Top = VP
        HP = HP + dpMobDate.Width + 1
        dpMobDate.Font = New Font("arial", 8)
        dpMobDate.Enabled = False
        dpMobDate.Tag = intI
        dpMobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
        MobdatePickers.Add(intI, dpMobDate)

        dpDemobDate.Value = valdemobdate
        dpDemobDate.Name = "DemobDate"
        dpDemobDate.Format = DateTimePickerFormat.Custom
        dpDemobDate.CustomFormat = "dd-MMM-yyyy"
        dpDemobDate.Height = 30
        dpDemobDate.Width = 100
        dpDemobDate.Left = HP
        dpDemobDate.Top = VP
        HP = HP + dpDemobDate.Width + 1
        dpDemobDate.Font = New Font("arial", 8)
        dpDemobDate.Enabled = False
        dpDemobDate.Tag = intI
        dpDemobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
        DemobDatePickers.Add(intI, dpDemobDate)

        txtQty.Text = valqty
        txtQty.Height = 30
        txtQty.Width = 30
        txtQty.Left = HP
        txtQty.Top = VP
        HP = HP + txtQty.Width + 1
        txtQty.Font = New Font("arial", 8)
        txtQty.Enabled = False
        txtQty.Tag = intI
        txtQty.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler txtQty.Validated, AddressOf HandleQty
        QtyTextBoxes.Add(intI, txtQty)


        txtHrsPerMonth.Text = valHPM
        txtHrsPerMonth.Height = 30
        txtHrsPerMonth.Width = 60
        txtHrsPerMonth.Left = HP
        txtHrsPerMonth.Top = VP
        HP = HP + txtHrsPerMonth.Width + 1
        txtHrsPerMonth.Font = New Font("arial", 8)
        txtHrsPerMonth.Enabled = False
        txtHrsPerMonth.Tag = intI
        txtHrsPerMonth.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        HrsPermonthTextBoxes.Add(intI, txtHrsPerMonth)

        'cmbDepPerc.Items.Add("Fixed")
        cmbDepPerc.Items.Add(2.75)
        cmbDepPerc.Items.Add(1.25)
        cmbDepPerc.Items.Add(0.5)
        cmbDepPerc.SelectedIndex = 0
        cmbDepPerc.Height = 30
        cmbDepPerc.Width = 60
        cmbDepPerc.Left = HP
        cmbDepPerc.Top = VP
        HP = HP + cmbDepPerc.Width + 1
        cmbDepPerc.Font = New Font("arial", 8)
        cmbDepPerc.Enabled = False
        cmbDepPerc.Tag = intI
        cmbDepPerc.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        DepPercComboboxes.Add(intI, cmbDepPerc)

        cmbShifts.Items.Add(1)
        cmbShifts.Items.Add(1.5)
        cmbShifts.Items.Add(2)
        cmbShifts.SelectedIndex = 0
        cmbShifts.Height = 30
        cmbShifts.Width = 45
        cmbShifts.Left = HP
        cmbShifts.Top = VP
        HP = HP + cmbShifts.Width + 1
        cmbShifts.Font = New Font("arial", 8)
        cmbShifts.Enabled = False
        cmbShifts.Tag = intI
        cmbShifts.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        ShiftsComboboxes.Add(intI, cmbShifts)

        txtRepValue.Text = valrepvalue
        txtRepValue.Height = 30
        txtRepValue.Width = 60
        txtRepValue.Left = HP
        txtRepValue.Top = VP
        'HP = HP + txtRepValue.Width + 1
        txtRepValue.Font = New Font("arial", 8)
        txtRepValue.Tag = intI
        txtRepValue.TabIndex = TabIndex
        TabIndex = TabIndex + 1
        txtRepValue.Visible = False
        txtRepValue.Enabled = False
        RepValueTextBoxes.Add(intI, txtRepValue)

        txtConcreteQty.Text = valconcqty
        txtConcreteQty.Height = 30
        txtConcreteQty.Width = 60
        txtConcreteQty.Left = HP
        txtConcreteQty.Top = VP
        txtConcreteQty.Font = New Font("arial", 8)
        txtConcreteQty.Enabled = False
        txtConcreteQty.Tag = intI
        txtConcreteQty.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        concreteqtyTextboxes.Add(intI, txtConcreteQty)
        If CategoryTextBoxes(intI).Text <> "Concreting" Then
            concreteqtyTextboxes(intI).Visible = False
        Else
            concreteqtyTextboxes(intI).Visible = True
            HP = HP + txtConcreteQty.Width + 1
        End If
        btnAddExtra.Text = "Add"
        btnAddExtra.Height = 20
        btnAddExtra.Width = 40
        btnAddExtra.Left = HP
        btnAddExtra.Top = VP
        btnAddExtra.Font = New Font("arial", 8)
        btnAddExtra.Enabled = False
        btnAddExtra.Tag = intI
        btnAddExtra.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
        ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
        AddHandler btnAddExtra.Click, AddressOf HandleBtnExtraClick
        AddButtons.Add(intI, btnAddExtra)
        If CategoryTextBoxes(intI).Text = "Conveyance" Or _
          CategoryTextBoxes(intI).Text = "Major Others" Then btnAddExtra.Visible = False
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtRepValue)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtConcreteQty)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
        'intI = intI + 1
        VP = VP + 20
        HP = 1
    End Sub
    Private Sub buildMinorControls(ByVal valCategory As String, ByVal valEquipsname As String, ByVal valCapacity As String, ByVal valMake As String, _
          ByVal valModel As String, ByVal valMobDate As String, ByVal valDemobDate As String, ByVal valQty As Integer, ByVal valHPM As Single, _
          ByVal valPurchvalue As Long, ByVal valchkd As Integer, ByVal cnt As Integer, ByVal valIsNew As Boolean)
        Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
        Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
        Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
        Dim txtHrsPerMonth As New TextBox
        Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox
        Dim txtPurchValue As New TextBox   'txtMaintPerc As New TextBox, chkSelected As New CheckBox
        Dim chkSelected As New CheckBox, IsNew As New CheckBox
        Dim btnAddExtra As New Button, chkval As Boolean
        Dim intI As Integer
        intI = cnt

        chkval = False
        chkSelected.Checked = valchkd
        chkSelected.Height = 30
        chkSelected.Width = 24
        chkSelected.Left = HP
        chkSelected.Top = VP - 3
        HP = HP + chkSelected.Width + 1
        chkSelected.Tag = intI
        chkSelected.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler chkSelected.CheckedChanged, AddressOf HandleMinorCheckboxStatus
        Checkboxes.Add(intI, chkSelected)

        txtCategory.Text = valCategory
        txtCategory.WordWrap = True
        txtCategory.Height = 50
        txtCategory.Width = 80
        txtCategory.TextAlign = HorizontalAlignment.Center
        txtCategory.Left = HP
        txtCategory.Top = VP
        HP = HP + txtCategory.Width + 1
        txtCategory.Font = New Font("arial", 8)
        txtCategory.Tag = intI
        txtCategory.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCategory.Enabled = False
        CategoryTextBoxes.Add(intI, txtCategory)

        txtEquipname.Text = valEquipsname
        txtEquipname.WordWrap = True
        txtEquipname.Height = 60
        txtEquipname.Width = 150
        txtEquipname.TextAlign = HorizontalAlignment.Center
        txtEquipname.Left = HP
        txtEquipname.Top = VP
        HP = HP + txtEquipname.Width + 1
        txtEquipname.Font = New Font("arial", 8)
        txtEquipname.Tag = intI
        txtEquipname.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtEquipname.Enabled = False
        EquipNameTextBoxes.Add(intI, txtEquipname)

        txtCapacity.Text = valCapacity
        txtCapacity.WordWrap = True
        txtCapacity.Height = 30
        txtCapacity.Width = 70
        txtCapacity.TextAlign = HorizontalAlignment.Center
        txtCapacity.Left = HP
        txtCapacity.Top = VP
        HP = HP + txtCapacity.Width + 1
        txtCapacity.Font = New Font("arial", 8)
        txtCapacity.Tag = intI
        txtCapacity.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCapacity.Enabled = False
        CapacityTextBoxes.Add(intI, txtCapacity)

        txtMakeModel.Text = valMake & " / " & valModel
        txtMakeModel.WordWrap = True
        txtMakeModel.Height = 50
        txtMakeModel.Width = 160
        txtMakeModel.TextAlign = HorizontalAlignment.Center
        txtMakeModel.Left = HP
        txtMakeModel.Top = VP
        HP = HP + txtMakeModel.Width + 1
        txtMakeModel.Font = New Font("arial", 8)
        txtMakeModel.Tag = intI
        txtMakeModel.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtMakeModel.Enabled = False
        MakeModelTextBoxes.Add(intI, txtMakeModel)

        dpMobDate.Value = valMobDate
        dpMobDate.Name = "MobDate"
        dpMobDate.Format = DateTimePickerFormat.Custom
        dpMobDate.CustomFormat = "dd-MMM-yyyy"
        dpMobDate.Height = 30
        dpMobDate.Width = 100
        dpMobDate.Left = HP
        dpMobDate.Top = VP
        HP = HP + dpMobDate.Width + 1
        dpMobDate.Font = New Font("arial", 8)
        dpMobDate.Enabled = False
        dpMobDate.Tag = intI
        dpMobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
        MobdatePickers.Add(intI, dpMobDate)

        dpDemobDate.Value = valDemobDate
        dpDemobDate.Name = "DemobDate"
        dpDemobDate.Format = DateTimePickerFormat.Custom
        dpDemobDate.CustomFormat = "dd-MMM-yyyy"
        dpDemobDate.Height = 30
        dpDemobDate.Width = 100
        dpDemobDate.Left = HP
        dpDemobDate.Top = VP
        HP = HP + dpDemobDate.Width + 1
        dpDemobDate.Font = New Font("arial", 8)
        dpDemobDate.Enabled = False
        dpDemobDate.Tag = intI
        dpDemobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
        DemobDatePickers.Add(intI, dpDemobDate)

        txtQty.Text = valQty
        txtQty.Height = 30
        txtQty.Width = 30
        txtQty.Left = HP
        txtQty.Top = VP
        HP = HP + txtQty.Width + 1
        txtQty.Font = New Font("arial", 8)
        txtQty.Enabled = False
        txtQty.Tag = intI
        txtQty.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler txtQty.Validated, AddressOf HandleQty
        QtyTextBoxes.Add(intI, txtQty)

        txtHrsPerMonth.Text = valHPM
        txtHrsPerMonth.Height = 30
        txtHrsPerMonth.Width = 60
        txtHrsPerMonth.Left = HP
        txtHrsPerMonth.Top = VP
        HP = HP + txtHrsPerMonth.Width + 5
        txtHrsPerMonth.Font = New Font("arial", 8)
        txtHrsPerMonth.Enabled = False
        txtHrsPerMonth.Tag = intI
        txtHrsPerMonth.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        HrsPermonthTextBoxes.Add(intI, txtHrsPerMonth)

        IsNew.Checked = valIsNew
        IsNew.Height = 30
        IsNew.Width = 40
        IsNew.Left = HP
        IsNew.Top = VP - 3
        HP = HP + IsNew.Width + 1
        IsNew.Tag = intI
        IsNew.TabIndex = TabIndex
        TabIndex = TabIndex + 1
        IsNew.Visible = True
        IsNew.Enabled = False
        AddHandler IsNew.CheckedChanged, AddressOf HandleNewMCStatus
        IsNewMC.Add(intI, IsNew)

        txtPurchValue.Text = valPurchvalue
        txtPurchValue.Height = 30
        txtPurchValue.Width = 80
        txtPurchValue.Left = HP
        txtPurchValue.Top = VP
        HP = HP + txtPurchValue.Width + 5
        txtPurchValue.Font = New Font("arial", 8)
        txtPurchValue.Tag = intI
        txtPurchValue.TabIndex = TabIndex
        TabIndex = TabIndex + 1
        txtPurchValue.Visible = True
        txtPurchValue.Enabled = False
        PurchvalTextBoxes.Add(intI, txtPurchValue)

        cmbDepPerc.Items.Add(2.75)
        cmbDepPerc.Items.Add(1.25)
        cmbDepPerc.Items.Add(0.5)
        cmbDepPerc.SelectedIndex = 0
        cmbDepPerc.Height = 30
        cmbDepPerc.Width = 45
        cmbDepPerc.Left = HP
        cmbDepPerc.Top = VP
        cmbDepPerc.Font = New Font("arial", 8)
        cmbDepPerc.Enabled = False
        cmbDepPerc.Visible = False
        cmbDepPerc.Tag = intI
        cmbDepPerc.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        DepPercComboboxes.Add(intI, cmbDepPerc)

        cmbShifts.Items.Add(1)
        cmbShifts.Items.Add(1.5)
        cmbShifts.Items.Add(2)
        cmbShifts.Items.Add(1)
        cmbShifts.SelectedIndex = 0
        cmbShifts.Height = 30
        cmbShifts.Width = 45
        cmbShifts.Left = HP
        cmbShifts.Top = VP
        cmbShifts.Font = New Font("arial", 8)
        cmbShifts.Enabled = False
        cmbShifts.Visible = False
        cmbShifts.Tag = intI
        cmbShifts.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        ShiftsComboboxes.Add(intI, cmbShifts)

        btnAddExtra.Text = "Add"
        btnAddExtra.Height = 20
        btnAddExtra.Width = 40
        btnAddExtra.Left = HP
        btnAddExtra.Top = VP
        btnAddExtra.Font = New Font("arial", 8)
        btnAddExtra.Enabled = False
        btnAddExtra.Tag = intI
        btnAddExtra.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
        ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
        AddHandler btnAddExtra.Click, AddressOf HandleMinorBtnExtraClick
        AddButtons.Add(intI, btnAddExtra)

        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(IsNew)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtPurchValue)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
        VP = VP + 20
        HP = 1
    End Sub
    Private Sub LoadLightingControlsInPage(ByVal mcategoryname As String)
        Dim intI As Integer, cnt As Integer
        Dim valcategory As String, valEquipsname As String, valCapacity As String, valMake As String, valModel As String
        Dim valMobDate As String, valDemobDate As String, valQty As Integer, valPPU As Single, valConnLoad As Single, valUF As Single
        Dim valchkd As Integer

        Me.tbcBdgetHeads.TabPages(0).Text = "Single Phase Equipments And " & mcategoryname
        HP = Me.Left + 2
        intI = 1
        HP = 1
        VP = 0
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        QtyTextBoxes.Clear()
        PowerPerUnitTextBoxes.Clear()
        ConnectLoadTextBoxes.Clear()
        Checkboxes.Clear()
        UtilityFactorTextBoxes.Clear()
        AddButtons.Clear()

        If lightingEquipsNames.Length > 0 Then
            For cnt = 0 To lightingEquipsNames.Length - 1
                valcategory = "Lighting"
                valEquipsname = lightingEquipsNames(cnt)
                valCapacity = lightingEquipsCapacity(cnt)
                valMake = lightingEquipsMake(cnt)
                valModel = lightingEquipsModel(cnt)
                valMobDate = lightingEquipsMobDate(cnt)
                valDemobDate = lightingEquipsDemobDate(cnt)
                valQty = lightingEquipsQty(cnt)
                valPPU = lightingEquipsPPU(cnt)
                valConnLoad = lightingEquipsCLPerMc(cnt)
                valUF = lightingEquipsUF(cnt)
                valchkd = lightingEquipsChkd(cnt)
                buildLightingcontrols(valcategory, valEquipsname, valCapacity, valMake, valModel, valMobDate, valDemobDate, valQty, valPPU, valConnLoad, valUF, valchkd, cnt)
            Next
        End If

        Dim j As Integer
        For j = 0 To AddButtons.Count - 1
            If Checkboxes(j).Checked Then AddButtons(j).Enabled = True
        Next

        Me.Button1.Enabled = True
        'End If
    End Sub
    Private Sub buildLightingcontrols(ByVal valcategory As String, ByVal valEquipsname As String, ByVal valCapacity As String, ByVal valMake As String, _
        ByVal valModel As String, ByVal valMobDate As String, ByVal valDemobDate As String, ByVal valQty As Integer, ByVal valPPU As Single, _
        ByVal valConnLoad As Single, ByVal valUF As Single, ByVal valchkd As Integer, ByVal cnt As Integer)

        Dim txtEquipname As New TextBox, txtCapacity As New TextBox
        Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
        Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
        Dim txtPowerPerUnit As New TextBox, txtConnectLoad As New TextBox, txtUtilityFactor As New TextBox
        Dim chkSelected As New CheckBox
        Dim btnAddExtra As New Button, chkval As Boolean
        Dim intI As Integer
        intI = cnt

        chkval = False
        chkSelected.Checked = valchkd
        chkSelected.Height = 30
        chkSelected.Width = 24
        chkSelected.Left = HP
        chkSelected.Top = VP - 3
        HP = HP + chkSelected.Width + 1
        chkSelected.Tag = intI
        chkSelected.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler chkSelected.CheckedChanged, AddressOf HandleLightingCheckboxStatus
        Checkboxes.Add(intI, chkSelected)

        txtEquipname.Text = valEquipsname
        txtEquipname.WordWrap = True
        txtEquipname.Height = 60
        txtEquipname.Width = 150
        txtEquipname.TextAlign = HorizontalAlignment.Center
        txtEquipname.Left = HP
        txtEquipname.Top = VP
        HP = HP + txtEquipname.Width + 1
        txtEquipname.Font = New Font("arial", 8)
        txtEquipname.Tag = intI
        txtEquipname.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtEquipname.Enabled = False
        EquipNameTextBoxes.Add(intI, txtEquipname)

        txtCapacity.Text = valCapacity
        txtCapacity.WordWrap = True
        txtCapacity.Height = 30
        txtCapacity.Width = 70
        txtCapacity.TextAlign = HorizontalAlignment.Center
        txtCapacity.Left = HP
        txtCapacity.Top = VP
        HP = HP + txtCapacity.Width + 1
        txtCapacity.Font = New Font("arial", 8)
        txtCapacity.Tag = intI
        txtCapacity.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCapacity.Enabled = False
        CapacityTextBoxes.Add(intI, txtCapacity)

        txtMakeModel.Text = valMake & " / " & valModel
        txtMakeModel.WordWrap = True
        txtMakeModel.Height = 50
        txtMakeModel.Width = 160
        txtMakeModel.TextAlign = HorizontalAlignment.Center
        txtMakeModel.Left = HP
        txtMakeModel.Top = VP
        HP = HP + txtMakeModel.Width + 1
        txtMakeModel.Font = New Font("arial", 8)
        txtMakeModel.Tag = intI
        txtMakeModel.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtMakeModel.Enabled = False
        MakeModelTextBoxes.Add(intI, txtMakeModel)

        dpMobDate.Value = valMobDate  'Today().Date      'mStartDate
        dpMobDate.Name = "MobDate"
        dpMobDate.Format = DateTimePickerFormat.Custom
        dpMobDate.CustomFormat = "dd-MMM-yyyy"
        dpMobDate.Height = 30
        dpMobDate.Width = 100
        dpMobDate.Left = HP
        dpMobDate.Top = VP
        HP = HP + dpMobDate.Width + 1
        dpMobDate.Font = New Font("arial", 8)
        dpMobDate.Enabled = False
        dpMobDate.Tag = intI
        dpMobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
        MobdatePickers.Add(intI, dpMobDate)

        dpDemobDate.Value = valDemobDate 'Today().Date    'menddate
        dpDemobDate.Name = "DemobDate"
        dpDemobDate.Format = DateTimePickerFormat.Custom
        dpDemobDate.CustomFormat = "dd-MMM-yyyy"
        dpDemobDate.Height = 30
        dpDemobDate.Width = 100
        dpDemobDate.Left = HP
        dpDemobDate.Top = VP
        HP = HP + dpDemobDate.Width + 1
        dpDemobDate.Font = New Font("arial", 8)
        dpDemobDate.Enabled = False
        dpDemobDate.Tag = intI
        dpDemobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
        DemobDatePickers.Add(intI, dpDemobDate)

        txtQty.Text = valQty
        txtQty.Height = 30
        txtQty.Width = 30
        txtQty.Left = HP
        txtQty.Top = VP
        HP = HP + txtQty.Width + 1
        txtQty.Font = New Font("arial", 8)
        txtQty.Enabled = False
        txtQty.Tag = intI
        txtQty.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler txtQty.Validated, AddressOf HandleQty
        QtyTextBoxes.Add(intI, txtQty)


        txtPowerPerUnit.Text = valPPU
        txtPowerPerUnit.Height = 30
        txtPowerPerUnit.Width = 60
        txtPowerPerUnit.Left = HP
        txtPowerPerUnit.Top = VP
        HP = HP + txtPowerPerUnit.Width + 1
        txtPowerPerUnit.Font = New Font("arial", 8)
        txtPowerPerUnit.Enabled = False
        txtPowerPerUnit.Tag = intI
        txtPowerPerUnit.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        PowerPerUnitTextBoxes.Add(intI, txtPowerPerUnit)

        txtConnectLoad.Text = valConnLoad
        txtConnectLoad.Height = 30
        txtConnectLoad.Width = 60
        txtConnectLoad.Left = HP
        txtConnectLoad.Top = VP
        HP = HP + txtConnectLoad.Width + 1
        txtConnectLoad.Font = New Font("arial", 8)
        txtConnectLoad.Enabled = False
        txtConnectLoad.Tag = intI
        txtConnectLoad.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        ConnectLoadTextBoxes.Add(intI, txtConnectLoad)

        txtUtilityFactor.Text = valUF
        txtUtilityFactor.Height = 30
        txtUtilityFactor.Width = 45
        txtUtilityFactor.Left = HP
        txtUtilityFactor.Top = VP
        HP = HP + txtUtilityFactor.Width + 1
        txtUtilityFactor.Font = New Font("arial", 8)
        txtUtilityFactor.Enabled = False
        txtUtilityFactor.Tag = intI
        txtUtilityFactor.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        UtilityFactorTextBoxes.Add(intI, txtUtilityFactor)

        btnAddExtra.Text = "Add"
        btnAddExtra.Height = 20
        btnAddExtra.Width = 40
        btnAddExtra.Left = HP
        btnAddExtra.Top = VP
        btnAddExtra.Font = New Font("arial", 8)
        btnAddExtra.Enabled = False
        btnAddExtra.Tag = intI
        btnAddExtra.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
        ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
        AddHandler btnAddExtra.Click, AddressOf HandleLightingBtnExtraClick
        AddButtons.Add(intI, btnAddExtra)

        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtPowerPerUnit)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtConnectLoad)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtUtilityFactor)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
        VP = VP + 20
        HP = 1

    End Sub
    Private Sub LoadBPFExpControlsInpage(ByVal mcategory As String)
        Dim intI As Integer, cnt As Integer
        Me.tbcBdgetHeads.TabPages(0).Text = " BPlant Fixed Expenses"

        HP = Me.Left + 2
        intI = 1
        HP = 1
        VP = 0
        CategoryTextBoxes.Clear()
        QtyTextBoxes.Clear()
        CostTextBoxes.Clear()
        RemarksTextBoxes.Clear()
        ClientBillingTextBoxes.Clear()
        CostPercTextBoxes.Clear()
        AmountTextBoxes.Clear()
        Checkboxes.Clear()
        mTabindex = 0

        If fixedBPExpCategoryNames.Length > 0 Then
            Dim valCategory As String, valQty As Integer, valCost As Single, valClientBill As Long, valAmount As Single, valchkd As Integer
            Dim valRemarks As String

            For cnt = 0 To fixedBPExpCategoryNames.Length - 1
                valCategory = fixedBPExpCategoryNames(cnt)
                valQty = fixedBPExpEquipsQty(cnt)
                valCost = fixedBPExpCost(cnt)
                valAmount = valQty * valCost
                valClientBill = mProjectvalue
                valchkd = fixedBPExpEquipsChkd(cnt)
                valRemarks = fixedBPExpRemarks(cnt)
                buildFixedBPExpControls(valCategory, valQty, valCost, valAmount, valClientBill, valRemarks, valchkd, cnt)

            Next
            Me.Refresh()
            Me.Button1.Enabled = True
        End If
    End Sub
    Private Sub buildFixedBPExpControls(ByVal valCategory As String, ByVal valQty As Integer, ByVal valCost As Single, ByVal valAmount As Single, _
        ByVal valClientBill As Long, ByVal valRemarks As String, ByVal valchkd As Integer, ByVal cnt As Integer)

        Dim txtCategory As New TextBox, txtQty As New TextBox, txtCost As New TextBox
        Dim txtRemarks As New TextBox, chkSelected As New CheckBox
        Dim txtClientBilling As New TextBox, txtCostperc As New TextBox, txtAmount As New TextBox
        Dim intI As Integer, chkval As Boolean

        intI = cnt
        chkval = False
        chkSelected.Checked = valchkd
        chkSelected.Height = 30
        chkSelected.Width = 24
        chkSelected.Left = HP
        chkSelected.Top = VP - 3
        HP = HP + chkSelected.Width + 1
        chkSelected.Tag = intI
        chkSelected.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler chkSelected.CheckedChanged, AddressOf HandleFexpCheckboxStatus
        Checkboxes.Add(intI, chkSelected)

        txtCategory.Text = valCategory
        txtCategory.WordWrap = True
        txtCategory.Height = 50
        txtCategory.Width = 220
        txtCategory.TextAlign = HorizontalAlignment.Center
        txtCategory.Left = HP
        txtCategory.Top = VP
        HP = HP + txtCategory.Width + 10
        txtCategory.Font = New Font("arial", 8)
        txtCategory.Tag = intI
        txtCategory.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCategory.Enabled = False
        CategoryTextBoxes.Add(intI, txtCategory)

        txtQty.Text = valQty
        txtQty.Height = 60
        txtQty.Width = 40
        txtQty.TextAlign = HorizontalAlignment.Center
        txtQty.Left = HP
        txtQty.Top = VP
        HP = HP + txtQty.Width + 5
        txtQty.Font = New Font("arial", 8)
        txtQty.Tag = intI
        txtQty.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtQty.Enabled = False
        AddHandler txtQty.Validated, AddressOf HandleFexpAmtCalc
        QtyTextBoxes.Add(intI, txtQty)

        txtCost.Text = valCost
        txtCost.WordWrap = True
        txtCost.Height = 30
        txtCost.Width = 50
        txtCost.TextAlign = HorizontalAlignment.Center
        txtCost.Left = HP
        txtCost.Top = VP
        HP = HP + txtCost.Width + 1
        txtCost.Font = New Font("arial", 8)
        txtCost.Tag = intI
        txtCost.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCost.Enabled = False
        AddHandler txtCost.Validated, AddressOf HandleFexpAmtCalc
        CostTextBoxes.Add(intI, txtCost)

        txtAmount.Text = valQty * valCost
        txtAmount.WordWrap = True
        txtAmount.Height = 30
        txtAmount.Width = 80
        txtAmount.TextAlign = HorizontalAlignment.Center
        txtAmount.Left = HP
        txtAmount.Top = VP
        HP = HP + txtAmount.Width + 5
        txtAmount.Font = New Font("arial", 8)
        txtAmount.Tag = intI
        txtAmount.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtAmount.Enabled = False
        AmountTextBoxes.Add(intI, txtAmount)

        txtClientBilling.Text = mProjectvalue
        txtClientBilling.WordWrap = True
        txtClientBilling.Height = 30
        txtClientBilling.Width = 80
        txtClientBilling.TextAlign = HorizontalAlignment.Center
        txtClientBilling.Left = HP
        txtClientBilling.Top = VP
        HP = HP + txtClientBilling.Width + 1
        txtClientBilling.Font = New Font("arial", 8)
        txtClientBilling.Tag = intI
        txtClientBilling.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtClientBilling.Enabled = False
        ClientBillingTextBoxes.Add(intI, txtClientBilling)

        txtCostperc.Text = System.Math.Round((valAmount / mProjectvalue) * 100, 3)
        txtCostperc.WordWrap = True
        txtCostperc.Height = 30
        txtCostperc.Width = 80
        txtCostperc.TextAlign = HorizontalAlignment.Center
        txtCostperc.Left = HP
        txtCostperc.Top = VP
        HP = HP + txtCostperc.Width + 1
        txtCostperc.Font = New Font("arial", 8)
        txtCostperc.Tag = intI
        txtCostperc.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCostperc.Enabled = False
        CostPercTextBoxes.Add(intI, txtCostperc)

        txtRemarks.Text = valRemarks
        txtRemarks.WordWrap = True
        txtRemarks.Height = 50
        txtRemarks.Width = 300
        txtRemarks.TextAlign = HorizontalAlignment.Center
        txtRemarks.Left = HP
        txtRemarks.Top = VP
        HP = HP + txtRemarks.Width + 1
        txtRemarks.Font = New Font("arial", 8)
        txtRemarks.Tag = intI
        txtRemarks.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtRemarks.Enabled = False
        RemarksTextBoxes.Add(intI, txtRemarks)

        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCost)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtRemarks)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtAmount)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtClientBilling)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCostperc)
        VP = VP + 20
        HP = 1
    End Sub

    Private Sub LoadFexpControlsInpage(ByVal mcategory As String)
        Dim intI As Integer, cnt As Integer
        Me.tbcBdgetHeads.TabPages(0).Text = "Fixed Expenses"
        HP = Me.Left + 2
        intI = 1
        HP = 1
        VP = 0
        CategoryTextBoxes.Clear()
        QtyTextBoxes.Clear()
        CostTextBoxes.Clear()
        RemarksTextBoxes.Clear()
        ClientBillingTextBoxes.Clear()
        CostPercTextBoxes.Clear()
        AmountTextBoxes.Clear()
        Checkboxes.Clear()
        mTabindex = 0
        If fixedExpCategoryNames.Length > 0 Then
            Dim valCategory As String, valQty As Integer, valCost As Single, valClientBill As Long, valAmount As Single, valchkd As Integer
            Dim valRemarks As String
            For cnt = 0 To fixedExpCategoryNames.Length - 1
                valCategory = fixedExpCategoryNames(cnt)
                valQty = fixedExpEquipsQty(cnt)
                valCost = fixedExpCost(cnt)
                valAmount = valQty * valCost
                valClientBill = mProjectvalue
                valchkd = fixedExpEquipsChkd(cnt)
                valRemarks = fixedExpRemarks(cnt)
                buildFixedExpControls(valCategory, valQty, valCost, valAmount, valClientBill, valRemarks, valchkd, cnt)
            Next

            Me.Refresh()
            Me.Button1.Enabled = True
        End If
    End Sub
    Private Sub buildFixedExpControls(ByVal valCategory As String, ByVal valQty As Integer, ByVal valCost As Single, ByVal valAmount As Single, _
        ByVal valClientBill As Long, ByVal valRemarks As String, ByVal valchkd As Integer, ByVal cnt As Integer)

        Dim txtCategory As New TextBox, txtQty As New TextBox, txtCost As New TextBox
        Dim txtRemarks As New TextBox, chkSelected As New CheckBox
        Dim txtClientBilling As New TextBox, txtCostperc As New TextBox, txtAmount As New TextBox
        Dim intI As Integer, chkval As Integer

        intI = cnt
        chkval = False
        chkSelected.Checked = valchkd
        chkSelected.Height = 30
        chkSelected.Width = 24
        chkSelected.Left = HP
        chkSelected.Top = VP - 3
        HP = HP + chkSelected.Width + 1
        chkSelected.Tag = intI
        chkSelected.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler chkSelected.CheckedChanged, AddressOf HandleFexpCheckboxStatus
        Checkboxes.Add(intI, chkSelected)

        txtCategory.Text = valCategory
        txtCategory.WordWrap = True
        txtCategory.Height = 50
        txtCategory.Width = 220
        txtCategory.TextAlign = HorizontalAlignment.Center
        txtCategory.Left = HP
        txtCategory.Top = VP
        HP = HP + txtCategory.Width + 10
        txtCategory.Font = New Font("arial", 8)
        txtCategory.Tag = intI
        txtCategory.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCategory.Enabled = False
        CategoryTextBoxes.Add(intI, txtCategory)

        txtQty.Text = valQty
        txtQty.Height = 60
        txtQty.Width = 40
        txtQty.TextAlign = HorizontalAlignment.Center
        txtQty.Left = HP
        txtQty.Top = VP
        HP = HP + txtQty.Width + 5
        txtQty.Font = New Font("arial", 8)
        txtQty.Tag = intI
        txtQty.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtQty.Enabled = False
        AddHandler txtQty.Validated, AddressOf HandleFexpAmtCalc
        QtyTextBoxes.Add(intI, txtQty)

        txtCost.Text = valCost
        txtCost.WordWrap = True
        txtCost.Height = 30
        txtCost.Width = 50
        txtCost.TextAlign = HorizontalAlignment.Center
        txtCost.Left = HP
        txtCost.Top = VP
        HP = HP + txtCost.Width + 1
        txtCost.Font = New Font("arial", 8)
        txtCost.Tag = intI
        txtCost.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCost.Enabled = False
        AddHandler txtCost.Validated, AddressOf HandleFexpAmtCalc
        CostTextBoxes.Add(intI, txtCost)

        txtAmount.Text = valQty * valCost
        txtAmount.WordWrap = True
        txtAmount.Height = 30
        txtAmount.Width = 80
        txtAmount.TextAlign = HorizontalAlignment.Center
        txtAmount.Left = HP
        txtAmount.Top = VP
        HP = HP + txtAmount.Width + 5
        txtAmount.Font = New Font("arial", 8)
        txtAmount.Tag = intI
        txtAmount.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtAmount.Enabled = False
        AmountTextBoxes.Add(intI, txtAmount)

        txtClientBilling.Text = mProjectvalue
        txtClientBilling.WordWrap = True
        txtClientBilling.Height = 30
        txtClientBilling.Width = 80
        txtClientBilling.TextAlign = HorizontalAlignment.Center
        txtClientBilling.Left = HP
        txtClientBilling.Top = VP
        HP = HP + txtClientBilling.Width + 1
        txtClientBilling.Font = New Font("arial", 8)
        txtClientBilling.Tag = intI
        txtClientBilling.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtClientBilling.Enabled = False
        ClientBillingTextBoxes.Add(intI, txtClientBilling)

        txtCostperc.Text = System.Math.Round((valAmount / mProjectvalue) * 100, 3)
        txtCostperc.WordWrap = True
        txtCostperc.Height = 30
        txtCostperc.Width = 80
        txtCostperc.TextAlign = HorizontalAlignment.Center
        txtCostperc.Left = HP
        txtCostperc.Top = VP
        HP = HP + txtCostperc.Width + 1
        txtCostperc.Font = New Font("arial", 8)
        txtCostperc.Tag = intI
        txtCostperc.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCostperc.Enabled = False
        CostPercTextBoxes.Add(intI, txtCostperc)

        txtRemarks.Text = valRemarks
        txtRemarks.WordWrap = True
        txtRemarks.Height = 50
        txtRemarks.Width = 300
        txtRemarks.TextAlign = HorizontalAlignment.Center
        txtRemarks.Left = HP
        txtRemarks.Top = VP
        HP = HP + txtRemarks.Width + 1
        txtRemarks.Font = New Font("arial", 8)
        txtRemarks.Tag = intI
        txtRemarks.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtRemarks.Enabled = False
        RemarksTextBoxes.Add(intI, txtRemarks)

        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCost)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtRemarks)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtAmount)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtClientBilling)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCostperc)
        VP = VP + 20
        HP = 1
        Me.Refresh()

    End Sub

    Private Sub LoadMinorControlsInpage()
        Dim intI As Integer, cnt As Integer
        Dim valCategory As String, valEquipsname As String, valCapacity As String, valMake As String, valModel As String
        Dim valMobDate As String, valDemobDate As String, valQty As Integer, valHPM As Single, valPurchvalue As Long, valIsNew As Boolean
        Dim valchkd As Integer
        mcategory = "Minor Equipments"
        Me.tbcBdgetHeads.TabPages(0).Text = "Minor Equipments"
        HP = Me.Left + 1
        intI = 1
        HP = 1
        VP = 0
        CategoryTextBoxes.Clear()
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        QtyTextBoxes.Clear()
        HrsPermonthTextBoxes.Clear()
        PurchvalTextBoxes.Clear()
        Checkboxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        DepPercComboboxes.Clear()
        ShiftsComboboxes.Clear()
        AddButtons.Clear()
        IsNewMC.Clear()

        If minorEquipsNames.Length > 0 Then
            mTabindex = 0
            For cnt = 0 To minorEquipsNames.Length - 1
                valCategory = "Minor Equipments"
                valEquipsname = minorEquipsNames(cnt)
                valCapacity = minorEquipsCapacity(cnt)
                valMake = minorEquipsMake(cnt)
                valModel = minorEquipsModel(cnt)
                valMobDate = minorEquipsMobDate(cnt)
                valDemobDate = minorEquipsDemobDate(cnt)
                valQty = minorEquipsQty(cnt)
                valHPM = minorEquipsHPM(cnt)
                valPurchvalue = minorEquipsNewCost(cnt)
                valchkd = minorEquipsChkd(cnt)
                valIsNew = minorIsNewMC(cnt)
                buildMinorControls(valCategory, valEquipsname, valCapacity, valMake, valModel, _
                     valMobDate, valDemobDate, valQty, valHPM, valPurchvalue, valchkd, cnt, valIsNew)

            Next
        End If

        Dim j As Integer
        For j = 0 To AddButtons.Count - 1
            If Checkboxes(j).Checked Then AddButtons(j).Enabled = True
        Next

        Me.Button1.Enabled = True
    End Sub
    
    Private Sub LoadFexpControlsInpage(ByVal mcategory)
        Dim intI As Integer
        Me.tbcBdgetHeads.TabPages(0).Text = "Fixed Expenses"
        strStatement = "Select * from FixedExpenses"
        If moledbConnection Is Nothing Then
            strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
        mDataSet = New DataSet()
        mOledbDataAdapter.Fill(mDataSet, "FixedExpesnes")
        Dim Machine As DataRow
        HP = Me.Left + 2
        intI = 1
        HP = 1
        VP = 0
        CategoryTextBoxes.Clear()
        QtyTextBoxes.Clear()
        CostTextBoxes.Clear()
        RemarksTextBoxes.Clear()
        ClientBillingTextBoxes.Clear()
        CostPercTextBoxes.Clear()
        AmountTextBoxes.Clear()

        If mDataSet.Tables("FixedExpenses").Rows.Count > 0 Then
            mTabindex = 1
            For Each Machine In mDataSet.Tables("FixedExpenses").Rows
                Dim txtCategory As New TextBox, txtQty As New TextBox, txtCost As New TextBox
                Dim txtRemarks As New TextBox, chkSelected As New CheckBox
                Dim txtClientBilling As New TextBox, txtCostperc As New TextBox, txtAmount As New TextBox

                chkSelected.Checked = False
                chkSelected.Height = 30
                chkSelected.Width = 24
                chkSelected.Left = HP
                chkSelected.Top = VP - 3
                HP = HP + chkSelected.Width + 1
                chkSelected.Tag = intI
                chkSelected.TabIndex = mTabindex
                mTabindex = mTabindex + 1
                AddHandler chkSelected.CheckedChanged, AddressOf HandleFexpCheckboxStatus
                Checkboxes.Add(intI, chkSelected)

                txtCategory.Text = Machine("Category").ToString
                txtCategory.WordWrap = True
                txtCategory.Height = 50
                txtCategory.Width = 80
                txtCategory.TextAlign = HorizontalAlignment.Center
                txtCategory.Left = HP
                txtCategory.Top = VP
                HP = HP + txtCategory.Width + 1
                txtCategory.Font = New Font("arial", 8)
                txtCategory.Tag = intI
                txtCategory.TabIndex = mTabindex
                mTabindex = mTabindex + 1
                txtCategory.Enabled = False
                CategoryTextBoxes.Add(intI, txtCategory)

                txtQty.Text = 1
                txtQty.Height = 60
                txtQty.Width = 150
                txtQty.TextAlign = HorizontalAlignment.Center
                txtQty.Left = HP
                txtQty.Top = VP
                HP = HP + txtQty.Width + 1
                txtQty.Font = New Font("arial", 8)
                txtQty.Tag = intI
                txtQty.TabIndex = mTabindex
                mTabindex = mTabindex + 1
                txtQty.Enabled = False
                AddHandler txtQty.Validated, AddressOf HandleFexpAmtCalc
                QtyTextBoxes.Add(intI, txtQty)

                txtCost.Text = Machine("Cost").ToString
                txtCost.WordWrap = True
                txtCost.Height = 30
                txtCost.Width = 70
                txtCost.TextAlign = HorizontalAlignment.Center
                txtCost.Left = HP
                txtCost.Top = VP
                HP = HP + txtCost.Width + 1
                txtCost.Font = New Font("arial", 8)
                txtCost.Tag = intI
                txtCost.TabIndex = mTabindex
                mTabindex = mTabindex + 1
                txtCost.Enabled = False
                AddHandler txtQty.Validated, AddressOf HandleFexpAmtCalc
                CostTextBoxes.Add(intI, txtCost)

                txtAmount.Text = Val(QtyTextBoxes(intI).Text) * Val(CostTextBoxes(intI).Text)
                txtAmount.WordWrap = True
                txtAmount.Height = 30
                txtAmount.Width = 70
                txtAmount.TextAlign = HorizontalAlignment.Center
                txtAmount.Left = HP
                txtAmount.Top = VP
                HP = HP + txtAmount.Width + 1
                txtAmount.Font = New Font("arial", 8)
                txtAmount.Tag = intI
                txtAmount.TabIndex = mTabindex
                mTabindex = mTabindex + 1
                txtAmount.Enabled = False
                AmountTextBoxes.Add(intI, txtAmount)

                txtClientBilling.Text = mProjectvalue
                txtClientBilling.WordWrap = True
                txtClientBilling.Height = 30
                txtClientBilling.Width = 70
                txtClientBilling.TextAlign = HorizontalAlignment.Center
                txtClientBilling.Left = HP
                txtClientBilling.Top = VP
                HP = HP + txtClientBilling.Width + 1
                txtClientBilling.Font = New Font("arial", 8)
                txtClientBilling.Tag = intI
                txtClientBilling.TabIndex = mTabindex
                mTabindex = mTabindex + 1
                txtClientBilling.Enabled = False
                ClientBillingTextBoxes.Add(intI, txtClientBilling)

                txtCostperc.Text = mProjectvalue / System.Math.Round(Val(AmountTextBoxes(intI).Text) * 100, 3)
                txtCostperc.WordWrap = True
                txtCostperc.Height = 30
                txtCostperc.Width = 70
                txtCostperc.TextAlign = HorizontalAlignment.Center
                txtCostperc.Left = HP
                txtCostperc.Top = VP
                HP = HP + txtCostperc.Width + 1
                txtCostperc.Font = New Font("arial", 8)
                txtCostperc.Tag = intI
                txtCostperc.TabIndex = mTabindex
                mTabindex = mTabindex + 1
                txtCostperc.Enabled = False
                CostPercTextBoxes.Add(intI, txtCostperc)

                txtRemarks.Text = Machine("Remarks").ToString
                txtRemarks.WordWrap = True
                txtRemarks.Height = 50
                txtRemarks.Width = 160
                txtRemarks.TextAlign = HorizontalAlignment.Center
                txtRemarks.Left = HP
                txtRemarks.Top = VP
                HP = HP + txtRemarks.Width + 1
                txtRemarks.Font = New Font("arial", 8)
                txtRemarks.Tag = intI
                txtRemarks.TabIndex = mTabindex
                mTabindex = mTabindex + 1
                txtRemarks.Enabled = False
                RemarksTextBoxes.Add(intI, txtRemarks)

                Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
                Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
                Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCost)
                Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtRemarks)
                Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
                Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtAmount)
                Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtClientBilling)
                Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCostperc)
                VP = VP + 20
                HP = 1
                '====================================================================================
            Next
            Me.Refresh()
            moledbCommand = Nothing
            mDataSet = Nothing
            Me.Button1.Enabled = True
        End If
    End Sub
    Private Sub buildhiredcontrols(ByVal valCategory As String, ByVal valEquipsname As String, ByVal valCapacity As String, _
    ByVal valMake As String, ByVal valModel As String, ByVal valMobDate As String, ByVal valDemobDate As String, _
    ByVal valQty As Integer, ByVal valHPM As Single, ByVal valHirevalue As Long, ByVal valchkd As Integer, ByVal cnt As Integer)

        Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
        Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
        Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
        Dim txtHrsPerMonth As New TextBox
        Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox, txtHireVal As New TextBox
        Dim chkSelected As New CheckBox   ', IsNew As New CheckBox
        Dim btnAddExtra As New Button, chkval As Integer

        intI = cnt
        chkval = False

        chkSelected.Checked = valchkd
        chkSelected.Height = 30
        chkSelected.Width = 24
        chkSelected.Left = HP
        chkSelected.Top = VP - 3
        HP = HP + chkSelected.Width + 1
        chkSelected.Tag = intI
        chkSelected.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler chkSelected.CheckedChanged, AddressOf HandleHireCheckboxStatus
        Checkboxes.Add(intI, chkSelected)

        txtCategory.Text = valCategory
        txtCategory.WordWrap = True
        txtCategory.Height = 50
        txtCategory.Width = 80
        txtCategory.TextAlign = HorizontalAlignment.Center
        txtCategory.Left = HP
        txtCategory.Top = VP
        HP = HP + txtCategory.Width + 1
        txtCategory.Font = New Font("arial", 8)
        txtCategory.Tag = intI
        txtCategory.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCategory.Enabled = False
        CategoryTextBoxes.Add(intI, txtCategory)

        txtEquipname.Text = valEquipsname
        txtEquipname.WordWrap = True
        txtEquipname.Height = 60
        txtEquipname.Width = 150
        txtEquipname.TextAlign = HorizontalAlignment.Center
        txtEquipname.Left = HP
        txtEquipname.Top = VP
        HP = HP + txtEquipname.Width + 1
        txtEquipname.Font = New Font("arial", 8)
        txtEquipname.Tag = intI
        txtEquipname.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtEquipname.Enabled = False
        EquipNameTextBoxes.Add(intI, txtEquipname)

        txtCapacity.Text = valCapacity
        txtCapacity.WordWrap = True
        txtCapacity.Height = 30
        txtCapacity.Width = 70
        txtCapacity.TextAlign = HorizontalAlignment.Center
        txtCapacity.Left = HP
        txtCapacity.Top = VP
        HP = HP + txtCapacity.Width + 1
        txtCapacity.Font = New Font("arial", 8)
        txtCapacity.Tag = intI
        txtCapacity.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCapacity.Enabled = False
        CapacityTextBoxes.Add(intI, txtCapacity)

        txtMakeModel.Text = valMake & " / " & valModel
        txtMakeModel.WordWrap = True
        txtMakeModel.Height = 50
        txtMakeModel.Width = 160
        txtMakeModel.TextAlign = HorizontalAlignment.Center
        txtMakeModel.Left = HP
        txtMakeModel.Top = VP
        HP = HP + txtMakeModel.Width + 1
        txtMakeModel.Font = New Font("arial", 8)
        txtMakeModel.Tag = intI
        txtMakeModel.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtMakeModel.Enabled = False
        MakeModelTextBoxes.Add(intI, txtMakeModel)

        dpMobDate.Value = valMobDate
        dpMobDate.Name = "MobDate"
        dpMobDate.Format = DateTimePickerFormat.Custom
        dpMobDate.CustomFormat = "dd-MMM-yyyy"
        dpMobDate.Height = 30
        dpMobDate.Width = 100
        dpMobDate.Left = HP
        dpMobDate.Top = VP
        HP = HP + dpMobDate.Width + 1
        dpMobDate.Font = New Font("arial", 8)
        dpMobDate.Enabled = False
        dpMobDate.Tag = intI
        dpMobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
        MobdatePickers.Add(intI, dpMobDate)


        dpDemobDate.Value = valDemobDate
        dpDemobDate.Name = "DemobDate"
        dpDemobDate.Format = DateTimePickerFormat.Custom
        dpDemobDate.CustomFormat = "dd-MMM-yyyy"
        dpDemobDate.Height = 30
        dpDemobDate.Width = 100
        dpDemobDate.Left = HP
        dpDemobDate.Top = VP
        HP = HP + dpDemobDate.Width + 1
        dpDemobDate.Font = New Font("arial", 8)
        dpDemobDate.Enabled = False
        dpDemobDate.Tag = intI
        dpDemobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
        DemobDatePickers.Add(intI, dpDemobDate)

        txtQty.Text = valQty
        txtQty.Height = 30
        txtQty.Width = 30
        txtQty.Left = HP
        txtQty.Top = VP
        HP = HP + txtQty.Width + 1
        txtQty.Font = New Font("arial", 8)
        txtQty.Enabled = False
        txtQty.Tag = intI
        txtQty.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler txtQty.Validated, AddressOf HandleQty
        QtyTextBoxes.Add(intI, txtQty)

        txtHrsPerMonth.Text = valHPM
        txtHrsPerMonth.Height = 30
        txtHrsPerMonth.Width = 60
        txtHrsPerMonth.Left = HP
        txtHrsPerMonth.Top = VP
        HP = HP + txtHrsPerMonth.Width + 5
        txtHrsPerMonth.Font = New Font("arial", 8)
        txtHrsPerMonth.Enabled = False
        txtHrsPerMonth.Tag = intI
        txtHrsPerMonth.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        HrsPermonthTextBoxes.Add(intI, txtHrsPerMonth)

        txtHireVal.Text = valHirevalue
        txtHireVal.Height = 30
        txtHireVal.Width = 80
        txtHireVal.Left = HP
        txtHireVal.Top = VP
        HP = HP + txtHireVal.Width + 5
        txtHireVal.Font = New Font("arial", 8)
        txtHireVal.Tag = intI
        txtHireVal.TabIndex = TabIndex
        TabIndex = TabIndex + 1
        txtHireVal.Visible = True
        txtHireVal.Enabled = False
        HireChargesTextBoxes.Add(intI, txtHireVal)

        cmbDepPerc.Items.Add(2.75)
        cmbDepPerc.Items.Add(1.25)
        cmbDepPerc.Items.Add(0.5)
        cmbDepPerc.SelectedIndex = 0
        cmbDepPerc.Height = 30
        cmbDepPerc.Width = 45
        cmbDepPerc.Left = HP
        cmbDepPerc.Top = VP
        cmbDepPerc.Font = New Font("arial", 8)
        cmbDepPerc.Visible = False
        cmbDepPerc.Enabled = False
        cmbDepPerc.Tag = intI
        cmbDepPerc.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        DepPercComboboxes.Add(intI, cmbDepPerc)

        cmbShifts.Items.Add(1)
        cmbShifts.Items.Add(1.5)
        cmbShifts.Items.Add(2)
        cmbShifts.Items.Add(1)
        cmbShifts.Items.Add(0)
        cmbShifts.SelectedIndex = 0
        cmbShifts.Height = 30
        cmbShifts.Width = 45
        cmbShifts.Left = HP
        cmbShifts.Top = VP
        cmbShifts.Font = New Font("arial", 8)
        cmbShifts.Enabled = False
        cmbShifts.Visible = False
        cmbShifts.Tag = intI
        cmbShifts.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        ShiftsComboboxes.Add(intI, cmbShifts)

        btnAddExtra.Text = "Add"
        btnAddExtra.Height = 20
        btnAddExtra.Width = 40
        btnAddExtra.Left = HP
        btnAddExtra.Top = VP
        btnAddExtra.Font = New Font("arial", 8)
        btnAddExtra.Enabled = False
        btnAddExtra.Tag = intI
        btnAddExtra.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
        ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
        AddHandler btnAddExtra.Click, AddressOf HandleHireBtnExtraClick
        AddButtons.Add(intI, btnAddExtra)

        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHireVal)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
        VP = VP + 20
        HP = 1
    End Sub

    Private Sub LoadHiredControlsInpage()
        Dim intI As Integer, cnt As Integer
        Dim valCategory As String, valEquipsname As String, valCapacity As String, valMake As String, valModel As String
        Dim valMobDate As String, valDemobDate As String, valQty As Integer, valHPM As Single, valHirevalue As Long
        Dim valchkd As Integer

        mcategory = "Hired Equipments"
        Me.tbcBdgetHeads.TabPages(0).Text = "Hired Equipments"
        HP = Me.Left + 2
        intI = 1
        HP = 1
        VP = 0

        CategoryTextBoxes.Clear()
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        QtyTextBoxes.Clear()
        HrsPermonthTextBoxes.Clear()
        HireChargesTextBoxes.Clear()
        Checkboxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        DepPercComboboxes.Clear()
        ShiftsComboboxes.Clear()
        AddButtons.Clear()
        If hiredEquipsNames.Length > 0 Then
            mTabindex = 0
            For cnt = 0 To hiredEquipsNames.Length - 1
                valCategory = hiredCategoryNames(cnt)
                valEquipsname = hiredEquipsNames(cnt)
                valCapacity = hiredEquipsCapacity(cnt)
                valMake = hiredEquipsMake(cnt)
                valModel = hiredEquipsModel(cnt)
                valMobDate = hiredEquipsMobDate(cnt)
                valDemobDate = hiredEquipsDemobDate(cnt)
                valQty = hiredEquipsQty(cnt)
                valHPM = hiredEquipsHPM(cnt)
                valHirevalue = hiredEquipsHireCharges(cnt)
                valchkd = hiredEquipsChkd(cnt)
                buildhiredcontrols(valCategory, valEquipsname, valCapacity, _
                    valMake, valModel, valMobDate, valDemobDate, valQty, valHPM, valHirevalue, valchkd, cnt)
            Next

            Dim j As Integer
            For j = 0 To AddButtons.Count - 1
                If Checkboxes(j).Checked Then AddButtons(j).Enabled = True
            Next
        End If
    End Sub
    Private Sub HandleCheckboxStatus(ByVal sender As Object, ByVal e As EventArgs)
        Dim chk As CheckBox = sender
        changeEnabled(chk.Checked, chk.Tag)
    End Sub
    Private Sub HandleHireCheckboxStatus(ByVal sender As Object, ByVal e As EventArgs)
        Dim chk As CheckBox = sender
        HireChangeEnabled(chk.Checked, chk.Tag)
    End Sub
    Private Sub HandleFexpCheckboxStatus(ByVal sender As Object, ByVal e As EventArgs)
        Dim chk As CheckBox = sender
        FexpChangeEnabled(chk.Checked, chk.Tag)
    End Sub
    Private Sub HandleFexpAmtCalc(ByVal sender As Object, ByVal e As EventArgs)
        Dim txtbox As TextBox = sender
        AmountTextBoxes(txtbox.Tag).Text = Val(QtyTextBoxes(txtbox.Tag).Text) * Val(CostTextBoxes(txtbox.Tag).Text)
        CostPercTextBoxes(txtbox.Tag).Text = System.Math.Round(Val(AmountTextBoxes(txtbox.Tag).Text) / mProjectvalue * 100, 3)
    End Sub
    Private Sub HandleMinorCheckboxStatus(ByVal sender As Object, ByVal e As EventArgs)
        Dim chk As CheckBox = sender
        MinorChangeEnabled(chk.Checked, chk.Tag)
    End Sub
    Private Sub HandleLightingCheckboxStatus(ByVal sender As Object, ByVal e As EventArgs)
        Dim chk As CheckBox = sender
        LightingChangeEnabled(chk.Checked, chk.Tag)
    End Sub
    Private Sub HandleNewMCStatus(ByVal sender As Object, ByVal e As EventArgs)
        Dim chk As CheckBox = sender
        Dim tag = chk.Tag
        PurchvalTextBoxes(tag).Enabled = IsNewMC(tag).Checked
    End Sub
    Private Sub HandleQty(ByVal sender As Object, ByVal e As EventArgs)
        Dim intTag As Integer
        Dim QtyText As TextBox = sender
        intTag = QtyText.Tag
        If Checkboxes(intTag).Checked Then
            If (Val(QtyTextBoxes(intTag).Text) = 0 Or Len(Trim(QtyTextBoxes(intTag).Text)) = 0) Then
                MsgBox("Error in Quantity entry  for " & vbNewLine & CategoryTextBoxes(intTag).Text & "," & EquipNameTextBoxes(intTag).Text & "," & _
                   CapacityTextBoxes(intTag).Text & "," & MakeModelTextBoxes(intTag).Text & " is missing", MsgBoxStyle.Critical, "Error in Qty Entry")
                QtyTextBoxes(intTag).Text = 1
                QtyTextBoxes(intTag).Focus()
            End If
        End If
    End Sub
    Private Sub LoadOneLightingControlSet(ByVal newtag As Integer, ByVal chkstatus As Boolean)
        Dim txtEquipname As New TextBox, txtCapacity As New TextBox
        Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
        Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
        Dim txtPowerPerUnit As New TextBox
        Dim txtConnectLoad As New TextBox, txtUtilityFactor As New TextBox
        Dim chkSelected As New CheckBox
        Dim btnAddExtra As New Button

        VP = EquipNameTextBoxes(newtag - 1).Top
        VP = VP + 20
        HP = 1
        mTabindex = 1

        chkSelected.Checked = chkstatus
        chkSelected.Height = 30
        chkSelected.Width = 24
        chkSelected.Left = HP
        chkSelected.Top = VP - 3
        HP = HP + chkSelected.Width + 1
        chkSelected.Tag = newtag
        AddHandler chkSelected.CheckedChanged, AddressOf HandleLightingCheckboxStatus
        Checkboxes.Add(newtag, chkSelected)

        txtEquipname.WordWrap = True
        txtEquipname.Height = 60
        txtEquipname.Width = 150
        txtEquipname.TextAlign = HorizontalAlignment.Center
        txtEquipname.Left = HP
        txtEquipname.Top = VP
        HP = HP + txtEquipname.Width + 1
        txtEquipname.Font = New Font("arial", 8)
        txtEquipname.Tag = newtag
        txtEquipname.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtEquipname.Enabled = False
        EquipNameTextBoxes.Add(newtag, txtEquipname)

        txtCapacity.WordWrap = True
        txtCapacity.Height = 30
        txtCapacity.Width = 70
        txtCapacity.TextAlign = HorizontalAlignment.Center
        txtCapacity.Left = HP
        txtCapacity.Top = VP
        HP = HP + txtCapacity.Width + 1
        txtCapacity.Font = New Font("arial", 8)
        txtCapacity.Tag = newtag
        txtCapacity.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCapacity.Enabled = False
        CapacityTextBoxes.Add(newtag, txtCapacity)

        txtMakeModel.WordWrap = True
        txtMakeModel.Height = 50
        txtMakeModel.Width = 160
        txtMakeModel.TextAlign = HorizontalAlignment.Center
        txtMakeModel.Left = HP
        txtMakeModel.Top = VP
        HP = HP + txtMakeModel.Width + 1
        txtMakeModel.Font = New Font("arial", 8)
        txtMakeModel.Tag = newtag
        txtMakeModel.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtMakeModel.Enabled = False
        MakeModelTextBoxes.Add(newtag, txtMakeModel)

        dpMobDate.Name = "MobDate"
        dpMobDate.Format = DateTimePickerFormat.Custom
        dpMobDate.CustomFormat = "dd-MMM-yyyy"
        dpMobDate.Height = 30
        dpMobDate.Width = 100
        dpMobDate.Left = HP
        dpMobDate.Top = VP
        HP = HP + dpMobDate.Width + 1
        dpMobDate.Font = New Font("arial", 8)
        dpMobDate.Enabled = False
        dpMobDate.Tag = newtag
        dpMobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
        MobdatePickers.Add(newtag, dpMobDate)

        dpDemobDate.Name = "DemobDate"
        dpDemobDate.Format = DateTimePickerFormat.Custom
        dpDemobDate.CustomFormat = "dd-MMM-yyyy"
        dpDemobDate.Height = 30
        dpDemobDate.Width = 100
        dpDemobDate.Left = HP
        dpDemobDate.Top = VP
        HP = HP + dpDemobDate.Width + 1
        dpDemobDate.Font = New Font("arial", 8)
        dpDemobDate.Enabled = False
        dpDemobDate.Tag = newtag
        dpDemobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
        DemobDatePickers.Add(newtag, dpDemobDate)

        txtQty.Height = 30
        txtQty.Width = 30
        txtQty.Left = HP
        txtQty.Top = VP
        HP = HP + txtQty.Width + 1
        txtQty.Font = New Font("arial", 8)
        txtQty.Enabled = False
        txtQty.Tag = newtag
        txtQty.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler txtQty.Validated, AddressOf HandleQty
        QtyTextBoxes.Add(newtag, txtQty)


        txtPowerPerUnit.Height = 30
        txtPowerPerUnit.Width = 60
        txtPowerPerUnit.Left = HP
        txtPowerPerUnit.Top = VP
        HP = HP + txtPowerPerUnit.Width + 1
        txtPowerPerUnit.Font = New Font("arial", 8)
        txtPowerPerUnit.Enabled = False
        txtPowerPerUnit.Tag = newtag
        txtPowerPerUnit.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        PowerPerUnitTextBoxes.Add(newtag, txtPowerPerUnit)

        txtConnectLoad.Height = 30
        txtConnectLoad.Width = 60
        txtConnectLoad.Left = HP
        txtConnectLoad.Top = VP
        HP = HP + txtConnectLoad.Width + 1
        txtConnectLoad.Font = New Font("arial", 8)
        txtConnectLoad.Tag = newtag
        txtConnectLoad.TabIndex = TabIndex
        TabIndex = TabIndex + 1
        txtConnectLoad.Visible = True
        txtConnectLoad.Enabled = False
        ConnectLoadTextBoxes.Add(newtag, txtConnectLoad)

        txtUtilityFactor.Height = 30
        txtUtilityFactor.Width = 45
        txtUtilityFactor.Left = HP
        txtUtilityFactor.Top = VP
        HP = HP + txtUtilityFactor.Width + 1
        txtUtilityFactor.Font = New Font("arial", 8)
        txtUtilityFactor.Tag = newtag
        txtUtilityFactor.TabIndex = TabIndex
        TabIndex = TabIndex + 1
        txtUtilityFactor.Visible = True
        txtUtilityFactor.Enabled = False
        UtilityFactorTextBoxes.Add(newtag, txtUtilityFactor)


        btnAddExtra.Text = "Add"
        btnAddExtra.Height = 20
        btnAddExtra.Width = 40
        btnAddExtra.Left = HP
        btnAddExtra.Top = VP
        btnAddExtra.Font = New Font("arial", 8)
        btnAddExtra.Enabled = False
        btnAddExtra.Tag = newtag
        btnAddExtra.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
        ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
        AddHandler btnAddExtra.Click, AddressOf HandleLightingBtnExtraClick
        AddButtons.Add(newtag, btnAddExtra)

        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtPowerPerUnit)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtConnectLoad)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtUtilityFactor)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
        Checkboxes(newtag).Checked = chkstatus
        Me.Refresh()
        Checkboxes(newtag).Checked = chkstatus
        VP = VP + 20
        HP = 1
        If chkSelected.Checked Then LightingChangeEnabled(True, newtag)
    End Sub
    Private Sub LoadOneMinorControlSet(ByVal newtag As Integer, ByVal chkstatus As Boolean)
        Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
        Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
        Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
        Dim txtHrsPerMonth As New TextBox
        Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox, txtPurchVal As New TextBox
        Dim chkSelected As New CheckBox, IsNew As New CheckBox
        Dim btnAddExtra As New Button

        VP = CategoryTextBoxes(newtag - 1).Top
        VP = VP + 20
        HP = 1
        mTabindex = 1

        chkSelected.Checked = chkstatus
        chkSelected.Height = 30
        chkSelected.Width = 24
        chkSelected.Left = HP
        chkSelected.Top = VP - 3
        HP = HP + chkSelected.Width + 1
        chkSelected.Tag = newtag
        AddHandler chkSelected.CheckedChanged, AddressOf HandleMinorCheckboxStatus
        Checkboxes.Add(newtag, chkSelected)

        txtCategory.WordWrap = True
        txtCategory.Height = 50
        txtCategory.Width = 80
        txtCategory.TextAlign = HorizontalAlignment.Center
        txtCategory.Left = HP
        txtCategory.Top = VP
        HP = HP + txtCategory.Width + 1
        txtCategory.Font = New Font("arial", 8)
        txtCategory.Tag = newtag
        txtCategory.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCategory.Enabled = False
        CategoryTextBoxes.Add(newtag, txtCategory)

        txtEquipname.WordWrap = True
        txtEquipname.Height = 60
        txtEquipname.Width = 150
        txtEquipname.TextAlign = HorizontalAlignment.Center
        txtEquipname.Left = HP
        txtEquipname.Top = VP
        HP = HP + txtEquipname.Width + 1
        txtEquipname.Font = New Font("arial", 8)
        txtEquipname.Tag = newtag
        txtEquipname.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtEquipname.Enabled = False
        EquipNameTextBoxes.Add(newtag, txtEquipname)

        txtCapacity.WordWrap = True
        txtCapacity.Height = 30
        txtCapacity.Width = 70
        txtCapacity.TextAlign = HorizontalAlignment.Center
        txtCapacity.Left = HP
        txtCapacity.Top = VP
        HP = HP + txtCapacity.Width + 1
        txtCapacity.Font = New Font("arial", 8)
        txtCapacity.Tag = newtag
        txtCapacity.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCapacity.Enabled = False
        CapacityTextBoxes.Add(newtag, txtCapacity)

        txtMakeModel.WordWrap = True
        txtMakeModel.Height = 50
        txtMakeModel.Width = 160
        txtMakeModel.TextAlign = HorizontalAlignment.Center
        txtMakeModel.Left = HP
        txtMakeModel.Top = VP
        HP = HP + txtMakeModel.Width + 1
        txtMakeModel.Font = New Font("arial", 8)
        txtMakeModel.Tag = newtag
        txtMakeModel.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtMakeModel.Enabled = False
        MakeModelTextBoxes.Add(newtag, txtMakeModel)

        dpMobDate.Name = "MobDate"
        dpMobDate.Format = DateTimePickerFormat.Custom
        dpMobDate.CustomFormat = "dd-MMM-yyyy"
        dpMobDate.Height = 30
        dpMobDate.Width = 100
        dpMobDate.Left = HP
        dpMobDate.Top = VP
        HP = HP + dpMobDate.Width + 1
        dpMobDate.Font = New Font("arial", 8)
        dpMobDate.Enabled = False
        dpMobDate.Tag = newtag
        dpMobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
        MobdatePickers.Add(newtag, dpMobDate)

        dpDemobDate.Name = "DemobDate"
        dpDemobDate.Format = DateTimePickerFormat.Custom
        dpDemobDate.CustomFormat = "dd-MMM-yyyy"
        dpDemobDate.Height = 30
        dpDemobDate.Width = 100
        dpDemobDate.Left = HP
        dpDemobDate.Top = VP
        HP = HP + dpDemobDate.Width + 1
        dpDemobDate.Font = New Font("arial", 8)
        dpDemobDate.Enabled = False
        dpDemobDate.Tag = newtag
        dpDemobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
        DemobDatePickers.Add(newtag, dpDemobDate)

        txtQty.Height = 30
        txtQty.Width = 30
        txtQty.Left = HP
        txtQty.Top = VP
        HP = HP + txtQty.Width + 1
        txtQty.Font = New Font("arial", 8)
        txtQty.Enabled = False
        txtQty.Tag = newtag
        txtQty.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler txtQty.Validated, AddressOf HandleQty
        QtyTextBoxes.Add(newtag, txtQty)


        txtHrsPerMonth.Height = 30
        txtHrsPerMonth.Width = 60
        txtHrsPerMonth.Left = HP
        txtHrsPerMonth.Top = VP
        HP = HP + txtHrsPerMonth.Width + 5
        txtHrsPerMonth.Font = New Font("arial", 8)
        txtHrsPerMonth.Enabled = False
        txtHrsPerMonth.Tag = newtag
        txtHrsPerMonth.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        HrsPermonthTextBoxes.Add(newtag, txtHrsPerMonth)

        IsNew.Checked = False
        IsNew.Height = 30
        IsNew.Width = 40
        IsNew.Left = HP
        IsNew.Top = VP - 3
        HP = HP + IsNew.Width + 1
        IsNew.Tag = newtag
        IsNew.TabIndex = TabIndex
        TabIndex = TabIndex + 1
        IsNew.Visible = True
        IsNew.Enabled = False
        AddHandler IsNew.CheckedChanged, AddressOf HandleNewMCStatus
        IsNewMC.Add(newtag, IsNew)

        'txtPurchVal.Text = MinorNewPurchVal(newtag - 1)
        txtPurchVal.Height = 30
        txtPurchVal.Width = 80
        txtPurchVal.Left = HP
        txtPurchVal.Top = VP
        HP = HP + txtPurchVal.Width + 1
        txtPurchVal.Font = New Font("arial", 8)
        txtPurchVal.Tag = newtag
        txtPurchVal.TabIndex = TabIndex
        TabIndex = TabIndex + 1
        txtPurchVal.Visible = True
        txtPurchVal.Enabled = False
        PurchvalTextBoxes.Add(newtag, txtPurchVal)

        cmbDepPerc.Items.Add(2.75)
        cmbDepPerc.Items.Add(1.25)
        cmbDepPerc.Items.Add(0.5)
        cmbDepPerc.Height = 30
        cmbDepPerc.Width = 45
        cmbDepPerc.Left = HP
        cmbDepPerc.Top = VP
        cmbDepPerc.Font = New Font("arial", 8)
        cmbDepPerc.Visible = False
        cmbDepPerc.Enabled = False
        cmbDepPerc.Tag = newtag
        cmbDepPerc.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        DepPercComboboxes.Add(newtag, cmbDepPerc)

        cmbShifts.Items.Add(1)
        cmbShifts.Items.Add(1.5)
        cmbShifts.Items.Add(2)
        cmbShifts.Items.Add(1)
        cmbShifts.Items.Add(0)
        cmbShifts.Text = ShiftsComboboxes(newtag - 1).Text
        cmbShifts.Height = 30
        cmbShifts.Width = 45
        cmbShifts.Left = HP
        cmbShifts.Top = VP
        cmbShifts.Font = New Font("arial", 8)
        cmbShifts.Visible = False
        cmbShifts.Enabled = False
        cmbShifts.Tag = newtag
        cmbShifts.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        ShiftsComboboxes.Add(newtag, cmbShifts)

        btnAddExtra.Text = "Add"
        btnAddExtra.Height = 20
        btnAddExtra.Width = 40
        btnAddExtra.Left = HP
        btnAddExtra.Top = VP
        btnAddExtra.Font = New Font("arial", 8)
        btnAddExtra.Enabled = False
        btnAddExtra.Tag = newtag
        btnAddExtra.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
        ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
        AddHandler btnAddExtra.Click, AddressOf HandleMinorBtnExtraClick
        AddButtons.Add(newtag, btnAddExtra)

        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(IsNew)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtPurchVal)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
        Checkboxes(newtag).Checked = chkstatus
        Me.Refresh()
        Checkboxes(newtag).Checked = chkstatus
        VP = VP + 20
        HP = 1
        If chkSelected.Checked Then MinorChangeEnabled(True, newtag)
    End Sub
    Private Sub LoadOneHireControlSet(ByVal newtag As Integer, ByVal chkstatus As Boolean)
        Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
        Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
        Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
        Dim txtHrsPerMonth As New TextBox
        Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox, txtHireVal As New TextBox
        Dim chkSelected As New CheckBox   ', IsNew As New CheckBox
        Dim btnAddExtra As New Button

        VP = CategoryTextBoxes(newtag - 1).Top
        VP = VP + 20
        HP = 1
        mTabindex = 1

        chkSelected.Checked = chkstatus
        chkSelected.Height = 30
        chkSelected.Width = 24
        chkSelected.Left = HP
        chkSelected.Top = VP - 3
        HP = HP + chkSelected.Width + 1
        chkSelected.Tag = newtag
        AddHandler chkSelected.CheckedChanged, AddressOf HandleHireCheckboxStatus
        Checkboxes.Add(newtag, chkSelected)

        txtCategory.WordWrap = True
        txtCategory.Height = 50
        txtCategory.Width = 80
        txtCategory.TextAlign = HorizontalAlignment.Center
        txtCategory.Left = HP
        txtCategory.Top = VP
        HP = HP + txtCategory.Width + 1
        txtCategory.Font = New Font("arial", 8)
        txtCategory.Tag = newtag
        txtCategory.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCategory.Enabled = False
        CategoryTextBoxes.Add(newtag, txtCategory)

        txtEquipname.WordWrap = True
        txtEquipname.Height = 60
        txtEquipname.Width = 150
        txtEquipname.TextAlign = HorizontalAlignment.Center
        txtEquipname.Left = HP
        txtEquipname.Top = VP
        HP = HP + txtEquipname.Width + 1
        txtEquipname.Font = New Font("arial", 8)
        txtEquipname.Tag = newtag
        txtEquipname.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtEquipname.Enabled = False
        EquipNameTextBoxes.Add(newtag, txtEquipname)

        txtCapacity.WordWrap = True
        txtCapacity.Height = 30
        txtCapacity.Width = 70
        txtCapacity.TextAlign = HorizontalAlignment.Center
        txtCapacity.Left = HP
        txtCapacity.Top = VP
        HP = HP + txtCapacity.Width + 1
        txtCapacity.Font = New Font("arial", 8)
        txtCapacity.Tag = newtag
        txtCapacity.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCapacity.Enabled = False
        CapacityTextBoxes.Add(newtag, txtCapacity)

        txtMakeModel.WordWrap = True
        txtMakeModel.Height = 50
        txtMakeModel.Width = 160
        txtMakeModel.TextAlign = HorizontalAlignment.Center
        txtMakeModel.Left = HP
        txtMakeModel.Top = VP
        HP = HP + txtMakeModel.Width + 1
        txtMakeModel.Font = New Font("arial", 8)
        txtMakeModel.Tag = newtag
        txtMakeModel.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtMakeModel.Enabled = False
        MakeModelTextBoxes.Add(newtag, txtMakeModel)

        dpMobDate.Name = "MobDate"
        dpMobDate.Format = DateTimePickerFormat.Custom
        dpMobDate.CustomFormat = "dd-MMM-yyyy"
        dpMobDate.Height = 30
        dpMobDate.Width = 100
        dpMobDate.Left = HP
        dpMobDate.Top = VP
        HP = HP + dpMobDate.Width + 1
        dpMobDate.Font = New Font("arial", 8)
        dpMobDate.Enabled = False
        dpMobDate.Tag = newtag
        dpMobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
        MobdatePickers.Add(newtag, dpMobDate)

        dpDemobDate.Name = "DemobDate"
        dpDemobDate.Format = DateTimePickerFormat.Custom
        dpDemobDate.CustomFormat = "dd-MMM-yyyy"
        dpDemobDate.Height = 30
        dpDemobDate.Width = 100
        dpDemobDate.Left = HP
        dpDemobDate.Top = VP
        HP = HP + dpDemobDate.Width + 1
        dpDemobDate.Font = New Font("arial", 8)
        dpDemobDate.Enabled = False
        dpDemobDate.Tag = newtag
        dpDemobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
        DemobDatePickers.Add(newtag, dpDemobDate)

        txtQty.Height = 30
        txtQty.Width = 30
        txtQty.Left = HP
        txtQty.Top = VP
        HP = HP + txtQty.Width + 1
        txtQty.Font = New Font("arial", 8)
        txtQty.Enabled = False
        txtQty.Tag = newtag
        txtQty.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler txtQty.Validated, AddressOf HandleQty
        QtyTextBoxes.Add(newtag, txtQty)


        txtHrsPerMonth.Height = 30
        txtHrsPerMonth.Width = 60
        txtHrsPerMonth.Left = HP
        txtHrsPerMonth.Top = VP
        HP = HP + txtHrsPerMonth.Width + 5
        txtHrsPerMonth.Font = New Font("arial", 8)
        txtHrsPerMonth.Enabled = False
        txtHrsPerMonth.Tag = newtag
        txtHrsPerMonth.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        HrsPermonthTextBoxes.Add(newtag, txtHrsPerMonth)


        txtHireVal.Height = 30
        txtHireVal.Width = 80
        txtHireVal.Left = HP
        txtHireVal.Top = VP
        HP = HP + txtHireVal.Width + 1
        txtHireVal.Font = New Font("arial", 8)
        txtHireVal.Tag = newtag
        txtHireVal.TabIndex = TabIndex
        TabIndex = TabIndex + 1
        txtHireVal.Visible = True
        txtHireVal.Enabled = False
        HireChargesTextBoxes.Add(newtag, txtHireVal)

        cmbDepPerc.Items.Add(2.75)
        cmbDepPerc.Items.Add(1.25)
        cmbDepPerc.Items.Add(0.5)
        cmbDepPerc.Height = 30
        cmbDepPerc.Width = 45
        cmbDepPerc.Left = HP
        cmbDepPerc.Top = VP
        cmbDepPerc.Font = New Font("arial", 8)
        cmbDepPerc.Visible = False
        cmbDepPerc.Enabled = False
        cmbDepPerc.Tag = newtag
        cmbDepPerc.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        DepPercComboboxes.Add(newtag, cmbDepPerc)

        cmbShifts.Items.Add(1)
        cmbShifts.Items.Add(1.5)
        cmbShifts.Items.Add(2)
        cmbShifts.Items.Add(1)
        cmbShifts.Items.Add(0)
        cmbShifts.Height = 30
        cmbShifts.Width = 45
        cmbShifts.Left = HP
        cmbShifts.Top = VP
        cmbShifts.Font = New Font("arial", 8)
        cmbShifts.Visible = False
        cmbShifts.Enabled = False
        cmbShifts.Tag = newtag
        cmbShifts.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        ShiftsComboboxes.Add(newtag, cmbShifts)

        btnAddExtra.Text = "Add"
        btnAddExtra.Height = 20
        btnAddExtra.Width = 40
        btnAddExtra.Left = HP
        btnAddExtra.Top = VP
        btnAddExtra.Font = New Font("arial", 8)
        btnAddExtra.Enabled = False
        btnAddExtra.Tag = newtag
        btnAddExtra.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
        ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
        AddHandler btnAddExtra.Click, AddressOf HandleHireBtnExtraClick
        AddButtons.Add(newtag, btnAddExtra)

        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHireVal)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
        Checkboxes(newtag).Checked = chkstatus
        VP = VP + 20
        HP = 1
        If chkSelected.Checked Then changeEnabled(True, newtag)
    End Sub
    Private Sub HandleHireBtnExtraClick(ByVal sender As Object, ByVal e As EventArgs)
        Dim K As Integer, cnts As Integer, J As Integer
        Dim btn As Button = sender
        K = btn.Tag
        cnts = CategoryTextBoxes.Count
        Dim category = CategoryTextBoxes(K).Text
        Dim chkstatus As Boolean = Checkboxes(K).Checked
        LoadOneHireControlSet(cnts, chkstatus)

        For J = cnts To K + 1 Step -1
            CategoryTextBoxes(J).Text = CategoryTextBoxes(J - 1).Text
            CategoryTextBoxes(J).Tag = J
            EquipNameTextBoxes(J).Text = EquipNameTextBoxes(J - 1).Text
            EquipNameTextBoxes(J).Tag = J
            CapacityTextBoxes(J).Text = CapacityTextBoxes(J - 1).Text
            CapacityTextBoxes(J).Tag = J
            MakeModelTextBoxes(J).Text = MakeModelTextBoxes(J - 1).Text
            MakeModelTextBoxes(J).Tag = J
            MobdatePickers(J).Value = MobdatePickers(J - 1).Value.Date
            MobdatePickers(J).Tag = J
            DemobDatePickers(J).Value = DemobDatePickers(J - 1).Value.Date
            DemobDatePickers(J).Tag = J
            QtyTextBoxes(J).Text = QtyTextBoxes(J - 1).Text
            QtyTextBoxes(J).Tag = J
            HrsPermonthTextBoxes(J).Text = HrsPermonthTextBoxes(J - 1).Text
            HrsPermonthTextBoxes(J).Tag = J
            HireChargesTextBoxes(J).Text = HireChargesTextBoxes(J - 1).Text
            HireChargesTextBoxes(J).Tag = J
            DepPercComboboxes(J).Text = DepPercComboboxes(J - 1).Text
            DepPercComboboxes(J).Tag = J
            ShiftsComboboxes(J).Text = ShiftsComboboxes(J - 1).Text
            ShiftsComboboxes(J).Tag = J
            Checkboxes(J).Checked = Checkboxes(J - 1).Checked
            Checkboxes(J).Tag = J
        Next
        CategoryTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        EquipNameTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        CapacityTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        MakeModelTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        MobdatePickers(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        DemobDatePickers(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        QtyTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        HrsPermonthTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        HireChargesTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        DepPercComboboxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        ShiftsComboboxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        If (Checkboxes(K + 1).Checked And Checkboxes(K).Checked) Then
            If (CategoryTextBoxes(K + 1).Text = CategoryTextBoxes(K).Text And _
                EquipNameTextBoxes(K + 1).Text = EquipNameTextBoxes(K).Text And _
                CapacityTextBoxes(K + 1).Text = CapacityTextBoxes(K).Text And _
                MakeModelTextBoxes(K + 1).Text = MakeModelTextBoxes(K).Text And _
                MobdatePickers(K + 1).Value.Date = MobdatePickers(K).Value.Date And _
                DemobDatePickers(K + 1).Value.Date = DemobDatePickers(K).Value.Date And _
                DepPercComboboxes(K + 1).Text = DepPercComboboxes(K).Text) Then
                MsgBox("Items " & K + 1 & " and " & K + 2 & " are duplicate entries. Please Check and correct")
            End If
        End If
        For J = 0 To cnts
            AddButtons(J).Enabled = Checkboxes(J).Checked
        Next
        Me.Refresh()
    End Sub
    Private Sub HandleLightingBtnExtraClick(ByVal sender As Object, ByVal e As EventArgs)
        Dim K As Integer, cnts As Integer, J As Integer
        Dim btn As Button = sender
        K = btn.Tag
        cnts = EquipNameTextBoxes.Count
        Dim category As String = "Lighting"
        Dim chkstatus As Boolean = Checkboxes(K).Checked
        LoadOneLightingControlSet(cnts, chkstatus)

        For J = cnts To K + 1 Step -1
            EquipNameTextBoxes(J).Text = EquipNameTextBoxes(J - 1).Text
            EquipNameTextBoxes(J).Tag = J
            CapacityTextBoxes(J).Text = CapacityTextBoxes(J - 1).Text
            CapacityTextBoxes(J).Tag = J
            MakeModelTextBoxes(J).Text = MakeModelTextBoxes(J - 1).Text
            MakeModelTextBoxes(J).Tag = J
            MobdatePickers(J).Value = MobdatePickers(J - 1).Value.Date
            MobdatePickers(J).Tag = J
            DemobDatePickers(J).Value = DemobDatePickers(J - 1).Value.Date
            DemobDatePickers(J).Tag = J
            QtyTextBoxes(J).Text = QtyTextBoxes(J - 1).Text
            QtyTextBoxes(J).Tag = J
            PowerPerUnitTextBoxes(J).Text = PowerPerUnitTextBoxes(J - 1).Text
            PowerPerUnitTextBoxes(J).Tag = J
            ConnectLoadTextBoxes(J).Text = ConnectLoadTextBoxes(J - 1).Text
            ConnectLoadTextBoxes(J).Tag = J
            UtilityFactorTextBoxes(J).Text = UtilityFactorTextBoxes(J - 1).Text
            UtilityFactorTextBoxes(J).Tag = J
            Checkboxes(J).Checked = Checkboxes(J - 1).Checked
            Checkboxes(J).Tag = J
        Next
        EquipNameTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        CapacityTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        MakeModelTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        MobdatePickers(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        DemobDatePickers(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        QtyTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        PowerPerUnitTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        ConnectLoadTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        UtilityFactorTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        If (Checkboxes(K + 1).Checked And Checkboxes(K).Checked) Then
            If EquipNameTextBoxes(K + 1).Text = EquipNameTextBoxes(K).Text And _
                CapacityTextBoxes(K + 1).Text = CapacityTextBoxes(K).Text And _
                MakeModelTextBoxes(K + 1).Text = MakeModelTextBoxes(K).Text And _
                MobdatePickers(K + 1).Value.Date = MobdatePickers(K).Value.Date And _
                DemobDatePickers(K + 1).Value.Date = DemobDatePickers(K).Value.Date Then
                MsgBox("Items " & K + 1 & " and " & K + 2 & " are duplicate entries. Please Check and correct")
                'TestValidity = False
            End If
        End If


        For J = 0 To cnts
            AddButtons(J).Enabled = Checkboxes(J).Checked
        Next
        Me.Refresh()
    End Sub
    Private Sub HandleMinorBtnExtraClick(ByVal sender As Object, ByVal e As EventArgs)
        Dim K As Integer, cnts As Integer, J As Integer, mcat As String
        Dim btn As Button = sender
        K = btn.Tag
        cnts = CategoryTextBoxes.Count
        Dim category = CategoryTextBoxes(K).Text
        Dim chkstatus As Boolean = Checkboxes(K).Checked
        LoadOneMinorControlSet(cnts, chkstatus)

        For J = cnts To K + 1 Step -1
            CategoryTextBoxes(J).Text = CategoryTextBoxes(J - 1).Text
            CategoryTextBoxes(J).Tag = J
            EquipNameTextBoxes(J).Text = EquipNameTextBoxes(J - 1).Text
            EquipNameTextBoxes(J).Tag = J
            CapacityTextBoxes(J).Text = CapacityTextBoxes(J - 1).Text
            CapacityTextBoxes(J).Tag = J
            MakeModelTextBoxes(J).Text = MakeModelTextBoxes(J - 1).Text
            MakeModelTextBoxes(J).Tag = J
            MobdatePickers(J).Value = MobdatePickers(J - 1).Value.Date
            MobdatePickers(J).Tag = J
            DemobDatePickers(J).Value = DemobDatePickers(J - 1).Value.Date
            DemobDatePickers(J).Tag = J
            QtyTextBoxes(J).Text = QtyTextBoxes(J - 1).Text
            QtyTextBoxes(J).Tag = J
            HrsPermonthTextBoxes(J).Text = HrsPermonthTextBoxes(J - 1).Text
            HrsPermonthTextBoxes(J).Tag = J
            PurchvalTextBoxes(J).Text = PurchvalTextBoxes(J - 1).Text
            PurchvalTextBoxes(J).Tag = J
            DepPercComboboxes(J).Text = DepPercComboboxes(J - 1).Text
            DepPercComboboxes(J).Tag = J
            ShiftsComboboxes(J).Text = ShiftsComboboxes(J - 1).Text
            ShiftsComboboxes(J).Tag = J
            Checkboxes(J).Checked = Checkboxes(J - 1).Checked
            Checkboxes(J).Tag = J
            IsNewMC(J).Checked = IsNewMC(J - 1).Checked
            IsNewMC(J).Tag = J
        Next
        CategoryTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        EquipNameTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        CapacityTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        MakeModelTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        MobdatePickers(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        DemobDatePickers(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        QtyTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        HrsPermonthTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        PurchvalTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        DepPercComboboxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        ShiftsComboboxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        If (Checkboxes(K + 1).Checked And Checkboxes(K).Checked) Then
            If (CategoryTextBoxes(K + 1).Text = CategoryTextBoxes(K).Text And _
                EquipNameTextBoxes(K + 1).Text = EquipNameTextBoxes(K).Text And _
                CapacityTextBoxes(K + 1).Text = CapacityTextBoxes(K).Text And _
                MakeModelTextBoxes(K + 1).Text = MakeModelTextBoxes(K).Text And _
                MobdatePickers(K + 1).Value.Date = MobdatePickers(K).Value.Date And _
                DemobDatePickers(K + 1).Value.Date = DemobDatePickers(K).Value.Date And _
                DepPercComboboxes(K + 1).Text = DepPercComboboxes(K).Text) Then
                MsgBox("Items " & K + 1 & " and " & K + 2 & " are duplicate entries. Please Check and correct")
            End If
        End If

        For J = 0 To cnts
            AddButtons(J).Enabled = Checkboxes(J).Checked
        Next
        Me.Refresh()
    End Sub
    Private Sub HandleBtnExtraClick(ByVal sender As Object, ByVal e As EventArgs)
        Dim K As Integer, cnts As Integer, J As Integer, mcat As String
        Dim btn As Button = sender
        K = btn.Tag
        cnts = CategoryTextBoxes.Count
        mcat = mcategory
        Dim category = EquipNameTextBoxes(K).Text
        Dim chkstatus As Boolean = Checkboxes(K).Checked

        LoadOneControlSet(cnts, chkstatus)

        For J = cnts To K + 1 Step -1
            CategoryTextBoxes(J).Text = CategoryTextBoxes(J - 1).Text
            CategoryTextBoxes(J).Tag = J
            EquipNameTextBoxes(J).Text = EquipNameTextBoxes(J - 1).Text
            EquipNameTextBoxes(J).Tag = J
            CapacityTextBoxes(J).Text = CapacityTextBoxes(J - 1).Text
            CapacityTextBoxes(J).Tag = J
            MakeModelTextBoxes(J).Text = MakeModelTextBoxes(J - 1).Text
            MakeModelTextBoxes(J).Tag = J
            MobdatePickers(J).Value = MobdatePickers(J - 1).Value.Date
            MobdatePickers(J).Tag = J
            DemobDatePickers(J).Value = DemobDatePickers(J - 1).Value.Date
            DemobDatePickers(J).Tag = J
            QtyTextBoxes(J).Text = QtyTextBoxes(J - 1).Text
            QtyTextBoxes(J).Tag = J
            HrsPermonthTextBoxes(J).Text = HrsPermonthTextBoxes(J - 1).Text
            HrsPermonthTextBoxes(J).Tag = J
            DepPercComboboxes(J).Text = DepPercComboboxes(J - 1).Text
            DepPercComboboxes(J).Tag = J
            ShiftsComboboxes(J).Text = ShiftsComboboxes(J - 1).Text
            ShiftsComboboxes(J).Tag = J
            Checkboxes(J).Checked = Checkboxes(J - 1).Checked
            Checkboxes(J).Tag = J
            concreteqtyTextboxes(J).Text = concreteqtyTextboxes(J - 1).Text
            concreteqtyTextboxes(J).Tag = J
            AddButtons(J).Tag = J
        Next
        CategoryTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        EquipNameTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        CapacityTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        MakeModelTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        MobdatePickers(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        DemobDatePickers(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        QtyTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        HrsPermonthTextBoxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        DepPercComboboxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        ShiftsComboboxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        concreteqtyTextboxes(K + 1).BackColor = Color.FromArgb(245, 224, 192)
        If (Checkboxes(K + 1).Checked And Checkboxes(K).Checked) Then
            If (CategoryTextBoxes(K + 1).Text = CategoryTextBoxes(K).Text And _
                EquipNameTextBoxes(K + 1).Text = EquipNameTextBoxes(K).Text And _
                CapacityTextBoxes(K + 1).Text = CapacityTextBoxes(K).Text And _
                MakeModelTextBoxes(K + 1).Text = MakeModelTextBoxes(K).Text And _
                MobdatePickers(K + 1).Value.Date = MobdatePickers(K).Value.Date And _
                DemobDatePickers(K + 1).Value.Date = DemobDatePickers(K).Value.Date And _
                DepPercComboboxes(K + 1).Text = DepPercComboboxes(K).Text) Then
                MsgBox("Items " & K + 1 & " and " & K + 2 & " are duplicate entries. Please Check and correct")
                'TestValidity = False
            End If
        End If

        Select Case mcat
            Case "Concreting"
                concreteitems = CategoryTextBoxes.Count
            Case "Conveyance"
                ConveyanceItems = CategoryTextBoxes.Count
            Case "Cranes"
                CraneItems = CategoryTextBoxes.Count
            Case "DG Sets"
                DGSetItems = CategoryTextBoxes.Count
            Case "Material Handling"
                MHItems = CategoryTextBoxes.Count
            Case "Non Concreting"
                NCItems = CategoryTextBoxes.Count
            Case "Major Others"
                MajorOtherItems = CategoryTextBoxes.Count
            Case "Minor Equipments"
                MinorItems = CategoryTextBoxes.Count
            Case "Hiredequipments"
                HireItems = CategoryTextBoxes.Count
            Case "Tr crane related exp"
                fexpItems = CategoryTextBoxes.Count
            Case "FixedExp - BP"
                BPFExpItems = CategoryTextBoxes.Count
            Case "Lighting"
                LightingItems = EquipNameTextBoxes.Count
        End Select
        For J = 0 To cnts
            AddButtons(J).Enabled = Checkboxes(J).Checked
        Next
        Me.Refresh()
    End Sub
    Private Sub LoadOneControlSet(ByVal newTag As Integer, ByVal chkstatus As Boolean)
        Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
        Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
        Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
        Dim txtHrsPerMonth As New TextBox
        Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox
        Dim txtRepValue As New TextBox, txtMaintPerc As New TextBox, chkSelected As New CheckBox
        Dim txtConcreteQty As New TextBox, btnAddExtra As New Button


        VP = CategoryTextBoxes(newTag - 1).Top
        VP = VP + 20
        HP = 1
        mTabindex = 1

        chkSelected.Checked = chkstatus
        chkSelected.Height = 30
        chkSelected.Width = 24
        chkSelected.Left = HP
        chkSelected.Top = VP - 3
        HP = HP + chkSelected.Width + 1
        chkSelected.Tag = newTag
        AddHandler chkSelected.CheckedChanged, AddressOf HandleCheckboxStatus
        Checkboxes.Add(newTag, chkSelected)

        txtCategory.WordWrap = True
        txtCategory.Height = 50
        txtCategory.Width = 65
        txtCategory.TextAlign = HorizontalAlignment.Center
        txtCategory.Left = HP
        txtCategory.Top = VP
        HP = HP + txtCategory.Width + 1
        txtCategory.Font = New Font("arial", 8)
        txtCategory.Tag = newTag
        txtCategory.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCategory.Enabled = False
        CategoryTextBoxes.Add(newTag, txtCategory)

        txtEquipname.WordWrap = True
        txtEquipname.Height = 60
        txtEquipname.Width = 150
        txtEquipname.TextAlign = HorizontalAlignment.Center
        txtEquipname.Left = HP
        txtEquipname.Top = VP
        HP = HP + txtEquipname.Width + 1
        txtEquipname.Font = New Font("arial", 8)
        txtEquipname.Tag = newTag
        txtEquipname.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtEquipname.Enabled = False
        EquipNameTextBoxes.Add(newTag, txtEquipname)

        txtCapacity.WordWrap = True
        txtCapacity.Height = 30
        txtCapacity.Width = 70
        txtCapacity.TextAlign = HorizontalAlignment.Center
        txtCapacity.Left = HP
        txtCapacity.Top = VP
        HP = HP + txtCapacity.Width + 1
        txtCapacity.Font = New Font("arial", 8)
        txtCapacity.Tag = newTag
        txtCapacity.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtCapacity.Enabled = False
        CapacityTextBoxes.Add(newTag, txtCapacity)

        txtMakeModel.WordWrap = True
        txtMakeModel.Height = 50
        txtMakeModel.Width = 160
        txtMakeModel.TextAlign = HorizontalAlignment.Center
        txtMakeModel.Left = HP
        txtMakeModel.Top = VP
        HP = HP + txtMakeModel.Width + 1
        txtMakeModel.Font = New Font("arial", 8)
        txtMakeModel.Tag = newTag
        txtMakeModel.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        txtMakeModel.Enabled = False
        MakeModelTextBoxes.Add(newTag, txtMakeModel)

        dpMobDate.Name = "MobDate"
        dpMobDate.Format = DateTimePickerFormat.Custom
        dpMobDate.CustomFormat = "dd-MMM-yyyy"
        dpMobDate.Height = 30
        dpMobDate.Width = 100
        dpMobDate.Left = HP
        dpMobDate.Top = VP
        HP = HP + dpMobDate.Width + 1
        dpMobDate.Font = New Font("arial", 8)
        dpMobDate.Enabled = False
        dpMobDate.Tag = newTag
        dpMobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
        MobdatePickers.Add(newTag, dpMobDate)


        dpDemobDate.Name = "DemobDate"
        dpDemobDate.Format = DateTimePickerFormat.Custom
        dpDemobDate.CustomFormat = "dd-MMM-yyyy"
        dpDemobDate.Height = 30
        dpDemobDate.Width = 100
        dpDemobDate.Left = HP
        dpDemobDate.Top = VP
        HP = HP + dpDemobDate.Width + 1
        dpDemobDate.Font = New Font("arial", 8)
        dpDemobDate.Enabled = False
        dpDemobDate.Tag = newTag
        dpDemobDate.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
        DemobDatePickers.Add(newTag, dpDemobDate)

        txtQty.Height = 30
        txtQty.Width = 30
        txtQty.Left = HP
        txtQty.Top = VP
        HP = HP + txtQty.Width + 1
        txtQty.Font = New Font("arial", 8)
        txtQty.Enabled = False
        txtQty.Tag = newTag + 1
        txtQty.TabIndex = mTabindex
        mTabindex = mTabindex
        AddHandler txtQty.Validated, AddressOf HandleQty
        QtyTextBoxes.Add(newTag, txtQty)


        txtHrsPerMonth.Height = 30
        txtHrsPerMonth.Width = 60
        txtHrsPerMonth.Left = HP
        txtHrsPerMonth.Top = VP
        HP = HP + txtHrsPerMonth.Width + 1
        txtHrsPerMonth.Font = New Font("arial", 8)
        txtHrsPerMonth.Enabled = False
        txtHrsPerMonth.Tag = newTag
        txtHrsPerMonth.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        HrsPermonthTextBoxes.Add(newTag, txtHrsPerMonth)

        cmbDepPerc.Items.Add(2.75)
        cmbDepPerc.Items.Add(1.25)
        cmbDepPerc.Items.Add(0.5)
        cmbDepPerc.SelectedIndex = 0
        cmbDepPerc.Height = 30
        cmbDepPerc.Width = 60
        cmbDepPerc.Left = HP
        cmbDepPerc.Top = VP
        HP = HP + cmbDepPerc.Width + 1
        cmbDepPerc.Font = New Font("arial", 8)
        cmbDepPerc.Enabled = False
        cmbDepPerc.TabIndex = mTabindex
        cmbDepPerc.Tag = newTag
        mTabindex = mTabindex + 1
        DepPercComboboxes.Add(newTag, cmbDepPerc)

        cmbShifts.Items.Add(1)
        cmbShifts.Items.Add(1.5)
        cmbShifts.Items.Add(2)
        cmbShifts.Items.Add(1)
        cmbShifts.Items.Add(0)
        cmbShifts.SelectedIndex = 0
        cmbShifts.Height = 30
        cmbShifts.Width = 45
        cmbShifts.Left = HP
        cmbShifts.Top = VP
        HP = HP + cmbShifts.Width + 1
        cmbShifts.Font = New Font("arial", 8)
        cmbShifts.Enabled = False
        cmbShifts.Tag = newTag
        cmbShifts.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        ShiftsComboboxes.Add(newTag, cmbShifts)

        txtConcreteQty.Text = mConcreteQty
        txtConcreteQty.Height = 30
        txtConcreteQty.Width = 60
        txtConcreteQty.Left = HP
        txtConcreteQty.Top = VP
        'HP = HP + txtConcreteQty.Width + 1
        txtConcreteQty.Font = New Font("arial", 8)
        txtConcreteQty.Enabled = False
        txtConcreteQty.Tag = newTag
        concreteqtyTextboxes.Add(newTag, txtConcreteQty)
        If CategoryTextBoxes(newTag - 1).Text <> "Concreting" Then
            concreteqtyTextboxes(newTag).Visible = False
        Else
            concreteqtyTextboxes(newTag).Visible = True
            HP = HP + txtConcreteQty.Width + 1
        End If

        btnAddExtra.Text = "Add"
        btnAddExtra.Height = 20
        btnAddExtra.Width = 40
        btnAddExtra.Left = HP
        btnAddExtra.Top = VP
        btnAddExtra.Font = New Font("arial", 8)
        btnAddExtra.Enabled = False
        btnAddExtra.Tag = newTag
        btnAddExtra.TabIndex = mTabindex
        mTabindex = mTabindex + 1
        Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
        ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
        AddHandler btnAddExtra.Click, AddressOf HandleBtnExtraClick
        AddButtons.Add(newTag, btnAddExtra)

        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtRepValue)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtConcreteQty)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
        Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
        Checkboxes(newTag).Checked = chkstatus
        VP = VP + 20
        HP = 1
        If chkSelected.Checked Then changeEnabled(True, newTag)

    End Sub
    Private Sub HandleCheckDates(ByVal sender As Object, ByVal e As EventArgs)
        Dim dpdate As DateTimePicker = sender, msgstr As String
        Dim intTag As Integer = dpdate.Tag
        If dpdate.Name = "MobDate" Then
            If Checkboxes(intTag).Checked Then
                If (MobdatePickers(intTag).Value.Date > DemobDatePickers(intTag).Value.Date) Or _
                    (MobdatePickers(intTag).Value.Date < mStartDate) Then
                    msgstr = "Error in Mobilisation Date for " & vbNewLine & CategoryTextBoxes(intTag).Text & "," & EquipNameTextBoxes(intTag).Text & "," & _
                       CapacityTextBoxes(intTag).Text & "," & MakeModelTextBoxes(intTag).Text & vbNewLine & vbNewLine & _
                    "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                      "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                    MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Dataentry")
                    dpdate.Value = mStartDate.Date
                    dpdate.Focus()
                    Exit Sub
                End If
            End If
        ElseIf dpdate.Name = "DemobDate" Then
            If Checkboxes(intTag).Checked Then
                If (DemobDatePickers(intTag).Value.Date < MobdatePickers(intTag).Value.Date) Or _
                    (DemobDatePickers(intTag).Value.Date > mEndDate) Then
                    msgstr = "Error in Demobilisation Date for " & vbNewLine & CategoryTextBoxes(intTag).Text & "," & EquipNameTextBoxes(intTag).Text & "," & _
                       CapacityTextBoxes(intTag).Text & "," & MakeModelTextBoxes(intTag).Text & vbNewLine & vbNewLine & _
                    "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                      "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                    MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Dataentry")
                    DemobDatePickers(intTag).Value = mStartDate.Date
                    DemobDatePickers(intTag).Focus()
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim intI As Integer
        Dim ch As CheckBox

        If Me.optLightingEquips.Checked Then
            cntItems = EquipNameTextBoxes.Count
        Else
            cntItems = CategoryTextBoxes.Count
        End If

        For intI = 0 To cntItems - 1
            ch = Checkboxes(intI)
            ch.Checked = True
            changeEnabled(True, intI)
        Next
        Button1.Enabled = True
    End Sub

    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim intI As Integer
        Dim ch As CheckBox
        If Me.optLightingEquips.Checked Then
            cntItems = EquipNameTextBoxes.Count
        Else
            cntItems = CategoryTextBoxes.Count
        End If
        For intI = 0 To cntItems - 1
            ch = Checkboxes(intI)
            ch.Checked = False
            changeEnabled(False, intI)
        Next
        Button1.Enabled = True

    End Sub

    Private Sub changeEnabled(ByVal status As Boolean, ByVal key As Integer)
        MobdatePickers(key).Enabled = status
        DemobDatePickers(key).Enabled = status
        QtyTextBoxes(key).Enabled = status
        HrsPermonthTextBoxes(key).Enabled = status
        DepPercComboboxes(key).Enabled = status
        ShiftsComboboxes(key).Enabled = status
        AddButtons(key).Enabled = status
        Me.Button1.Enabled = True

    End Sub
    Private Sub MinorChangeEnabled(ByVal status As Boolean, ByVal Key As Integer)
        MobdatePickers(Key).Enabled = status
        DemobDatePickers(Key).Enabled = status
        QtyTextBoxes(Key).Enabled = status
        HrsPermonthTextBoxes(Key).Enabled = status
        IsNewMC(Key).Enabled = status
        AddButtons(Key).Enabled = status
        Me.Button1.Enabled = True
    End Sub
    Private Sub LightingChangeEnabled(ByVal status As Boolean, ByVal Key As Integer)
        MobdatePickers(Key).Enabled = status
        DemobDatePickers(Key).Enabled = status
        QtyTextBoxes(Key).Enabled = status
        AddButtons(Key).Enabled = status
        Me.Button1.Enabled = True
    End Sub
    Private Sub HireChangeEnabled(ByVal status As Boolean, ByVal key As Integer)
        MobdatePickers(key).Enabled = status
        DemobDatePickers(key).Enabled = status
        QtyTextBoxes(key).Enabled = status
        HrsPermonthTextBoxes(key).Enabled = status
        HireChargesTextBoxes(key).Enabled = status
        AddButtons(key).Enabled = status
        Me.Button1.Enabled = True
    End Sub
    Private Sub FexpChangeEnabled(ByVal status As Boolean, ByVal key As Integer)
        QtyTextBoxes(key).Enabled = status
        CostTextBoxes(key).Enabled = status
        RemarksTextBoxes(key).Enabled = status
        Me.Button1.Enabled = True
    End Sub
    Private Sub SaveMinorEquipments()
        Dim intI As Integer ', mcategory As String
        Dim Counts As Integer = 0, RepeatFactor As Integer  'intj As Integer,
        Currenttab = Me.tbcBdgetHeads.SelectedIndex
        If Len(Trim(xlFilename)) = 0 Then
            MsgBox("Filename not specified to save. cannot Save data. Start the application again.")
            Application.Exit()
        End If

        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With

        Me.btnQuit.Enabled = False
        Me.btnClose.Enabled = False
        Dim cnt As Integer = 0
        mcategory = "Minor Equipments"

        MinorCheckedItems = 0
        MinorItems = minorEquipsNames.Length

        For intI = 0 To MinorItems - 1
            If minorEquipsChkd(intI) = 1 Then
                MinorCheckedItems = MinorCheckedItems + 1
            End If
        Next
        DeleteRecordsFromAddedItemsTable(mcategory)
        xlWorksheet = xlWorkbook.Sheets("Minor Eqpts")
        xlWorksheet.Activate()
        RepeatFactor = MinorItems
        RecordsInserted(getSheetNo(xlWorksheet.Name)) = 0
        SaveDataInMinorEquipsSheet("Minor Eqpts", cnt, RepeatFactor)
        lblmessage.Visible = False
        Button1.Enabled = False
    End Sub
    Private Sub SaveLightingBudget(ByVal msheetname As String, ByVal items As Integer)
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, mtablename As String = ""
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet
        Dim txtmake As String, txtmodel As String, intk As Integer, Exists As Integer = 0
        For Each ws In xlWorkbook.Worksheets
            If UCase(ws.Name) = UCase(msheetname) Then
                'index = ws.Index
                Exists = 1
                Exit For
            End If
        Next
        If Not Exists = 1 Then
            CopyPowerReqTemplate()
        End If
        Dim intI As Integer, intJ As Integer = 0
        lblmessage.Text = "Lighting accessories details  Being Saved. Please wait..."
        lblmessage.Visible = True
        Me.Refresh()

        xlWorksheet = xlWorkbook.Sheets(msheetname)
        xlWorksheet.Activate()
        DeleteRecordsFromAddedItemsTable("Lighting")
        mtablename = GetTablename("Lighting")
        getCategoryShortname(xlWorksheet)
        For intI = 0 To items - 1
            InsertCommand = ""
            txtmake = lightingEquipsMake(intI)    ' & , Strings.InStr(MakeModelTextBoxes(intI).Text, " /") - 1))
            txtmodel = lightingEquipsModel(intI)  ', InStr(MakeModelTextBoxes(intI).Text, "/ ") + 1))

            InsertCommand = "INSERT INTO " & mtablename & " Values (" & "'Lighting',"
            InsertCommand = InsertCommand & "'" & lightingEquipsNames(intI) & "', "
            InsertCommand = InsertCommand & "'" & lightingEquipsCapacity(intI) & "', "
            InsertCommand = InsertCommand & "'" & txtmake & "', "
            InsertCommand = InsertCommand & "'" & txtmodel & "', "
            InsertCommand = InsertCommand & "'" & lightingEquipsMobDate(intI) & "', "
            InsertCommand = InsertCommand & "'" & lightingEquipsDemobDate(intI) & "', "
            InsertCommand = InsertCommand & lightingEquipsQty(intI) & ", "
            InsertCommand = InsertCommand & lightingEquipsChkd(intI) & ", "
            InsertCommand = InsertCommand & lightingEquipsPPU(intI) & ", "
            InsertCommand = InsertCommand & lightingEquipsCLPerMc(intI) & ", "
            InsertCommand = InsertCommand & lightingEquipsUF(intI) & ")"
            Try
                If (moledbConnection1.State.ToString().Equals("Closed")) Then
                    moledbConnection1.Open()
                End If
                moleDBInsertCommand = New OleDbCommand
                moleDBInsertCommand.CommandType = CommandType.Text
                moleDBInsertCommand.CommandText = InsertCommand
                moleDBInsertCommand.Connection = moledbConnection1
                moleDBInsertCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            Finally
                moleDBInsertCommand = Nothing
                'moledbconnection3.Close()
            End Try
        Next
        Button1.Enabled = False
    End Sub
    Private Sub SaveFixedExpenses(ByVal msheetname As String, ByVal items As Integer)
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, mtablename As String = ""
        Dim intI As Integer, intJ As Integer = 0
        lblmessage.Text = "Fixed Expenses Data Being Saved. Please wait..."
        lblmessage.Visible = True
        Me.Refresh()
        xlWorksheet = xlWorkbook.Sheets(msheetname)
        xlWorksheet.Activate()
        DeleteRecordsFromAddedItemsTable("FixedExp")
        mtablename = GetTablename("FixedExp")
        getCategoryShortname(xlWorksheet)

        For intI = 0 To items - 1
            If fixedExpEquipsChkd(intI) = 1 Then
                'intJ = intJ + 1
                xlRange = xlWorksheet.Range(Category_Shortname & "Item1").Offset(intI, 0)
                xlRange.Value = fixedExpCategoryNames(intI)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = fixedExpEquipsQty(intI)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = fixedExpCost(intI)
                xlRange = xlRange.Offset(0, 2)
                xlRange.Value = mProjectvalue
                xlRange = xlRange.Offset(0, 2)
                xlRange.Value = ""
            Else
                xlRange = xlWorksheet.Range(Category_Shortname & "Item1").Offset(intI, 1)
                xlRange.Value = 0
                xlRange = xlRange.Offset(0, 5)
                xlRange.Value = ""
                xlRange.Cells.Application.ActiveCell.RowHeight = 0
            End If
            InsertCommand = ""
            InsertCommand = "INSERT INTO " & mtablename & " Values ("
            InsertCommand = InsertCommand & "'" & fixedExpCategoryNames(intI) & "', "
            InsertCommand = InsertCommand & fixedExpEquipsQty(intI) & ", "
            InsertCommand = InsertCommand & fixedExpCost(intI) & ", "
            InsertCommand = InsertCommand & fixedExpEquipsQty(intI) * fixedExpCost(intI) & ", "
            InsertCommand = InsertCommand & "'" & fixedExpRemarks(intI) & "', "
            InsertCommand = InsertCommand & fixedExpEquipsChkd(intI) & ", "
            InsertCommand = InsertCommand & fixedExpProjValue(intI) & ", "
            InsertCommand = InsertCommand & fixedExpEquipsCostPerc(intI) & ")"
            Try
                If (moledbConnection1.State.ToString().Equals("Closed")) Then
                    moledbConnection1.Open()
                End If
                moleDBInsertCommand = New OleDbCommand
                moleDBInsertCommand.CommandType = CommandType.Text
                moleDBInsertCommand.CommandText = InsertCommand
                moleDBInsertCommand.Connection = moledbConnection1
                moleDBInsertCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            Finally
                moleDBInsertCommand = Nothing
            End Try

        Next
        Dim r As Integer
        xlRange = xlWorksheet.Range(Category_Shortname & "Item1")
        While UCase(xlRange.Value) <> "TOTAL"
            'MsgBox(xlRange.Address & "...." & xlRange.Value)
            xlRange = xlRange.Offset(1, 0)
        End While
        xlRange.Cells.Application.ActiveCell.RowHeight = 23
        'For intJ = 1 To items
        '    If xlRange.Value = 0 Or Len(Trim(xlRange.Value)) = 0 Then
        '        xlRange.Select()
        '        r = xlRange.Application.ActiveCell.EntireRow.RowHeight = 0
        '    End If
        '    xlRange = xlRange.Offset(1, 0)
        'Next
        lblmessage.Visible = False
        Button1.Enabled = False
    End Sub
    Private Sub SaveFixedBPExpenses(ByVal msheetname As String, ByVal items As Integer)
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, mtablename As String = ""
        Dim intI As Integer, intJ As Integer = 0
        lblmessage.Text = "Fixed Expenses - BPlant details Being Saved. Please wait..."
        lblmessage.Visible = True
        Me.Refresh()
        xlWorksheet = xlWorkbook.Sheets(msheetname)
        xlWorksheet.Activate()
        DeleteRecordsFromAddedItemsTable("BPFixed - Exp")
        mtablename = GetTablename("BPFixed - Exp")
        getCategoryShortname(xlWorksheet)

        For intI = 0 To items - 1
            If fixedBPExpEquipsChkd(intI) = 1 Then
                'intJ = intJ + 1
                xlRange = xlWorksheet.Range(Category_Shortname & "Item1").Offset(intI, 0)
                'MsgBox(xlRange.Address)
                xlRange.Value = fixedBPExpCategoryNames(intI)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = fixedBPExpEquipsQty(intI)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = fixedBPExpCost(intI)
                xlRange = xlRange.Offset(0, 2)
                xlRange.Value = mProjectvalue
                xlRange = xlRange.Offset(0, 2)
                xlRange.Value = ""
            Else
                xlRange = xlWorksheet.Range(Category_Shortname & "Item1").Offset(intI, 1)
                xlRange.Value = 0
                xlRange = xlRange.Offset(0, 5)
                xlRange.Value = ""
                xlRange.Cells.Application.ActiveCell.RowHeight = 0
            End If
            InsertCommand = ""
            InsertCommand = "INSERT INTO " & mtablename & " Values ("
            InsertCommand = InsertCommand & "'" & fixedBPExpCategoryNames(intI) & "', "
            InsertCommand = InsertCommand & fixedBPExpEquipsQty(intI) & ", "
            InsertCommand = InsertCommand & fixedBPExpCost(intI) & ", "
            InsertCommand = InsertCommand & fixedBPExpEquipsQty(intI) * fixedExpCost(intI) & ", "
            InsertCommand = InsertCommand & "'" & fixedBPExpRemarks(intI) & "', "
            InsertCommand = InsertCommand & fixedBPExpEquipsChkd(intI) & ", "
            InsertCommand = InsertCommand & fixedBPExpProjValue(intI) & ", "
            InsertCommand = InsertCommand & fixedBPExpEquipsCostPerc(intI) & ")"
            Try
                If (moledbConnection1.State.ToString().Equals("Closed")) Then
                    moledbConnection1.Open()
                End If
                moleDBInsertCommand = New OleDbCommand
                moleDBInsertCommand.CommandType = CommandType.Text
                moleDBInsertCommand.CommandText = InsertCommand
                moleDBInsertCommand.Connection = moledbConnection1
                moleDBInsertCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            Finally
                moleDBInsertCommand = Nothing
            End Try
        Next
        Dim r As Integer
        xlRange = xlWorksheet.Range(Category_Shortname & "Item1")
        While UCase(xlRange.Value) <> "TOTAL"
            'MsgBox(xlRange.Address & "...." & xlRange.Value)
            xlRange = xlRange.Offset(1, 0)
        End While
        xlRange.Cells.Application.ActiveCell.RowHeight = 23        'For intJ = 1 To items
        '    If xlRange.Value = 0 Or Len(Trim(xlRange.Value)) = 0 Then
        '        xlRange.Select()
        '        r = xlRange.Application.ActiveCell.EntireRow.RowHeight = 0
        '    End If
        '    xlRange = xlRange.Offset(1, 0)
        'Next
        lblmessage.Visible = False
        Button1.Enabled = False
    End Sub
    Private Sub SaveHiredEquipments()
        Dim intI As Integer  ', mcategory As String
        Dim Counts As Integer = 0, RepeatFactor As Integer  'intj As Integer,
        Currenttab = Me.tbcBdgetHeads.SelectedIndex
        If Len(Trim(xlFilename)) = 0 Then
            MsgBox("Filename not specified to save. cannot Save data. Start the application again.")
            Application.Exit()
        End If

        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With

        Me.btnQuit.Enabled = False
        Me.btnClose.Enabled = False
        Dim cnt As Integer = 0
        mcategory = hiredCategoryNames(0)

        HireItems = hiredEquipsNames.Length
        HiredCheckedItems = 0
        For intI = 0 To HireItems - 1
            If hiredEquipsChkd(intI) = 1 Then
                HiredCheckedItems = HiredCheckedItems + 1
            End If
        Next
        DeleteRecordsFromAddedItemsTable(mcategory)
        xlWorksheet = xlWorkbook.Sheets("external Hire")
        xlWorksheet.Activate()
        RepeatFactor = HireItems
        RecordsInserted(getSheetNo(xlWorksheet.Name)) = 0
        RecordsInserted(getSheetNo("External Others")) = 0

        SaveHEDataInSheet("external Hire", cnt, RepeatFactor)
        lblmessage.Visible = False
        Button1.Enabled = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim frmSaveMsg As Form
        frmSaveMsg = New Form2A()
        frmSaveMsg.ShowDialog()
        frmSaveMsg = Nothing
        Dim msheetname As String

        If answer = vbNo Then
            lblmessage.Text = "Saving canceled."
            DataSaved = False
            Me.lblmessage.Visible = True
            Me.Refresh()
            Exit Sub
        End If

        starttime = Now()

        lblmessage.Text = "Validating details before saving.... Pleas wait for few minutes"
        lblmessage.Visible = True
        Me.Refresh()
        If Not ValidateForSave() Then
            lblmessage.Text = "Validation Error. Please correct and  click 'Quit Data Entry' button again"
            Me.btnQuit.Enabled = True
            Me.Refresh()
            DataSaved = False
            Exit Sub
        End If
        Me.tbcBdgetHeads.Enabled = False
        Me.Panel1.Enabled = False
        Me.Panel3.Enabled = False
        Me.btnQuit.Enabled = False
        Me.btnClose.Enabled = False

        Dim intI As Integer
        Dim cnt As Integer
        'Dim Counts As Integer = 0 'Prevcategory As String, RepeatFactor As Integer  'intj As Integer,
        Currenttab = Me.tbcBdgetHeads.SelectedIndex

        If Len(Trim(xlFilename)) = 0 Then
            MsgBox("Filename not specified to save. Cannot Save data. Start the application again.")
            Application.Exit()
        End If

        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With

        Me.btnQuit.Enabled = False
        Me.btnClose.Enabled = False

        If optMajConcrete.Checked Then
            optMajConcrete.Checked = False
            Panel2.Visible = True
        End If
        If Me.optMajConvyance.Checked Then
            optMajConvyance.Checked = False
            Panel2.Visible = True
        End If
        If optMajCrane.Checked Then
            optMajCrane.Checked = False
            Panel2.Visible = True
        End If
        If optMajDGSets.Checked Then
            Panel2.Visible = True
        End If
        If optMajMH.Checked Then
            optMajMH.Checked = False
            Panel2.Visible = True
        End If
        If optMajNc.Checked Then
            optMajNc.Checked = False
            Panel2.Visible = True
        End If
        If optMajOthers.Checked Then
            optMajOthers.Checked = False
            Panel2.Visible = True
        End If
        If optMinorEquips.Checked Then
            optMinorEquips.Checked = False
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If
        If Me.optHireEquipments.Checked Then
            optHireEquipments.Checked = False
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If
        If Me.optLightingEquips.Checked Then
            optLightingEquips.Checked = False
            Panel7.Left = 1
            Panel7.Top = 15
            Panel7.Height = 45
            Me.Label283.Text = "Conn" & vbNewLine & "Load"
            Me.Label282.Text = "Utility" & vbNewLine & "Factor"
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        End If
        If Me.optFixedExp.Checked Then
            optFixedExp.Checked = False
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        End If
        If Me.optBPFixedExp.Checked Then
            optBPFixedExp.Checked = False
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        End If

        concretecheckeditems = 0
        For intI = 0 To concEquipsNames.Length - 1
            If concEquipsChkd(intI) = 1 Then
                concretecheckeditems = concretecheckeditems + 1
            End If
        Next
        DeleteRecordsFromAddedItemsTable("Concreting")
        xlWorksheet = xlWorkbook.Sheets("Concreting")
        xlWorksheet.Activate()
        RepeatFactor = ConcretingItems - 1
        RecordsInserted(getSheetNo(xlWorksheet.Name)) = 0
        cnt = 0
        DataSaved = True
        SaveDataInConcSheet("Concreting", cnt, RepeatFactor)
        If Not DataSaved Then
            MsgBox("Error in saving Concreting equipments details. Saving operation is terrminated")
            btnQuit.Enabled = True
            Exit Sub
        End If
        Concretesaved = True

        ConveyanceItems = convEquipsNames.Length '- 1
        conveyancecheckeditems = 0
        For intI = 0 To ConveyanceItems - 1
            If convEquipsChkd(intI) = 1 Then
                conveyancecheckeditems = conveyancecheckeditems + 1
            End If
        Next
        DeleteRecordsFromAddedItemsTable("Conveyance")
        xlWorksheet = xlWorkbook.Sheets("Conveyance")
        xlWorksheet.Activate()
        RepeatFactor = ConveyanceItems - 1
        RecordsInserted(getSheetNo(xlWorksheet.Name)) = 0
        cnt = 0
        DataSaved = True
        SaveDataInConvSheet("Conveyance", cnt, RepeatFactor)
        ConveyanceSaved = True

        CraneItems = craneEquipsNames.Length '- 1
        cranecheckeditems = 0
        For intI = 0 To CraneItems - 1
            If craneEquipsChkd(intI) = 1 Then
                cranecheckeditems = cranecheckeditems + 1
            End If
        Next
        DeleteRecordsFromAddedItemsTable("Cranes")
        xlWorksheet = xlWorkbook.Sheets("Cranes")
        xlWorksheet.Activate()
        RepeatFactor = CraneItems - 1
        RecordsInserted(getSheetNo(xlWorksheet.Name)) = 0
        SaveDataInCraneSheet("Cranes", cnt, RepeatFactor)
        CranesSaved = True

        DGSetItems = dgsetsEquipsNames.Length '- 1
        dgsetscheckeditems = 0
        For intI = 0 To DGSetItems - 1
            If dgsetsEquipsChkd(intI) = 1 Then
                dgsetscheckeditems = dgsetscheckeditems + 1
            End If
        Next
        DeleteRecordsFromAddedItemsTable("DG Sets")
        xlWorksheet = xlWorkbook.Sheets("DG Sets")
        xlWorksheet.Activate()
        RepeatFactor = DGSetItems - 1
        RecordsInserted(getSheetNo(xlWorksheet.Name)) = 0
        SaveDataInDgSetsSheet("DG Sets", cnt, RepeatFactor)
        dgsetssaved = True

        MHItems = MHEquipsNames.Length '- 1
        mhcheckeditems = 0
        For intI = 0 To MHItems - 1
            If MHEquipsChkd(intI) = 1 Then
                mhcheckeditems = mhcheckeditems + 1
            End If
        Next
        DeleteRecordsFromAddedItemsTable("Material Handling")
        xlWorksheet = xlWorkbook.Sheets("Material Handling")
        xlWorksheet.Activate()
        RepeatFactor = MHItems - 1
        RecordsInserted(getSheetNo(xlWorksheet.Name)) = 0
        SaveDataInMHSheet("Material Handling", cnt, RepeatFactor)
        MHSaved = True

        NCItems = NCEquipsNames.Length '- 1
        nccheckeditems = 0
        For intI = 0 To NCItems - 1
            If NCEquipsChkd(intI) = 1 Then
                nccheckeditems = nccheckeditems + 1
            End If
        Next
        DeleteRecordsFromAddedItemsTable("Non Concreting")
        xlWorksheet = xlWorkbook.Sheets("Non Concreting")
        xlWorksheet.Activate()
        RepeatFactor = NCItems - 1
        RecordsInserted(getSheetNo(xlWorksheet.Name)) = 0
        SaveDataInNCSheet("Non Concreting", cnt, RepeatFactor)
        NCSaved = True

        MajorOtherItems = majOthersEquipsNames.Length    ' - 1
        majorotherscheckeditems = 0
        For intI = 0 To MajorOtherItems - 1
            If majOthersEquipsChkd(intI) = 1 Then
                majorotherscheckeditems = majorotherscheckeditems + 1
            End If
        Next
        DeleteRecordsFromAddedItemsTable("Major Others")
        xlWorksheet = xlWorkbook.Sheets("Major Others")
        xlWorksheet.Activate()
        RepeatFactor = MajorOtherItems - 1
        RecordsInserted(getSheetNo(xlWorksheet.Name)) = 0
        SaveDataInMajOthersSheet("Major Others", cnt, RepeatFactor)
        MajOthersSaved = True

        SaveMinorEquipments()
        MinorEquipsSaved = True

        SaveHiredEquipments()
        HiredEquipsSaved = True

        msheetname = "Tr crane related exp"
        fexpItems = fixedExpCategoryNames.Length
        SaveFixedExpenses(msheetname, fexpItems)

        msheetname = "Bplant related exp"
        BPFExpItems = fixedBPExpCategoryNames.Length
        SaveFixedBPExpenses(msheetname, BPFExpItems)
        fixedBpExpSaved = True


        msheetname = "PowerReqmt"
        LightingItems = lightingCategoryNames.Length
        SaveLightingBudget(msheetname, LightingItems)
        LightingSaved = True

        Button1.Enabled = False

    End Sub
    Private Sub DeleteRecordsFromAddedItemsTable(ByVal mcategory As String)
        Dim DeleteCommand As String
        DeleteCommand = "Delete From  " & GetTablename(mcategory)
        Try
            If (moledbConnection1.State.ToString().Equals("Closed")) Then
                moledbConnection1.Open()
            End If
            moledbCommand = New OleDbCommand
            moledbCommand.CommandType = CommandType.Text
            moledbCommand.CommandText = DeleteCommand
            moledbCommand.Connection = moledbConnection1
            moledbCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.ToString())
        Finally
            moledbCommand = Nothing
        End Try
    End Sub
    Private Sub CopyPowerReqTemplate()
        IfExistsDelete("PowerReqmt")
        Dim worksheet1 As Microsoft.Office.Interop.Excel.Worksheet = xlWorkbook.Worksheets("PowerReqmtTemplate")
        worksheet1.Copy(Before:=xlWorkbook.Worksheets("PowerReqmtTemplate"))
        Dim intk As Integer = worksheet1.Index
        intk = intk - 1
        worksheet1 = xlWorkbook.Worksheets(intk)
        worksheet1.Name = "PowerReqmt"
        xlRange = worksheet1.Range("PowerReq_RatePerUnit")
        xlRange.Value = PowerCostPerUnit
        worksheet1.Visible = True
    End Sub
    Private Sub IfExistsDelete(ByVal sheetname As String)
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet, index As Integer
        For Each ws In xlWorkbook.Worksheets
            If UCase(ws.Name) = UCase(sheetname) Then
                index = ws.Index
                xlApp.DisplayAlerts = False
                CType(xlWorkbook.Sheets(index), Microsoft.Office.Interop.Excel.Worksheet).Delete()
                xlWorkbook.Save()
                xlWorkbook.Close((Microsoft.Office.Interop.Excel.XlSaveAction.xlSaveChanges))
                xlWorkbook = Nothing
                xlWorkbook = xlApp.Workbooks.Open(Currentfile)
                xlApp.DisplayAlerts = True
                Exit For
            End If
        Next
    End Sub
    Private Sub deleteOlddata(ByVal sheetname As String)
        Dim intI As Integer, msheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim cols As Integer
        Dim mCommand As New OleDbCommand
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlWorksheet.Activate()
        For intI = 70 To 2 Step -1
            xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(intI, 0)
            xlRange.Select()
            xlRange.Application.ActiveCell.EntireRow.Delete()
        Next
        xlWorksheet = xlWorkbook.Worksheets(sheetname)
        xlWorksheet.Select()
        getCategoryShortname(xlWorksheet)
        If Category_Shortname = "Concrete_" Or Category_Shortname = "Min_" Or Category_Shortname = "Ext_" Or _
             Category_Shortname = "ExtOthers_" Then
            cols = 34
        ElseIf Category_Shortname = "PowerGen_" Then
            cols = 29
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
        For intI = 1 To 100
            xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(intI, 0)
            xlRange.Select()
            xlRange.Application.ActiveCell.EntireRow.RowHeight = 27
        Next


        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = 0
    End Sub
    Private Function PromptToSave(ByVal eqptCategory As String) As Boolean
        Dim Ans As Integer
        Dim msgprompt As String = "If you have made changes to the " & eqptCategory & " Equipments selection, " & vbNewLine & _
                                  "you have to save before going to select other category equipments. " & vbNewLine & _
                                  "Do you want to Save now?"
        If Not concFirsttime Then
            If Button1.Enabled = True Then
                Ans = MsgBox(msgprompt, MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Save Data Before going to other category")
                If Ans = vbYes Then Return True Else Return False
            End If
        End If
    End Function
    Private Function ValidForSave(ByVal sheetname As String, ByVal Fromval As Integer, ByVal Toval As Integer) As Boolean
        Dim intI As Integer, intj As Integer
        ValidForSave = True
        intj = 1

        For intI = Fromval To Fromval + Toval - 1
            If Checkboxes(intI).Checked Then
                If (MobdatePickers(intI).Value.Date > DemobDatePickers(intI).Value.Date) Or _
                    (MobdatePickers(intI).Value.Date < mStartDate) Then
                    Dim msgstr = "Error in Mob. Date for " & CategoryTextBoxes(intI).Text & "," & EquipNameTextBoxes(intI).Text & "," & _
                       CapacityTextBoxes(intI).Text & "," & MakeModelTextBoxes(intI).Text & vbNewLine & _
                    "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                      "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                    MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                    MobdatePickers(intI).Value = mStartDate.Date
                    MobdatePickers(intI).Focus()
                    ValidForSave = False
                    Button1.Enabled = True
                    Exit Function
                End If

                If (DemobDatePickers(intI).Value.Date < MobdatePickers(intI).Value.Date) Or _
                    (DemobDatePickers(intI).Value.Date > mEndDate) Then
                    Dim msgstr = "Error in Demob. Date for " & CategoryTextBoxes(intI).Text & "," & EquipNameTextBoxes(intI).Text & "," & _
                       CapacityTextBoxes(intI).Text & "," & MakeModelTextBoxes(intI).Text & vbNewLine & _
                    "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                      "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                    MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                    DemobDatePickers(intI).Value = mStartDate.Date
                    DemobDatePickers(intI).Focus()
                    ValidForSave = False
                    Button1.Enabled = True
                    Exit Function
                End If

                If (Val(QtyTextBoxes(intI).Text) = 0 Or Len(Trim(QtyTextBoxes(intI).Text)) = 0) Then
                    MsgBox("Quantity for " & CategoryTextBoxes(intI).Text & "," & EquipNameTextBoxes(intI).Text & "," & _
                       CapacityTextBoxes(intI).Text & "," & MakeModelTextBoxes(intI).Text & " is missing", MsgBoxStyle.Critical, "Error in Qty Entry")
                    QtyTextBoxes(intI).Text = 1
                    QtyTextBoxes(intI).Focus()
                    ValidForSave = False
                    Button1.Enabled = True
                    Exit Function
                End If
                If sheetname = "Minor Eqpts" Then
                    For intj = 2 To CategoryTextBoxes.Count - 1
                        If (Checkboxes(intj).Checked And Checkboxes(intj - 1).Checked) Then
                            If (CategoryTextBoxes(intj).Text = CategoryTextBoxes(intj - 1).Text And _
                                EquipNameTextBoxes(intj).Text = EquipNameTextBoxes(intj - 1).Text And _
                                CapacityTextBoxes(intj).Text = CapacityTextBoxes(intj - 1).Text And _
                                MakeModelTextBoxes(intj).Text = MakeModelTextBoxes(intj - 1).Text And _
                                MobdatePickers(intj).Value.Date = MobdatePickers(intj - 1).Value.Date And _
                                DemobDatePickers(intj).Value.Date = DemobDatePickers(intj - 1).Value.Date And _
                                DepPercComboboxes(intj).Text = DepPercComboboxes(intj - 1).Text) And _
                                (IsNewMC(intj).Checked And IsNewMC(intj - 1).Checked) Then
                                MsgBox("Items " & intj - 1 & " and " & intj & " are duplicate entries. Please Check. Data not saved")
                                ValidForSave = False
                                Button1.Enabled = True
                                Exit Function
                            End If
                        ElseIf (Checkboxes(intj).Checked And Checkboxes(intj + 1).Checked) Then
                            If (CategoryTextBoxes(intj).Text = CategoryTextBoxes(intj + 1).Text And _
                                    EquipNameTextBoxes(intj).Text = EquipNameTextBoxes(intj + 1).Text And _
                                    CapacityTextBoxes(intj).Text = CapacityTextBoxes(intj + 1).Text And _
                                    MakeModelTextBoxes(intj).Text = MakeModelTextBoxes(intj + 1).Text And _
                                    MobdatePickers(intj).Value.Date = MobdatePickers(intj + 1).Value.Date And _
                                    DemobDatePickers(intj).Value.Date = DemobDatePickers(intj + 1).Value.Date And _
                                    DepPercComboboxes(intj).Text = DepPercComboboxes(intj + 1).Text) And _
                                    (IsNewMC(intj).Checked And IsNewMC(intj + 1).Checked) Then
                                MsgBox("Items " & intj & " and " & intj + 1 & " are duplicate entries. Please Check. Data not saved")
                                ValidForSave = False
                                Button1.Enabled = True
                                Exit Function
                            End If
                        End If
                    Next
                ElseIf sheetname = "PowerReqmt" Then
                    If Not intj = 1 Then
                        If (Checkboxes(intj).Checked And Checkboxes(intj - 1).Checked) Then
                            If (EquipNameTextBoxes(intj).Text = EquipNameTextBoxes(intj - 1).Text And _
                                CapacityTextBoxes(intj).Text = CapacityTextBoxes(intj - 1).Text And _
                                MakeModelTextBoxes(intj).Text = MakeModelTextBoxes(intj - 1).Text And _
                                MobdatePickers(intj).Value.Date = MobdatePickers(intj - 1).Value.Date And _
                                DemobDatePickers(intj).Value.Date = DemobDatePickers(intj - 1).Value.Date) Then
                                MsgBox("Items " & intj - 1 & " and " & intj & " are duplicate entries. Please Check. Data not saved")
                                ValidForSave = False
                                Button1.Enabled = True
                                Exit Function
                            End If
                        End If
                    ElseIf Not intj >= Checkboxes.Count Then
                        If (Checkboxes(intj).Checked And Checkboxes(intj + 1).Checked) Then
                            If (EquipNameTextBoxes(intj).Text = EquipNameTextBoxes(intj + 1).Text And _
                                CapacityTextBoxes(intj).Text = CapacityTextBoxes(intj + 1).Text And _
                                MakeModelTextBoxes(intj).Text = MakeModelTextBoxes(intj + 1).Text And _
                                MobdatePickers(intj).Value.Date = MobdatePickers(intj + 1).Value.Date And _
                                DemobDatePickers(intj).Value.Date = DemobDatePickers(intj + 1).Value.Date) Then
                                MsgBox("Items " & intj & " and " & intj + 1 & " are duplicate entries. Please Check. Data not saved")
                                ValidForSave = False
                                Button1.Enabled = True
                                Exit Function
                            End If
                        End If
                    End If
                Else
                    For intj = 2 To CategoryTextBoxes.Count - 1
                        If (Checkboxes(intj).Checked And Checkboxes(intj - 1).Checked) Then
                            If (CategoryTextBoxes(intj).Text = CategoryTextBoxes(intj - 1).Text And _
                                EquipNameTextBoxes(intj).Text = EquipNameTextBoxes(intj - 1).Text And _
                                CapacityTextBoxes(intj).Text = CapacityTextBoxes(intj - 1).Text And _
                                MakeModelTextBoxes(intj).Text = MakeModelTextBoxes(intj - 1).Text And _
                                MobdatePickers(intj).Value.Date = MobdatePickers(intj - 1).Value.Date And _
                                DemobDatePickers(intj).Value.Date = DemobDatePickers(intj - 1).Value.Date And _
                                DepPercComboboxes(intj).Text = DepPercComboboxes(intj - 1).Text) Then
                                MsgBox("Items " & intj - 1 & " and " & intj & " are duplicate entries. Please Check. Data not saved")
                                ValidForSave = False
                                Button1.Enabled = True
                                Exit Function
                            End If
                        ElseIf (Checkboxes(intj).Checked And Checkboxes(intj + 1).Checked) Then
                            If (CategoryTextBoxes(intj).Text = CategoryTextBoxes(intj + 1).Text And _
                                    EquipNameTextBoxes(intj).Text = EquipNameTextBoxes(intj + 1).Text And _
                                    CapacityTextBoxes(intj).Text = CapacityTextBoxes(intj + 1).Text And _
                                    MakeModelTextBoxes(intj).Text = MakeModelTextBoxes(intj + 1).Text And _
                                    MobdatePickers(intj).Value.Date = MobdatePickers(intj + 1).Value.Date And _
                                    DemobDatePickers(intj).Value.Date = DemobDatePickers(intj + 1).Value.Date And _
                                    DepPercComboboxes(intj).Text = DepPercComboboxes(intj + 1).Text) Then
                                MsgBox("Items " & intj & " and " & intj + 1 & " are duplicate entries. Please Check. Data not saved")
                                ValidForSave = False
                                Button1.Enabled = True
                                Exit Function
                            End If
                        End If
                    Next
                End If
            End If
        Next
    End Function
    Private Function ValidateForSave()
        Dim intI As Integer, intj As Integer
        Dim mMakeModel As String
        ValidateForSave = True
        intj = 1
        Dim mcategories(0 To 11) As String, CategoryItems(0 To 11) As Integer
        mcategories(0) = "Concreting" : CategoryItems(0) = ConcretingItems
        mcategories(1) = "Conveyance" : CategoryItems(1) = ConveyanceItems
        mcategories(2) = "Cranes" : CategoryItems(2) = CraneItems
        mcategories(3) = "DG Sets" : CategoryItems(3) = DGSetItems
        mcategories(4) = "Material Handling" : CategoryItems(4) = MHItems
        mcategories(5) = "Non Concreting" : CategoryItems(5) = NCItems
        mcategories(6) = "Major Others" : CategoryItems(6) = MajorOtherItems
        mcategories(7) = "Minor Equipments" : CategoryItems(7) = MinorItems
        mcategories(8) = "HiredEquipments" : CategoryItems(8) = HireItems
        mcategories(9) = "FixedExp" : CategoryItems(9) = fexpItems
        mcategories(10) = "FixedExp - BP" : CategoryItems(10) = BPFExpItems
        mcategories(11) = "lighting" : CategoryItems(11) = LightingItems

        For intI = 0 To 8
            Select Case intI
                Case 0
                    For intj = 0 To CategoryItems(intI) - 1
                        If concEquipsChkd(intj) = 1 Then
                            mMakeModel = concEquipsMake(intj) & "/" & concEquipsModel(intj)
                            If (concEquipsMobDate(intj) > concEquipsDemobDate(intj)) Or _
                                (concEquipsMobDate(intj) < mStartDate) Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & concEquipsNames(intj) & "," & _
                                   concEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            If (concEquipsDemobDate(intj) < concEquipsMobDate(intj)) Or _
                                (concEquipsDemobDate(intj) > mEndDate) Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & concEquipsNames(intj) & "," & _
                                   concEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            If (concEquipsQty(intj) = 0) Then   'Or Len(Trim(conce)) = 0) Then
                                MsgBox("Quantity for " & mcategories(intI) & "," & concEquipsNames(intj) & "," & _
                                   concEquipsCapacity(intj) & "," & mMakeModel & " is missing", MsgBoxStyle.Critical, "Error in Qty Entry")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next intj

                    For intj = 1 To CategoryItems(intI) - 2
                        If (concEquipsChkd(intj) = 1 And concEquipsChkd(intj - 1) = 1) Then
                            If (concEquipsNames(intj) = concEquipsNames(intj - 1) And _
                                concEquipsCapacity(intj) = concEquipsCapacity(intj - 1) And _
                                concEquipsMake(intj) = concEquipsMake(intj - 1) And _
                                concEquipsModel(intj) = concEquipsModel(intj - 1) And _
                                concEquipsMobDate(intj) = concEquipsMobDate(intj - 1) And _
                                concEquipsDemobDate(intj) = concEquipsDemobDate(intj - 1) And _
                                concEquipsDepPerc(intj) = concEquipsDepPerc(intj - 1)) Then
                                MsgBox("Items " & intj - 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        ElseIf (concEquipsChkd(intj) = 1 And concEquipsChkd(intj + 1) = 1) Then
                            If (concEquipsNames(intj) = concEquipsNames(intj + 1) And _
                                concEquipsCapacity(intj) = concEquipsCapacity(intj + 1) And _
                                concEquipsMake(intj) = concEquipsMake(intj + 1) And _
                                concEquipsModel(intj) = concEquipsModel(intj + 1) And _
                                concEquipsMobDate(intj) = concEquipsMobDate(intj + 1) And _
                                concEquipsDemobDate(intj) = concEquipsDemobDate(intj + 1) And _
                                concEquipsDepPerc(intj) = concEquipsDepPerc(intj + 1)) Then
                                MsgBox("Items " & intj + 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next
                Case 1
                    For intj = 0 To CategoryItems(intI) - 1
                        If convEquipsChkd(intj) = 1 Then
                            mMakeModel = convEquipsMake(intj) & "/" & convEquipsModel(intj)
                            If DateValue(convEquipsMobDate(intj)) > DateValue(convEquipsDemobDate(intj)) Or _
                                DateValue(convEquipsMobDate(intj)) < mStartDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & convEquipsNames(intj) & "," & _
                                   convEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            If DateValue(convEquipsMobDate(intj)) > DateValue(convEquipsDemobDate(intj)) Or _
                                DateValue(convEquipsDemobDate(intj)) > mEndDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & convEquipsNames(intj) & "," & _
                                                                  convEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                                               "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                                                                 "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If

                        If (convEquipsQty(intj) = 0) Then   'Or Len(Trim(conce)) = 0) Then
                            MsgBox("Quantity for " & mcategories(intI) & "," & convEquipsNames(intj) & "," & _
                               convEquipsCapacity(intj) & "," & mMakeModel & " is missing", MsgBoxStyle.Critical, "Error in Qty Entry")
                            ValidateForSave = False
                            Exit Function
                        End If
                    Next intj

                    For intj = 1 To CategoryItems(intI) - 2
                        If (convEquipsChkd(intj) = 1 And convEquipsChkd(intj - 1) = 1) Then
                            If (convEquipsNames(intj) = convEquipsNames(intj - 1) And _
                                convEquipsCapacity(intj) = convEquipsCapacity(intj - 1) And _
                                convEquipsMake(intj) = convEquipsMake(intj - 1) And _
                                convEquipsModel(intj) = convEquipsModel(intj - 1) And _
                                convEquipsMobDate(intj) = convEquipsMobDate(intj - 1) And _
                                convEquipsDemobDate(intj) = convEquipsDemobDate(intj - 1) And _
                                convEquipsDepPerc(intj) = convEquipsDepPerc(intj - 1)) Then
                                MsgBox("Items " & intj - 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        ElseIf (convEquipsChkd(intj) = 1 And convEquipsChkd(intj + 1) = 1) Then
                            If (convEquipsNames(intj) = convEquipsNames(intj + 1) And _
                                convEquipsCapacity(intj) = convEquipsCapacity(intj + 1) And _
                                convEquipsMake(intj) = convEquipsMake(intj + 1) And _
                                convEquipsModel(intj) = convEquipsModel(intj + 1) And _
                                convEquipsMobDate(intj) = convEquipsMobDate(intj + 1) And _
                                convEquipsDemobDate(intj) = convEquipsDemobDate(intj + 1) And _
                                convEquipsDepPerc(intj) = convEquipsDepPerc(intj + 1)) Then
                                MsgBox("Items " & intj + 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next
                Case 2
                    For intj = 0 To CategoryItems(intI) - 1
                        If craneEquipsChkd(intj) = 1 Then
                            mMakeModel = craneEquipsMake(intj) & "/" & craneEquipsModel(intj)
                            'If (craneEquipsMobDate(intj) > craneEquipsDemobDate(intj)) Then
                            '    If (craneEquipsMobDate(intj) < mStartDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & craneEquipsNames(intj) & "," & _
                            '            craneEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '            "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                            '            "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If

                            If DateValue(craneEquipsMobDate(intj)) > DateValue(craneEquipsDemobDate(intj)) Or _
                                DateValue(craneEquipsMobDate(intj)) < mStartDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & craneEquipsNames(intj) & "," & _
                                   craneEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            'If (craneEquipsDemobDate(intj) < craneEquipsMobDate(intj)) Then
                            '    If (craneEquipsDemobDate(intj) > mEndDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & craneEquipsNames(intj) & "," & _
                            '           craneEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If
                            If DateValue(craneEquipsDemobDate(intj)) < DateValue(craneEquipsMobDate(intj)) Or _
                                DateValue(craneEquipsDemobDate(intj)) > mEndDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & craneEquipsNames(intj) & "," & _
                                   craneEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            If (craneEquipsQty(intj) = 0) Then   'Or Len(Trim(conce)) = 0) Then
                                MsgBox("Quantity for " & mcategories(intI) & "," & craneEquipsNames(intj) & "," & _
                                   craneEquipsCapacity(intj) & "," & mMakeModel & " is missing", MsgBoxStyle.Critical, "Error in Qty Entry")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next intj

                    For intj = 1 To CategoryItems(intI) - 2
                        If (craneEquipsChkd(intj) = 1 And craneEquipsChkd(intj - 1) = 1) Then
                            If (craneEquipsNames(intj) = craneEquipsNames(intj - 1) And _
                                craneEquipsCapacity(intj) = craneEquipsCapacity(intj - 1) And _
                                craneEquipsMake(intj) = craneEquipsMake(intj - 1) And _
                                craneEquipsModel(intj) = craneEquipsModel(intj - 1) And _
                                craneEquipsMobDate(intj) = craneEquipsMobDate(intj - 1) And _
                                craneEquipsDemobDate(intj) = craneEquipsDemobDate(intj - 1) And _
                                craneEquipsDepPerc(intj) = craneEquipsDepPerc(intj - 1)) Then
                                MsgBox("Items " & intj - 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        ElseIf (craneEquipsChkd(intj) = 1 And craneEquipsChkd(intj + 1) = 1) Then
                            If (craneEquipsNames(intj) = craneEquipsNames(intj + 1) And _
                                craneEquipsCapacity(intj) = craneEquipsCapacity(intj + 1) And _
                                craneEquipsMake(intj) = craneEquipsMake(intj + 1) And _
                                craneEquipsModel(intj) = craneEquipsModel(intj + 1) And _
                                craneEquipsMobDate(intj) = craneEquipsMobDate(intj + 1) And _
                                craneEquipsDemobDate(intj) = craneEquipsDemobDate(intj + 1) And _
                                craneEquipsDepPerc(intj) = craneEquipsDepPerc(intj + 1)) Then
                                MsgBox("Items " & intj + 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next
                Case 3
                    For intj = 0 To CategoryItems(intI) - 1
                        If dgsetsEquipsChkd(intj) = 1 Then
                            mMakeModel = dgsetsEquipsMake(intj) & "/" & dgsetsEquipsModel(intj)
                            'If (dgsetsEquipsMobDate(intj) > dgsetsEquipsDemobDate(intj)) Then
                            '    If (dgsetsEquipsMobDate(intj) < mStartDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & dgsetsEquipsNames(intj) & "," & _
                            '           dgsetsEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If
                            If DateValue(dgsetsEquipsMobDate(intj)) > DateValue(dgsetsEquipsDemobDate(intj)) Or _
                                DateValue(dgsetsEquipsMobDate(intj)) < mStartDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & dgsetsEquipsNames(intj) & "," & _
                                   dgsetsEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            'If (dgsetsEquipsDemobDate(intj) < dgsetsEquipsMobDate(intj)) Then
                            '    If (dgsetsEquipsDemobDate(intj) > mEndDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & dgsetsEquipsNames(intj) & "," & _
                            '           dgsetsEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If
                            If DateValue(dgsetsEquipsDemobDate(intj)) < DateValue(dgsetsEquipsMobDate(intj)) Or _
                               DateValue(dgsetsEquipsDemobDate(intj)) > mEndDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & dgsetsEquipsNames(intj) & "," & _
                                   dgsetsEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            If (dgsetsEquipsQty(intj) = 0) Then   'Or Len(Trim(conce)) = 0) Then
                                MsgBox("Quantity for " & mcategories(intI) & "," & dgsetsEquipsNames(intj) & "," & _
                                   dgsetsEquipsCapacity(intj) & "," & mMakeModel & " is missing", MsgBoxStyle.Critical, "Error in Qty Entry")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next intj

                    For intj = 1 To CategoryItems(intI) - 2
                        If (dgsetsEquipsChkd(intj) = 1 And dgsetsEquipsChkd(intj - 1) = 1) Then
                            If (dgsetsEquipsNames(intj) = dgsetsEquipsNames(intj - 1) And _
                                dgsetsEquipsCapacity(intj) = dgsetsEquipsCapacity(intj - 1) And _
                                dgsetsEquipsMake(intj) = dgsetsEquipsMake(intj - 1) And _
                                dgsetsEquipsModel(intj) = dgsetsEquipsModel(intj - 1) And _
                                dgsetsEquipsMobDate(intj) = dgsetsEquipsMobDate(intj - 1) And _
                                dgsetsEquipsDemobDate(intj) = dgsetsEquipsDemobDate(intj - 1) And _
                                dgsetsEquipsDepPerc(intj) = dgsetsEquipsDepPerc(intj - 1)) Then
                                MsgBox("Items " & intj - 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        ElseIf (dgsetsEquipsChkd(intj) = 1 And dgsetsEquipsChkd(intj + 1) = 1) Then
                            If (dgsetsEquipsNames(intj) = dgsetsEquipsNames(intj + 1) And _
                                dgsetsEquipsCapacity(intj) = dgsetsEquipsCapacity(intj + 1) And _
                                dgsetsEquipsMake(intj) = dgsetsEquipsMake(intj + 1) And _
                                dgsetsEquipsModel(intj) = dgsetsEquipsModel(intj + 1) And _
                                dgsetsEquipsMobDate(intj) = dgsetsEquipsMobDate(intj + 1) And _
                                dgsetsEquipsDemobDate(intj) = dgsetsEquipsDemobDate(intj + 1) And _
                                dgsetsEquipsDepPerc(intj) = dgsetsEquipsDepPerc(intj + 1)) Then
                                MsgBox("Items " & intj + 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next
                Case 4
                    For intj = 0 To CategoryItems(intI) - 1
                        If MHEquipsChkd(intj) = 1 Then
                            mMakeModel = MHEquipsMake(intj) & "/" & MHEquipsModel(intj)
                            'If (MHEquipsMobDate(intj) > MHEquipsDemobDate(intj)) Then
                            '    If (MHEquipsMobDate(intj) < mStartDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & MHEquipsNames(intj) & "," & _
                            '           MHEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If

                            If DateValue(MHEquipsMobDate(intj)) > DateValue(MHEquipsDemobDate(intj)) Or _
                                DateValue(MHEquipsMobDate(intj)) < mStartDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & MHEquipsNames(intj) & "," & _
                                   MHEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            'If (MHEquipsDemobDate(intj) < MHEquipsMobDate(intj)) Then
                            '    If (MHEquipsDemobDate(intj) > mEndDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & MHEquipsNames(intj) & "," & _
                            '           MHEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If
                            If DateValue(MHEquipsDemobDate(intj)) < DateValue(MHEquipsMobDate(intj)) Or _
                                DateValue(MHEquipsDemobDate(intj)) > mEndDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & MHEquipsNames(intj) & "," & _
                                   MHEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            If (MHEquipsQty(intj) = 0) Then   'Or Len(Trim(conce)) = 0) Then
                                MsgBox("Quantity for " & mcategories(intI) & "," & MHEquipsNames(intj) & "," & _
                                   MHEquipsCapacity(intj) & "," & mMakeModel & " is missing", MsgBoxStyle.Critical, "Error in Qty Entry")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next intj

                    For intj = 1 To CategoryItems(intI) - 2
                        If (MHEquipsChkd(intj) = 1 And MHEquipsChkd(intj - 1) = 1) Then
                            If (MHEquipsNames(intj) = MHEquipsNames(intj - 1) And _
                                MHEquipsCapacity(intj) = MHEquipsCapacity(intj - 1) And _
                                MHEquipsMake(intj) = MHEquipsMake(intj - 1) And _
                                MHEquipsModel(intj) = MHEquipsModel(intj - 1) And _
                                MHEquipsMobDate(intj) = MHEquipsMobDate(intj - 1) And _
                                MHEquipsDemobDate(intj) = MHEquipsDemobDate(intj - 1) And _
                                MHEquipsDepPerc(intj) = MHEquipsDepPerc(intj - 1)) Then
                                MsgBox("Items " & intj - 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        ElseIf (MHEquipsChkd(intj) = 1 And MHEquipsChkd(intj + 1) = 1) Then
                            If (MHEquipsNames(intj) = MHEquipsNames(intj + 1) And _
                                MHEquipsCapacity(intj) = MHEquipsCapacity(intj + 1) And _
                                MHEquipsMake(intj) = MHEquipsMake(intj + 1) And _
                                MHEquipsModel(intj) = MHEquipsModel(intj + 1) And _
                                MHEquipsMobDate(intj) = MHEquipsMobDate(intj + 1) And _
                                MHEquipsDemobDate(intj) = MHEquipsDemobDate(intj + 1) And _
                                MHEquipsDepPerc(intj) = MHEquipsDepPerc(intj + 1)) Then
                                MsgBox("Items " & intj + 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next
                Case 5
                    For intj = 0 To CategoryItems(intI) - 1
                        If NCEquipsChkd(intj) = 1 Then
                            mMakeModel = NCEquipsMake(intj) & "/" & NCEquipsModel(intj)

                            'If (NCEquipsMobDate(intj) > NCEquipsDemobDate(intj)) Then
                            '    If (NCEquipsMobDate(intj) < mStartDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & NCEquipsNames(intj) & "," & _
                            '           NCEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If
                            If DateValue(NCEquipsMobDate(intj)) > DateValue(NCEquipsDemobDate(intj)) Or _
                                DateValue(NCEquipsMobDate(intj)) < mStartDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & NCEquipsNames(intj) & "," & _
                                   NCEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            'If (NCEquipsDemobDate(intj) < NCEquipsMobDate(intj)) Then
                            '    If (NCEquipsDemobDate(intj) > mEndDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & NCEquipsNames(intj) & "," & _
                            '           NCEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If
                            If DateValue(NCEquipsDemobDate(intj)) < DateValue(NCEquipsMobDate(intj)) Or _
                                DateValue(NCEquipsDemobDate(intj)) > mEndDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & NCEquipsNames(intj) & "," & _
                                   NCEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            If (NCEquipsQty(intj) = 0) Then   'Or Len(Trim(conce)) = 0) Then
                                MsgBox("Quantity for " & mcategories(intI) & "," & NCEquipsNames(intj) & "," & _
                                   NCEquipsCapacity(intj) & "," & mMakeModel & " is missing", MsgBoxStyle.Critical, "Error in Qty Entry")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next intj

                    For intj = 1 To CategoryItems(intI) - 2
                        If (NCEquipsChkd(intj) = 1 And NCEquipsChkd(intj - 1) = 1) Then
                            If (NCEquipsNames(intj) = NCEquipsNames(intj - 1) And _
                                NCEquipsCapacity(intj) = NCEquipsCapacity(intj - 1) And _
                                NCEquipsMake(intj) = NCEquipsMake(intj - 1) And _
                                NCEquipsModel(intj) = NCEquipsModel(intj - 1) And _
                                NCEquipsMobDate(intj) = NCEquipsMobDate(intj - 1) And _
                                NCEquipsDemobDate(intj) = NCEquipsDemobDate(intj - 1) And _
                                NCEquipsDepPerc(intj) = NCEquipsDepPerc(intj - 1)) Then
                                MsgBox("Items " & intj - 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        ElseIf (NCEquipsChkd(intj) = 1 And NCEquipsChkd(intj + 1) = 1) Then
                            If (NCEquipsNames(intj) = NCEquipsNames(intj + 1) And _
                                NCEquipsCapacity(intj) = NCEquipsCapacity(intj + 1) And _
                                NCEquipsMake(intj) = NCEquipsMake(intj + 1) And _
                                NCEquipsModel(intj) = NCEquipsModel(intj + 1) And _
                                NCEquipsMobDate(intj) = NCEquipsMobDate(intj + 1) And _
                                NCEquipsDemobDate(intj) = NCEquipsDemobDate(intj + 1) And _
                                NCEquipsDepPerc(intj) = NCEquipsDepPerc(intj + 1)) Then
                                MsgBox("Items " & intj + 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next
                Case 6
                    For intj = 0 To CategoryItems(intI) - 1
                        If majOthersEquipsChkd(intj) = 1 Then
                            mMakeModel = majOthersEquipsMake(intj) & "/" & majOthersEquipsModel(intj)
                            'If (majOthersEquipsMobDate(intj) > majOthersEquipsDemobDate(intj)) Then
                            '    If (majOthersEquipsMobDate(intj) < mStartDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & majOthersEquipsNames(intj) & "," & _
                            '           majOthersEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If
                            If DateValue(majOthersEquipsMobDate(intj)) > DateValue(majOthersEquipsDemobDate(intj)) Or _
                                DateValue(majOthersEquipsMobDate(intj)) < mStartDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & majOthersEquipsNames(intj) & "," & _
                                   majOthersEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            'If (majOthersEquipsDemobDate(intj) < majOthersEquipsMobDate(intj)) Then
                            '    If (majOthersEquipsDemobDate(intj) > mEndDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & majOthersEquipsNames(intj) & "," & _
                            '           majOthersEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If

                            If DateValue(majOthersEquipsDemobDate(intj)) < DateValue(majOthersEquipsMobDate(intj)) Or _
                                DateValue(majOthersEquipsDemobDate(intj)) > mEndDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & majOthersEquipsNames(intj) & "," & _
                                   majOthersEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            If (majOthersEquipsQty(intj) = 0) Then   'Or Len(Trim(conce)) = 0) Then
                                MsgBox("Quantity for " & mcategories(intI) & "," & majOthersEquipsNames(intj) & "," & _
                                   majOthersEquipsCapacity(intj) & "," & mMakeModel & " is missing", MsgBoxStyle.Critical, "Error in Qty Entry")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next intj

                    For intj = 1 To CategoryItems(intI) - 2
                        If (majOthersEquipsChkd(intj) = 1 And majOthersEquipsChkd(intj - 1) = 1) Then
                            If (majOthersEquipsNames(intj) = majOthersEquipsNames(intj - 1) And _
                                majOthersEquipsCapacity(intj) = majOthersEquipsCapacity(intj - 1) And _
                                majOthersEquipsMake(intj) = majOthersEquipsMake(intj - 1) And _
                                majOthersEquipsModel(intj) = majOthersEquipsModel(intj - 1) And _
                                majOthersEquipsMobDate(intj) = majOthersEquipsMobDate(intj - 1) And _
                                majOthersEquipsDemobDate(intj) = majOthersEquipsDemobDate(intj - 1) And _
                                majOthersEquipsDepPerc(intj) = majOthersEquipsDepPerc(intj - 1)) Then
                                MsgBox("Items " & intj - 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        ElseIf (majOthersEquipsChkd(intj) = 1 And majOthersEquipsChkd(intj + 1) = 1) Then
                            If (majOthersEquipsNames(intj) = majOthersEquipsNames(intj + 1) And _
                                majOthersEquipsCapacity(intj) = majOthersEquipsCapacity(intj + 1) And _
                                majOthersEquipsMake(intj) = majOthersEquipsMake(intj + 1) And _
                                majOthersEquipsModel(intj) = majOthersEquipsModel(intj + 1) And _
                                majOthersEquipsMobDate(intj) = majOthersEquipsMobDate(intj + 1) And _
                                majOthersEquipsDemobDate(intj) = majOthersEquipsDemobDate(intj + 1) And _
                                majOthersEquipsDepPerc(intj) = majOthersEquipsDepPerc(intj + 1)) Then
                                MsgBox("Items " & intj + 1 & " and " & intj & " in " & mcategories(intI) & "Cataegory are duplicate entries. " & _
                                    "Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next
                Case 7
                    For intj = 0 To CategoryItems(intI) - 1
                        If minorEquipsChkd(intj) = 1 Then
                            mMakeModel = minorEquipsMake(intj) & "/" & minorEquipsModel(intj)
                            'If (minorEquipsMobDate(intj) > minorEquipsDemobDate(intj)) Then
                            '    If (minorEquipsMobDate(intj) < mStartDate) Then '
                            '        Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & minorEquipsNames(intj) & "," & _
                            '           minorEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If

                            If DateValue(minorEquipsMobDate(intj)) > DateValue(minorEquipsDemobDate(intj)) Or _
                                DateValue(minorEquipsMobDate(intj)) < mStartDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & minorEquipsNames(intj) & "," & _
                                   minorEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            'If (minorEquipsDemobDate(intj) < minorEquipsMobDate(intj)) Then
                            '    If (minorEquipsDemobDate(intj) > mEndDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & minorEquipsNames(intj) & "," & _
                            '           minorEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If

                            If DateValue(minorEquipsDemobDate(intj)) < DateValue(minorEquipsMobDate(intj)) Or _
                                DateValue(minorEquipsDemobDate(intj)) > mEndDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & mcategories(intI) & "," & minorEquipsNames(intj) & "," & _
                                   minorEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            If (minorEquipsQty(intj) = 0) Then   'Or Len(Trim(conce)) = 0) Then
                                MsgBox("Quantity for " & mcategories(intI) & "," & minorEquipsNames(intj) & "," & _
                                   minorEquipsCapacity(intj) & "," & mMakeModel & " is missing", MsgBoxStyle.Critical, "Error in Qty Entry")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next intj

                    For intj = 1 To CategoryItems(intI) - 2
                        If (minorEquipsChkd(intj) = 1 And minorEquipsChkd(intj - 1) = 1) Then
                            If (mcategories(intI) = mcategories(intI - 1) And _
                                minorEquipsNames(intj) = minorEquipsNames(intj - 1) And _
                                minorEquipsCapacity(intj) = minorEquipsCapacity(intj - 1) And _
                                minorEquipsMake(intj) = minorEquipsMake(intj - 1) And _
                                minorEquipsModel(intj) = minorEquipsModel(intj - 1) And _
                                minorEquipsMobDate(intj) = minorEquipsMobDate(intj - 1) And _
                                minorEquipsDemobDate(intj) = minorEquipsDemobDate(intj - 1) And _
                                minorEquipsDepPerc(intj) = minorEquipsDepPerc(intj - 1) And _
                                minorIsNewMC(intj) = minorIsNewMC(intj - 1)) Then
                                MsgBox("Items " & intj - 1 & " and " & intj & " in " & mcategories(intI) & " category are duplicate entries. Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        ElseIf (minorEquipsChkd(intj) = 1 And minorEquipsChkd(intj + 1) = 1) Then
                            If (mcategories(intI) = mcategories(intI + 1) And _
                                minorEquipsNames(intj) = minorEquipsNames(intj + 1) And _
                                minorEquipsCapacity(intj) = minorEquipsCapacity(intj + 1) And _
                                minorEquipsMake(intj) = minorEquipsMake(intj + 1) And _
                                minorEquipsModel(intj) = minorEquipsModel(intj + 1) And _
                                minorEquipsMobDate(intj) = minorEquipsMobDate(intj + 1) And _
                                minorEquipsDemobDate(intj) = minorEquipsDemobDate(intj + 1) And _
                                minorEquipsDepPerc(intj) = minorEquipsDepPerc(intj + 1) And _
                                minorIsNewMC(intj) = minorIsNewMC(intj + 1)) Then
                                MsgBox("Items " & intj + 1 & " and " & intj & " in " & mcategories(intI) & " category are duplicate entries. Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next
                Case 8
                    For intj = 0 To CategoryItems(intI) - 1
                        If hiredEquipsChkd(intj) = 1 Then
                            mMakeModel = hiredEquipsMake(intj) & "/" & hiredEquipsModel(intj)
                            'If (hiredEquipsMobDate(intj) > hiredEquipsDemobDate(intj)) Then
                            '    If (hiredEquipsMobDate(intj) < mStartDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & hiredCategoryNames(intj) & "," & hiredEquipsNames(intj) & "," & _
                            '           hiredEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If

                            If DateValue(hiredEquipsMobDate(intj)) > DateValue(hiredEquipsDemobDate(intj)) Or _
                                DateValue(hiredEquipsMobDate(intj)) < mStartDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & hiredCategoryNames(intj) & "," & hiredEquipsNames(intj) & "," & _
                                   hiredEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            'If (hiredEquipsDemobDate(intj) < hiredEquipsMobDate(intj)) Then
                            '    If (hiredEquipsDemobDate(intj) > mEndDate) Then
                            '        Dim msgstr = "Error in Mob. Date for " & hiredCategoryNames(intj) & "," & hiredEquipsNames(intj) & "," & _
                            '           hiredEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                            '        "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                            '          "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                            '        MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                            '        ValidateForSave = False
                            '        Exit Function
                            '    End If

                            If DateValue(hiredEquipsDemobDate(intj)) < DateValue(hiredEquipsMobDate(intj)) Or _
                                DateValue(hiredEquipsDemobDate(intj)) > mEndDate.Date Then
                                Dim msgstr = "Error in Mob. Date for " & hiredCategoryNames(intj) & "," & hiredEquipsNames(intj) & "," & _
                                   hiredEquipsCapacity(intj) & "," & mMakeModel & vbNewLine & _
                                "DeMobilisation Date must be less than Mobilisation Date" & vbNewLine & _
                                  "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                                MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                                ValidateForSave = False
                                Exit Function
                            End If

                            If (hiredEquipsQty(intj) = 0) Then   'Or Len(Trim(conce)) = 0) Then
                                MsgBox("Quantity for " & hiredCategoryNames(intj) & "," & hiredEquipsNames(intj) & "," & _
                                   hiredEquipsCapacity(intj) & "," & mMakeModel & " is missing", MsgBoxStyle.Critical, "Error in Qty Entry")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next intj
                    For intj = 1 To CategoryItems(intI) - 2
                        If (hiredEquipsChkd(intj) = 1 And hiredEquipsChkd(intj - 1) = 1) Then
                            If (mcategories(intI) = mcategories(intI - 1) And _
                             hiredEquipsNames(intj) = hiredEquipsNames(intj - 1) And _
                             hiredEquipsCapacity(intj) = hiredEquipsCapacity(intj - 1) And _
                             hiredEquipsMake(intj) = hiredEquipsMake(intj - 1) And _
                             hiredEquipsModel(intj) = hiredEquipsModel(intj - 1) And _
                             hiredEquipsMobDate(intj) = hiredEquipsMobDate(intj - 1) And _
                             hiredEquipsDemobDate(intj) = hiredEquipsDemobDate(intj - 1) And _
                             hiredEquipsDepPerc(intj) = hiredEquipsDepPerc(intj - 1)) Then
                                MsgBox("Items " & intj - 1 & " and " & intj & " in " & mcategories(intI) & " category are duplicate entries. Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        ElseIf (hiredEquipsChkd(intj) = 1 And hiredEquipsChkd(intj + 1) = 1) Then
                            If (mcategories(intI) = mcategories(intI + 1) And _
                             hiredEquipsNames(intj) = hiredEquipsNames(intj + 1) And _
                             hiredEquipsCapacity(intj) = hiredEquipsCapacity(intj + 1) And _
                             hiredEquipsMake(intj) = hiredEquipsMake(intj + 1) And _
                             hiredEquipsModel(intj) = hiredEquipsModel(intj + 1) And _
                             hiredEquipsMobDate(intj) = hiredEquipsMobDate(intj + 1) And _
                             hiredEquipsDemobDate(intj) = hiredEquipsDemobDate(intj + 1) And _
                             hiredEquipsDepPerc(intj) = hiredEquipsDepPerc(intj + 1)) Then
                                MsgBox("Items " & intj + 1 & " and " & intj & " in " & mcategories(intI) & " category are duplicate entries. Please Check. Data not saved")
                                ValidateForSave = False
                                Exit Function
                            End If
                        End If
                    Next
            End Select
        Next
    End Function
    Private Function TestValidity(ByVal Fromval As Integer, ByVal Toval As Integer) As Boolean
        Dim intI As Integer, intj As Integer
        TestValidity = True
        intj = 1

        For intI = Fromval - 1 To Fromval + Toval - 2
            If Checkboxes(intI).Checked Then
                If (MobdatePickers(intI).Value.Date > DemobDatePickers(intI).Value.Date) Or _
                    (MobdatePickers(intI).Value.Date < mStartDate) Then
                    Dim msgstr = "Error in Mob. Date for " & CategoryTextBoxes(intI).Text & "," & EquipNameTextBoxes(intI).Text & "," & _
                       CapacityTextBoxes(intI).Text & "," & MakeModelTextBoxes(intI).Text & vbNewLine & _
                    "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                      "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                    MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                    MobdatePickers(intI).Value = mStartDate.Date
                    MobdatePickers(intI).Focus()
                    TestValidity = False
                    Button1.Enabled = True
                    Exit Function
                End If

                If (DemobDatePickers(intI).Value.Date < MobdatePickers(intI).Value.Date) Or _
                    (DemobDatePickers(intI).Value.Date > mEndDate) Then
                    Dim msgstr = "Error in Demob. Date for " & CategoryTextBoxes(intI).Text & "," & EquipNameTextBoxes(intI).Text & "," & _
                       CapacityTextBoxes(intI).Text & "," & MakeModelTextBoxes(intI).Text & vbNewLine & _
                    "Mobilisation Date must be less than De-Mobilisation Date" & vbNewLine & _
                      "And mobilisation and demobilisation dates must fall within Projet Start and End dates"
                    MsgBox(msgstr, MsgBoxStyle.Critical, "Error in Entry")
                    DemobDatePickers(intI).Value = mStartDate.Date
                    DemobDatePickers(intI).Focus()
                    TestValidity = False
                    Button1.Enabled = True
                    Exit Function
                End If

                If (Val(QtyTextBoxes(intI).Text) = 0 Or Len(Trim(QtyTextBoxes(intI).Text)) = 0) Then
                    MsgBox("Quantity for " & CategoryTextBoxes(intI).Text & "," & EquipNameTextBoxes(intI).Text & "," & _
                       CapacityTextBoxes(intI).Text & "," & MakeModelTextBoxes(intI).Text & " is missing", MsgBoxStyle.Critical, "Error in Qty Entry")
                    QtyTextBoxes(intI).Text = 1
                    QtyTextBoxes(intI).Focus()
                    TestValidity = False
                    Button1.Enabled = True
                    Exit Function
                End If

                For intj = 1 To CategoryTextBoxes.Count - 1
                    If (Checkboxes(intj).Checked And Checkboxes(intj - 1).Checked) Then
                        If (CategoryTextBoxes(intj).Text = CategoryTextBoxes(intj - 1).Text And _
                            EquipNameTextBoxes(intj).Text = EquipNameTextBoxes(intj - 1).Text And _
                            CapacityTextBoxes(intj).Text = CapacityTextBoxes(intj - 1).Text And _
                            MakeModelTextBoxes(intj).Text = MakeModelTextBoxes(intj - 1).Text And _
                            MobdatePickers(intj).Value.Date = MobdatePickers(intj - 1).Value.Date And _
                            DemobDatePickers(intj).Value.Date = DemobDatePickers(intj - 1).Value.Date And _
                            DepPercComboboxes(intj).Text = DepPercComboboxes(intj - 1).Text) Then
                            MsgBox("Items " & intj - 1 & " and " & intj & " are duplicate entries. Please Check. Data not saved")
                            TestValidity = False
                            Button1.Enabled = True
                            Exit Function
                        End If
                    End If
                Next
                For intj = 0 To CategoryTextBoxes.Count - 2
                    If (Checkboxes(intj).Checked And Checkboxes(intj + 1).Checked) Then
                        If (CategoryTextBoxes(intj).Text = CategoryTextBoxes(intj + 1).Text And _
                                EquipNameTextBoxes(intj).Text = EquipNameTextBoxes(intj + 1).Text And _
                                CapacityTextBoxes(intj).Text = CapacityTextBoxes(intj + 1).Text And _
                                MakeModelTextBoxes(intj).Text = MakeModelTextBoxes(intj + 1).Text And _
                                MobdatePickers(intj).Value.Date = MobdatePickers(intj + 1).Value.Date And _
                                DemobDatePickers(intj).Value.Date = DemobDatePickers(intj + 1).Value.Date And _
                                DepPercComboboxes(intj).Text = DepPercComboboxes(intj + 1).Text) Then
                            MsgBox("Items " & intj & " and " & intj + 1 & " are duplicate entries. Please Check. Data not saved")
                            TestValidity = False
                            Button1.Enabled = True
                            Exit Function
                        End If
                    End If
                Next
                'End If
            End If
        Next
    End Function
    Private Sub WriteToConcEquipsArray(ByVal mitems As Integer)
        Dim intI As Integer
        ReDim concEquipsNames(mitems - 1)
        ReDim concEquipsCapacity(mitems - 1)
        ReDim concEquipsMake(mitems - 1)
        ReDim concEquipsModel(mitems - 1)
        ReDim concEquipsMobDate(mitems - 1)
        ReDim concEquipsDemobDate(mitems - 1)
        ReDim concEquipsQty(mitems - 1)
        ReDim concEquipsChkd(mitems - 1)
        ReDim concEquipsHPM(mitems - 1)
        ReDim concEquipsDepPerc(mitems - 1)
        ReDim concEquipsShifts(mitems - 1)
        ReDim concEquipsConcQty(mitems - 1)

        For intI = 0 To mitems - 1
            concEquipsNames(intI) = EquipNameTextBoxes(intI).Text
            concEquipsCapacity(intI) = CapacityTextBoxes(intI).Text
            concEquipsMake(intI) = Mid(MakeModelTextBoxes(intI).Text, 1, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) - 2)
            concEquipsModel(intI) = Mid(MakeModelTextBoxes(intI).Text, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) + 2)
            concEquipsMobDate(intI) = MobdatePickers(intI).Value.Date
            concEquipsDemobDate(intI) = DemobDatePickers(intI).Value.Date
            concEquipsQty(intI) = Val(QtyTextBoxes(intI).Text)
            concEquipsHPM(intI) = Val(HrsPermonthTextBoxes(intI).Text)
            concEquipsDepPerc(intI) = Val(DepPercComboboxes(intI).Text)
            concEquipsShifts(intI) = Val(ShiftsComboboxes(intI).Text)
            'concEquipsRepValue(intI) = Val(RepValueTextBoxes(intI).Text)
            concEquipsConcQty(intI) = Val(concreteqtyTextboxes(intI).Text)
            concEquipsChkd(intI) = IIf(Checkboxes(intI).Checked, 1, 0)
        Next

    End Sub
    Private Sub WriteToConveyanceEquipsArray(ByVal mitems As Integer)
        Dim intI As Integer
        ReDim convEquipsNames(mitems - 1)
        ReDim convEquipsCapacity(mitems - 1)
        ReDim convEquipsMake(mitems - 1)
        ReDim convEquipsModel(mitems - 1)
        ReDim convEquipsMobDate(mitems - 1)
        ReDim convEquipsDemobDate(mitems - 1)
        ReDim convEquipsQty(mitems - 1)
        ReDim convEquipsChkd(mitems - 1)
        ReDim convEquipsHPM(mitems - 1)
        ReDim convEquipsDepPerc(mitems - 1)
        ReDim convEquipsShifts(mitems - 1)
        ReDim convEquipsConcQty(mitems - 1)

        For intI = 0 To mitems - 1
            convEquipsNames(intI) = EquipNameTextBoxes(intI).Text
            convEquipsCapacity(intI) = CapacityTextBoxes(intI).Text
            convEquipsMake(intI) = Mid(MakeModelTextBoxes(intI).Text, 1, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) - 2)
            convEquipsModel(intI) = Mid(MakeModelTextBoxes(intI).Text, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) + 2)
            convEquipsMobDate(intI) = MobdatePickers(intI).Value.Date
            convEquipsDemobDate(intI) = DemobDatePickers(intI).Value.Date
            convEquipsQty(intI) = Val(QtyTextBoxes(intI).Text)
            convequipsHPM(intI) = Val(HrsPermonthTextBoxes(intI).Text)
            convEquipsDepPerc(intI) = Val(DepPercComboboxes(intI).Text)
            convEquipsShifts(intI) = Val(ShiftsComboboxes(intI).Text)
            'convEquipsRepValue(intI) = Val(RepValueTextBoxes(intI).Text)
            convEquipsConcQty(intI) = Val(concreteqtyTextboxes(intI).Text)
            convEquipsChkd(intI) = IIf(Checkboxes(intI).Checked, 1, 0)
        Next
    End Sub
    Private Sub WriteToCraneEquipsArray(ByVal mitems As Integer)
        Dim intI As Integer
        ReDim craneEquipsNames(mitems - 1)
        ReDim craneEquipsCapacity(mitems - 1)
        ReDim craneEquipsMake(mitems - 1)
        ReDim craneEquipsModel(mitems - 1)
        ReDim craneEquipsMobDate(mitems - 1)
        ReDim craneEquipsDemobDate(mitems - 1)
        ReDim craneEquipsQty(mitems - 1)
        ReDim craneEquipsChkd(mitems - 1)
        ReDim craneEquipsHPM(mitems - 1)
        ReDim craneEquipsDepPerc(mitems - 1)
        ReDim craneEquipsRepValue(mitems - 1)
        ReDim craneEquipsShifts(mitems - 1)
        ReDim craneEquipsConcQty(mitems - 1)

        For intI = 0 To mitems - 1
            craneEquipsNames(intI) = EquipNameTextBoxes(intI).Text
            craneEquipsCapacity(intI) = CapacityTextBoxes(intI).Text
            craneEquipsMake(intI) = Mid(MakeModelTextBoxes(intI).Text, 1, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) - 2)
            craneEquipsModel(intI) = Mid(MakeModelTextBoxes(intI).Text, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) + 2)
            craneEquipsMobDate(intI) = MobdatePickers(intI).Value.Date
            craneEquipsDemobDate(intI) = DemobDatePickers(intI).Value.Date
            craneEquipsQty(intI) = Val(QtyTextBoxes(intI).Text)
            craneEquipsHPM(intI) = Val(HrsPermonthTextBoxes(intI).Text)
            craneEquipsDepPerc(intI) = Val(DepPercComboboxes(intI).Text)
            craneEquipsShifts(intI) = Val(ShiftsComboboxes(intI).Text)
            'craneEquipsRepValue(intI) = Val(RepValueTextBoxes(intI).Text)
            craneEquipsConcQty(intI) = Val(concreteqtyTextboxes(intI).Text)
            craneEquipsChkd(intI) = IIf(Checkboxes(intI).Checked, 1, 0)
        Next
    End Sub
    Private Sub WriteToDgsetsEquipsArray(ByVal mitems As Integer)
        Dim intI As Integer
        ReDim dgsetsEquipsNames(mitems - 1)
        ReDim dgsetsEquipsCapacity(mitems - 1)
        ReDim dgsetsEquipsMake(mitems - 1)
        ReDim dgsetsEquipsModel(mitems - 1)
        ReDim dgsetsEquipsMobDate(mitems - 1)
        ReDim dgsetsEquipsDemobDate(mitems - 1)
        ReDim dgsetsEquipsQty(mitems - 1)
        ReDim dgsetsEquipsChkd(mitems - 1)
        ReDim dgsetsEquipsHPM(mitems - 1)
        ReDim dgsetsEquipsDepPerc(mitems - 1)
        ReDim dgsetsEquipsRepValue(mitems - 1)
        ReDim dgsetsEquipsShifts(mitems - 1)
        ReDim dgsetsEquipsConcQty(mitems - 1)

        For intI = 0 To mitems - 1
            dgsetsEquipsNames(intI) = EquipNameTextBoxes(intI).Text
            dgsetsEquipsCapacity(intI) = CapacityTextBoxes(intI).Text
            dgsetsEquipsMake(intI) = Mid(MakeModelTextBoxes(intI).Text, 1, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) - 2)
            dgsetsEquipsModel(intI) = Mid(MakeModelTextBoxes(intI).Text, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) + 2)
            dgsetsEquipsMobDate(intI) = MobdatePickers(intI).Value.Date
            dgsetsEquipsDemobDate(intI) = DemobDatePickers(intI).Value.Date
            dgsetsEquipsQty(intI) = Val(QtyTextBoxes(intI).Text)
            dgsetsEquipsHPM(intI) = Val(HrsPermonthTextBoxes(intI).Text)
            dgsetsEquipsDepPerc(intI) = Val(DepPercComboboxes(intI).Text)
            dgsetsEquipsShifts(intI) = Val(ShiftsComboboxes(intI).Text)
            'dgsetsEquipsRepValue(intI) = Val(RepValueTextBoxes(intI).Text)
            dgsetsEquipsConcQty(intI) = Val(concreteqtyTextboxes(intI).Text)
            dgsetsEquipsChkd(intI) = IIf(Checkboxes(intI).Checked, 1, 0)
        Next
    End Sub
    Private Sub WriteToMHEquipsArray(ByVal mitems As Integer)
        Dim intI As Integer
        ReDim MhEquipsNames(mitems - 1)
        ReDim MhEquipsCapacity(mitems - 1)
        ReDim MhEquipsMake(mitems - 1)
        ReDim MhEquipsModel(mitems - 1)
        ReDim MhEquipsMobDate(mitems - 1)
        ReDim MhEquipsDemobDate(mitems - 1)
        ReDim MhEquipsQty(mitems - 1)
        ReDim MhEquipsChkd(mitems - 1)
        ReDim MhEquipsHPM(mitems - 1)
        ReDim MhEquipsDepPerc(mitems - 1)
        ReDim MhEquipsRepValue(mitems - 1)
        ReDim MhEquipsShifts(mitems - 1)
        ReDim MhEquipsConcQty(mitems - 1)

        For intI = 0 To mitems - 1
            MhEquipsNames(intI) = EquipNameTextBoxes(intI).Text
            MhEquipsCapacity(intI) = CapacityTextBoxes(intI).Text
            MhEquipsMake(intI) = Mid(MakeModelTextBoxes(intI).Text, 1, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) - 2)
            MhEquipsModel(intI) = Mid(MakeModelTextBoxes(intI).Text, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) + 2)
            MhEquipsMobDate(intI) = MobdatePickers(intI).Value.Date
            MhEquipsDemobDate(intI) = DemobDatePickers(intI).Value.Date
            MhEquipsQty(intI) = Val(QtyTextBoxes(intI).Text)
            MhEquipsHPM(intI) = Val(HrsPermonthTextBoxes(intI).Text)
            MhEquipsDepPerc(intI) = Val(DepPercComboboxes(intI).Text)
            MhEquipsShifts(intI) = Val(ShiftsComboboxes(intI).Text)
            MHEquipsConcQty(intI) = Val(concreteqtyTextboxes(intI).Text)
            MHEquipsChkd(intI) = IIf(Checkboxes(intI).Checked, 1, 0)
        Next
    End Sub
    Private Sub WriteToNCEquipsArray(ByVal mitems As Integer)
        Dim intI As Integer
        ReDim NCEquipsNames(mitems - 1)
        ReDim NCEquipsCapacity(mitems - 1)
        ReDim NCEquipsMake(mitems - 1)
        ReDim NCEquipsModel(mitems - 1)
        ReDim NCEquipsMobDate(mitems - 1)
        ReDim NCEquipsDemobDate(mitems - 1)
        ReDim NCEquipsQty(mitems - 1)
        ReDim NCEquipsChkd(mitems - 1)
        ReDim NCEquipsHPM(mitems - 1)
        ReDim NCEquipsDepPerc(mitems - 1)
        ReDim NCEquipsRepValue(mitems - 1)
        ReDim NCEquipsShifts(mitems - 1)
        ReDim NCEquipsConcQty(mitems - 1)

        For intI = 0 To mitems - 1
            NCEquipsNames(intI) = EquipNameTextBoxes(intI).Text
            NCEquipsCapacity(intI) = CapacityTextBoxes(intI).Text
            NCEquipsMake(intI) = Mid(MakeModelTextBoxes(intI).Text, 1, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) - 2)
            NCEquipsModel(intI) = Mid(MakeModelTextBoxes(intI).Text, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) + 2)
            NCEquipsMobDate(intI) = MobdatePickers(intI).Value.Date
            NCEquipsDemobDate(intI) = DemobDatePickers(intI).Value.Date
            NCEquipsQty(intI) = Val(QtyTextBoxes(intI).Text)
            NCEquipsHPM(intI) = Val(HrsPermonthTextBoxes(intI).Text)
            NCEquipsDepPerc(intI) = Val(DepPercComboboxes(intI).Text)
            NCEquipsShifts(intI) = Val(ShiftsComboboxes(intI).Text)
            NCEquipsConcQty(intI) = Val(concreteqtyTextboxes(intI).Text)
            NCEquipsChkd(intI) = IIf(Checkboxes(intI).Checked, 1, 0)
        Next
    End Sub
    Private Sub WriteToMajOtherEquipsArray(ByVal mitems As Integer)
        Dim intI As Integer
        ReDim majOthersEquipsNames(mitems - 1)
        ReDim majOthersEquipsCapacity(mitems - 1)
        ReDim majOthersEquipsMake(mitems - 1)
        ReDim majOthersEquipsModel(mitems - 1)
        ReDim majOthersEquipsMobDate(mitems - 1)
        ReDim majOthersEquipsDemobDate(mitems - 1)
        ReDim majOthersEquipsQty(mitems - 1)
        ReDim majOthersEquipsChkd(mitems - 1)
        ReDim majOthersEquipsHPM(mitems - 1)
        ReDim majOthersEquipsDepPerc(mitems - 1)
        ReDim majOthersEquipsRepValue(mitems - 1)
        ReDim majOthersEquipsShifts(mitems - 1)
        ReDim majOthersEquipsConcQty(mitems - 1)

        For intI = 0 To mitems - 1
            majOthersEquipsNames(intI) = EquipNameTextBoxes(intI).Text
            majOthersEquipsCapacity(intI) = CapacityTextBoxes(intI).Text
            majOthersEquipsMake(intI) = Mid(MakeModelTextBoxes(intI).Text, 1, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) - 2)
            majOthersEquipsModel(intI) = Mid(MakeModelTextBoxes(intI).Text, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) + 2)
            majOthersEquipsMobDate(intI) = MobdatePickers(intI).Value.Date
            majOthersEquipsDemobDate(intI) = DemobDatePickers(intI).Value.Date
            majOthersEquipsQty(intI) = Val(QtyTextBoxes(intI).Text)
            majOthersEquipsHPM(intI) = Val(HrsPermonthTextBoxes(intI).Text)
            majOthersEquipsDepPerc(intI) = Val(DepPercComboboxes(intI).Text)
            majOthersEquipsShifts(intI) = Val(ShiftsComboboxes(intI).Text)
            majOthersEquipsConcQty(intI) = Val(concreteqtyTextboxes(intI).Text)
            majOthersEquipsChkd(intI) = IIf(Checkboxes(intI).Checked, 1, 0)
        Next
    End Sub
    Private Sub WriteToHiredEquipsArray(ByVal mitems As Integer)
        Dim intI As Integer
        ReDim hiredCategoryNames(mitems - 1)
        ReDim hiredEquipsNames(mitems - 1)
        ReDim hiredEquipsCapacity(mitems - 1)
        ReDim hiredEquipsMake(mitems - 1)
        ReDim hiredEquipsModel(mitems - 1)
        ReDim hiredEquipsMobDate(mitems - 1)
        ReDim hiredEquipsDemobDate(mitems - 1)
        ReDim hiredEquipsQty(mitems - 1)
        ReDim hiredEquipsChkd(mitems - 1)
        ReDim hiredEquipsHPM(mitems - 1)
        ReDim hiredEquipsDepPerc(mitems - 1)
        ReDim hiredEquipsHireCharges(mitems - 1)
        ReDim hiredEquipsShifts(mitems - 1)

        For intI = 0 To mitems - 1
            hiredCategoryNames(intI) = CategoryTextBoxes(intI).Text
            hiredEquipsNames(intI) = EquipNameTextBoxes(intI).Text
            hiredEquipsCapacity(intI) = CapacityTextBoxes(intI).Text
            hiredEquipsMake(intI) = Mid(MakeModelTextBoxes(intI).Text, 1, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) - 2)
            hiredEquipsModel(intI) = Mid(MakeModelTextBoxes(intI).Text, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) + 2)
            hiredEquipsMobDate(intI) = MobdatePickers(intI).Value.Date
            hiredEquipsDemobDate(intI) = DemobDatePickers(intI).Value.Date
            hiredEquipsQty(intI) = Val(QtyTextBoxes(intI).Text)
            hiredEquipsHPM(intI) = Val(HrsPermonthTextBoxes(intI).Text)
            hiredEquipsDepPerc(intI) = Val(DepPercComboboxes(intI).Text)
            hiredEquipsShifts(intI) = Val(ShiftsComboboxes(intI).Text)
            hiredEquipsHireCharges(intI) = Val(HireChargesTextBoxes(intI).Text)
            hiredEquipsChkd(intI) = IIf(Checkboxes(intI).Checked, 1, 0)
        Next
    End Sub
    Private Sub WriteToMinorEquipsArray(ByVal mitems As Integer)
        Dim intI As Integer
        ReDim minorEquipsNames(mitems - 1)
        ReDim minorEquipsCapacity(mitems - 1)
        ReDim minorEquipsMake(mitems - 1)
        ReDim minorEquipsModel(mitems - 1)
        ReDim minorEquipsMobDate(mitems - 1)
        ReDim minorEquipsDemobDate(mitems - 1)
        ReDim minorEquipsQty(mitems - 1)
        ReDim minorEquipsChkd(mitems - 1)
        ReDim minorEquipsHPM(mitems - 1)
        ReDim minorEquipsDepPerc(mitems - 1)
        ReDim minorEquipsNewCost(mitems - 1)
        ReDim minorEquipsShifts(mitems - 1)
        ReDim minorIsNewMC(mitems - 1)

        For intI = 0 To mitems - 1
            minorEquipsNames(intI) = EquipNameTextBoxes(intI).Text
            minorEquipsCapacity(intI) = CapacityTextBoxes(intI).Text
            minorEquipsMake(intI) = Mid(MakeModelTextBoxes(intI).Text, 1, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) - 2)
            minorEquipsModel(intI) = Mid(MakeModelTextBoxes(intI).Text, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) + 2)
            minorEquipsMobDate(intI) = MobdatePickers(intI).Value.Date
            minorEquipsDemobDate(intI) = DemobDatePickers(intI).Value.Date
            minorEquipsQty(intI) = Val(QtyTextBoxes(intI).Text)
            minorEquipsHPM(intI) = Val(HrsPermonthTextBoxes(intI).Text)
            minorEquipsDepPerc(intI) = Val(DepPercComboboxes(intI).Text)
            minorEquipsShifts(intI) = Val(ShiftsComboboxes(intI).Text)
            minorEquipsNewCost(intI) = IsNewMC(intI).Checked
            minorEquipsChkd(intI) = IIf(Checkboxes(intI).Checked, 1, 0)
        Next
    End Sub
    Private Sub WriteToFixedExpArray(ByVal mitems As Integer)
        Dim intI As Integer
        ReDim fixedExpCategoryNames(mitems - 1)
        ReDim fixedExpEquipsQty(mitems - 1)
        ReDim fixedExpCost(mitems - 1)
        ReDim fixedExpAmount(mitems - 1)
        ReDim fixedExpRemarks(mitems - 1)
        ReDim fixedExpEquipsChkd(mitems - 1)
        ReDim fixedExpProjValue(mitems - 1)
        ReDim fixedExpEquipsCostPerc(mitems - 1)

        For intI = 0 To mitems - 1
            fixedExpCategoryNames(intI) = CategoryTextBoxes(intI).Text
            fixedExpEquipsQty(intI) = QtyTextBoxes(intI).Text
            fixedExpCost(intI) = CostTextBoxes(intI).Text
            fixedExpAmount(intI) = AmountTextBoxes(intI).Text
            fixedExpRemarks(intI) = RemarksTextBoxes(intI).Text
            fixedExpEquipsChkd(intI) = IIf(Checkboxes(intI).Checked, 1, 0)
            fixedExpProjValue(intI) = ClientBillingTextBoxes(intI).Text
            fixedExpEquipsCostPerc(intI) = CostPercTextBoxes(intI).Text
        Next
    End Sub
    Private Sub WriteToFixedBPExpArray(ByVal mitems As Integer)
        Dim intI As Integer
        ReDim fixedBPExpCategoryNames(mitems - 1)
        ReDim fixedBPExpEquipsQty(mitems - 1)
        ReDim fixedBPExpCost(mitems - 1)
        ReDim fixedBPExpAmount(mitems - 1)
        ReDim fixedBPExpRemarks(mitems - 1)
        ReDim fixedBPExpEquipsChkd(mitems - 1)
        ReDim fixedBPExpProjValue(mitems - 1)
        ReDim fixedBPExpEquipsCostPerc(mitems - 1)

        For intI = 0 To mitems - 1
            fixedBPExpCategoryNames(intI) = CategoryTextBoxes(intI).Text
            fixedBPExpEquipsQty(intI) = QtyTextBoxes(intI).Text
            fixedBPExpCost(intI) = CostTextBoxes(intI).Text
            fixedBPExpAmount(intI) = AmountTextBoxes(intI).Text
            fixedBPExpRemarks(intI) = RemarksTextBoxes(intI).Text
            fixedBPExpEquipsChkd(intI) = IIf(Checkboxes(intI).Checked, 1, 0)
            fixedBPExpProjValue(intI) = ClientBillingTextBoxes(intI).Text
            fixedBPExpEquipsCostPerc(intI) = CostPercTextBoxes(intI).Text
        Next
    End Sub
    Private Sub WriteToLightingEquipsArray(ByVal mitems As Integer)
        Dim intI As Integer

        ReDim lightingCategoryNames(mitems - 1)
        ReDim lightingEquipsNames(mitems - 1)
        ReDim lightingEquipsCapacity(mitems - 1)
        ReDim lightingEquipsMake(mitems - 1)
        ReDim lightingEquipsModel(mitems - 1)
        ReDim lightingEquipsMobDate(mitems - 1)
        ReDim lightingEquipsDemobDate(mitems - 1)
        ReDim lightingEquipsQty(mitems - 1)
        ReDim lightingEquipsChkd(mitems - 1)
        ReDim lightingEquipsPPU(mitems - 1)
        ReDim lightingEquipsCLPerMc(mitems - 1)
        ReDim lightingEquipsUF(mitems - 1)

        For intI = 0 To mitems - 1
            lightingCategoryNames(intI) = CategoryTextBoxes(intI).Text
            lightingEquipsNames(intI) = EquipNameTextBoxes(intI).Text
            lightingEquipsCapacity(intI) = CapacityTextBoxes(intI).Text
            lightingEquipsMake(intI) = Mid(MakeModelTextBoxes(intI).Text, 1, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) - 2)
            lightingEquipsModel(intI) = Mid(MakeModelTextBoxes(intI).Text, InStr(MakeModelTextBoxes(intI).Text, "/", CompareMethod.Text) + 2)
            lightingEquipsMobDate(intI) = MobdatePickers(intI).Value.Date
            lightingEquipsDemobDate(intI) = DemobDatePickers(intI).Value.Date
            lightingEquipsQty(intI) = Val(QtyTextBoxes(intI).Text)
            lightingEquipsChkd(intI) = IIf(Checkboxes(intI).Checked, 1, 0)
            lightingEquipsPPU(intI) = Val(PowerPerUnitTextBoxes(intI).Text)
            lightingEquipsCLPerMc(intI) = Val(ConnectLoadTextBoxes(intI).Text)
            lightingEquipsUF(intI) = Val(UtilityFactorTextBoxes(intI).Text)
        Next
    End Sub

    Private Sub SaveDataInSheet(ByVal sheetname As String, ByVal Fromval As Integer, ByVal Toval As Integer)
        Dim intI As Integer, intj As Integer
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, Querystring As String = "", mdataset As DataSet, madapter As OleDbDataAdapter
        Dim Machine As DataRow
        deleteOlddata(sheetname)
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        RowCount = 0
        intj = 1

        Me.lblmessage.Text = "Major Concreting Equipments Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()
        For intI = Fromval To Fromval + Toval - 1
            If concEquipsChkd(intI) = 1 Then
                If Category_Shortname = "Concrete_" Then
                    xlRange = xlWorksheet.Range(Category_Shortname & "ConcreteQty")
                    xlRange.Value = mConcreteQty
                End If
                xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
                If Len(Trim(xlRange.Value)) = 0 Then
                    xlRange.Value = mMainTitle1
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
                    xlRange.Value = mMainTitle2
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
                    'xlRange.Value = mMainTitle3
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
                End If

                ComputeValues(intI) '===============

                xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
                xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
                xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"

                SheetNo = getSheetNo(xlWorksheet.Name)
                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = RecordsInserted(SheetNo) + 1
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = concEquipsNames(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDrive
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = CapacityTextBoxes(intI).Text
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = MakeModelTextBoxes(intI).Text
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = QtyTextBoxes(intI).Text
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1) ''
                xlRange.Value = MobdatePickers(intI).Value.Date.ToString()
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = DemobDatePickers(intI).Value.Date.ToString()
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange = xlRange.Offset(0, 1)
                If sheetname = "Minor Eqpts" Then ' Check from here
                    xlRange.Value = MinorEquipmentCost
                Else
                    xlRange.Value = txtRepvalue
                End If
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDepreciation
                xlRange.NumberFormat = "#0.00"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = HrsPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtFuelPerHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = FuelCostperLtr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtPowerperHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = PowerCostPerUnit
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtOprCostPerMCPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtshifts
                xlRange.NumberFormat = "#.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtConsumablesPerMonth '* Val(Me.txtMajorEquipQty.Text))
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                If UCase(Category_Shortname) = UCase("Concrete_") Then
                    xlRange = xlRange.Offset(0, 4)
                    xlRange.Value = IIf(mConcreteQty = 0, 0, mConcreteQty)
                ElseIf CategoryTextBoxes(intI).Text = "Minor Equipments" Then
                    xlRange = xlRange.Offset(0, 4)
                    xlRange.Value = NewMCCost
                End If
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlWorksheet = xlWorkbook.Sheets(sheetname)
                xlWorksheet.Activate()
                getCategoryShortname(xlWorksheet)
                SheetNo = getSheetNo(sheetname)
                RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1
                If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
            End If
            InsertCommand = ""
            mcategory = CategoryTextBoxes(intI).Text
            If mcategory <> "Minor Equipments" Then
                InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
                InsertCommand = InsertCommand & "'" & EquipNameTextBoxes(intI).Text & "', "
                InsertCommand = InsertCommand & "'" & CapacityTextBoxes(intI).Text & "', "
                InsertCommand = InsertCommand & "'" & txtmake & "', "
                InsertCommand = InsertCommand & "'" & txtmodel & "', "
                InsertCommand = InsertCommand & "'" & MobdatePickers(intI).Text & "', "
                InsertCommand = InsertCommand & "'" & DemobDatePickers(intI).Text & "', "
                InsertCommand = InsertCommand & QtyTextBoxes(intI).Text & ", "
                InsertCommand = InsertCommand & IIf(Checkboxes(intI).Checked, 1, 0) & ", "
                InsertCommand = InsertCommand & Val(HrsPermonthTextBoxes(intI).Text) & ", "
                InsertCommand = InsertCommand & txtDepreciation & ", "
                InsertCommand = InsertCommand & txtRepvalue & ", "
                InsertCommand = InsertCommand & Val(ShiftsComboboxes(intI).Text) & ", "
                InsertCommand = InsertCommand & txtMaintCostperMC_PerMonth & ", "
                InsertCommand = InsertCommand & Val(concreteqtyTextboxes(intI).Text) & ", "
                InsertCommand = InsertCommand & "'" & txtDrive & "', "

                Querystring = "Select * from MajorEquipments where categoryname = '" & mcategory & "' and "
                Querystring = Querystring & "EquipmentName = '" & EquipNameTextBoxes(intI).Text & "'  and "
                Querystring = Querystring & "Capacity = '" & CapacityTextBoxes(intI).Text & "'  and "
                Querystring = Querystring & "Make = '" & txtmake & "'  and "
                Querystring = Querystring & "Model = '" & txtmodel & "'"

                madapter = New OleDbDataAdapter(Querystring, moledbConnection)
                mDataSet = New DataSet
                madapter.Fill(mDataSet, "MajorEqupts")
                If mDataSet.Tables("MajorEqupts").Rows.Count > 0 Then
                    For Each Machine In mDataSet.Tables("MajorEqupts").Rows
                        InsertCommand = InsertCommand & Val(Machine("PowerPerUnit(HP)").ToString()) & ", "
                        InsertCommand = InsertCommand & Val(Machine("ConnectedLoadPerMC").ToString()) & ", "
                        InsertCommand = InsertCommand & Val(Machine("UtilityFactor").ToString()) & ")"
                    Next
                Else
                    InsertCommand = InsertCommand & Val("0") & ", "
                    InsertCommand = InsertCommand & Val("0") & ", "
                    InsertCommand = InsertCommand & Val("0") & ")"
                End If
                mDataSet = Nothing
                madapter = Nothing
                Querystring = ""

            Else
                InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
                InsertCommand = InsertCommand & "'" & EquipNameTextBoxes(intI).Text & "', "
                InsertCommand = InsertCommand & "'" & CapacityTextBoxes(intI).Text & "', "
                InsertCommand = InsertCommand & "'" & txtmake & "', "
                InsertCommand = InsertCommand & "'" & txtmodel & "', "
                InsertCommand = InsertCommand & "'" & MobdatePickers(intI).Text & "', "
                InsertCommand = InsertCommand & "'" & DemobDatePickers(intI).Text & "', "
                InsertCommand = InsertCommand & QtyTextBoxes(intI).Text & ", "
                InsertCommand = InsertCommand & IIf(Checkboxes(intI).Checked, 1, 0) & ", "
                InsertCommand = InsertCommand & Val(HrsPermonthTextBoxes(intI).Text) & ", "
                InsertCommand = InsertCommand & RAndMPercentage & ", "
                InsertCommand = InsertCommand & Val(ShiftsComboboxes(intI).Text) & ", "
                InsertCommand = InsertCommand & MinorEquipmentCost & ", "
                InsertCommand = InsertCommand & IsNewMC(intI).Checked & ", "
                InsertCommand = InsertCommand & "'" & txtDrive & "', "

                Querystring = "Select * from MinorEquipments where categoryname = '" & mcategory & "' and "
                Querystring = Querystring & "EquipmentName = '" & EquipNameTextBoxes(intI).Text & "' and "
                Querystring = Querystring & "Capacity = '" & CapacityTextBoxes(intI).Text & "' and "
                Querystring = Querystring & "Make = '" & txtmake & "' and "
                Querystring = Querystring & "Model = '" & txtmodel & "'"

                madapter = New OleDbDataAdapter(Querystring, moledbConnection)
                mDataSet = New DataSet
                madapter.Fill(mDataSet, "MinorEqupts")
                If mDataSet.Tables("MinorEqupts").Rows.Count > 0 Then
                    For Each Machine In mDataSet.Tables("MinorEqupts").Rows
                        InsertCommand = InsertCommand & Val(Machine("PowerPerUnit(HP)").ToString()) & ", "
                        InsertCommand = InsertCommand & Val(Machine("ConnectedLoadPerMC").ToString()) & ", "
                        InsertCommand = InsertCommand & Val(Machine("UtilityFactor").ToString()) & ")"
                    Next
                Else
                    InsertCommand = InsertCommand & Val("0") & ", "
                    InsertCommand = InsertCommand & Val("0") & ", "
                    InsertCommand = InsertCommand & Val("0") & ")"
                End If

                mDataSet = Nothing
                madapter = Nothing
                Querystring = ""
            End If
            Try
                If (moledbConnection1.State.ToString().Equals("Closed")) Then
                    moledbConnection1.Open()
                End If
                moleDBInsertCommand = New OleDbCommand
                moleDBInsertCommand.CommandType = CommandType.Text
                moleDBInsertCommand.CommandText = InsertCommand
                moleDBInsertCommand.Connection = moledbConnection1
                moleDBInsertCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            Finally
                moleDBInsertCommand = Nothing
            End Try
            intj = intj + 1
        Next
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = RecordsInserted(getSheetNo(sheetname))
    End Sub
    Private Sub SaveDataInConcSheet(ByVal sheetname As String, ByVal Fromval As Integer, ByVal Toval As Integer)
        Dim intI As Integer, intj As Integer
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, mcommand As OleDbCommand, Querystring As String = ""
        Dim mdataset As DataSet, madapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim AllValid As Boolean
        Dim currentcategory As String 'txtpurchval As Long
        Dim msgstring As String, strsql As String
        Dim PPU As Single, ConnLoad As Single, UF As Single

        deleteOlddata(sheetname)
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        RowCount = 0

        Me.lblmessage.Text = "Major (Concreting) Equipments Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()


        For intI = Fromval To Fromval + Toval
            AllValid = True
            msgstring = "The following were not entered/selected" & vbNewLine
            If concEquipsChkd(intI) = 1 Then
                If (concEquipsQty(intI) = 0) Then
                    msgstring = msgstring & "*** Equipment Quantity not entered for the " & intI & _
                        IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                    AllValid = False
                End If
            End If
            If (concEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If (concEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If Not AllValid Then
                DataSaved = False
                Exit Sub
            End If

            txtMonths = System.Math.Round((concEquipsDemobDate(intI).Date - concEquipsMobDate(intI).Date).Days / 30, 0)
            txtmake = concEquipsMake(intI)
            txtmodel = concEquipsModel(intI)
            strsql = "categoryname = 'Concreting' And EquipmentName = '" & concEquipsNames(intI) & "' and " & _
               "Make ='" & txtmake & "' and Model ='" & txtmodel & "' and Capacity ='" & concEquipsCapacity(intI) & "'"
            strStatement = "Select * from MajorEquipments where " & strsql
            If moledbConnection Is Nothing Then
                strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
                moledbConnection = New OleDbConnection(strConnection)
            End If
            If (moledbConnection.State.ToString().Equals("Closed")) Then
                moledbConnection.Open()
            End If
            mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
            mdataset = New DataSet()

            mOledbDataAdapter.Fill(mdataset, "Equipments")
            For Each Machine In mdataset.Tables("Equipments").Rows
                txtRepvalue = Val(Machine("Repvalue").ToString())    'concEquipsRepValue(intI)
                txtDepreciation = concEquipsDepPerc(intI)
                HrsPerMonth = concEquipsHPM(intI)
                txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
                txtPowerperHr = Val(Machine("Power_PerHour").ToString())
                txtOprCostPerMCPerMonth = Val(Machine("OperatorCost_PerMonth").ToString())
                If concEquipsDepPerc(intI) = 2.75 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_275").ToString())
                ElseIf concEquipsDepPerc(intI) = 1.25 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_125").ToString())
                Else
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPerc_050").ToString())
                End If
                If (txtFuelPerHr = 0) Then
                    txtDrive = "Electrical"
                Else
                    txtDrive = "Fuel (Diesel)"
                End If
                txtshifts = concEquipsShifts(intI)
                txtMaintCostperMC_PerMonth = txtRepvalue * (txtMaintPercPerMC_PerMonth / 100)
                PPU = Val(Machine("PowerPerUnit(HP)").ToString())
                ConnLoad = Val(Machine("ConnectedLoadPerMC").ToString())
                UF = Val(Machine("UtilityFactor").ToString())
            Next
            txtFuelperUnitPerMonth = txtFuelPerHr * HrsPerMonth
            txtFuelCostPerMonth = txtFuelperUnitPerMonth * concEquipsQty(intI) * FuelCostperLtr
            txtFuelCostProject = txtFuelCostPerMonth * txtMonths
            txtPowerCostperMonth = txtPowerperHr * HrsPerMonth * concEquipsQty(intI) * PowerCostPerUnit
            txtPowerCostProject = txtPowerCostperMonth * txtMonths
            txtOprCostPerMonth = txtOprCostPerMCPerMonth * concEquipsQty(intI) * txtshifts
            txtOprCostProject = txtOprCostPerMonth * txtMonths
            txtConsumablesPerMonth = txtMaintCostperMC_PerMonth * concEquipsQty(intI)
            txtConsumablesProject = txtConsumablesPerMonth * txtMonths
            OperatingCost_MajorEquips = txtHirecharges + txtFuelCostProject + txtPowerCostProject + _
                  txtOprCostProject + txtConsumablesProject
            mOledbDataAdapter = Nothing
            mdataset = Nothing

            If concEquipsChkd(intI) = 1 Then
                xlRange = xlWorksheet.Range(Category_Shortname & "ConcreteQty")
                If Len(Trim(xlRange.Value)) = 0 Then
                    xlRange.Value = mConcreteQty
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
                    'If Len(Trim(xlRange.Value)) = 0 Then
                    xlRange.Value = mMainTitle1
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
                    xlRange.Value = mMainTitle2
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
                    'xlRange.Value = mMainTitle3
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
                    xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
                    xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                    xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
                End If

                SheetNo = getSheetNo(xlWorksheet.Name)
                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = RecordsInserted(SheetNo) + 1
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = concEquipsNames(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDrive
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = concEquipsCapacity(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = concEquipsMake(intI) & " / " & concEquipsModel(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = concEquipsQty(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1) ''
                xlRange.Value = concEquipsMobDate(intI).Date.ToString()
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = concEquipsDemobDate(intI).Date.ToString()
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtRepvalue
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDepreciation
                xlRange.NumberFormat = "#0.00"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = HrsPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtFuelPerHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = FuelCostperLtr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtPowerperHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = PowerCostPerUnit
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtOprCostPerMCPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtshifts
                xlRange.NumberFormat = "#.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtConsumablesPerMonth '* Val(Me.txtMajorEquipQty.Text))
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = mConcreteQty
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlWorksheet = xlWorkbook.Sheets(sheetname)
                xlWorksheet.Activate()
                getCategoryShortname(xlWorksheet)
                SheetNo = getSheetNo(sheetname)
                RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1
                If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
            End If
            InsertCommand = ""
            mcategory = "Concreting"
            InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
            InsertCommand = InsertCommand & "'" & concEquipsNames(intI) & "', "
            InsertCommand = InsertCommand & "'" & concEquipsCapacity(intI) & "', "
            InsertCommand = InsertCommand & "'" & txtmake & "', "
            InsertCommand = InsertCommand & "'" & txtmodel & "', "
            InsertCommand = InsertCommand & "'" & concEquipsMobDate(intI) & "', "
            InsertCommand = InsertCommand & "'" & concEquipsDemobDate(intI) & "', "
            InsertCommand = InsertCommand & concEquipsQty(intI) & ", "
            InsertCommand = InsertCommand & concEquipsChkd(intI) & ", "
            InsertCommand = InsertCommand & concEquipsHPM(intI) & ", "
            InsertCommand = InsertCommand & txtDepreciation & ", "
            InsertCommand = InsertCommand & txtRepvalue & ", "
            InsertCommand = InsertCommand & concEquipsShifts(intI) & ", "
            InsertCommand = InsertCommand & txtMaintCostperMC_PerMonth & ", "
            InsertCommand = InsertCommand & concEquipsConcQty(intI) & ", "
            InsertCommand = InsertCommand & "'" & txtDrive & "', "
            InsertCommand = InsertCommand & PPU & ", "
            InsertCommand = InsertCommand & ConnLoad & ", "
            InsertCommand = InsertCommand & UF & ")"

            Try
                If (moledbConnection1.State.ToString().Equals("Closed")) Then
                    moledbConnection1.Open()
                End If
                moleDBInsertCommand = New OleDbCommand
                moleDBInsertCommand.CommandType = CommandType.Text
                moleDBInsertCommand.CommandText = InsertCommand
                moleDBInsertCommand.Connection = moledbConnection1
                moleDBInsertCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            Finally
                moleDBInsertCommand = Nothing
            End Try
        Next
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = RecordsInserted(getSheetNo(sheetname))
    End Sub
    Private Sub SaveDataInConvSheet(ByVal sheetname As String, ByVal Fromval As Integer, ByVal Toval As Integer)
        Dim intI As Integer, intj As Integer
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, mcommand As OleDbCommand, Querystring As String = ""
        Dim mdataset As DataSet, madapter As OleDbDataAdapter
        Dim Machine As DataRow, AllValid As Boolean
        Dim currentcategory As String, msgstring As String, strsql As String
        Dim PPU As Single, Connload As Single, UF As Single

        deleteOlddata(sheetname)
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        RowCount = 0

        Me.lblmessage.Text = "Major (Conveyance) Equipments Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()


        For intI = Fromval To Fromval + Toval
            AllValid = True
            msgstring = "The following were not entered/selected" & vbNewLine
            If convEquipsChkd(intI) = 1 Then
                If convEquipsQty(intI) = 0 Then
                    msgstring = msgstring & "*** Equipment Quantity not entered for the " & intI & _
                        IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                    AllValid = False
                End If
            End If
            If (convEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If (convEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If Not AllValid Then
                DataSaved = False
                Exit Sub
            End If
            txtMonths = System.Math.Round((DateValue(convEquipsDemobDate(intI)) - DateValue(convEquipsMobDate(intI))).Days / 30, 0)
            txtmake = convEquipsMake(intI)
            txtmodel = convEquipsModel(intI)
            strsql = "categoryname = 'Conveyance' And EquipmentName = '" & convEquipsNames(intI) & "' and " & _
               "Make ='" & txtmake & "' and Model ='" & txtmodel & "' and Capacity ='" & convEquipsCapacity(intI) & "'"
            strStatement = "Select * from MajorEquipments where " & strsql
            If moledbConnection Is Nothing Then
                strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
                moledbConnection = New OleDbConnection(strConnection)
            End If
            If (moledbConnection.State.ToString().Equals("Closed")) Then
                moledbConnection.Open()
            End If
            mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
            mdataset = New DataSet()

            mOledbDataAdapter.Fill(mdataset, "Equipments")
            For Each Machine In mdataset.Tables("Equipments").Rows
                txtRepvalue = convEquipsRepValue(intI)
                txtDepreciation = convEquipsDepPerc(intI)
                HrsPerMonth = convEquipsHPM(intI)
                txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
                txtPowerperHr = Val(Machine("Power_PerHour").ToString())
                txtOprCostPerMCPerMonth = Val(Machine("OperatorCost_PerMonth").ToString())
                If convEquipsDepPerc(intI) = 2.75 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_275").ToString())
                ElseIf convEquipsDepPerc(intI) = 1.25 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_125").ToString())
                Else
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPerc_050").ToString())
                End If
                If (txtFuelPerHr = 0) Then
                    txtDrive = "Electrical"
                Else
                    txtDrive = "Fuel (Diesel)"
                End If
                txtshifts = convEquipsShifts(intI)
                txtMaintCostperMC_PerMonth = txtRepvalue * (txtMaintPercPerMC_PerMonth / 100)
                PPU = Val(Machine("PowerPerUnit(HP)").ToString())
                Connload = Val(Machine("ConnectedLoadPerMC").ToString())
                UF = Val(Machine("UtilityFactor").ToString())
            Next
            txtFuelperUnitPerMonth = txtFuelPerHr * HrsPerMonth
            txtFuelCostPerMonth = txtFuelperUnitPerMonth * convEquipsQty(intI) * FuelCostperLtr
            txtFuelCostProject = txtFuelCostPerMonth * txtMonths
            txtPowerCostperMonth = txtPowerperHr * HrsPerMonth * convEquipsQty(intI) * PowerCostPerUnit
            txtPowerCostProject = txtPowerCostperMonth * txtMonths
            txtOprCostPerMonth = txtOprCostPerMCPerMonth * convEquipsQty(intI) * txtshifts
            txtOprCostProject = txtOprCostPerMonth * txtMonths
            txtConsumablesPerMonth = txtMaintCostperMC_PerMonth * convEquipsQty(intI)
            txtConsumablesProject = txtConsumablesPerMonth * txtMonths
            OperatingCost_MajorEquips = txtHirecharges + txtFuelCostProject + txtPowerCostProject + _
                  txtOprCostProject + txtConsumablesProject
            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
            xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
            xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
            xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
            'End If
            mOledbDataAdapter = Nothing
            mdataset = Nothing

            If convEquipsChkd(intI) = 1 Then
                xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
                If Len(Trim(xlRange.Value)) = 0 Then
                    xlRange.Value = mMainTitle1
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
                    xlRange.Value = mMainTitle2
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
                    'xlRange.Value = mMainTitle3
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
                    xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
                    xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                    xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
                End If

                SheetNo = getSheetNo(xlWorksheet.Name)
                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = RecordsInserted(SheetNo) + 1
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = convEquipsNames(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDrive
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = convEquipsCapacity(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = convEquipsMake(intI) & " / " & convEquipsModel(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = convEquipsQty(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1) ''
                xlRange.Value = DateValue(convEquipsMobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = DateValue(convEquipsDemobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtRepvalue
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDepreciation
                xlRange.NumberFormat = "#0.00"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = HrsPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtFuelPerHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = FuelCostperLtr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtPowerperHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = PowerCostPerUnit
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtOprCostPerMCPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtshifts
                xlRange.NumberFormat = "#.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtConsumablesPerMonth '* Val(Me.txtMajorEquipQty.Text))
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlWorksheet = xlWorkbook.Sheets(sheetname)
                xlWorksheet.Activate()
                getCategoryShortname(xlWorksheet)
                SheetNo = getSheetNo(sheetname)
                RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1
                If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
            End If
            InsertCommand = ""
            mcategory = "Conveyance"
            InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
            InsertCommand = InsertCommand & "'" & convEquipsNames(intI) & "', "
            InsertCommand = InsertCommand & "'" & convEquipsCapacity(intI) & "', "
            InsertCommand = InsertCommand & "'" & txtmake & "', "
            InsertCommand = InsertCommand & "'" & txtmodel & "', "
            InsertCommand = InsertCommand & "'" & convEquipsMobDate(intI) & "', "
            InsertCommand = InsertCommand & "'" & convEquipsDemobDate(intI) & "', "
            InsertCommand = InsertCommand & convEquipsQty(intI) & ", "
            InsertCommand = InsertCommand & convEquipsChkd(intI) & ", "
            InsertCommand = InsertCommand & convEquipsHPM(intI) & ", "
            InsertCommand = InsertCommand & txtDepreciation & ", "
            InsertCommand = InsertCommand & txtRepvalue & ", "
            InsertCommand = InsertCommand & convEquipsShifts(intI) & ", "
            InsertCommand = InsertCommand & txtMaintCostperMC_PerMonth & ", "
            InsertCommand = InsertCommand & convEquipsConcQty(intI) & ", "
            InsertCommand = InsertCommand & "'" & txtDrive & "', "
            InsertCommand = InsertCommand & PPU & ", "
            InsertCommand = InsertCommand & Connload & ", "
            InsertCommand = InsertCommand & UF & ")"

            Try
                If (moledbConnection1.State.ToString().Equals("Closed")) Then
                    moledbConnection1.Open()
                End If
                moleDBInsertCommand = New OleDbCommand
                moleDBInsertCommand.CommandType = CommandType.Text
                moleDBInsertCommand.CommandText = InsertCommand
                moleDBInsertCommand.Connection = moledbConnection1
                moleDBInsertCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            Finally
                moleDBInsertCommand = Nothing
            End Try
        Next
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = RecordsInserted(getSheetNo(sheetname))
    End Sub
    Private Sub SaveDataInCraneSheet(ByVal sheetname As String, ByVal Fromval As Integer, ByVal Toval As Integer)
        Dim intI As Integer, intj As Integer
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, mcommand As OleDbCommand, Querystring As String = ""
        Dim mdataset As DataSet, madapter As OleDbDataAdapter
        Dim Machine As DataRow, AllValid As Boolean
        Dim currentcategory As String, msgstring As String, strsql As String
        Dim PPU As Single, Connload As Single, UF As Single

        deleteOlddata(sheetname)
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        RowCount = 0

        Me.lblmessage.Text = "Major (Crane) Equipments Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()


        For intI = Fromval To Fromval + Toval
            AllValid = True
            msgstring = "The following were not entered/selected" & vbNewLine
            If craneEquipsChkd(intI) = 1 Then
                If craneEquipsQty(intI) = 0 Then
                    msgstring = msgstring & "*** Equipment Quantity not entered for the " & intI & _
                        IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                    AllValid = False
                End If
            End If

            If (craneEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If (craneEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If Not AllValid Then
                DataSaved = False
                Exit Sub
            End If
            txtMonths = System.Math.Round((DateValue(craneEquipsDemobDate(intI)) - DateValue(craneEquipsMobDate(intI))).Days / 30, 0)
            txtmake = craneEquipsMake(intI)
            txtmodel = craneEquipsModel(intI)
            strsql = "categoryname = 'Cranes' And EquipmentName = '" & craneEquipsNames(intI) & "' and " & _
               "Make ='" & txtmake & "' and Model ='" & txtmodel & "' and Capacity ='" & craneEquipsCapacity(intI) & "'"
            strStatement = "Select * from MajorEquipments where " & strsql
            'End If
            If moledbConnection Is Nothing Then
                strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
                moledbConnection = New OleDbConnection(strConnection)
            End If
            If (moledbConnection.State.ToString().Equals("Closed")) Then
                moledbConnection.Open()
            End If
            mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
            mdataset = New DataSet()

            mOledbDataAdapter.Fill(mdataset, "Equipments")
            For Each Machine In mdataset.Tables("Equipments").Rows
                txtRepvalue = Val(Machine("Repvalue").ToString())  'craneEquipsRepValue(intI)
                txtDepreciation = craneEquipsDepPerc(intI)
                HrsPerMonth = craneEquipsHPM(intI)
                txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
                txtPowerperHr = Val(Machine("Power_PerHour").ToString())
                txtOprCostPerMCPerMonth = Val(Machine("OperatorCost_PerMonth").ToString())
                If craneEquipsDepPerc(intI) = 2.75 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_275").ToString())
                ElseIf craneEquipsDepPerc(intI) = 1.25 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_125").ToString())
                Else
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPerc_050").ToString())
                End If
                If (txtFuelPerHr = 0) Then
                    txtDrive = "Electrical"
                Else
                    txtDrive = "Fuel (Diesel)"
                End If
                txtshifts = craneEquipsShifts(intI)
                txtMaintCostperMC_PerMonth = txtRepvalue * (txtMaintPercPerMC_PerMonth / 100)
                PPU = Val(Machine("PowerPerUnit(HP)").ToString())
                Connload = Val(Machine("ConnectedLoadPerMC").ToString())
                UF = Val(Machine("UtilityFactor").ToString())
            Next
            txtFuelperUnitPerMonth = txtFuelPerHr * HrsPerMonth
            txtFuelCostPerMonth = txtFuelperUnitPerMonth * craneEquipsQty(intI) * FuelCostperLtr
            txtFuelCostProject = txtFuelCostPerMonth * txtMonths
            txtPowerCostperMonth = txtPowerperHr * HrsPerMonth * craneEquipsQty(intI) * PowerCostPerUnit
            txtPowerCostProject = txtPowerCostperMonth * txtMonths
            txtOprCostPerMonth = txtOprCostPerMCPerMonth * craneEquipsQty(intI) * txtshifts
            txtOprCostProject = txtOprCostPerMonth * txtMonths
            txtConsumablesPerMonth = txtMaintCostperMC_PerMonth * craneEquipsQty(intI)
            txtConsumablesProject = txtConsumablesPerMonth * txtMonths
            OperatingCost_MajorEquips = txtHirecharges + txtFuelCostProject + txtPowerCostProject + _
                  txtOprCostProject + txtConsumablesProject
            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
            xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
            xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
            xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
            'End If
            mOledbDataAdapter = Nothing
            mdataset = Nothing

            If craneEquipsChkd(intI) = 1 Then
                xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
                If Len(Trim(xlRange.Value)) = 0 Then
                    xlRange.Value = mMainTitle1
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
                    xlRange.Value = mMainTitle2
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
                    'xlRange.Value = mMainTitle3
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
                    xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
                    xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                    xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
                End If

                SheetNo = getSheetNo(xlWorksheet.Name)
                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = RecordsInserted(SheetNo) + 1
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = craneEquipsNames(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDrive
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = craneEquipsCapacity(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = craneEquipsMake(intI) & " / " & craneEquipsModel(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = craneEquipsQty(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1) ''
                xlRange.Value = DateValue(craneEquipsMobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = DateValue(craneEquipsDemobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtRepvalue
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDepreciation
                xlRange.NumberFormat = "#0.00"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = HrsPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtFuelPerHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = FuelCostperLtr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtPowerperHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = PowerCostPerUnit
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtOprCostPerMCPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtshifts
                xlRange.NumberFormat = "#.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtConsumablesPerMonth '* Val(Me.txtMajorEquipQty.Text))
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlWorksheet = xlWorkbook.Sheets(sheetname)
                xlWorksheet.Activate()
                getCategoryShortname(xlWorksheet)
                SheetNo = getSheetNo(sheetname)
                RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1
                If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
            End If
            InsertCommand = ""
            mcategory = "Cranes"
            InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
            InsertCommand = InsertCommand & "'" & craneEquipsNames(intI) & "', "
            InsertCommand = InsertCommand & "'" & craneEquipsCapacity(intI) & "', "
            InsertCommand = InsertCommand & "'" & txtmake & "', "
            InsertCommand = InsertCommand & "'" & txtmodel & "', "
            InsertCommand = InsertCommand & "'" & craneEquipsMobDate(intI) & "', "
            InsertCommand = InsertCommand & "'" & craneEquipsDemobDate(intI) & "', "
            InsertCommand = InsertCommand & craneEquipsQty(intI) & ", "
            InsertCommand = InsertCommand & craneEquipsChkd(intI) & ", "
            InsertCommand = InsertCommand & craneEquipsHPM(intI) & ", "
            InsertCommand = InsertCommand & txtDepreciation & ", "
            InsertCommand = InsertCommand & txtRepvalue & ", "
            InsertCommand = InsertCommand & craneEquipsShifts(intI) & ", "
            InsertCommand = InsertCommand & txtMaintCostperMC_PerMonth & ", "
            InsertCommand = InsertCommand & craneEquipsConcQty(intI) & ", "
            InsertCommand = InsertCommand & "'" & txtDrive & "', "
            InsertCommand = InsertCommand & PPU & ", "
            InsertCommand = InsertCommand & Connload & ", "
            InsertCommand = InsertCommand & UF & ")"

            Try
                If (moledbConnection1.State.ToString().Equals("Closed")) Then
                    moledbConnection1.Open()
                End If
                moleDBInsertCommand = New OleDbCommand
                moleDBInsertCommand.CommandType = CommandType.Text
                moleDBInsertCommand.CommandText = InsertCommand
                moleDBInsertCommand.Connection = moledbConnection1
                moleDBInsertCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            Finally
                moleDBInsertCommand = Nothing
            End Try
        Next
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = RecordsInserted(getSheetNo(sheetname))
    End Sub
    Private Sub SaveDataInMinorEquipsSheet(ByVal sheetname As String, ByVal Fromval As Integer, ByVal Toval As Integer)
        Dim intI As Integer, intj As Integer
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, Querystring As String = ""
        Dim mdataset As DataSet, madapter As OleDbDataAdapter
        Dim Machine As DataRow, AllValid As Boolean
        Dim currentcategory As String, msgstring As String, strsql As String
        Dim PPU As Single, Connload As Single, UF As Single

        deleteOlddata(sheetname)
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        RowCount = 0

        Me.lblmessage.Text = "Minor Equipments Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()


        For intI = Fromval To Fromval + Toval - 1
            AllValid = True
            msgstring = "The following were not entered/selected" & vbNewLine
            If minorEquipsChkd(intI) = 1 Then
                If minorEquipsQty(intI) = 0 Then
                    msgstring = msgstring & "*** Equipment Quantity not entered for the " & intI & _
                        IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                    AllValid = False
                End If
            End If

            If (minorEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If (minorEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If Not AllValid Then
                DataSaved = False
                Exit Sub
            End If
            txtMonths = System.Math.Round((DateValue(minorEquipsDemobDate(intI)) - DateValue(minorEquipsMobDate(intI))).Days / 30, 0)
            txtmake = minorEquipsMake(intI)
            txtmodel = minorEquipsModel(intI)
            strsql = "categoryname = 'Minor equipments' And EquipmentName = '" & minorEquipsNames(intI) & "' and " & _
               "Make ='" & txtmake & "' and Model ='" & txtmodel & "' and Capacity ='" & minorEquipsCapacity(intI) & "'"
            strStatement = "Select * from MinorEquipments where " & strsql
            If moledbConnection Is Nothing Then
                strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
                moledbConnection = New OleDbConnection(strConnection)
            End If
            If (moledbConnection.State.ToString().Equals("Closed")) Then
                moledbConnection.Open()
            End If
            mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
            mdataset = New DataSet()
            mOledbDataAdapter.Fill(mdataset, "Equipments")
            For Each Machine In mdataset.Tables("Equipments").Rows
                txtDepreciation = 0
                HrsPerMonth = Val(Machine("Hrs_PerMonth").ToString())
                MinorEquipmentCost = Val(Machine("CostOfNewEquipment").ToString())
                RAndMPercentage = Val(Machine("RandMPercentage").ToString())
                txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
                txtPowerperHr = Val(Machine("Power_PerHour").ToString())
                txtOprCostPerMCPerMonth = 0
                If (txtFuelPerHr = 0) Then
                    txtDrive = "Electrical"
                Else
                    txtDrive = "Fuel (Diesel)"
                End If
                txtshifts = 0
                txtMaintCostperMC_PerMonth = (MinorEquipmentCost * RAndMPercentage / 100)
                PPU = Val(Machine("PowerPerUnit(HP)").ToString())
                Connload = Val(Machine("ConnectedLoadPerMC").ToString())
                UF = Val(Machine("UtilityFactor").ToString())
                If minorIsNewMC(intI) = 1 Then NewMCCost = MinorEquipmentCost * minorEquipsQty(intI) Else NewMCCost = 0
            Next
            txtFuelperUnitPerMonth = 0
            txtFuelperUnitPerMonth = txtFuelPerHr * HrsPerMonth
            txtFuelCostPerMonth = txtFuelperUnitPerMonth * minorEquipsQty(intI) * FuelCostperLtr
            txtFuelCostProject = txtFuelCostPerMonth * txtMonths
            txtPowerCostperMonth = txtPowerperHr * HrsPerMonth * minorEquipsQty(intI) * PowerCostPerUnit
            txtPowerCostProject = txtPowerCostperMonth * txtMonths
            txtOprCostPerMonth = txtOprCostPerMCPerMonth * minorEquipsQty(intI) * txtshifts
            txtOprCostProject = txtOprCostPerMonth * txtMonths
            txtConsumablesPerMonth = txtMaintCostperMC_PerMonth * minorEquipsQty(intI)
            txtConsumablesProject = txtConsumablesPerMonth * txtMonths
            OperatingCost_MajorEquips = txtHirecharges + txtFuelCostProject + txtPowerCostProject + _
            txtOprCostProject + txtConsumablesProject
            'End If
            mOledbDataAdapter = Nothing
            mdataset = Nothing

            If minorEquipsChkd(intI) = 1 Then
                xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
                If Len(Trim(xlRange.Value)) = 0 Then
                    xlRange.Value = mMainTitle1
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
                    xlRange.Value = mMainTitle2
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
                    'xlRange.Value = mMainTitle3
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
                    xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
                    xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                    xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
                End If

                SheetNo = getSheetNo(xlWorksheet.Name)
                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = RecordsInserted(SheetNo) + 1
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = minorEquipsNames(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDrive
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = minorEquipsCapacity(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtmake & " / " & txtmodel
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = minorEquipsQty(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1) ''
                xlRange.Value = DateValue(minorEquipsMobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = DateValue(minorEquipsDemobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtRepvalue
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDepreciation
                xlRange.NumberFormat = "#0.00"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = HrsPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtFuelPerHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = FuelCostperLtr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtPowerperHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = PowerCostPerUnit
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtOprCostPerMCPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtshifts
                xlRange.NumberFormat = "#.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtConsumablesPerMonth '* Val(Me.txtMajorEquipQty.Text))
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = MinorEquipmentCost
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlWorksheet = xlWorkbook.Sheets(sheetname)
                xlWorksheet.Activate()
                getCategoryShortname(xlWorksheet)
                SheetNo = getSheetNo(sheetname)
                RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1
                If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
            End If

            InsertCommand = ""
            mcategory = "Minor Equipments"
            InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
            InsertCommand = InsertCommand & "'" & minorEquipsNames(intI) & "', "
            InsertCommand = InsertCommand & "'" & minorEquipsCapacity(intI) & "', "
            InsertCommand = InsertCommand & "'" & txtmake & "', "
            InsertCommand = InsertCommand & "'" & txtmodel & "', "
            InsertCommand = InsertCommand & "'" & minorEquipsMobDate(intI) & "', "
            InsertCommand = InsertCommand & "'" & minorEquipsDemobDate(intI) & "', "
            InsertCommand = InsertCommand & minorEquipsQty(intI) & ", "
            InsertCommand = InsertCommand & minorEquipsChkd(intI) & ", "
            InsertCommand = InsertCommand & minorEquipsHPM(intI) & ", "
            InsertCommand = InsertCommand & txtDepreciation & ", "
            InsertCommand = InsertCommand & minorEquipsShifts(intI) & ", "
            InsertCommand = InsertCommand & MinorEquipmentCost & ", "
            'InsertCommand = InsertCommand & txtMaintCostperMC_PerMonth & ", "
            InsertCommand = InsertCommand & minorIsNewMC(intI) & ", "
            InsertCommand = InsertCommand & "'" & txtDrive & "', "
            InsertCommand = InsertCommand & PPU & ", "
            InsertCommand = InsertCommand & Connload & ", "
            InsertCommand = InsertCommand & UF & ")"
            Try
                If (moledbConnection1.State.ToString().Equals("Closed")) Then
                    moledbConnection1.Open()
                End If
                moleDBInsertCommand = New OleDbCommand
                moleDBInsertCommand.CommandType = CommandType.Text
                moleDBInsertCommand.CommandText = InsertCommand
                moleDBInsertCommand.Connection = moledbConnection1
                moleDBInsertCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString() + vbNewLine + InsertCommand)
            Finally
                moleDBInsertCommand = Nothing
            End Try
        Next
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = RecordsInserted(getSheetNo(sheetname))
    End Sub
    Private Sub SaveDataInDgSetsSheet(ByVal sheetname As String, ByVal Fromval As Integer, ByVal Toval As Integer)
        Dim intI As Integer, intj As Integer
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, mcommand As OleDbCommand, Querystring As String = ""
        Dim mdataset As DataSet, madapter As OleDbDataAdapter
        Dim Machine As DataRow, AllValid As Boolean
        Dim currentcategory As String, msgstring As String, strsql As String
        Dim PPU As Single, Connload As Single, UF As Single

        deleteOlddata(sheetname)
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        RowCount = 0

        Me.lblmessage.Text = "Major (DG Sets) Equipments Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()


        For intI = Fromval To Fromval + Toval
            AllValid = True
            msgstring = "The following were not entered/selected" & vbNewLine
            If dgsetsEquipsChkd(intI) = 1 Then
                If dgsetsEquipsQty(intI) = 0 Then
                    msgstring = msgstring & "*** Equipment Quantity not entered for the " & intI & _
                        IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                    AllValid = False
                End If
            End If

            If (dgsetsEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If (dgsetsEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If Not AllValid Then
                DataSaved = False
                Exit Sub
            End If
            txtMonths = System.Math.Round((DateValue(dgsetsEquipsDemobDate(intI)) - DateValue(dgsetsEquipsMobDate(intI))).Days / 30, 0)
            txtmake = dgsetsEquipsMake(intI)
            txtmodel = dgsetsEquipsModel(intI)
            strsql = "categoryname = 'DG Sets' And EquipmentName = '" & dgsetsEquipsNames(intI) & "' and " & _
               "Make ='" & txtmake & "' and Model ='" & txtmodel & "' and Capacity ='" & dgsetsEquipsCapacity(intI) & "'"
            strStatement = "Select * from MajorEquipments where " & strsql
            If moledbConnection Is Nothing Then
                strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
                moledbConnection = New OleDbConnection(strConnection)
            End If
            If (moledbConnection.State.ToString().Equals("Closed")) Then
                moledbConnection.Open()
            End If
            mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
            mdataset = New DataSet()

            mOledbDataAdapter.Fill(mdataset, "Equipments")
            'Dim Machine As DataRow
            For Each Machine In mdataset.Tables("Equipments").Rows
                txtRepvalue = Val(Machine("Repvalue").ToString())  'dgsetsEquipsRepValue(intI)
                txtDepreciation = dgsetsEquipsDepPerc(intI)
                HrsPerMonth = dgsetsEquipsHPM(intI)
                txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
                txtPowerperHr = Val(Machine("Power_PerHour").ToString())
                txtOprCostPerMCPerMonth = Val(Machine("OperatorCost_PerMonth").ToString())
                If dgsetsEquipsDepPerc(intI) = 2.75 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_275").ToString())
                ElseIf dgsetsEquipsDepPerc(intI) = 1.25 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_125").ToString())
                Else
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPerc_050").ToString())
                End If
                If (txtFuelPerHr = 0) Then
                    txtDrive = "Electrical"
                Else
                    txtDrive = "Fuel (Diesel)"
                End If
                txtshifts = dgsetsEquipsShifts(intI)
                txtMaintCostperMC_PerMonth = txtRepvalue * (txtMaintPercPerMC_PerMonth / 100)
                PPU = Val(Machine("PowerPerUnit(HP)").ToString())
                Connload = Val(Machine("ConnectedLoadPerMC").ToString())
                UF = Val(Machine("UtilityFactor").ToString())
            Next
            txtFuelperUnitPerMonth = txtFuelPerHr * HrsPerMonth
            txtFuelCostPerMonth = txtFuelperUnitPerMonth * dgsetsEquipsQty(intI) * FuelCostperLtr
            txtFuelCostProject = txtFuelCostPerMonth * txtMonths
            txtPowerCostperMonth = txtPowerperHr * HrsPerMonth * dgsetsEquipsQty(intI) * PowerCostPerUnit
            txtPowerCostProject = txtPowerCostperMonth * txtMonths
            txtOprCostPerMonth = txtOprCostPerMCPerMonth * dgsetsEquipsQty(intI) * txtshifts
            txtOprCostProject = txtOprCostPerMonth * txtMonths
            txtConsumablesPerMonth = txtMaintCostperMC_PerMonth * dgsetsEquipsQty(intI)
            txtConsumablesProject = txtConsumablesPerMonth * txtMonths
            OperatingCost_MajorEquips = txtHirecharges + txtFuelCostProject + txtPowerCostProject + _
                  txtOprCostProject + txtConsumablesProject
            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
            xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
            xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
            xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
            'End If
            mOledbDataAdapter = Nothing
            mdataset = Nothing

            If dgsetsEquipsChkd(intI) = 1 Then
                xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
                If Len(Trim(xlRange.Value)) = 0 Then
                    xlRange.Value = mMainTitle1
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
                    xlRange.Value = mMainTitle2
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
                    'xlRange.Value = mMainTitle3
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
                    xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
                    xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                    xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
                End If

                SheetNo = getSheetNo(xlWorksheet.Name)
                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = RecordsInserted(SheetNo) + 1
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = dgsetsEquipsNames(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDrive
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = dgsetsEquipsCapacity(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = dgsetsEquipsMake(intI) & " / " & dgsetsEquipsModel(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = dgsetsEquipsQty(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1) ''
                xlRange.Value = DateValue(dgsetsEquipsMobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = DateValue(dgsetsEquipsDemobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtRepvalue
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDepreciation
                xlRange.NumberFormat = "#0.00"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = HrsPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtFuelPerHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = FuelCostperLtr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtPowerperHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = PowerCostPerUnit
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtOprCostPerMCPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtshifts
                xlRange.NumberFormat = "#.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtConsumablesPerMonth '* Val(Me.txtMajorEquipQty.Text))
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlWorksheet = xlWorkbook.Sheets(sheetname)
                xlWorksheet.Activate()
                getCategoryShortname(xlWorksheet)
                SheetNo = getSheetNo(sheetname)
                RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1
                If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
            End If

            InsertCommand = ""
            mcategory = "DG Sets"
            InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
            InsertCommand = InsertCommand & "'" & dgsetsEquipsNames(intI) & "', "
            InsertCommand = InsertCommand & "'" & dgsetsEquipsCapacity(intI) & "', "
            InsertCommand = InsertCommand & "'" & txtmake & "', "
            InsertCommand = InsertCommand & "'" & txtmodel & "', "
            InsertCommand = InsertCommand & "'" & dgsetsEquipsMobDate(intI) & "', "
            InsertCommand = InsertCommand & "'" & dgsetsEquipsDemobDate(intI) & "', "
            InsertCommand = InsertCommand & dgsetsEquipsQty(intI) & ", "
            InsertCommand = InsertCommand & dgsetsEquipsChkd(intI) & ", "
            InsertCommand = InsertCommand & dgsetsEquipsHPM(intI) & ", "
            InsertCommand = InsertCommand & txtDepreciation & ", "
            InsertCommand = InsertCommand & txtRepvalue & ", "
            InsertCommand = InsertCommand & dgsetsEquipsShifts(intI) & ", "
            InsertCommand = InsertCommand & txtMaintCostperMC_PerMonth & ", "
            InsertCommand = InsertCommand & dgsetsEquipsConcQty(intI) & ", "
            InsertCommand = InsertCommand & "'" & txtDrive & "', "
            InsertCommand = InsertCommand & PPU & ", "
            InsertCommand = InsertCommand & Connload & ", "
            InsertCommand = InsertCommand & UF & ")"

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
                MsgBox(ex.ToString())
            Finally
                moleDBInsertCommand = Nothing
            End Try
        Next
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = RecordsInserted(getSheetNo(sheetname))
    End Sub
    Private Sub SaveDataInMHSheet(ByVal sheetname As String, ByVal Fromval As Integer, ByVal Toval As Integer)
        Dim intI As Integer, intj As Integer
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, mcommand As OleDbCommand, Querystring As String = ""
        Dim mdataset As DataSet, madapter As OleDbDataAdapter
        Dim Machine As DataRow, AllValid As Boolean
        Dim currentcategory As String, msgstring As String, strsql As String
        Dim PPU As Single, Connload As Single, UF As Single

        deleteOlddata(sheetname)
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        RowCount = 0

        Me.lblmessage.Text = "Major (Matl. Handling) Equipments Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()


        For intI = Fromval To Fromval + Toval
            AllValid = True
            msgstring = "The following were not entered/selected" & vbNewLine
            If MHEquipsChkd(intI) = 1 Then
                If MHEquipsQty(intI) = 0 Then
                    msgstring = msgstring & "*** Equipment Quantity not entered for the " & intI & _
                        IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                    AllValid = False
                End If
            End If

            If (MHEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If (MHEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If Not AllValid Then
                DataSaved = False
                Exit Sub
            End If
            txtMonths = System.Math.Round((DateValue(MHEquipsDemobDate(intI)) - DateValue(MHEquipsMobDate(intI))).Days / 30, 0)
            txtmake = MHEquipsMake(intI)
            txtmodel = MHEquipsModel(intI)
            strsql = "categoryname = 'Material Handling' And EquipmentName = '" & MHEquipsNames(intI) & "' and " & _
               "Make ='" & txtmake & "' and Model ='" & txtmodel & "' and Capacity ='" & MHEquipsCapacity(intI) & "'"
            strStatement = "Select * from MajorEquipments where " & strsql
            If moledbConnection Is Nothing Then
                strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
                moledbConnection = New OleDbConnection(strConnection)
            End If
            If (moledbConnection.State.ToString().Equals("Closed")) Then
                moledbConnection.Open()
            End If
            mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
            mdataset = New DataSet()

            mOledbDataAdapter.Fill(mdataset, "Equipments")
            For Each Machine In mdataset.Tables("Equipments").Rows
                txtRepvalue = Val(Machine("Repvalue").ToString())  'MHEquipsRepValue(intI)
                txtDepreciation = MHEquipsDepPerc(intI)
                HrsPerMonth = MHEquipsHPM(intI)
                txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
                txtPowerperHr = Val(Machine("Power_PerHour").ToString())
                txtOprCostPerMCPerMonth = Val(Machine("OperatorCost_PerMonth").ToString())
                If MHEquipsDepPerc(intI) = 2.75 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_275").ToString())
                ElseIf MHEquipsDepPerc(intI) = 1.25 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_125").ToString())
                Else
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPerc_050").ToString())
                End If
                If (txtFuelPerHr = 0) Then
                    txtDrive = "Electrical"
                Else
                    txtDrive = "Fuel (Diesel)"
                End If
                txtshifts = MHEquipsShifts(intI)
                txtMaintCostperMC_PerMonth = txtRepvalue * (txtMaintPercPerMC_PerMonth / 100)
                PPU = Val(Machine("PowerPerUnit(HP)").ToString())
                Connload = Val(Machine("ConnectedLoadPerMC").ToString())
                UF = Val(Machine("UtilityFactor").ToString())
            Next
            txtFuelperUnitPerMonth = txtFuelPerHr * HrsPerMonth
            txtFuelCostPerMonth = txtFuelperUnitPerMonth * MHEquipsQty(intI) * FuelCostperLtr
            txtFuelCostProject = txtFuelCostPerMonth * txtMonths
            txtPowerCostperMonth = txtPowerperHr * HrsPerMonth * MHEquipsQty(intI) * PowerCostPerUnit
            txtPowerCostProject = txtPowerCostperMonth * txtMonths
            txtOprCostPerMonth = txtOprCostPerMCPerMonth * MHEquipsQty(intI) * txtshifts
            txtOprCostProject = txtOprCostPerMonth * txtMonths
            txtConsumablesPerMonth = txtMaintCostperMC_PerMonth * MHEquipsQty(intI)
            txtConsumablesProject = txtConsumablesPerMonth * txtMonths
            OperatingCost_MajorEquips = txtHirecharges + txtFuelCostProject + txtPowerCostProject + _
                  txtOprCostProject + txtConsumablesProject
            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
            xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
            xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
            xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
            'End If
            mOledbDataAdapter = Nothing
            mdataset = Nothing

            If MHEquipsChkd(intI) = 1 Then
                xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
                If Len(Trim(xlRange.Value)) = 0 Then
                    xlRange.Value = mMainTitle1
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
                    xlRange.Value = mMainTitle2
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
                    'xlRange.Value = mMainTitle3
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
                    xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
                    xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                    xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
                End If

                SheetNo = getSheetNo(xlWorksheet.Name)
                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = RecordsInserted(SheetNo) + 1
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = MHEquipsNames(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDrive
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = MHEquipsCapacity(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = MHEquipsMake(intI) & " / " & MHEquipsModel(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = MHEquipsQty(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1) ''
                xlRange.Value = DateValue(MHEquipsMobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = DateValue(MHEquipsDemobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtRepvalue
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDepreciation
                xlRange.NumberFormat = "#0.00"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = HrsPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtFuelPerHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = FuelCostperLtr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtPowerperHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = PowerCostPerUnit
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtOprCostPerMCPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtshifts
                xlRange.NumberFormat = "#.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtConsumablesPerMonth '* Val(Me.txtMajorEquipQty.Text))
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlWorksheet = xlWorkbook.Sheets(sheetname)
                xlWorksheet.Activate()
                getCategoryShortname(xlWorksheet)
                SheetNo = getSheetNo(sheetname)
                RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1
                If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
            End If

            InsertCommand = ""
            mcategory = "Material Handling"
            InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
            InsertCommand = InsertCommand & "'" & MHEquipsNames(intI) & "', "
            InsertCommand = InsertCommand & "'" & MHEquipsCapacity(intI) & "', "
            InsertCommand = InsertCommand & "'" & txtmake & "', "
            InsertCommand = InsertCommand & "'" & txtmodel & "', "
            InsertCommand = InsertCommand & "'" & MHEquipsMobDate(intI) & "', "
            InsertCommand = InsertCommand & "'" & MHEquipsDemobDate(intI) & "', "
            InsertCommand = InsertCommand & MHEquipsQty(intI) & ", "
            InsertCommand = InsertCommand & MHEquipsChkd(intI) & ", "
            InsertCommand = InsertCommand & MHEquipsHPM(intI) & ", "
            InsertCommand = InsertCommand & txtDepreciation & ", "
            InsertCommand = InsertCommand & txtRepvalue & ", "
            InsertCommand = InsertCommand & MHEquipsShifts(intI) & ", "
            InsertCommand = InsertCommand & txtMaintCostperMC_PerMonth & ", "
            InsertCommand = InsertCommand & MHEquipsConcQty(intI) & ", "
            InsertCommand = InsertCommand & "'" & txtDrive & "', "
            InsertCommand = InsertCommand & PPU & ", "
            InsertCommand = InsertCommand & Connload & ", "
            InsertCommand = InsertCommand & UF & ")"

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
                MsgBox(ex.ToString())
            Finally
                moleDBInsertCommand = Nothing
            End Try
        Next
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = RecordsInserted(getSheetNo(sheetname))
    End Sub
    Private Sub SaveDataInNCSheet(ByVal sheetname As String, ByVal Fromval As Integer, ByVal Toval As Integer)
        Dim intI As Integer, intj As Integer
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, mcommand As OleDbCommand, Querystring As String = ""
        Dim mdataset As DataSet, madapter As OleDbDataAdapter
        Dim Machine As DataRow, AllValid As Boolean
        Dim currentcategory As String, msgstring As String, strsql As String
        Dim PPU As Single, Connload As Single, UF As Single

        deleteOlddata(sheetname)
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        RowCount = 0

        Me.lblmessage.Text = "Major (Non Concreting) Equipments Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()


        For intI = Fromval To Fromval + Toval
            AllValid = True
            msgstring = "The following were not entered/selected" & vbNewLine
            If NCEquipsChkd(intI) = 1 Then
                If NCEquipsQty(intI) = 0 Then
                    msgstring = msgstring & "*** Equipment Quantity not entered for the " & intI & _
                        IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                    AllValid = False
                End If
            End If

            If (NCEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If (NCEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If Not AllValid Then
                DataSaved = False
                Exit Sub
            End If
            txtMonths = System.Math.Round((DateValue(NCEquipsDemobDate(intI)) - DateValue(NCEquipsMobDate(intI))).Days / 30, 0)
            txtmake = NCEquipsMake(intI)
            txtmodel = NCEquipsModel(intI)
            strsql = "categoryname = 'Non Concreting' And EquipmentName = '" & NCEquipsNames(intI) & "' and " & _
               "Make ='" & txtmake & "' and Model ='" & txtmodel & "' and Capacity ='" & NCEquipsCapacity(intI) & "'"
            strStatement = "Select * from MajorEquipments where " & strsql
            If moledbConnection Is Nothing Then
                strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
                moledbConnection = New OleDbConnection(strConnection)
            End If
            If (moledbConnection.State.ToString().Equals("Closed")) Then
                moledbConnection.Open()
            End If
            mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
            mdataset = New DataSet()

            mOledbDataAdapter.Fill(mdataset, "Equipments")
            For Each Machine In mdataset.Tables("Equipments").Rows
                txtRepvalue = Val(Machine("Repvalue").ToString())  'NCEquipsRepValue(intI)
                txtDepreciation = NCEquipsDepPerc(intI)
                HrsPerMonth = NCEquipsHPM(intI)
                txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
                txtPowerperHr = Val(Machine("Power_PerHour").ToString())
                txtOprCostPerMCPerMonth = Val(Machine("OperatorCost_PerMonth").ToString())
                If NCEquipsDepPerc(intI) = 2.75 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_275").ToString())
                ElseIf NCEquipsDepPerc(intI) = 1.25 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_125").ToString())
                Else
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPerc_050").ToString())
                End If
                If (txtFuelPerHr = 0) Then
                    txtDrive = "Electrical"
                Else
                    txtDrive = "Fuel (Diesel)"
                End If
                txtshifts = NCEquipsShifts(intI)
                txtMaintCostperMC_PerMonth = txtRepvalue * (txtMaintPercPerMC_PerMonth / 100)
                PPU = Val(Machine("PowerPerUnit(HP)").ToString())
                Connload = Val(Machine("ConnectedLoadPerMC").ToString())
                UF = Val(Machine("UtilityFactor").ToString())
            Next
            txtFuelperUnitPerMonth = txtFuelPerHr * HrsPerMonth
            txtFuelCostPerMonth = txtFuelperUnitPerMonth * NCEquipsQty(intI) * FuelCostperLtr
            txtFuelCostProject = txtFuelCostPerMonth * txtMonths
            txtPowerCostperMonth = txtPowerperHr * HrsPerMonth * NCEquipsQty(intI) * PowerCostPerUnit
            txtPowerCostProject = txtPowerCostperMonth * txtMonths
            txtOprCostPerMonth = txtOprCostPerMCPerMonth * NCEquipsQty(intI) * txtshifts
            txtOprCostProject = txtOprCostPerMonth * txtMonths
            txtConsumablesPerMonth = txtMaintCostperMC_PerMonth * NCEquipsQty(intI)
            txtConsumablesProject = txtConsumablesPerMonth * txtMonths
            OperatingCost_MajorEquips = txtHirecharges + txtFuelCostProject + txtPowerCostProject + _
                  txtOprCostProject + txtConsumablesProject
            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
            xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
            xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
            xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
            'End If
            mOledbDataAdapter = Nothing
            mdataset = Nothing

            If NCEquipsChkd(intI) = 1 Then
                xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
                If Len(Trim(xlRange.Value)) = 0 Then
                    xlRange.Value = mMainTitle1
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
                    xlRange.Value = mMainTitle2
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
                    'xlRange.Value = mMainTitle3
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
                    xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
                    xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                    xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
                End If
                SheetNo = getSheetNo(xlWorksheet.Name)
                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = RecordsInserted(SheetNo) + 1
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = NCEquipsNames(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDrive
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = NCEquipsCapacity(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = NCEquipsMake(intI) & " / " & NCEquipsModel(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = NCEquipsQty(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1) ''
                xlRange.Value = DateValue(NCEquipsMobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = DateValue(NCEquipsDemobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtRepvalue
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDepreciation
                xlRange.NumberFormat = "#0.00"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = HrsPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtFuelPerHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = FuelCostperLtr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtPowerperHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = PowerCostPerUnit
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtOprCostPerMCPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtshifts
                xlRange.NumberFormat = "#.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtConsumablesPerMonth '* Val(Me.txtMajorEquipQty.Text))
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlWorksheet = xlWorkbook.Sheets(sheetname)
                xlWorksheet.Activate()
                getCategoryShortname(xlWorksheet)
                SheetNo = getSheetNo(sheetname)
                RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1
                'MsgBox(RecordsInserted(SheetNo))
                If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
            End If

            InsertCommand = ""
            mcategory = "Non Concreting"
            InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
            InsertCommand = InsertCommand & "'" & NCEquipsNames(intI) & "', "
            InsertCommand = InsertCommand & "'" & NCEquipsCapacity(intI) & "', "
            InsertCommand = InsertCommand & "'" & txtmake & "', "
            InsertCommand = InsertCommand & "'" & txtmodel & "', "
            InsertCommand = InsertCommand & "'" & NCEquipsMobDate(intI) & "', "
            InsertCommand = InsertCommand & "'" & NCEquipsDemobDate(intI) & "', "
            InsertCommand = InsertCommand & NCEquipsQty(intI) & ", "
            InsertCommand = InsertCommand & NCEquipsChkd(intI) & ", "
            InsertCommand = InsertCommand & NCEquipsHPM(intI) & ", "
            InsertCommand = InsertCommand & txtDepreciation & ", "
            InsertCommand = InsertCommand & txtRepvalue & ", "
            InsertCommand = InsertCommand & NCEquipsShifts(intI) & ", "
            InsertCommand = InsertCommand & txtMaintCostperMC_PerMonth & ", "
            InsertCommand = InsertCommand & NCEquipsConcQty(intI) & ", "
            InsertCommand = InsertCommand & "'" & txtDrive & "', "
            InsertCommand = InsertCommand & PPU & ", "
            InsertCommand = InsertCommand & Connload & ", "
            InsertCommand = InsertCommand & UF & ")"

            Try
                If (moledbConnection1.State.ToString().Equals("Closed")) Then
                    moledbConnection1.Open()
                End If
                moleDBInsertCommand = New OleDbCommand
                moleDBInsertCommand.CommandType = CommandType.Text
                moleDBInsertCommand.CommandText = InsertCommand
                moleDBInsertCommand.Connection = moledbConnection1
                moleDBInsertCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            Finally
                moleDBInsertCommand = Nothing
            End Try
        Next
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = RecordsInserted(getSheetNo(sheetname))
    End Sub
    Private Sub SaveDataInMajOthersSheet(ByVal sheetname As String, ByVal Fromval As Integer, ByVal Toval As Integer)
        Dim intI As Integer, intj As Integer
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, mcommand As OleDbCommand, Querystring As String = ""
        Dim mdataset As DataSet, madapter As OleDbDataAdapter
        Dim Machine As DataRow, AllValid As Boolean
        Dim currentcategory As String, msgstring As String, strsql As String
        Dim PPU As Single, Connload As Single, UF As Single

        deleteOlddata(sheetname)
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        RowCount = 0

        Me.lblmessage.Text = "Major (Other) Equipments Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()

        For intI = Fromval To Fromval + Toval
            AllValid = True
            msgstring = "The following were not entered/selected" & vbNewLine
            If majOthersEquipsChkd(intI) = 1 Then
                If majOthersEquipsQty(intI) = 0 Then
                    msgstring = msgstring & "*** Equipment Quantity not entered for the " & intI & _
                        IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                    AllValid = False
                End If
            End If

            If (majOthersEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If (majOthersEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If Not AllValid Then
                DataSaved = False
                Exit Sub
            End If
            txtMonths = System.Math.Round((DateValue(majOthersEquipsDemobDate(intI)) - DateValue(majOthersEquipsMobDate(intI))).Days / 30, 0)
            txtmake = majOthersEquipsMake(intI)
            txtmodel = majOthersEquipsModel(intI)
            strsql = "categoryname = 'Major Others' And EquipmentName = '" & majOthersEquipsNames(intI) & "' and " & _
               "Make ='" & txtmake & "' and Model ='" & txtmodel & "' and Capacity ='" & majOthersEquipsCapacity(intI) & "'"
            strStatement = "Select * from MajorEquipments where " & strsql
            If moledbConnection Is Nothing Then
                strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
                moledbConnection = New OleDbConnection(strConnection)
            End If
            If (moledbConnection.State.ToString().Equals("Closed")) Then
                moledbConnection.Open()
            End If
            mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
            mdataset = New DataSet()

            mOledbDataAdapter.Fill(mdataset, "Equipments")
            'Dim Machine As DataRow
            For Each Machine In mdataset.Tables("Equipments").Rows
                txtRepvalue = Val(Machine("Repvalue").ToString())  'majOthersEquipsRepValue(intI)
                txtDepreciation = MajOthersDepPerc(intI)
                HrsPerMonth = majOthersEquipsHPM(intI)
                txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
                txtPowerperHr = Val(Machine("Power_PerHour").ToString())
                txtOprCostPerMCPerMonth = Val(Machine("OperatorCost_PerMonth").ToString())
                If MajOthersDepPerc(intI) = 2.75 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_275").ToString())
                ElseIf MajOthersDepPerc(intI) = 1.25 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_125").ToString())
                Else
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPerc_050").ToString())
                End If
                If (txtFuelPerHr = 0) Then
                    txtDrive = "Electrical"
                Else
                    txtDrive = "Fuel (Diesel)"
                End If
                txtshifts = MajOthersShifts(intI)
                txtMaintCostperMC_PerMonth = txtRepvalue * (txtMaintPercPerMC_PerMonth / 100)
                PPU = Val(Machine("PowerPerUnit(HP)").ToString())
                Connload = Val(Machine("ConnectedLoadPerMC").ToString())
                UF = Val(Machine("UtilityFactor").ToString())
            Next
            txtFuelperUnitPerMonth = txtFuelPerHr * HrsPerMonth
            txtFuelCostPerMonth = txtFuelperUnitPerMonth * majOthersEquipsQty(intI) * FuelCostperLtr
            txtFuelCostProject = txtFuelCostPerMonth * txtMonths
            txtPowerCostperMonth = txtPowerperHr * HrsPerMonth * majOthersEquipsQty(intI) * PowerCostPerUnit
            txtPowerCostProject = txtPowerCostperMonth * txtMonths
            txtOprCostPerMonth = txtOprCostPerMCPerMonth * majOthersEquipsQty(intI) * txtshifts
            txtOprCostProject = txtOprCostPerMonth * txtMonths
            txtConsumablesPerMonth = txtMaintCostperMC_PerMonth * majOthersEquipsQty(intI)
            txtConsumablesProject = txtConsumablesPerMonth * txtMonths
            OperatingCost_MajorEquips = txtHirecharges + txtFuelCostProject + txtPowerCostProject + _
                  txtOprCostProject + txtConsumablesProject
            xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
            xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
            xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
            xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
            'End If
            mOledbDataAdapter = Nothing
            mdataset = Nothing

            If majOthersEquipsChkd(intI) = 1 Then
                xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
                If Len(Trim(xlRange.Value)) = 0 Then
                    xlRange.Value = mMainTitle1
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
                    xlRange.Value = mMainTitle2
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
                    'xlRange.Value = mMainTitle3
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
                    xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
                    xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                    xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
                End If
                SheetNo = getSheetNo(xlWorksheet.Name)
                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = RecordsInserted(SheetNo) + 1
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = majOthersEquipsNames(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDrive
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = majOthersEquipsCapacity(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = majOthersEquipsMake(intI) & " / " & majOthersEquipsModel(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = majOthersEquipsQty(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1) ''
                xlRange.Value = DateValue(majOthersEquipsMobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = DateValue(majOthersEquipsDemobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtRepvalue
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDepreciation
                xlRange.NumberFormat = "#0.00"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = HrsPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtFuelPerHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = FuelCostperLtr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtPowerperHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = PowerCostPerUnit
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtOprCostPerMCPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtshifts
                xlRange.NumberFormat = "#.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtConsumablesPerMonth '* Val(Me.txtMajorEquipQty.Text))
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlWorksheet = xlWorkbook.Sheets(sheetname)
                xlWorksheet.Activate()
                getCategoryShortname(xlWorksheet)
                SheetNo = getSheetNo(sheetname)
                RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1
                If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
            End If

            InsertCommand = ""
            mcategory = "Major Others"
            InsertCommand = "INSERT INTO " & GetTablename(mcategory) & " Values ("
            InsertCommand = InsertCommand & "'" & majOthersEquipsNames(intI) & "', "
            InsertCommand = InsertCommand & "'" & majOthersEquipsCapacity(intI) & "', "
            InsertCommand = InsertCommand & "'" & txtmake & "', "
            InsertCommand = InsertCommand & "'" & txtmodel & "', "
            InsertCommand = InsertCommand & "'" & majOthersEquipsMobDate(intI) & "', "
            InsertCommand = InsertCommand & "'" & majOthersEquipsDemobDate(intI) & "', "
            InsertCommand = InsertCommand & majOthersEquipsQty(intI) & ", "
            InsertCommand = InsertCommand & majOthersEquipsChkd(intI) & ", "
            InsertCommand = InsertCommand & majOthersEquipsHPM(intI) & ", "
            InsertCommand = InsertCommand & txtDepreciation & ", "
            InsertCommand = InsertCommand & txtRepvalue & ", "
            InsertCommand = InsertCommand & MajOthersShifts(intI) & ", "
            InsertCommand = InsertCommand & txtMaintCostperMC_PerMonth & ", "
            InsertCommand = InsertCommand & majOthersEquipsConcQty(intI) & ", "
            InsertCommand = InsertCommand & "'" & txtDrive & "', "
            InsertCommand = InsertCommand & PPU & ", "
            InsertCommand = InsertCommand & Connload & ", "
            InsertCommand = InsertCommand & UF & ")"
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
                MsgBox(ex.ToString())
            Finally
                moleDBInsertCommand = Nothing
                'moledbconnection3.Close()
            End Try
        Next
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = RecordsInserted(getSheetNo(sheetname))
    End Sub

    Private Sub AddRowInPoweGenCostSheet(ByVal intI As Integer)
        Dim msheetname As String = "PowerGen Cost", cursheetno As Integer
        xlWorksheet = xlWorkbook.Sheets(msheetname)
        xlWorksheet.Activate()
        cursheetno = getSheetNo(xlWorksheet.Name)
        getCategoryShortname(xlWorksheet)
        RowCount = 0
        intj = 1
        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
        xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
        xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
        xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
        xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(cursheetno) + 1, 0)
        xlRange.Value = RecordsInserted(cursheetno) + 1
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.RowHeight = 33
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = EquipNameTextBoxes(intI).Text
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.RowHeight = 33
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = CapacityTextBoxes(intI).Text
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.RowHeight = 33
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = QtyTextBoxes(intI).Text
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.RowHeight = 33
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 1) ''
        xlRange.Value = MobdatePickers(intI).Value.Date.ToString()
        xlRange.NumberFormat = "dd-mmm-yyyy"
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.RowHeight = 33
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = DemobDatePickers(intI).Value.Date.ToString()
        xlRange.NumberFormat = "dd-mmm-yyyy"
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.RowHeight = 33
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 2)
        If msheetname = "Minor Eqpts" Then ' Check from here
            xlRange.Value = MinorEquipmentCost
        Else
            xlRange.Value = txtRepvalue
        End If
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.RowHeight = 33
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = txtDepreciation  ' calcualted as Repvalue * depperc
        xlRange.NumberFormat = "#0.00"
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = HrsPerMonth
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = txtFuelPerHr   ' take from Equipmentsmasters.mdb
        xlRange.NumberFormat = "##0.0#"
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.RowHeight = 33
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = FuelCostperLtr   ' take from global variable 
        xlRange.NumberFormat = "##0.0#"
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.RowHeight = 33
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = txtOprCostPerMCPerMonth  ' from master database 
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = txtshifts   ' from addedequipments database 
        xlRange.NumberFormat = "#.0#"
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.RowHeight = 33
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = txtConsumablesPerMonth '* Val(Me.txtMajorEquipQty.Text))  ' calculated
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = "Remarks"
        xlRange.Cells.Font.Size = 12
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        RecordsInserted(cursheetno) = RecordsInserted(cursheetno) + 1
        If RecordsInserted(cursheetno) > 1 Then FillFormulas(cursheetno, RecordsInserted(cursheetno))
        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = RecordsInserted(cursheetno)
    End Sub
    Private Sub SaveHEDataInSheet(ByVal sheetname As String, ByVal Fromval As Integer, ByVal Toval As Integer)
        Dim intI As Integer, intj As Integer, currentsheetname As String = sheetname
        Dim InsertCommand As String    ', DeleteCommand As String
        Dim moleDBInsertCommand As OleDbCommand, Querystring As String
        Dim mdataset As DataSet, madapter As OleDbDataAdapter
        Dim Machine As DataRow, AllValid As Boolean
        Dim currentcategory As String, msgstring As String, strsql As String
        Dim PPU As Single, Connload As Single, UF As Single

        deleteOlddata(sheetname)
        deleteOlddata("External Others")
        xlWorksheet = xlWorkbook.Sheets(sheetname)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        RowCount = 0

        'intj = 1
        Me.lblmessage.Text = "External Hired equipments Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()

        xlWorksheet = xlWorkbook.Sheets("external Hire")
        xlWorkbook.Activate()
        xlRange = xlWorksheet.Range("Ext_Conv_HirechargesPerMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Conv_HirechargesTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Conv_FuelPerMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Conv_FuelProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Conv_FuelCostperMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Conv_FuelCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Conv_PowerCostperMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Conv_PowerCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Conv_OperatorCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Conv_consumablesProjectTotal")
        xlRange.Value = 0

        xlWorksheet = xlWorkbook.Sheets("external Hire")
        xlWorkbook.Activate()
        xlRange = xlWorksheet.Range("Ext_Excav_HirechargesPerMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Excav_HirechargesTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Excav_FuelPerMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Excav_FuelProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Excav_FuelCostperMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Excav_FuelCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Excav_PowerCostperMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Excav_PowerCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Excav_OperatorCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Excav_consumablesProjectTotal")
        xlRange.Value = 0

        xlWorksheet = xlWorkbook.Sheets("external Hire")
        xlWorkbook.Activate()
        xlRange = xlWorksheet.Range("Ext_Mhandle_HirechargesPerMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Mhandle_HirechargesTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Mhandle_FuelPerMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Mhandle_FuelProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Mhandle_FuelCostperMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Mhandle_FuelCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Mhandle_PowerCostperMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Mhandle_PowerCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Mhandle_OperatorCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Mhandle_consumablesProjectTotal")
        xlRange.Value = 0

        xlWorksheet = xlWorkbook.Sheets("external Hire")
        xlWorkbook.Activate()
        xlRange = xlWorksheet.Range("Ext_Gensets_HirechargesPerMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Gensets_HirechargesTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Gensets_FuelPerMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Gensets_FuelProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Gensets_FuelCostperMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Gensets_FuelCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Gensets_PowerCostperMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Gensets_PowerCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Gensets_OperatorCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_Gensets_consumablesProjectTotal")
        xlRange.Value = 0

        xlWorksheet = xlWorkbook.Sheets("external Hire")
        xlWorkbook.Activate()
        xlRange = xlWorksheet.Range("Ext_MTransport_HirechargesPerMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_MTransport_HirechargesTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_MTransport_FuelPerMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_MTransport_FuelProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_MTransport_FuelCostperMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_MTransport_FuelCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_MTransport_PowerCostperMonthTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_MTransport_PowerCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_MTransport_OperatorCostProjectTotal")
        xlRange.Value = 0
        xlRange = xlWorksheet.Range("Ext_MTransport_consumablesProjectTotal")
        xlRange.Value = 0

        For intI = Fromval To Fromval + Toval - 1
            AllValid = True
            msgstring = "The following were not entered/selected" & vbNewLine
            If hiredEquipsChkd(intI) = 1 Then
                If hiredEquipsQty(intI) = 0 Then
                    msgstring = msgstring & "*** Equipment Quantity not entered for the " & intI & _
                        IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                    AllValid = False
                End If
            End If

            If (hiredEquipsMobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If (hiredEquipsDemobDate(intI).ToString() = "") Then
                msgstring = msgstring & "*** Equipment Demobilisation date  not specfied for the " & intI & _
                     IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
                AllValid = False
            End If
            If Not AllValid Then
                DataSaved = False
                Exit Sub
            End If
            txtMonths = System.Math.Round((DateValue(hiredEquipsDemobDate(intI)) - DateValue(hiredEquipsMobDate(intI))).Days / 30, 0)
            txtmake = hiredEquipsMake(intI)
            txtmodel = hiredEquipsModel(intI)
            strsql = "EquipmentName = '" & hiredEquipsNames(intI) & "' and " & _
               "Make ='" & txtmake & "' and Model ='" & txtmodel & "' and Capacity ='" & hiredEquipsCapacity(intI) & "'"
            strStatement = "Select * from HiredEquipments where " & strsql
            If moledbConnection Is Nothing Then
                strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
                moledbConnection = New OleDbConnection(strConnection)
            End If
            If (moledbConnection.State.ToString().Equals("Closed")) Then
                moledbConnection.Open()
            End If
            mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
            mdataset = New DataSet()
            mOledbDataAdapter.Fill(mdataset, "Equipments")
            'Dim Machine As DataRow
            'MsgBox(mdataset.Tables("Equipments").Rows.Count)
            For Each Machine In mdataset.Tables("Equipments").Rows
                txtRepvalue = hiredEquipsHireCharges(intI)
                txtDepreciation = 0
                HrsPerMonth = Val(Machine("Hrs_PerMonth").ToString())
                txtHirecharges = Val(Machine("Hirecharges").ToString())
                RAndMPercentage = Val(Machine("RandMPercentage").ToString())
                txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
                txtPowerperHr = Val(Machine("Power_PerHour").ToString())
                txtOprCostPerMCPerMonth = Val(Machine("OperatorCost_PerMonth").ToString()) '
                If (txtFuelPerHr = 0) Then
                    txtDrive = "Electrical"
                Else
                    txtDrive = "Fuel (Diesel)"
                End If
                txtshifts = 0
                txtMaintCostperMC_PerMonth = (txtHirecharges * RAndMPercentage / 100)

            Next
            'txtFuelperUnitPerMonth = 0
            If hiredCategoryNames(intI) = "HiredConveyance" Then
                txtFuelperUnitPerMonth = HrsPerMonth / txtFuelPerHr '/ HrsPerMonth
            Else
                txtFuelperUnitPerMonth = txtFuelPerHr * HrsPerMonth
            End If
            'txtFuelperUnitPerMonth = txtFuelPerHr * HrsPerMonth
            txtFuelCostPerMonth = txtFuelperUnitPerMonth * hiredEquipsQty(intI) * FuelCostperLtr
            txtFuelCostProject = txtFuelCostPerMonth * txtMonths
            txtPowerCostperMonth = txtPowerperHr * HrsPerMonth * hiredEquipsQty(intI) * PowerCostPerUnit
            txtPowerCostProject = txtPowerCostperMonth * txtMonths
            txtOprCostPerMonth = txtOprCostPerMCPerMonth * hiredEquipsQty(intI) * txtshifts
            txtOprCostProject = txtOprCostPerMonth * txtMonths
            txtConsumablesPerMonth = txtMaintCostperMC_PerMonth * hiredEquipsQty(intI)
            txtConsumablesProject = txtConsumablesPerMonth * txtMonths
            OperatingCost_MajorEquips = txtHirecharges + txtFuelCostProject + txtPowerCostProject + _
            txtOprCostProject + txtConsumablesProject
            'End If
            mOledbDataAdapter = Nothing
            mdataset = Nothing

            If hiredEquipsChkd(intI) = 1 Then
                If hiredCategoryNames(intI) = "Others" Then
                    currentsheetname = "External Others"
                Else
                    currentsheetname = "external Hire"
                End If
                xlWorksheet = xlWorkbook.Sheets(currentsheetname)
                xlWorksheet.Activate()
                getCategoryShortname(xlWorksheet)
                'MsgBox(CategoryTextBoxes(intI).Text)
                'If CategoryTextBoxes(intI).Text = "Others" Then
                xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
                If Len(Trim(xlRange.Value)) = 0 Then
                    xlRange.Value = mMainTitle1
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
                    xlRange.Value = mMainTitle2
                    xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
                    'xlRange.Value = mMainTitle3
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
                    xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
                    xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                    xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
                End If

                SheetNo = getSheetNo(xlWorksheet.Name)
                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = RecordsInserted(SheetNo) + 1
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = hiredEquipsNames(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDrive
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = hiredEquipsCapacity(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtmake & " / " & txtmodel
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = hiredEquipsQty(intI)
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1) ''
                xlRange.Value = DateValue(hiredEquipsMobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = DateValue(hiredEquipsDemobDate(intI))
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtRepvalue
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtDepreciation
                xlRange.NumberFormat = "#0.00"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 4)
                xlRange.Value = HrsPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtFuelPerHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = System.Math.Round(txtFuelperUnitPerMonth, 0)
                xlRange.NumberFormat = "###0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = FuelCostperLtr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtPowerperHr
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = PowerCostPerUnit
                xlRange.NumberFormat = "##0.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtOprCostPerMCPerMonth
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = txtshifts
                xlRange.NumberFormat = "#.0#"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlRange.Offset(0, 3)
                xlRange.Value = txtConsumablesPerMonth '* Val(Me.txtMajorEquipQty.Text))
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                If (UCase(hiredCategoryNames(intI)) = UCase("HiredConveyance")) Then
                    xlRange = xlWorksheet.Range("Ext_Conv_HirechargesPerMonthTotal")
                    xlRange.Value = xlRange.Value + txtHirecharges * hiredEquipsQty(intI)
                    xlRange = xlWorksheet.Range("Ext_Conv_HirechargesTotal")
                    xlRange.Value = xlRange.Value + txtRepvalue * hiredEquipsQty(intI) * txtMonths
                    xlRange = xlWorksheet.Range("Ext_Conv_FuelPerMonthTotal")
                    xlRange.Value = xlRange.Value + txtFuelperUnitPerMonth * hiredEquipsQty(intI)
                    xlRange = xlWorksheet.Range("Ext_Conv_FuelProjectTotal")
                    xlRange.Value = xlRange.Value + txtFuelperUnitPerMonth * hiredEquipsQty(intI) * txtMonths
                    xlRange = xlWorksheet.Range("Ext_Conv_FuelCostperMonthTotal")
                    xlRange.Value = xlRange.Value + txtFuelCostPerMonth
                    xlRange = xlWorksheet.Range("Ext_Conv_FuelCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtFuelCostProject
                    xlRange = xlWorksheet.Range("Ext_Conv_PowerCostperMonthTotal")
                    xlRange.Value = xlRange.Value + txtPowerCostperMonth
                    xlRange = xlWorksheet.Range("Ext_Conv_PowerCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtPowerCostProject
                    xlRange = xlWorksheet.Range("Ext_Conv_OperatorCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtOprCostProject
                    xlRange = xlWorksheet.Range("Ext_Conv_consumablesProjectTotal")
                    xlRange.Value = xlRange.Value + txtConsumablesProject
                ElseIf (UCase(hiredCategoryNames(intI)) = UCase("Excav / Earthwork")) Then
                    xlRange = xlWorksheet.Range("Ext_Excav_HirechargesPerMonthTotal")
                    xlRange.Value = xlRange.Value + txtRepvalue * hiredEquipsQty(intI)
                    xlRange = xlWorksheet.Range("Ext_Excav_HirechargesTotal")
                    xlRange.Value = xlRange.Value + txtRepvalue * hiredEquipsQty(intI) * txtMonths
                    xlRange = xlWorksheet.Range("Ext_Excav_FuelPerMonthTotal")
                    xlRange.Value = xlRange.Value + txtFuelperUnitPerMonth * hiredEquipsQty(intI)
                    xlRange = xlWorksheet.Range("Ext_Excav_FuelProjectTotal")
                    xlRange.Value = xlRange.Value + txtFuelperUnitPerMonth * hiredEquipsQty(intI) * txtMonths
                    xlRange = xlWorksheet.Range("Ext_Excav_FuelCostperMonthTotal")
                    xlRange.Value = xlRange.Value + txtFuelCostPerMonth
                    xlRange = xlWorksheet.Range("Ext_Excav_FuelCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtFuelCostProject
                    xlRange = xlWorksheet.Range("Ext_Excav_PowerCostperMonthTotal")
                    xlRange.Value = xlRange.Value + txtPowerCostperMonth
                    xlRange = xlWorksheet.Range("Ext_Excav_PowerCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtPowerCostProject
                    xlRange = xlWorksheet.Range("Ext_Excav_OperatorCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtOprCostProject
                    xlRange = xlWorksheet.Range("Ext_Excav_consumablesProjectTotal")
                    xlRange.Value = xlRange.Value + txtConsumablesProject
                ElseIf (UCase(hiredCategoryNames(intI)) = UCase("Gensets")) Then
                    xlRange = xlWorksheet.Range("Ext_Gensets_HirechargesPerMonthTotal")
                    xlRange.Value = xlRange.Value + txtRepvalue * hiredEquipsQty(intI)
                    xlRange = xlWorksheet.Range("Ext_Gensets_HirechargesTotal")
                    xlRange.Value = xlRange.Value + txtRepvalue * hiredEquipsQty(intI) * txtMonths
                    xlRange = xlWorksheet.Range("Ext_Gensets_FuelPerMonthTotal")
                    xlRange.Value = xlRange.Value + txtFuelperUnitPerMonth * hiredEquipsQty(intI)
                    xlRange = xlWorksheet.Range("Ext_Gensets_FuelProjectTotal")
                    xlRange.Value = xlRange.Value + txtFuelperUnitPerMonth * hiredEquipsQty(intI) * txtMonths
                    xlRange = xlWorksheet.Range("Ext_Gensets_FuelCostperMonthTotal")
                    xlRange.Value = xlRange.Value + txtFuelCostPerMonth
                    xlRange = xlWorksheet.Range("Ext_Gensets_FuelCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtFuelCostProject
                    xlRange = xlWorksheet.Range("Ext_Gensets_PowerCostperMonthTotal")
                    xlRange.Value = xlRange.Value + txtPowerCostperMonth
                    xlRange = xlWorksheet.Range("Ext_Gensets_PowerCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtPowerCostProject
                    xlRange = xlWorksheet.Range("Ext_Gensets_OperatorCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtOprCostProject
                    xlRange = xlWorksheet.Range("Ext_Gensets_consumablesProjectTotal")
                    xlRange.Value = xlRange.Value + txtConsumablesProject
                ElseIf (UCase(hiredCategoryNames(intI)) = UCase("Matl Handling")) Then
                    xlRange = xlWorksheet.Range("Ext_Mhandle_HirechargesPerMonthTotal")
                    xlRange.Value = xlRange.Value + txtRepvalue * hiredEquipsQty(intI)
                    xlRange = xlWorksheet.Range("Ext_Mhandle_HirechargesTotal")
                    xlRange.Value = xlRange.Value + txtRepvalue * hiredEquipsQty(intI) * txtMonths
                    xlRange = xlWorksheet.Range("Ext_Mhandle_FuelPerMonthTotal")
                    xlRange.Value = xlRange.Value + txtFuelperUnitPerMonth * hiredEquipsQty(intI)
                    xlRange = xlWorksheet.Range("Ext_Mhandle_FuelProjectTotal")
                    xlRange.Value = xlRange.Value + txtFuelperUnitPerMonth * hiredEquipsQty(intI) * txtMonths
                    xlRange = xlWorksheet.Range("Ext_Mhandle_FuelCostperMonthTotal")
                    xlRange.Value = xlRange.Value + txtFuelCostPerMonth
                    xlRange = xlWorksheet.Range("Ext_Mhandle_FuelCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtFuelCostProject
                    xlRange = xlWorksheet.Range("Ext_Mhandle_PowerCostperMonthTotal")
                    xlRange.Value = xlRange.Value + txtPowerCostperMonth
                    xlRange = xlWorksheet.Range("Ext_Mhandle_PowerCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtPowerCostProject
                    xlRange = xlWorksheet.Range("Ext_Mhandle_OperatorCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtOprCostProject
                    xlRange = xlWorksheet.Range("Ext_Mhandle_consumablesProjectTotal")
                    xlRange.Value = xlRange.Value + txtConsumablesProject
                ElseIf (UCase(hiredCategoryNames(intI)) = UCase("Matl Transport")) Then
                    xlRange = xlWorksheet.Range("Ext_MTransport_HirechargesPerMonthTotal")
                    xlRange.Value = xlRange.Value + txtRepvalue * hiredEquipsQty(intI)
                    xlRange = xlWorksheet.Range("Ext_MTransport_HirechargesTotal")
                    xlRange.Value = xlRange.Value + txtRepvalue * hiredEquipsQty(intI) * txtMonths
                    xlRange = xlWorksheet.Range("Ext_MTransport_FuelPerMonthTotal")
                    xlRange.Value = xlRange.Value + txtFuelperUnitPerMonth * hiredEquipsQty(intI)
                    xlRange = xlWorksheet.Range("Ext_MTransport_FuelProjectTotal")
                    xlRange.Value = xlRange.Value + txtFuelperUnitPerMonth * hiredEquipsQty(intI) * txtMonths
                    xlRange = xlWorksheet.Range("Ext_MTransport_FuelCostperMonthTotal")
                    xlRange.Value = xlRange.Value + txtFuelCostPerMonth
                    xlRange = xlWorksheet.Range("Ext_MTransport_FuelCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtFuelCostProject
                    xlRange = xlWorksheet.Range("Ext_MTransport_PowerCostperMonthTotal")
                    xlRange.Value = xlRange.Value + txtPowerCostperMonth
                    xlRange = xlWorksheet.Range("Ext_MTransport_PowerCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtPowerCostProject
                    xlRange = xlWorksheet.Range("Ext_MTransport_OperatorCostProjectTotal")
                    xlRange.Value = xlRange.Value + txtOprCostProject
                    xlRange = xlWorksheet.Range("Ext_MTransport_consumablesProjectTotal")
                    xlRange.Value = xlRange.Value + txtConsumablesProject
                End If

                SheetNo = getSheetNo(xlWorksheet.Name)
                RecordsInserted(SheetNo) = RecordsInserted(SheetNo) + 1
                If RecordsInserted(SheetNo) > 1 Then FillFormulas(SheetNo, RecordsInserted(SheetNo))
            End If
            InsertCommand = ""
            InsertCommand = "INSERT INTO HiredEquips Values ("
            InsertCommand = InsertCommand & "'" & hiredCategoryNames(intI) & "', "
            InsertCommand = InsertCommand & "'" & hiredEquipsNames(intI) & "', "
            InsertCommand = InsertCommand & "'" & hiredEquipsCapacity(intI) & "', "
            InsertCommand = InsertCommand & "'" & txtmake & "', "
            InsertCommand = InsertCommand & "'" & txtmodel & "', "
            InsertCommand = InsertCommand & "'" & hiredEquipsMobDate(intI) & "', "
            InsertCommand = InsertCommand & "'" & hiredEquipsDemobDate(intI) & "', "
            InsertCommand = InsertCommand & hiredEquipsQty(intI) & ", "
            InsertCommand = InsertCommand & hiredEquipsChkd(intI) & ", "
            InsertCommand = InsertCommand & hiredEquipsHPM(intI) & ", "
            InsertCommand = InsertCommand & txtDepreciation & ", "
            InsertCommand = InsertCommand & hiredEquipsShifts(intI) & ", "
            InsertCommand = InsertCommand & hiredEquipsHireCharges(intI) & ")"

            Try
                If (moledbConnection1.State.ToString().Equals("Closed")) Then
                    moledbConnection1.Open()
                End If
                moleDBInsertCommand = New OleDbCommand
                moleDBInsertCommand.CommandType = CommandType.Text
                moleDBInsertCommand.CommandText = InsertCommand
                moleDBInsertCommand.Connection = moledbConnection1
                moleDBInsertCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            Finally
                moleDBInsertCommand = Nothing
            End Try
        Next

        xlWorksheet = xlWorkbook.Sheets("external Hire")
        getCategoryShortname(xlWorksheet)
        SheetNo = getSheetNo("external Hire")
        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = RecordsInserted(SheetNo)

        xlWorksheet = xlWorkbook.Sheets("External Others")
        getCategoryShortname(xlWorksheet)
        SheetNo = getSheetNo("External Others")
        xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
        xlRange.Value = RecordsInserted(SheetNo)
    End Sub
    Private Sub SaveLightingDataInSheet(ByVal sheetname As String)
        Dim strconnection1 As String, SelectCommand As String
        Dim madapter As OleDbDataAdapter, mdataset As New DataSet, mrow As DataRow
        Dim LightingCategories As String = "LightingEquipments", FormulaString As String
        'Dim Index As Integer, 
        Dim cntr As Integer = 0, intI As Integer
        xlWorksheet = xlWorkbook.Worksheets(sheetname)
        xlWorksheet.Activate()
        Category_Shortname = "PowerReq_"

        Me.lblmessage.Text = "Lighting Power Requirement Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()

        xlRange = xlWorksheet.Range(Category_Shortname & "Client")
        If Len(Trim(xlRange.Value)) = 0 Then
            xlRange.Value = mMainTitle1
            xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
            xlRange.Value = mMainTitle2
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
            'xlRange.Value = mMainTitle3
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            xlRange = xlWorksheet.Range(Category_Shortname & "Client")
            xlRange.Value = mClient
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            xlRange = xlWorksheet.Range(Category_Shortname & "Location")
            xlRange.Value = mLocation
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            xlRange = xlWorksheet.Range(Category_Shortname & "StartDate")
            xlRange.Value = mStartDate.Date.ToString()
            xlRange.NumberFormat = "dd-mmm-yyyy"
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            xlRange = xlWorksheet.Range(Category_Shortname & "EndDate")
            xlRange.Value = mEndDate.Date.ToString()
            xlRange.NumberFormat = "dd-mmm-yyyy"
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            xlRange = xlWorksheet.Range(Category_Shortname & "ProjectValue")
            xlRange.Value = mProjectvalue
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
        End If

        If moledbConnection1 Is Nothing Then
            strconnection1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
            moledbConnection1 = New OleDbConnection(strconnection1)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection1.Open()
        End If
        SelectCommand = "Select * from " & LightingCategories & " where ChkboxNo = 1"
        madapter = New OleDbDataAdapter(SelectCommand, moledbConnection1)
        madapter.Fill(mdataset, "LightingEquipments")
        'xlWorksheet = xlWorkbook.Worksheets("PowerReqmt")
        cntr = 0
        xlRange = xlWorksheet.Range("PowerReq_LightingTotal")
        xlRange = xlRange.Offset(-1, 0)
        xlRange.Select()
        x1 = xlRange.Address


        While (xlRange.Address) <> xlWorksheet.Range("PowerReq_LightingStart").Address
            xlRange.Application.ActiveCell.EntireRow.Delete()
            xlRange = xlWorksheet.Range("PowerReq_LightingTotal")
            xlRange = xlRange(-1, 0)
            xlRange.Select()
        End While

        If xlRange.Value = "" Then
            If mdataset.Tables("LightingEquipments").Rows.Count > 0 Then
                For Each mrow In mdataset.Tables("LightingEquipments").Rows
                    cntr = cntr + 1
                    xlRange = xlWorksheet.Range("PowerReq_LightingTotal")
                    xlRange.Select()
                    If cntr > 1 Then
                        xlApp.DisplayAlerts = False
                        xlRange.Application.ActiveCell.EntireRow.Insert()
                        xlRange = xlWorksheet.Range(xlRange.Application.ActiveCell.Address)
                    Else
                        xlRange = xlWorksheet.Range("PowerReq_LightingTotal").Offset(-1, 0)
                        xlRange.Select()
                    End If
                    xlRange.Value = cntr
                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = mrow("Description").ToString()
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = mrow("Capacity").ToString()
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = mrow("Make").ToString() & "/" & mrow("Model").ToString()
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = Val(mrow("Qty").ToString())
                    xlRange.NumberFormat = "##0"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = mrow("PowerPerUnit").ToString()
                    xlRange.NumberFormat = "##0.000"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = mrow("MobDate").ToString
                    xlRange.NumberFormat = "dd-mmm-yyyy"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.ColumnWidth = 25
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = mrow("DemobDate").ToString()
                    xlRange.NumberFormat = "dd-mmm-yyyy"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.ColumnWidth = 25
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = xlRange.Offset(0, 2)
                    xlRange.Formula = Val(mrow("ConnectedLoadPerMC").ToString())
                    xlRange.NumberFormat = "##0.000"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = Val(mrow("Utilityfactor").ToString())
                    xlRange.NumberFormat = "##0.000"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33

                    xlRange = xlWorksheet.Range("$A$" & xlRange.Application.ActiveCell.Row)
                    xlRange.Select()

                    xlRange = xlWorksheet.Range(xlRange.Address, xlRange.Offset(0, 12).Address)
                    With xlRange
                        .BorderAround(, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin)
                        .Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    End With
                Next
                Dim Times As Integer = 0

                For Times = 1 To cntr - 1
                    xlRange = xlWorksheet.Range("PowerReq__LightingFormula1")
                    xlRange.Copy()
                    xlRange = xlRange.Offset(Times, 0)
                    xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
                    With xlRange
                        .Font.Bold = False
                        .Font.Size = 12
                        '.Interior.Pattern = 1
                    End With
                Next

                For Times = 1 To cntr - 1
                    xlRange = xlWorksheet.Range("PowerReq__LightingFormula2")
                    xlRange.Copy()
                    xlRange = xlRange.Offset(Times, 0)
                    xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
                    With xlRange
                        .Font.Bold = False
                        .Font.Size = 12
                        '.Interior.Pattern = 1
                    End With
                Next

                For Times = 1 To cntr - 1
                    xlRange = xlWorksheet.Range("PowerReq__LightingFormula3")
                    xlRange.Copy()
                    xlRange = xlRange.Offset(Times, 0)
                    xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
                    With xlRange
                        .Font.Bold = False
                        .Font.Size = 12
                        '.Interior.Pattern = 1
                    End With
                Next

                xlRange = xlWorksheet.Range("PowerReq__LightingFormula3").Offset(cntr, 0)
                xlRange.Select()
                FormulaString = "=sum(R[-" & cntr & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                With xlRange
                    .Font.Bold = True
                    .Font.Size = 12
                    '.Interior.Pattern = 1
                End With

                For Times = 1 To cntr - 1
                    xlRange = xlWorksheet.Range("PowerReq__LightingFormula4")
                    xlRange.Copy()
                    xlRange = xlRange.Offset(Times, 0)
                    xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
                    With xlRange
                        .Font.Bold = False
                        .Font.Size = 12
                        '.Interior.Pattern = 1
                    End With
                Next

                xlRange = xlWorksheet.Range("PowerReq__LightingFormula4").Offset(cntr, 0)
                xlRange.Select()
                FormulaString = "=sum(R[-" & cntr & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                With xlRange
                    .Font.Bold = True
                    .Font.Size = 12
                    '.Interior.Pattern = 1
                End With
            End If
            'Next
            madapter = Nothing
            mdataset = Nothing
        End If
        'Next
        'End If
    End Sub

    Private Sub ComputeValues(ByVal intI As Integer)
        Dim AllValid As Boolean
        Dim currentcategory As String 'txtpurchval As Long
        AllValid = True
        Dim msgstring As String, strsql As String
        msgstring = "The following were not entered/selected" & vbNewLine
        If (EquipNameTextBoxes(intI).Text = "") Then
            msgstring = msgstring & "*** Equipment Name not entered for the " & intI & _
                 IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If

        If (CapacityTextBoxes(intI).Text = "") Then
            msgstring = msgstring & "*** Equipment capacity not entered for the " & intI & _
                 IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If
        If (MakeModelTextBoxes(intI).Text = "") Then
            msgstring = msgstring & "*** Equipment Make & Model not entered for the " & intI & _
              IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If
        If (QtyTextBoxes(intI).Text = "" Or Len(Trim(QtyTextBoxes(intI).Text)) = 0) Then
            msgstring = msgstring & "*** Equipment Quantity not entered for the " & intI & _
                IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If
        If (MobdatePickers(intI).Value.ToString() = "") Then
            msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                 IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If
        If (DemobDatePickers(intI).Value.ToString() = "") Then
            msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                 IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If
        If Not AllValid Then Exit Sub

        lblmessage.Visible = True
        txtMonths = System.Math.Round((DemobDatePickers(intI).Value.Date - MobdatePickers(intI).Value.Date).Days / 30, 0)

        txtmake = Strings.Trim(Strings.Left(MakeModelTextBoxes(intI).Text, Strings.InStr(MakeModelTextBoxes(intI).Text, " /") - 1))
        txtmodel = Strings.Trim(Strings.Mid(MakeModelTextBoxes(intI).Text, InStr(MakeModelTextBoxes(intI).Text, "/ ") + 1))
        strsql = "categoryname = '" & CategoryTextBoxes(intI).Text & "' And EquipmentName = '" & EquipNameTextBoxes(intI).Text & "' and " & _
           "Make ='" & txtmake & "' and Model ='" & txtmodel & "' and Capacity ='" & CapacityTextBoxes(intI).Text & "'"
        If CategoryTextBoxes(intI).Text = "Minor Equipments" Then
            currentcategory = "Minor"
            strStatement = "Select * from MinorEquipments where " & strsql
        ElseIf CategoryTextBoxes(intI).Text = "HiredConveyance" Or CategoryTextBoxes(intI).Text = "Excav / Earthwork" Or _
            CategoryTextBoxes(intI).Text = "Matl Handling" Or CategoryTextBoxes(intI).Text = "Matl Transport" Or _
            CategoryTextBoxes(intI).Text = "Others" Or CategoryTextBoxes(intI).Text = "Gensets" Then
            currentcategory = "Hired"
            strStatement = "Select * from HiredEquipments where " & strsql
        Else
            currentcategory = "Major"
            strStatement = "Select * from MajorEquipments where " & strsql
        End If
        If moledbConnection Is Nothing Then
            strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
        mDataSet = New DataSet()

        mOledbDataAdapter.Fill(mDataSet, "Equipments")

        Dim Machine As DataRow
        If currentcategory = "Major" Then
            For Each Machine In mDataSet.Tables("Equipments").Rows
                txtRepvalue = Val(Machine("Repvalue").ToString())
                If (DepPercComboboxes(intI).Text = "Fixed") Then
                    txtDepreciation = Val(Machine("DepreciationPercentage").ToString())
                Else
                    txtDepreciation = Val(DepPercComboboxes(intI).Text)
                End If
                HrsPerMonth = Val(HrsPermonthTextBoxes(intI).Text)
                txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
                txtPowerperHr = Val(Machine("Power_PerHour").ToString())
                txtOprCostPerMCPerMonth = Val(Machine("OperatorCost_PerMonth").ToString())
                If Val(DepPercComboboxes(intI).Text) = 2.75 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_275").ToString())
                ElseIf Val(DepPercComboboxes(intI).Text) = 1.25 Then
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPer_125").ToString())
                Else
                    txtMaintPercPerMC_PerMonth = Val(Machine("RAndMPerc_050").ToString())
                End If
                If (txtFuelPerHr = 0) Then
                    txtDrive = "Electrical"
                Else
                    txtDrive = "Fuel (Diesel)"
                End If
                If (ShiftsComboboxes(intI).Text = "") Then
                    ShiftsComboboxes(intI).SelectedIndex = 0
                End If
                txtshifts = Val(ShiftsComboboxes(intI).Text)
                txtMaintCostperMC_PerMonth = txtRepvalue * (txtMaintPercPerMC_PerMonth / 100)
            Next
        ElseIf currentcategory = "Minor" Then
            For Each Machine In mDataSet.Tables("Equipments").Rows
                txtDepreciation = 0
                HrsPerMonth = Val(HrsPermonthTextBoxes(intI).Text)
                MinorEquipmentCost = Val(PurchvalTextBoxes(intI).Text)
                RAndMPercentage = Val(Machine("RandMPercentage").ToString())
                txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
                txtPowerperHr = Val(Machine("Power_PerHour").ToString())
                txtOprCostPerMCPerMonth = 0
                If (txtFuelPerHr = 0) Then
                    txtDrive = "Electrical"
                Else
                    txtDrive = "Fuel (Diesel)"
                End If
                txtshifts = 0
                txtMaintCostperMC_PerMonth = (MinorEquipmentCost * RAndMPercentage / 100)
            Next
            If IsNewMC(intI).Checked Then NewMCCost = MinorEquipmentCost * Val(QtyTextBoxes(intI).Text) Else NewMCCost = 0
        ElseIf currentcategory = "Hired" Then
            For Each Machine In mDataSet.Tables("Equipments").Rows
                txtRepvalue = Val(HireChargesTextBoxes(intI).Text)
                txtDepreciation = 0
                HrsPerMonth = Val(HrsPermonthTextBoxes(intI).Text)
                MinorEquipmentCost = 0
                RAndMPercentage = 0
                'MsgBox(HrsPerMonth)
                txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
                txtPowerperHr = Val(Machine("Power_PerHour").ToString())
                txtOprCostPerMCPerMonth = 0
                If (txtFuelPerHr = 0) Then
                    txtDrive = "Electrical"
                Else
                    txtDrive = "Fuel (Diesel)"
                End If
                txtshifts = 0
                txtMaintCostperMC_PerMonth = (MinorEquipmentCost * RAndMPercentage / 100)
            Next
        End If
        mOledbDataAdapter = Nothing

        txtFuelperUnitPerMonth = 0
        If UCase(CategoryTextBoxes(intI).Text) = UCase("Conveyance") Or UCase(EquipNameTextBoxes(intI).Text) = UCase("Tipper") Or _
              UCase(EquipNameTextBoxes(intI).Text) = UCase("Truck") Or UCase(CategoryTextBoxes(intI).Text) = UCase("HiredConveyance") Or _
              UCase(CategoryTextBoxes(intI).Text) = UCase("Matl Transport") Then
            txtFuelperUnitPerMonth = HrsPerMonth / txtFuelPerHr
        Else
            txtFuelperUnitPerMonth = txtFuelPerHr * HrsPerMonth
        End If
        txtFuelCostPerMonth = txtFuelperUnitPerMonth * QtyTextBoxes(intI).Text * FuelCostperLtr
        txtFuelCostProject = txtFuelCostPerMonth * txtMonths
        txtPowerCostperMonth = txtPowerperHr * HrsPerMonth * Val(QtyTextBoxes(intI).Text) * PowerCostPerUnit
        txtPowerCostProject = txtPowerCostperMonth * txtMonths
        txtOprCostPerMonth = txtOprCostPerMCPerMonth * Val(QtyTextBoxes(intI).Text) * txtshifts
        txtOprCostProject = txtOprCostPerMonth * txtMonths
        txtConsumablesPerMonth = txtMaintCostperMC_PerMonth * Val(QtyTextBoxes(intI).Text)
        txtConsumablesProject = txtConsumablesPerMonth * txtMonths
        OperatingCost_MajorEquips = txtHirecharges + txtFuelCostProject + txtPowerCostProject + _
              txtOprCostProject + txtConsumablesProject
    End Sub
    Private Sub ComputeHEValues(ByVal intI As Integer)
        Dim AllValid As Boolean
        AllValid = True
        Dim msgstring As String, strsql As String
        msgstring = "The following were not entered/selected" & vbNewLine
        If (CategoryTextBoxes(intI).Text = "") Then
            msgstring = msgstring & "*** Equipment catefory not selected for the " & intI & _
                IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If
        If UCase(CategoryTextBoxes(intI).Text) = UCase("Concrete") And (concreteqtyTextboxes(intI).Text = "" Or _
            Val(concreteqtyTextboxes(intI).Text) = 0) Then
            msgstring = msgstring & "*** Concrete Qunatity is not entered or is zero for the " & intI & _
                 IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If
        If (EquipNameTextBoxes(intI).Text = "") Then
            msgstring = msgstring & "*** Equipment Name not entered for the " & intI & _
                 IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If

        If (CapacityTextBoxes(intI).Text = "") Then
            msgstring = msgstring & "*** Equipment capacity not entered for the " & intI & _
                 IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If
        If (MakeModelTextBoxes(intI).Text = "") Then
            msgstring = msgstring & "*** Equipment Make & Model not entered for the " & intI & _
              IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If
        If (QtyTextBoxes(intI).Text = "" Or Len(Trim(QtyTextBoxes(intI).Text)) = 0) Then
            msgstring = msgstring & "*** Equipment Quantity not entered for the " & intI & _
                IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If
        If (MobdatePickers(intI).Value.ToString() = "") Then
            msgstring = msgstring & "*** Equipment mobilisation date  not specfied for the " & intI & _
                 IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If
        If (DemobDatePickers(intI).Value.ToString() = "") Then
            msgstring = msgstring & "*** Equipment Demobilisation date  not specfied for the " & intI & _
                 IIf(intI = 1, "st", IIf(intI = 2, "nd", IIf(intI = 3, "rd", "th"))) & "item" & vbNewLine
            AllValid = False
        End If
        If Not AllValid Then Exit Sub

        lblmessage.Visible = True
        txtMonths = System.Math.Round((DemobDatePickers(intI).Value.Date - MobdatePickers(intI).Value.Date).Days / 30, 0)

        txtmake = Strings.Trim(Strings.Left(MakeModelTextBoxes(intI).Text, Strings.InStr(MakeModelTextBoxes(intI).Text, " /") - 1))
        txtmodel = Strings.Trim(Strings.Mid(MakeModelTextBoxes(intI).Text, InStr(MakeModelTextBoxes(intI).Text, "/ ") + 1))
        strsql = "categoryname = '" & CategoryTextBoxes(intI).Text & "' And EquipmentName = '" & EquipNameTextBoxes(intI).Text & "' and " & _
           "Make ='" & txtmake & "' and Model ='" & txtmodel & "' and Capacity ='" & CapacityTextBoxes(intI).Text & "'"
        strStatement = "Select * from HiredEquipments where " & strsql
        'MsgBox(moledbConnection.ToString)
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

        mOledbDataAdapter.Fill(mDataSet, "Equipments")

        Dim Machine As DataRow
        For Each Machine In mDataSet.Tables("Equipments").Rows
            txtDepreciation = 0
            HrsPerMonth = Val(HrsPermonthTextBoxes(intI).Text)
            txtFuelPerHr = Val(Machine("Fuel_PerHour").ToString())
            txtPowerperHr = Val(Machine("Power_PerHour").ToString())
            txtOprCostPerMCPerMonth = 0
            If (txtFuelPerHr = 0) Then
                txtDrive = "Electrical"
            Else
                txtDrive = "Fuel (Diesel)"
            End If
            txtshifts = 0
            txtMaintCostperMC_PerMonth = 0
        Next
        mOledbDataAdapter = Nothing

        If UCase(CategoryTextBoxes(intI).Text) = UCase("Conveyance") Or UCase(EquipNameTextBoxes(intI).Text) = UCase("Tipper") Or _
              UCase(EquipNameTextBoxes(intI).Text) = UCase("Truck") Then
            txtFuelperUnitPerMonth = HrsPerMonth / txtFuelPerHr
        Else
            txtFuelperUnitPerMonth = txtFuelPerHr * HrsPerMonth
        End If
        txtFuelCostPerMonth = txtFuelperUnitPerMonth * QtyTextBoxes(intI).Text * FuelCostperLtr
        txtFuelCostProject = txtFuelCostPerMonth * txtMonths
        txtPowerCostperMonth = txtPowerperHr * HrsPerMonth * Val(QtyTextBoxes(intI).Text) * PowerCostPerUnit
        txtPowerCostProject = txtPowerCostperMonth * txtMonths
        txtOprCostPerMonth = txtOprCostPerMCPerMonth * Val(QtyTextBoxes(intI).Text) * txtshifts
        txtOprCostProject = txtOprCostPerMonth * txtMonths
        txtConsumablesPerMonth = txtMaintCostperMC_PerMonth * Val(QtyTextBoxes(intI).Text)
        txtConsumablesProject = txtConsumablesPerMonth * txtMonths
        OperatingCost_MajorEquips = txtHirecharges + txtFuelCostProject + txtPowerCostProject + _
              txtOprCostProject + txtConsumablesProject
        Me.Button1.Enabled = True
    End Sub
    Private Function getCategory(ByVal addedcat As String) As String
        Select Case addedcat
            Case "MajorConcreteEquips"
                Return "Concreting"
            Case "MajorConveyanceEquips"
                Return "Conveyance"
            Case "MajorCraneEquips"
                Return "Cranes"
            Case "MajorDGSetsEquips"
                Return "DG Sets"
            Case "MajorMaterialhandlingEquips"
                Return "Material Handling"
            Case "MajorNonConcreteEquips"
                Return "Non Concreting"
            Case "MajorOthers"
                Return "MajOthers"
        End Select
    End Function

    Private Sub AddRowsinPowergenCost()
        Dim oAdapter1 As OleDbDataAdapter, ds1 As DataSet
        Dim oAdapter2 As OleDbDataAdapter, ds2 As DataSet, mcat As String
        Dim Categories() As String = {"MajorConcreteEquips", "MajorConveyanceEquips", "MajorCraneEquips", "MajorDGSetsEquips", "MajorMaterialhandlingEquips", _
        "MajorNonConcreteEquips", "MajorOthers", "MinorEquips", "HiredEquips"}
        Dim row1 As DataRow, Row2 As DataRow, txtDepreciationPerc As Single, txtFuelperLtr As Single
        Dim txtcategory As String, txtEname As String, txtCapacity As String
        Dim msheetname As String = "PowerGen Cost", cursheetno As Integer
        Dim txtMdate As String, txtDMDate As String
        txtcategory = ""
        txtEname = ""
        txtMdate = ""
        txtDMDate = ""
        txtCapacity = ""

        Dim intK As Integer  ' , intL As Integer
        For intK = 0 To Categories.Length - 1
            mcat = Categories(intK)
            ' mcat = "HiredEquips"
            Dim str1 As String = "Select * from " & mcat & " where chkboxno = 1 and right(Capacity,3) = 'KVA'"
            Dim str2 As String
            If moledbConnection1 Is Nothing Then
                strconnection1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
                moledbConnection1 = New OleDbConnection(strconnection1)
            End If
            If (moledbConnection1.State.ToString().Equals("Closed")) Then
                moledbConnection1.Open()
            End If
            oAdapter2 = New OleDbDataAdapter(str1, moledbConnection1)
            ds2 = New DataSet()
            oAdapter2.Fill(ds2, "Added")

            If ds2.Tables("Added").Rows.Count > 0 Then
                For Each Row2 In ds2.Tables("Added").Rows
                    str2 = "Select * from "
                    If intK <= 6 Then
                        str2 = str2 & "MajorEquipments where Equipmentname = '" & Row2("Description").ToString & "' And Capacity ='" & _
                           Row2("Capacity").ToString() & "' And Make ='" & Row2("Make").ToString() & "' And Model ='" & Row2("Model").ToString() & "'"
                    ElseIf intK = 7 Then
                        str2 = str2 & "MinorEquipments where Equipmentname = '" & Row2("Description").ToString & "' And Capacity ='" & _
                           Row2("Capacity").ToString() & "' And Make ='" & Row2("Make").ToString() & "' And Model ='" & Row2("Model").ToString() & "'"
                    ElseIf intK = 8 Then
                        str2 = str2 & "HiredEquipments where Equipmentname = '" & Row2("Description").ToString & "' And Capacity ='" & _
                           Row2("Capacity").ToString() & "' And Make ='" & Row2("Make").ToString() & "' And Model ='" & Row2("Model").ToString() & "'"
                    End If
                    If moledbConnection Is Nothing Then
                        strconnection1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
                        moledbConnection = New OleDbConnection(strconnection1)
                    End If
                    If (moledbConnection.State.ToString().Equals("Closed")) Then
                        moledbConnection.Open()
                    End If
                    oAdapter1 = New OleDbDataAdapter(str2, moledbConnection)
                    ds1 = New DataSet()
                    oAdapter1.Fill(ds1, "Master")
                    If ds1.Tables("Master").Rows.Count > 0 Then
                        For Each row1 In ds1.Tables("Master").Rows
                            txtcategory = getCategory(mcat)
                            txtEname = Row2("Description").ToString()
                            txtCapacity = Row2("Capacity").ToString()
                            txtmake = Row2("Make").ToString()
                            txtmodel = Row2("Model").ToString()
                            txtMdate = Row2("MobDate").ToString()
                            txtDMDate = Row2("DemobDate").ToString()
                            Dim days As TimeSpan = DateValue(txtMdate).Date - DateValue(txtDMDate).Date
                            'txtMonths = Round(days / 30, 0)
                            If intK <> 8 Then
                                txtRepvalue = Val(row1("Repvalue").ToString())
                            Else
                                txtRepvalue = Val(row1("Hirecharges").ToString())
                            End If

                            txtDepreciation = 0
                            HrsPerMonth = Val(Row2("HrsperMonth").ToString())
                            txtFuelPerHr = Val(row1("Fuel_PerHour").ToString())
                            If intK <= 6 Then
                                txtDepreciationPerc = Val(Row2("DepPerc").ToString())
                                If txtDepreciationPerc = 2.75 Then
                                    txtDepreciation = Val(row1("RAndMPer_275").ToString())
                                ElseIf txtDepreciationPerc = 1.25 Then
                                    txtDepreciation = Val(row1("RAndMPer_125").ToString())
                                Else
                                    txtDepreciation = Val(row1("RAndMPerc_050").ToString())
                                End If
                            End If
                            txtOprCostPerMCPerMonth = Val(row1("OperatorCost_PerMonth").ToString())
                            txtshifts = Val(Row2("Shifts").ToString())
                            txtMaintCostperMC_PerMonth = txtRepvalue * (txtMaintPercPerMC_PerMonth / 100)
                            txtFuelperLtr = FuelCostperLtr
                        Next
                        txtOprCostPerMonth = txtOprCostPerMCPerMonth * Val(Row2("Qty").ToString()) * txtshifts
                        txtConsumablesPerMonth = txtMaintCostperMC_PerMonth * Val(Row2("Qty").ToString())
                    End If
                    xlWorksheet = xlWorkbook.Sheets(msheetname)
                    xlWorksheet.Activate()
                    cursheetno = getSheetNo(xlWorksheet.Name)
                    getCategoryShortname(xlWorksheet)
                    RowCount = 0
                    intj = 1
                    xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
                    xlRange.Value = "Fuel Cost for all mc/month @Rs. " & FuelCostperLtr & " per Lt"
                    xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                    xlRange.Value = "Opr Cost for all m/c "  'with " & Me.cmbShifts.Text & " shift(s) per day"
                    xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(RecordsInserted(cursheetno) + 1, 0)
                    xlRange.Value = RecordsInserted(cursheetno) + 1
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = txtcategory
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = txtEname
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = txtCapacity
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = Val(Row2("Qty").ToString())
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 1) ''
                    xlRange.Value = DateValue(txtMdate)
                    xlRange.NumberFormat = "dd-mmm-yyyy"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = DateValue(txtDMDate)
                    xlRange.NumberFormat = "dd-mmm-yyyy"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 2)
                    If msheetname = "Minor Eqpts" Then ' Check from here
                        xlRange.Value = MinorEquipmentCost
                    Else
                        xlRange.Value = txtRepvalue
                    End If
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = txtDepreciation  ' calcualted as Repvalue * depperc
                    xlRange.NumberFormat = "#0.00"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 3)
                    xlRange.Value = HrsPerMonth
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = txtFuelPerHr   ' take from Equipmentsmasters.mdb
                    xlRange.NumberFormat = "##0.0#"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 4)
                    xlRange.Value = FuelCostperLtr   ' take from global variable 
                    xlRange.NumberFormat = "##0.0#"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 3)
                    xlRange.Value = txtOprCostPerMCPerMonth  ' from master database 
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 1)
                    xlRange.Value = txtshifts   ' from addedequipments database 
                    xlRange.NumberFormat = "#.0#"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.RowHeight = 33
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 3)
                    xlRange.Value = txtConsumablesPerMonth '* Val(Me.txtMajorEquipQty.Text))  ' calculated
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    xlRange = xlRange.Offset(0, 4)
                    xlRange.Value = "Remarks"
                    xlRange.Cells.Font.Size = 12
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    RecordsInserted(cursheetno) = RecordsInserted(cursheetno) + 1
                    'MsgBox(SheetNo)
                    If RecordsInserted(cursheetno) > 1 Then FillFormulas(cursheetno, RecordsInserted(cursheetno))
                    xlRange = xlWorksheet.Range(Category_Shortname & "RecordsTotal")
                    xlRange.Value = RecordsInserted(cursheetno)
                Next
            End If
        Next
    End Sub
    Private Sub FillFormulas(ByVal sheetno As Integer, ByVal record As Integer)

        xlWorksheet = xlWorkbook.Sheets(sheetno)
        xlWorksheet.Activate()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "Months").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)

        xlRange = xlWorksheet.Range(Category_Shortname & "Depreciation").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        If Category_Shortname = "Ext_" Or Category_Shortname = "ExtOthers_" Then
            xlRange = xlWorksheet.Range(Category_Shortname & "Hire_Charges_PerMonth").Offset(1, 0)
            xlRange.Copy()
            xlRange = xlRange.Offset(record - 1, 0)
            xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        End If

        xlRange = xlWorksheet.Range(Category_Shortname & "Hire_Charges").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Per_Mc").Offset(1, 0)
        If xlWorksheet.Name = "PowerGen Cost" Then
            xlRange.Copy()
            xlRange = xlRange.Offset(record - 1, 0)
            xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        Else
            If Not (xlWorksheet.Name = "external Hire") Then   'Or xlWorksheet.Name = "PowerGen Cost") Then
                xlRange.Copy()
                xlRange = xlRange.Offset(record - 1, 0)
                If Category_Shortname = "Ext_" Then
                    If hiredCategoryNames(record - 1) = "HiredConveyance" Then
                        xlRange.Formula = "=Round(RC[-2]/RC[-1],0)"
                    End If
                ElseIf Category_Shortname = "Conv_" Then
                    xlRange.Formula = "=Round(RC[-2]/RC[-1],0)"
                ElseIf Category_Shortname = "MH_" Then
                    If (MHEquipsNames(record - 1) = "Tipper" Or MHEquipsNames(record - 1) = "Truck") Then
                        xlRange.Formula = "=Round(RC[-2]/RC[-1],0)"
                    End If
                Else
                    xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
                End If
            End If
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        End If
        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Total").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_For_Project").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_Project").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        If Not xlWorksheet.Name = "PowerGen Cost" Then
            xlRange = xlWorksheet.Range(Category_Shortname & "Power_Per_Month").Offset(1, 0)
            xlRange.Copy()
            xlRange = xlRange.Offset(record - 1, 0)
            xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

            xlRange = xlWorksheet.Range(Category_Shortname & "Power_Cost_Project").Offset(1, 0)
            xlRange.Copy()
            xlRange = xlRange.Offset(record - 1, 0)
            xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        End If

        xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Project").Offset(1, 0)
        xlRange.Copy()
        'MsgBox(xlRange.Formula)
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Consumables_Project").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        'If Not (Category_Shortname = "Ext_" Or Category_Shortname = "Min_") Then

        xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Month").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        'End If

        xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Project").Offset(1, 0)
        xlRange.Copy()
        xlRange = xlRange.Offset(record - 1, 0)
        xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    End Sub
    Private Sub GetTotals()
        Dim FormulaString As String
        For SheetNo = 2 To 8
            With xlApp
                If xlWorkbook Is Nothing Then
                    xlWorkbook = .Workbooks.Open(xlFilename)
                End If
            End With
            xlWorksheet = xlWorkbook.Sheets(SheetNo)
            xlWorksheet.Select()

            If RecordsInserted(SheetNo) > 1 Then
                getCategoryShortname(xlWorksheet)
                xlRange = xlWorksheet.Range(Category_Shortname & "Equip_Name")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = "Total"
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "Depreciation")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "Hire_Charges")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Per_Mc")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Total")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_For_Project")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_Project")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "Power_Per_Month")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "Power_Cost_Project")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Project")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "Consumables_Project")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Month")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Project")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
                xlRange.FormulaR1C1 = FormulaString
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

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
                    .Font.Size = 12
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
    Private Sub GetPowerGenTotal()
        Dim FormulaString As String, msheetno As Integer
        With xlApp
            If xlWorkbook Is Nothing Then
                xlWorkbook = .Workbooks.Open(xlFilename)
            End If
        End With
        xlWorksheet = xlWorkbook.Worksheets("PowerGen Cost")
        xlWorksheet.Select()
        msheetno = getSheetNo(xlWorksheet.Name)

        getCategoryShortname(xlWorksheet)
        SheetNo = getSheetNo(xlWorksheet.Name)
        If RecordsInserted(SheetNo) <= 1 Then Exit Sub
        'If RecordsInserted(SheetNo) = 1 Then
        '    xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(1, 0)
        '    x1 = xlRange.Address
        '    xlRange.Select()
        '    'xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
        '    xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight)
        '    x2 = xlRange.Address
        '    xlRange = xlWorksheet.Range(x1, x2)
        '    RangeCols = xlRange.Columns.Count() - 1
        '    With xlRange
        '        .BorderAround(, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
        '        .Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        '        .Interior.ColorIndex = 2
        '        .Font.Bold = False
        '        .Font.Size = 12
        '        '.Interior.Pattern = 1
        '    End With
        'End If
        xlRange = xlWorksheet.Range(Category_Shortname & "Equip_Name")
        'If xlRange.Value = "" Or Len(Trim(xlRange.Value)) = 0 Then Exit Sub
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        xlRange.Select()
        'If xlRange.Value = "" Or Len(Trim(xlRange.Value)) = 0 Then Exit Sub
        xlRange.Value = "Total"
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Months")
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter


        xlRange = xlWorksheet.Range(Category_Shortname & "Depreciation")
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Hire_Charges")
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Per_Mc")
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Per_Month_Total")
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_For_Project")
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_per_Month")
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Fuel_Cost_Project")
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Per_Month_2Shift")
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "OprCost_Project")
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "Consumables_Project")
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Month")
        'MsgBox(xlRange.Address)
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        'MsgBox(xlRange.Address)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

        xlRange = xlWorksheet.Range(Category_Shortname & "OperatingCost_Project")
        xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
        FormulaString = "=sum(R[-" & RecordsInserted(SheetNo) & "]C:R[-1]C"
        xlRange.FormulaR1C1 = FormulaString
        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

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
            .Font.Size = 12
            '.Interior.Pattern = 1
        End With

        xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(1, 0)
        x1 = xlRange.Address
        'MsgBox(xlRange.Address)
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
    End Sub
    Private Sub PopulateSelectedMajorPowerEquipments()
        Dim strconnection1 As String, SelectCommand As String, FormulaString As String
        Dim madapter As OleDbDataAdapter, mdataset As New DataSet, mrow As DataRow
        Dim majorCategories() As String = {"MajorConcreteEquips", "MajorConveyanceEquips", "MajorCraneEquips", _
             "MajorDGSetsEquips", "MajorMaterialhandlingEquips", "MajorNonConcreteEquips", "MajorOthers"}
        Dim Index As Integer, cntr As Integer = 0

        Dim strSql As String ', mOledbDataAdapter3 As OleDbDataAdapter
        Dim InsertCommand As String ', Machine As DataRow
        Dim moleDbInsertComamnd As OleDbCommand
        Dim moledbDataSet3 As New DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False

        Me.lblmessage.Text = "Major Power Equipments Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()
        xlWorksheet = xlWorkbook.Worksheets("PowerReqmt")
        xlWorksheet.Select()
        getCategoryShortname(xlWorksheet)
        xlRange = xlWorksheet.Range(Category_Shortname & "Client")
        DeleteRecordsFromAddedItemsTable("PowerEquips")
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
        If moledbConnection Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        If moledbConnection1 Is Nothing Then
            strconnection1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
            moledbConnection1 = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection1.Open()
        End If


        For Index = 0 To majorCategories.Length - 1
            strSql = "Select * from " & majorCategories(Index) & "  where Drive = 'Electrical' and  ChkBoxNo = 1"
            madapter = New OleDbDataAdapter(strSql, moledbConnection1)
            mdataset = New DataSet
            madapter.Fill(mdataset, "MajorElectEquipments")
            If mdataset.Tables("MajorElectEquipments").Rows.Count > 0 Then
                For Each mrow In mdataset.Tables("MajorElectEquipments").Rows
                    InsertCommand = ""
                    InsertCommand = "INSERT INTO MajorPowerEquipments values('" & mrow("Description").ToString() & "',"
                    InsertCommand = InsertCommand & "'" & mrow("Capacity").ToString() & "',"
                    InsertCommand = InsertCommand & "'" & mrow("Make").ToString() & "',"
                    InsertCommand = InsertCommand & "'" & mrow("Model").ToString() & "',"
                    InsertCommand = InsertCommand & "'" & mrow("Mobdate").ToString() & "',"
                    InsertCommand = InsertCommand & "'" & mrow("Demobdate").ToString() & "',"
                    InsertCommand = InsertCommand & Val(mrow("Qty").ToString()) & ","
                    InsertCommand = InsertCommand & Val(mrow("PowerPerUnit(HP)").ToString()) & ","
                    InsertCommand = InsertCommand & Val(mrow("ConnectedLoadPerMC").ToString()) & ","
                    InsertCommand = InsertCommand & Val(mrow("UtilityFactor").ToString()) & ")"

                    Try
                        If (moledbConnection1.State.ToString().Equals("Closed")) Then
                            moledbConnection1.Open()
                        End If
                        moleDbInsertComamnd = New OleDbCommand
                        moleDbInsertComamnd.CommandType = CommandType.Text
                        moleDbInsertComamnd.CommandText = InsertCommand
                        moleDbInsertComamnd.Connection = moledbConnection1
                        moleDbInsertComamnd.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.ToString())
                    Finally
                        moleDbInsertComamnd = Nothing
                        'moledbconnection3.Close()
                    End Try
                    moleDbInsertComamnd = Nothing
                Next
            End If
            madapter = Nothing
            mdataset = Nothing
        Next

        If moledbConnection1 Is Nothing Then
            strconnection1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
            moledbConnection1 = New OleDbConnection(strconnection1)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection1.Open()
        End If
        SelectCommand = "Select * from MajorPowerEquipments"
        madapter = New OleDbDataAdapter(SelectCommand, moledbConnection1)
        mdataset = New DataSet
        madapter.Fill(mdataset, "MajorPowerEquipments")
        xlWorksheet = xlWorkbook.Worksheets("PowerReqmt")
        cntr = 0
        If mdataset.Tables("MajorPowerEquipments").Rows.Count > 0 Then
            For Each mrow In mdataset.Tables("MajorPowerEquipments").Rows
                cntr = cntr + 1
                xlRange = xlWorksheet.Range("PowerReq_MajEquipsTotal")
                xlRange.Select()
                If cntr > 1 Then
                    xlApp.DisplayAlerts = False
                    xlRange.Application.ActiveCell.EntireRow.Insert()
                    xlRange = xlWorksheet.Range(xlRange.Application.ActiveCell.Address)
                Else
                    xlRange = xlWorksheet.Range("PowerReq_MajEquipsTotal").Offset(-1, 0)
                    xlRange.Select()
                End If
                xlRange.Value = cntr
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = mrow("Description").ToString()
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = mrow("Capacity").ToString()
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = mrow("Make").ToString() & "/" & mrow("Model").ToString()
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = Val(mrow("Qty").ToString())
                xlRange.NumberFormat = "##0"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = Val(mrow("PowerPerUnit(HP)").ToString())
                xlRange.NumberFormat = "##0.000"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = mrow("MobDate").ToString
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = mrow("DemobDate").ToString()
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 2)
                xlRange.Formula = Val(mrow("ConnectedLoadPerMC").ToString())
                xlRange.NumberFormat = "##0.000"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = Val(mrow("UtilityFactor").ToString())
                xlRange.NumberFormat = "##0.000"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range("$A$" & xlRange.Application.ActiveCell.Row)
                xlRange.Select()

                xlRange = xlWorksheet.Range(xlRange.Address, xlRange.Offset(0, 12).Address)
                With xlRange
                    .BorderAround(, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin)
                    .Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                End With
                With xlRange
                    .Font.Size = 12
                    .Font.Bold = False
                End With
            Next

            Dim Times As Integer = 0
            xlRange = xlWorksheet.Range("PowerReq__MajEquipsFormula1")
            xlRange.Cells.Font.Size = 12
            xlRange.Cells.RowHeight = 33
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            For Times = 1 To cntr - 1
                xlRange = xlWorksheet.Range("PowerReq__MajEquipsFormula1")
                xlRange.Copy()
                xlRange = xlRange.Offset(Times, 0)
                xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
                With xlRange
                    .Font.Bold = False
                    .Font.Size = 12
                    '.Interior.Pattern = 1
                End With
            Next
            xlRange = xlWorksheet.Range("PowerReq__MajEquipsFormula2")
            xlRange.Cells.Font.Size = 12
            xlRange.Cells.RowHeight = 33
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            For Times = 1 To cntr - 1
                xlRange = xlWorksheet.Range("PowerReq__MajEquipsFormula2")
                xlRange.Copy()
                xlRange = xlRange.Offset(Times, 0)
                xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
                With xlRange
                    .Font.Bold = False
                    .Font.Size = 12
                    '.Interior.Pattern = 1
                End With
            Next
            xlRange = xlWorksheet.Range("PowerReq__MajEquipsFormula3")
            xlRange.Cells.Font.Size = 12
            xlRange.Cells.RowHeight = 33
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            For Times = 1 To cntr - 1
                xlRange = xlWorksheet.Range("PowerReq__MajEquipsFormula3")
                xlRange.Copy()
                xlRange = xlRange.Offset(Times, 0)
                xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
                With xlRange
                    .Font.Bold = False
                    .Font.Size = 12
                    '.Interior.Pattern = 1
                End With
            Next

            xlRange = xlWorksheet.Range("PowerReq__MajEquipsFormula3").Offset(cntr, 0)
            xlRange.Select()
            FormulaString = "=sum(R[-" & cntr & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
        End If
        madapter = Nothing
        mdataset = Nothing
    End Sub
    Private Sub PopulateSelectedMinorPowerEquipments()
        Dim strconnection1 As String, SelectCommand As String, FormulaString As String
        Dim madapter As OleDbDataAdapter, mdataset As New DataSet, mrow As DataRow

        Me.lblmessage.Text = "Minor Power Equipments Data being saved. Please wait...."
        Me.lblmessage.Visible = True
        Me.Refresh()
        xlRange = xlWorksheet.Range(Category_Shortname & "Client")
        If Len(Trim(xlRange.Value)) = 0 Then
            xlRange.Value = mMainTitle1
            xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
            xlRange.Value = mMainTitle2
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
            'xlRange.Value = mMainTitle3
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            xlRange = xlWorksheet.Range(Category_Shortname & "Client")
            xlRange.Value = mClient
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            xlRange = xlWorksheet.Range(Category_Shortname & "Location")
            xlRange.Value = mLocation
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            xlRange = xlWorksheet.Range(Category_Shortname & "StartDate")
            xlRange.Value = mStartDate.Date.ToString()
            xlRange.NumberFormat = "dd-mmm-yyyy"
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            xlRange = xlWorksheet.Range(Category_Shortname & "EndDate")
            xlRange.Value = mEndDate.Date.ToString()
            xlRange.NumberFormat = "dd-mmm-yyyy"
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            xlRange = xlWorksheet.Range(Category_Shortname & "ProjectValue")
            xlRange.Value = mProjectvalue
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
        End If

        If moledbConnection1 Is Nothing Then
            strconnection1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
            moledbConnection1 = New OleDbConnection(strconnection1)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection1.Open()
        End If
        SelectCommand = "Select * from MinorEquips where Drive = 'Electrical' and ChkBoxNo=1"
        madapter = New OleDbDataAdapter(SelectCommand, moledbConnection1)
        mdataset = New DataSet
        madapter.Fill(mdataset, "MinorPowerEquipments")
        xlWorksheet = xlWorkbook.Worksheets("PowerReqmt")
        Dim cntr = 0
        If mdataset.Tables("MinorPowerEquipments").Rows.Count > 0 Then
            For Each mrow In mdataset.Tables("MinorPowerEquipments").Rows
                cntr = cntr + 1
                xlRange = xlWorksheet.Range("PowerReq_MinorEquipsTotal")
                xlRange.Select()
                If cntr > 1 Then
                    xlApp.DisplayAlerts = False
                    xlRange.Application.ActiveCell.EntireRow.Insert()
                    xlRange = xlWorksheet.Range(xlRange.Application.ActiveCell.Address)
                Else
                    xlRange = xlWorksheet.Range("PowerReq_MinorEquipsTotal").Offset(-1, 0)
                    xlRange.Select()
                End If
                xlRange.Value = cntr
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = mrow("Description").ToString()
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = mrow("Capacity").ToString()
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = mrow("Make").ToString() & "/" & mrow("Model").ToString()
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = Val(mrow("Qty").ToString())
                xlRange.NumberFormat = "##0"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = Val(mrow("PowerPerUnit(HP)").ToString())
                xlRange.NumberFormat = "##0.000"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = mrow("MobDate").ToString
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = mrow("DemobDate").ToString()
                xlRange.NumberFormat = "dd-mmm-yyyy"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 2)
                xlRange.Formula = Val(mrow("ConnectedLoadPerMC").ToString())
                xlRange.NumberFormat = "##0.000"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                xlRange = xlRange.Offset(0, 1)
                xlRange.Value = Val(mrow("UtilityFactor").ToString())
                xlRange.NumberFormat = "##0.000"
                xlRange.Cells.Font.Size = 12
                xlRange.Cells.RowHeight = 33
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                xlRange = xlWorksheet.Range("$A$" & xlRange.Application.ActiveCell.Row)
                xlRange.Select()

                xlRange = xlWorksheet.Range(xlRange.Address, xlRange.Offset(0, 12).Address)
                With xlRange
                    .BorderAround(, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin)
                    .Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                End With
                With xlRange
                    .Font.Size = 12
                    .Font.Bold = False
                End With
            Next

            Dim Times As Integer = 0

            For Times = 1 To cntr - 1
                xlRange = xlWorksheet.Range("PowerReq__MinorEquipsFormula1")
                xlRange.Copy()
                xlRange = xlRange.Offset(Times, 0)
                xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
                With xlRange
                    .Font.Bold = False
                    .Font.Size = 12
                    '.Interior.Pattern = 1
                End With
            Next

            For Times = 1 To cntr - 1
                xlRange = xlWorksheet.Range("PowerReq__MinorEquipsFormula2")
                xlRange.Copy()
                xlRange = xlRange.Offset(Times, 0)
                xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
                With xlRange
                    .Font.Bold = False
                    .Font.Size = 12
                    '.Interior.Pattern = 1
                End With
            Next

            For Times = 1 To cntr - 1
                xlRange = xlWorksheet.Range("PowerReq__MinorEquipsFormula3")
                xlRange.Copy()
                xlRange = xlRange.Offset(Times, 0)
                xlRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
                With xlRange
                    .Font.Bold = False
                    .Font.Size = 12
                    '.Interior.Pattern = 1
                End With
            Next

            xlRange = xlWorksheet.Range("PowerReq__MinorEquipsFormula3").Offset(cntr, 0)
            xlRange.Select()
            FormulaString = "=sum(R[-" & cntr & "]C:R[-1]C"
            xlRange.FormulaR1C1 = FormulaString
        End If
        madapter = Nothing
        mdataset = Nothing
    End Sub
    Private Sub btnQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuit.Click
        Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
        DataSaved = True
        Me.Button1_Click(sender, e)
        If Not DataSaved Then
            Exit Sub
        End If

        Me.lblmessage.Text = "Final processing in progress. Please wait till the Processing is Completed......"
        Me.lblmessage.Visible = True
        Me.Refresh()
        AddRowsinPowergenCost()

        CopyPowerReqTemplate()
        PopulateSelectedMajorPowerEquipments()
        PopulateSelectedMinorPowerEquipments()
        SaveLightingDataInSheet("PowerReqmt")

        GetTotals()
        GetTotals_MinorEquipments()
        GetTotals_ExternalEquipments()
        GetPowerGenTotal()

        SetRangeNamesForTotals()

        For Each xlworksheet In xlWorkbook.Worksheets
            If xlworksheet.Visible Then
                xlworksheet.Activate()
                'getCategoryShortname(xlworksheet)
                'MsgBox(Category_Shortname)
                xlRange = xlworksheet.Range("A1")
                xlRange.Select()
            End If
        Next


        If Not moledbConnection.State.ToString().Equals("Closed") Then
            moledbConnection.Close()
            moledbConnection = Nothing
        End If
        If Not moledbConnection1.State.ToString().Equals("Closed") Then
            moledbConnection1.Close()
            moledbConnection1 = Nothing
        End If
        If Not xlWorkbook Is Nothing Then
            xlworksheet = xlWorkbook.Sheets.Item(1)
            xlworksheet.Select()
            xlApp.CalculateBeforeSave = True
            xlApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic
            xlApp.CalculateFull()
            xlWorkbook.Save()
            xlWorkbook.Close()
            xlWorkbook = Nothing
        End If
        endtime = Now()
        Dim timeelapsed As String = (endtime - starttime).Minutes & " Minutes " & (endtime - starttime).Seconds & " Seconds"
        MsgBox("Time taken for Processing " & timeelapsed)
        Me.Close()
        xlApp.Quit()
        xlApp = Nothing
        System.GC.Collect()
        lblmessage.Visible = False
        frmProjectDetails.Show()
    End Sub
    Private Sub GetTotals_ExternalEquipments()
        Dim FormulaString As String
        Dim intI As Integer, currentsheetname1 As String, currentsheetname2 As String
        currentsheetname2 = "external Hire"
        For intI = 1 To 2
            'Dim no As Integer = getSheetNo(Extsheets(intI - 1))
            If intI = 1 Then
                xlWorksheet = xlWorkbook.Sheets("external Hire")
            Else
                xlWorksheet = xlWorkbook.Sheets("External Others")
            End If
            getCategoryShortname(xlWorksheet)
            SheetNo = getSheetNo(xlWorksheet.Name)
            xlWorksheet.Select()
            If RecordsInserted(SheetNo) > 1 Then
                xlRange = xlWorksheet.Range(Category_Shortname & "Equip_Name")
                xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
                xlRange.Value = "Total"
                xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

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
                    .Font.Size = 12
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
                'xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight)
                'xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                x2 = xlRange.Address
                xlRange = xlWorksheet.Range(x1, x2)
                With xlRange
                    .Font.Size = 12
                    .Font.Bold = True
                End With
            End If
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
            Category_Shortname = "Min_"
            xlRange = xlWorksheet.Range(Category_Shortname & "Equip_Name")
            xlRange = xlRange.Offset(RecordsInserted(SheetNo) + 1, 0)
            xlRange.Value = "Total"
            xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

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
            xlRange = xlRange.Offset(1, 0)
            x2 = xlRange.Address
            xlRange = xlWorksheet.Range(x1, x2)
            RangeCols = xlRange.Columns.Count() - 1
            With xlRange
                .BorderAround(, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium)
                .Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .Interior.ColorIndex = 2
                .Font.Bold = False
                .Font.Size = 12
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

            xlRange = xlWorksheet.Range(Category_Shortname & "Slno")
            xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
            xlRange = xlRange.Offset(2, 1)
            x1 = xlRange.Address
            Dim r1 As Integer = xlRange.Application.ActiveCell.Row
            For intI = r1 To 10
                xlRange = xlWorksheet.Range(Category_Shortname & "SlNo").Offset(intI, 0)
                xlRange.Select()
                xlRange.Application.ActiveCell.EntireRow.Delete()
            Next
        End If
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
                Rangenamestring = CatShortnames(intI) & "HireChargesTotal"
                xlWorkbook.Names.Item(Rangenamestring).Delete()
                'MsgBox(CatShortnames(intI) & "Hire_Charges")
                xlRange = .Range(CatShortnames(intI) & "Hire_Charges")

                xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                xlRange.Name = Rangenamestring

                If CatShortnames(intI) = "Ext_" Then    'Or CatShortnames(intI) = "ExtOthers" Then
                    Rangenamestring = CatShortnames(intI) & "HireChargesPerMonthTotal"
                    xlWorkbook.Names.Item(Rangenamestring).Delete()
                    xlRange = .Range(CatShortnames(intI) & "Hire_Charges_PerMonth")
                    xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
                    xlRange.Name = Rangenamestring
                End If

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
        Rangenamestring = "PowerReq__PowerCost3Total"
        xlWorkbook.Names.Item(Rangenamestring).Delete()
        xlWorksheet = xlWorkbook.Worksheets("PowerReqmt")
        xlWorksheet.Select()
        If mWorkingMode = "New" Then
            xlWorksheet.Names.Item("PowerReq__PowerCost3Total").Delete()
        End If
        xlRange = xlWorksheet.Range("PowerReq__LightingFormula3")
        xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
        xlRange = xlRange.Offset(0, 12)
        xlRange.Name = Rangenamestring

        Rangenamestring = "PowerReq__TotalPowerCost"
        xlWorkbook.Names.Item(Rangenamestring).Delete()
        xlWorksheet = xlWorkbook.Worksheets("PowerReqmt")
        xlWorksheet.Select()
        If mWorkingMode = "New" Then
            On Error Resume Next
            xlWorksheet.Names.Item("PowerReq__TotalPowerCost").Delete()
        End If
        xlRange = xlWorksheet.Range("PowerReq__LightingFormula3")
        xlRange = xlRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
        xlRange = xlRange.Offset(2, 12)
        xlRange.Name = Rangenamestring
    End Sub

    Private Sub optMajConcrete_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMajConcrete.CheckedChanged
        Dim intL As Integer
        If optMinorEquips.Checked = True Then
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel4.BringToFront()
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optHireEquipments.Checked Then
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optFixedExp.Checked Or optBPFixedExp.Checked Then
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        ElseIf Me.optLightingEquips.Checked Then
            Panel7.Left = 1
            Panel7.Top = 20
            Panel7.Height = 30
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        Else
            Panel2.Left = 1
            Panel2.Top = 20
            Panel2.Height = 30
            Panel2.Visible = True
            Panel2.BringToFront()
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If
        If Not optMajConcrete.Checked Then
            Dim intK As Integer, L As Integer
            ConcretingItems = CategoryTextBoxes.Count
            L = ConcretingItems - 1  'CategoryTextBoxes.Count - 1
            ReDim Preserve ConcreteChecked(L)
            ReDim Preserve ConcreteMobdate(L)
            ReDim Preserve ConcreteDemobdate(L)
            ReDim Preserve ConcreteQty(L)
            ReDim Preserve ConcreteHrs(L)
            ReDim Preserve ConcreteDepPerc(L)
            ReDim Preserve ConcreteShifts(L)
            ReDim concCategory(L)
            ReDim concEName(L)
            ReDim concCapacity(L)
            ReDim concMakeModel(L)
            ReDim concMDate(L)
            ReDim concDMdate(L)
            ReDim concQuantity(L)
            ReDim concHrsPerMonth(L)
            ReDim concDep(L)
            ReDim concShift(L)
            For intK = 0 To L
                concCategory(intK) = CategoryTextBoxes(intK).Text
                concEName(intK) = EquipNameTextBoxes(intK).Text
                concCapacity(intK) = CapacityTextBoxes(intK).Text
                concMakeModel(intK) = MakeModelTextBoxes(intK).Text
                concMDate(intK) = MobdatePickers(intK).Value.Date
                concDMdate(intK) = DemobDatePickers(intK).Value.Date
                concQuantity(intK) = Val(QtyTextBoxes(intK).Text)
                concHrsPerMonth(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                concDep(intK) = Val(DepPercComboboxes(intK).Text)
                concShift(intK) = Val(ShiftsComboboxes(intK).Text)
            Next
            For intK = 0 To L
                If Checkboxes(intK).Checked Then
                    ConcreteChecked(intK) = 1
                    ConcreteMobdate(intK) = MobdatePickers(intK).Value.Date
                    ConcreteDemobdate(intK) = DemobDatePickers(intK).Value.Date
                    ConcreteQty(intK) = Val(QtyTextBoxes(intK).Text)
                    ConcreteHrs(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                    ConcreteDepPerc(intK) = Val(DepPercComboboxes(intK).Text)
                    ConcreteShifts(intK) = Val(ShiftsComboboxes(intK).Text)
                Else
                    ConcreteChecked(intK) = 0
                End If
            Next
            ConcretingItems = CategoryTextBoxes.Count
            WriteToConcEquipsArray(concreteitems)
            Exit Sub
        End If
        Button1.Enabled = True
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        mcategory = "Concreting"
        SelectedCategory = mcategory
        If Not concFirsttime Then
            LoadControlsFromConcArray(ConcretingItems)
            For intL = 0 To concreteitems - 1
                If ConcreteChecked(intL) = 1 Then
                    Checkboxes(intL).Checked = True
                    MobdatePickers(intL).Value = ConcreteMobdate(intL)
                    DemobDatePickers(intL).Value = ConcreteDemobdate(intL)
                    QtyTextBoxes(intL).Text = ConcreteQty(intL)
                    HrsPermonthTextBoxes(intL).Text = ConcreteHrs(intL)
                    DepPercComboboxes(intL).Text = ConcreteDepPerc(intL)
                    ShiftsComboboxes(intL).Text = ConcreteShifts(intL)
                Else
                    Checkboxes(intL).Checked = False
                End If
                'Checkboxes(intL + 1).Checked = IIf(ConcreteChecked(intL) = 1, True, False)
            Next
            Me.Refresh()
        Else
            mcategory = "Concreting"
            SelectedCategory = mcategory
            LoadControlsInpage(mcategory)
            For intL = 0 To Checkboxes.Count - 1
                If Checkboxes(intL).Checked Then
                    changeEnabled(True, intL)
                End If
            Next
            concFirsttime = False
        End If
        concreteitems = CategoryTextBoxes.Count
    End Sub
    Private Sub LoadControlsFromConcArray(ByVal mitems As Integer)
        Dim intI As Integer
        HP = 1
        VP = 0
        Dim mcategoryname As String = "Concreting"
        Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
        CategoryTextBoxes.Clear()
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        QtyTextBoxes.Clear()
        HrsPermonthTextBoxes.Clear()
        RepValueTextBoxes.Clear()
        concreteqtyTextboxes.Clear()
        Checkboxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        DepPercComboboxes.Clear()
        ShiftsComboboxes.Clear()
        AddButtons.Clear()
        mTabindex = 1
        For intI = 0 To mitems - 1
            Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
            Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
            Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
            Dim txtHrsPerMonth As New TextBox
            Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox
            Dim txtRepValue As New TextBox, txtMaintPerc As New TextBox, chkSelected As New CheckBox
            Dim txtConcreteQty As New TextBox   ', cmbDrive As New ComboBox
            Dim btnAddExtra As New Button

            chkSelected.Checked = False
            chkSelected.Height = 30
            chkSelected.Width = 24
            chkSelected.Left = HP
            chkSelected.Top = VP - 3
            HP = HP + chkSelected.Width + 1
            chkSelected.Tag = intI
            chkSelected.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler chkSelected.CheckedChanged, AddressOf HandleCheckboxStatus
            Checkboxes.Add(intI, chkSelected)

            txtCategory.Text = concCategory(intI)
            'txtCategory.Multiline = True
            txtCategory.WordWrap = True
            txtCategory.Height = 50
            txtCategory.Width = 65
            txtCategory.TextAlign = HorizontalAlignment.Center
            txtCategory.Left = HP
            txtCategory.Top = VP
            HP = HP + txtCategory.Width + 1
            txtCategory.Font = New Font("arial", 8)
            txtCategory.Tag = intI
            txtCategory.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCategory.Enabled = False
            CategoryTextBoxes.Add(intI, txtCategory)

            txtEquipname.Text = concEName(intI)
            ' txtEquipname.Multiline = True
            txtEquipname.WordWrap = True
            txtEquipname.Height = 60
            txtEquipname.Width = 150
            txtEquipname.TextAlign = HorizontalAlignment.Center
            txtEquipname.Left = HP
            txtEquipname.Top = VP
            HP = HP + txtEquipname.Width + 1
            txtEquipname.Font = New Font("arial", 8)
            txtEquipname.Tag = intI
            txtEquipname.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtEquipname.Enabled = False
            EquipNameTextBoxes.Add(intI, txtEquipname)

            txtCapacity.Text = concCapacity(intI)
            txtCapacity.WordWrap = True
            txtCapacity.Height = 30
            txtCapacity.Width = 70
            txtCapacity.TextAlign = HorizontalAlignment.Center
            txtCapacity.Left = HP
            txtCapacity.Top = VP
            HP = HP + txtCapacity.Width + 1
            txtCapacity.Font = New Font("arial", 8)
            txtCapacity.Tag = intI
            txtCapacity.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCapacity.Enabled = False
            CapacityTextBoxes.Add(intI, txtCapacity)

            txtMakeModel.Text = concMakeModel(intI)
            txtMakeModel.WordWrap = True
            txtMakeModel.Height = 50
            txtMakeModel.Width = 160
            txtMakeModel.TextAlign = HorizontalAlignment.Center
            txtMakeModel.Left = HP
            txtMakeModel.Top = VP
            HP = HP + txtMakeModel.Width + 1
            txtMakeModel.Font = New Font("arial", 8)
            txtMakeModel.Tag = intI
            txtMakeModel.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtMakeModel.Enabled = False
            MakeModelTextBoxes.Add(intI, txtMakeModel)

            dpMobDate.Value = concMDate(intI).Date
            dpMobDate.Name = "MobDate"
            dpMobDate.Format = DateTimePickerFormat.Custom
            dpMobDate.CustomFormat = "dd-MMM-yyyy"
            dpMobDate.Height = 30
            dpMobDate.Width = 100
            dpMobDate.Left = HP
            dpMobDate.Top = VP
            HP = HP + dpMobDate.Width + 1
            dpMobDate.Font = New Font("arial", 8)
            dpMobDate.Enabled = False
            dpMobDate.Tag = intI
            dpMobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
            MobdatePickers.Add(intI, dpMobDate)

            dpDemobDate.Value = concDMdate(intI)
            dpDemobDate.Name = "DemobDate"
            dpDemobDate.Format = DateTimePickerFormat.Custom
            dpDemobDate.CustomFormat = "dd-MMM-yyyy"
            dpDemobDate.Height = 30
            dpDemobDate.Width = 100
            dpDemobDate.Left = HP
            dpDemobDate.Top = VP
            HP = HP + dpDemobDate.Width + 1
            dpDemobDate.Font = New Font("arial", 8)
            dpDemobDate.Enabled = False
            dpDemobDate.Tag = intI
            dpDemobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
            DemobDatePickers.Add(intI, dpDemobDate)

            txtQty.Text = concQuantity(intI)
            txtQty.Height = 30
            txtQty.Width = 30
            txtQty.Left = HP
            txtQty.Top = VP
            HP = HP + txtQty.Width + 1
            txtQty.Font = New Font("arial", 8)
            txtQty.Enabled = False
            txtQty.Tag = intI
            txtQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler txtQty.Validated, AddressOf HandleQty
            QtyTextBoxes.Add(intI, txtQty)


            txtHrsPerMonth.Text = concHrsPerMonth(intI)
            txtHrsPerMonth.Height = 30
            txtHrsPerMonth.Width = 60
            txtHrsPerMonth.Left = HP
            txtHrsPerMonth.Top = VP
            HP = HP + txtHrsPerMonth.Width + 1
            txtHrsPerMonth.Font = New Font("arial", 8)
            txtHrsPerMonth.Enabled = False
            txtHrsPerMonth.Tag = intI
            txtHrsPerMonth.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            HrsPermonthTextBoxes.Add(intI, txtHrsPerMonth)

            cmbDepPerc.Items.Add(2.75)
            cmbDepPerc.Items.Add(1.25)
            cmbDepPerc.Items.Add(0.5)
            cmbDepPerc.Text = concDep(intI)
            cmbDepPerc.Height = 30
            cmbDepPerc.Width = 60
            cmbDepPerc.Left = HP
            cmbDepPerc.Top = VP
            HP = HP + cmbDepPerc.Width + 1
            cmbDepPerc.Font = New Font("arial", 8)
            cmbDepPerc.Enabled = False
            cmbDepPerc.Tag = intI
            cmbDepPerc.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            DepPercComboboxes.Add(intI, cmbDepPerc)

            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(1.5)
            cmbShifts.Items.Add(2)
            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(0)
            cmbShifts.Text = concShift(intI)
            cmbShifts.Height = 30
            cmbShifts.Width = 45
            cmbShifts.Left = HP
            cmbShifts.Top = VP
            HP = HP + cmbShifts.Width + 1
            cmbShifts.Font = New Font("arial", 8)
            cmbShifts.Enabled = False
            cmbShifts.Tag = intI
            cmbShifts.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            ShiftsComboboxes.Add(intI, cmbShifts)

            txtConcreteQty.Text = mConcreteQty
            txtConcreteQty.Height = 30
            txtConcreteQty.Width = 60
            txtConcreteQty.Left = HP
            txtConcreteQty.Top = VP
            HP = HP + txtConcreteQty.Width + 1
            txtConcreteQty.Font = New Font("arial", 8)
            txtConcreteQty.Enabled = False
            txtConcreteQty.Tag = intI
            txtConcreteQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            concreteqtyTextboxes.Add(intI, txtConcreteQty)
            If CategoryTextBoxes(intI).Text <> "Concreting" Then concreteqtyTextboxes(intI).Visible = False

            btnAddExtra.Text = "Add"
            btnAddExtra.Height = 20
            btnAddExtra.Width = 40
            btnAddExtra.Left = HP
            btnAddExtra.Top = VP
            btnAddExtra.Font = New Font("arial", 8)
            btnAddExtra.Enabled = False
            btnAddExtra.Tag = intI
            btnAddExtra.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
            ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
            AddHandler btnAddExtra.Click, AddressOf HandleBtnExtraClick
            AddButtons.Add(intI, btnAddExtra)
            If CategoryTextBoxes(intI).Text = "Conveyance" Or _
                 CategoryTextBoxes(intI).Text = "Major Others" Then btnAddExtra.Visible = False


            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtConcreteQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
            VP = VP + 20
            HP = 1
        Next
    End Sub
    Private Sub LoadControlsFromLightingArray(ByVal mItems As Integer)
        Dim intI As Integer
        HP = 1
        VP = 0
        Dim mcategoryname As String = "Lighting"
        Me.tbcBdgetHeads.TabPages(0).Text = "Single Phase Equipments And " & mcategoryname
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        QtyTextBoxes.Clear()
        PowerPerUnitTextBoxes.Clear()
        ConnectLoadTextBoxes.Clear()
        Checkboxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        UtilityFactorTextBoxes.Clear()
        AddButtons.Clear()
        mTabindex = 1
        For intI = 0 To mItems - 1
            Dim txtEquipname As New TextBox, txtCapacity As New TextBox
            Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
            Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
            Dim txtPowerPerUnit As New TextBox, txtConnectLoad As New TextBox, txtUtilityFactor As New TextBox
            Dim chkSelected As New CheckBox
            Dim btnAddExtra As New Button

            chkSelected.Checked = False
            chkSelected.Height = 30
            chkSelected.Width = 24
            chkSelected.Left = HP
            chkSelected.Top = VP - 3
            HP = HP + chkSelected.Width + 1
            chkSelected.Tag = intI
            chkSelected.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler chkSelected.CheckedChanged, AddressOf HandleLightingCheckboxStatus
            Checkboxes.Add(intI, chkSelected)

            txtEquipname.Text = LightEName(intI)
            txtEquipname.WordWrap = True
            txtEquipname.Height = 60
            txtEquipname.Width = 150
            txtEquipname.TextAlign = HorizontalAlignment.Center
            txtEquipname.Left = HP
            txtEquipname.Top = VP
            HP = HP + txtEquipname.Width + 1
            txtEquipname.Font = New Font("arial", 8)
            txtEquipname.Tag = intI
            txtEquipname.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtEquipname.Enabled = False
            EquipNameTextBoxes.Add(intI, txtEquipname)

            txtCapacity.Text = LightCapacity(intI)
            txtCapacity.WordWrap = True
            txtCapacity.Height = 30
            txtCapacity.Width = 70
            txtCapacity.TextAlign = HorizontalAlignment.Center
            txtCapacity.Left = HP
            txtCapacity.Top = VP
            HP = HP + txtCapacity.Width + 1
            txtCapacity.Font = New Font("arial", 8)
            txtCapacity.Tag = intI
            txtCapacity.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCapacity.Enabled = False
            CapacityTextBoxes.Add(intI, txtCapacity)

            txtMakeModel.Text = LightMakeModel(intI)
            txtMakeModel.WordWrap = True
            txtMakeModel.Height = 50
            txtMakeModel.Width = 160
            txtMakeModel.TextAlign = HorizontalAlignment.Center
            txtMakeModel.Left = HP
            txtMakeModel.Top = VP
            HP = HP + txtMakeModel.Width + 1
            txtMakeModel.Font = New Font("arial", 8)
            txtMakeModel.Tag = intI
            txtMakeModel.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtMakeModel.Enabled = False
            MakeModelTextBoxes.Add(intI, txtMakeModel)

            dpMobDate.Value = LightMDate(intI).Date
            dpMobDate.Name = "MobDate"
            dpMobDate.Format = DateTimePickerFormat.Custom
            dpMobDate.CustomFormat = "dd-MMM-yyyy"
            dpMobDate.Height = 30
            dpMobDate.Width = 100
            dpMobDate.Left = HP
            dpMobDate.Top = VP
            HP = HP + dpMobDate.Width + 1
            dpMobDate.Font = New Font("arial", 8)
            dpMobDate.Enabled = False
            dpMobDate.Tag = intI
            dpMobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
            MobdatePickers.Add(intI, dpMobDate)


            dpDemobDate.Value = LightDMDate(intI)
            dpDemobDate.Name = "DemobDate"
            dpDemobDate.Format = DateTimePickerFormat.Custom
            dpDemobDate.CustomFormat = "dd-MMM-yyyy"
            dpDemobDate.Height = 30
            dpDemobDate.Width = 100
            dpDemobDate.Left = HP
            dpDemobDate.Top = VP
            HP = HP + dpDemobDate.Width + 1
            dpDemobDate.Font = New Font("arial", 8)
            dpDemobDate.Enabled = False
            dpDemobDate.Tag = intI
            dpDemobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
            DemobDatePickers.Add(intI, dpDemobDate)

            txtQty.Text = LightQuantity(intI)
            txtQty.Height = 30
            txtQty.Width = 30
            txtQty.Left = HP
            txtQty.Top = VP
            HP = HP + txtQty.Width + 1
            txtQty.Font = New Font("arial", 8)
            txtQty.Enabled = False
            txtQty.Tag = intI
            txtQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler txtQty.Validated, AddressOf HandleQty
            QtyTextBoxes.Add(intI, txtQty)

            txtPowerPerUnit.Text = LightPowerPerUnit(intI)
            txtPowerPerUnit.Height = 30
            txtPowerPerUnit.Width = 60
            txtPowerPerUnit.Left = HP
            txtPowerPerUnit.Top = VP
            HP = HP + txtPowerPerUnit.Width + 1
            txtPowerPerUnit.Font = New Font("arial", 8)
            txtPowerPerUnit.Enabled = False
            txtPowerPerUnit.Tag = intI
            txtPowerPerUnit.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            PowerPerUnitTextBoxes.Add(intI, txtPowerPerUnit)

            txtConnectLoad.Text = LightConnectLoad(intI)
            txtConnectLoad.Height = 30
            txtConnectLoad.Width = 60
            txtConnectLoad.Left = HP
            txtConnectLoad.Top = VP
            HP = HP + txtConnectLoad.Width + 1
            txtConnectLoad.Font = New Font("arial", 8)
            txtConnectLoad.Enabled = False
            txtConnectLoad.Tag = intI
            txtConnectLoad.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            ConnectLoadTextBoxes.Add(intI, txtConnectLoad)

            txtUtilityFactor.Text = LightUtilityFactor(intI)
            txtUtilityFactor.Height = 30
            txtUtilityFactor.Width = 45
            txtUtilityFactor.Left = HP
            txtUtilityFactor.Top = VP
            HP = HP + txtUtilityFactor.Width + 1
            txtUtilityFactor.Font = New Font("arial", 8)
            txtUtilityFactor.Enabled = False
            txtUtilityFactor.Tag = intI
            txtUtilityFactor.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            UtilityFactorTextBoxes.Add(intI, txtUtilityFactor)

            btnAddExtra.Text = "Add"
            btnAddExtra.Height = 20
            btnAddExtra.Width = 40
            btnAddExtra.Left = HP
            btnAddExtra.Top = VP
            btnAddExtra.Font = New Font("arial", 8)
            btnAddExtra.Enabled = False
            btnAddExtra.Tag = intI
            btnAddExtra.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
            ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
            AddHandler btnAddExtra.Click, AddressOf HandleLightingBtnExtraClick
            AddButtons.Add(intI, btnAddExtra)

            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtPowerPerUnit)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtConnectLoad)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtUtilityFactor)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
            VP = VP + 20
            HP = 1
        Next
    End Sub
    Private Sub LoadControlsFromConvArray(ByVal mitems As Integer)
        Dim intI As Integer
        HP = 1
        VP = 0
        Dim mcategoryname As String = "Conveyance"
        Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
        CategoryTextBoxes.Clear()
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        QtyTextBoxes.Clear()
        HrsPermonthTextBoxes.Clear()
        concreteqtyTextboxes.Clear()
        Checkboxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        DepPercComboboxes.Clear()
        ShiftsComboboxes.Clear()
        AddButtons.Clear()
        mTabindex = 1
        For intI = 0 To mitems - 1
            Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
            Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
            Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
            Dim txtHrsPerMonth As New TextBox
            Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox
            Dim chkSelected As New CheckBox
            Dim txtConcreteQty As New TextBox   ', cmbDrive As New ComboBox
            Dim btnAddExtra As New Button

            chkSelected.Checked = False
            chkSelected.Height = 30
            chkSelected.Width = 24
            chkSelected.Left = HP
            chkSelected.Top = VP - 3
            HP = HP + chkSelected.Width + 1
            chkSelected.Tag = intI
            chkSelected.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler chkSelected.CheckedChanged, AddressOf HandleCheckboxStatus
            Checkboxes.Add(intI, chkSelected)

            txtCategory.Text = convCategory(intI)
            txtCategory.WordWrap = True
            txtCategory.Height = 50
            txtCategory.Width = 65
            txtCategory.TextAlign = HorizontalAlignment.Center
            txtCategory.Left = HP
            txtCategory.Top = VP
            HP = HP + txtCategory.Width + 1
            txtCategory.Font = New Font("arial", 8)
            txtCategory.Tag = intI
            txtCategory.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCategory.Enabled = False
            CategoryTextBoxes.Add(intI, txtCategory)

            txtEquipname.Text = convEName(intI)
            txtEquipname.WordWrap = True
            txtEquipname.Height = 60
            txtEquipname.Width = 150
            txtEquipname.TextAlign = HorizontalAlignment.Center
            txtEquipname.Left = HP
            txtEquipname.Top = VP
            HP = HP + txtEquipname.Width + 1
            txtEquipname.Font = New Font("arial", 8)
            txtEquipname.Tag = intI
            txtEquipname.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtEquipname.Enabled = False
            EquipNameTextBoxes.Add(intI, txtEquipname)

            txtCapacity.Text = convCapacity(intI)
            txtCapacity.WordWrap = True
            txtCapacity.Height = 30
            txtCapacity.Width = 70
            txtCapacity.TextAlign = HorizontalAlignment.Center
            txtCapacity.Left = HP
            txtCapacity.Top = VP
            HP = HP + txtCapacity.Width + 1
            txtCapacity.Font = New Font("arial", 8)
            txtCapacity.Tag = intI
            txtCapacity.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCapacity.Enabled = False
            CapacityTextBoxes.Add(intI, txtCapacity)

            txtMakeModel.Text = convMakeModel(intI)
            txtMakeModel.WordWrap = True
            txtMakeModel.Height = 50
            txtMakeModel.Width = 160
            txtMakeModel.TextAlign = HorizontalAlignment.Center
            txtMakeModel.Left = HP
            txtMakeModel.Top = VP
            HP = HP + txtMakeModel.Width + 1
            txtMakeModel.Font = New Font("arial", 8)
            txtMakeModel.Tag = intI
            txtMakeModel.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtMakeModel.Enabled = False
            MakeModelTextBoxes.Add(intI, txtMakeModel)

            dpMobDate.Value = convMDate(intI).Date
            dpMobDate.Name = "MobDate"
            dpMobDate.Format = DateTimePickerFormat.Custom
            dpMobDate.CustomFormat = "dd-MMM-yyyy"
            dpMobDate.Height = 30
            dpMobDate.Width = 100
            dpMobDate.Left = HP
            dpMobDate.Top = VP
            HP = HP + dpMobDate.Width + 1
            dpMobDate.Font = New Font("arial", 8)
            dpMobDate.Enabled = False
            dpMobDate.Tag = intI
            dpMobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
            MobdatePickers.Add(intI, dpMobDate)

            dpDemobDate.Value = convDMdate(intI)
            dpDemobDate.Name = "DemobDate"
            dpDemobDate.Format = DateTimePickerFormat.Custom
            dpDemobDate.CustomFormat = "dd-MMM-yyyy"
            dpDemobDate.Height = 30
            dpDemobDate.Width = 100
            dpDemobDate.Left = HP
            dpDemobDate.Top = VP
            HP = HP + dpDemobDate.Width + 1
            dpDemobDate.Font = New Font("arial", 8)
            dpDemobDate.Enabled = False
            dpDemobDate.Tag = intI
            dpDemobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
            DemobDatePickers.Add(intI, dpDemobDate)

            txtQty.Text = convQuantity(intI)
            txtQty.Height = 30
            txtQty.Width = 30
            txtQty.Left = HP
            txtQty.Top = VP
            HP = HP + txtQty.Width + 1
            txtQty.Font = New Font("arial", 8)
            txtQty.Enabled = False
            txtQty.Tag = intI
            txtQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler txtQty.Validated, AddressOf HandleQty
            QtyTextBoxes.Add(intI, txtQty)

            txtHrsPerMonth.Text = convHrsPerMonth(intI)
            txtHrsPerMonth.Height = 30
            txtHrsPerMonth.Width = 60
            txtHrsPerMonth.Left = HP
            txtHrsPerMonth.Top = VP
            HP = HP + txtHrsPerMonth.Width + 1
            txtHrsPerMonth.Font = New Font("arial", 8)
            txtHrsPerMonth.Enabled = False
            txtHrsPerMonth.Tag = intI
            txtHrsPerMonth.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            HrsPermonthTextBoxes.Add(intI, txtHrsPerMonth)

            cmbDepPerc.Items.Add(2.75)
            cmbDepPerc.Items.Add(1.25)
            cmbDepPerc.Items.Add(0.5)
            cmbDepPerc.Text = convDep(intI)
            cmbDepPerc.Height = 30
            cmbDepPerc.Width = 60
            cmbDepPerc.Left = HP
            cmbDepPerc.Top = VP
            HP = HP + cmbDepPerc.Width + 1
            cmbDepPerc.Font = New Font("arial", 8)
            cmbDepPerc.Enabled = False
            cmbDepPerc.Tag = intI
            cmbDepPerc.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            DepPercComboboxes.Add(intI, cmbDepPerc)
            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(1.5)
            cmbShifts.Items.Add(2)
            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(0)
            cmbShifts.Text = convShift(intI)
            cmbShifts.Height = 30
            cmbShifts.Width = 45
            cmbShifts.Left = HP
            cmbShifts.Top = VP
            HP = HP + cmbShifts.Width + 1
            cmbShifts.Font = New Font("arial", 8)
            cmbShifts.Enabled = False
            cmbShifts.Tag = intI
            cmbShifts.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            ShiftsComboboxes.Add(intI, cmbShifts)

            txtConcreteQty.Text = mConcreteQty
            txtConcreteQty.Height = 30
            txtConcreteQty.Width = 60
            txtConcreteQty.Left = HP
            txtConcreteQty.Top = VP
            'HP = HP + txtConcreteQty.Width + 1
            txtConcreteQty.Font = New Font("arial", 8)
            txtConcreteQty.Enabled = False
            txtConcreteQty.Tag = intI
            txtConcreteQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            concreteqtyTextboxes.Add(intI, txtConcreteQty)
            If CategoryTextBoxes(intI).Text <> "Concreting" Then concreteqtyTextboxes(intI).Visible = False

            btnAddExtra.Text = "Add"
            btnAddExtra.Height = 20
            btnAddExtra.Width = 40
            btnAddExtra.Left = HP
            btnAddExtra.Top = VP
            btnAddExtra.Font = New Font("arial", 8)
            btnAddExtra.Enabled = False
            btnAddExtra.Tag = intI
            btnAddExtra.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
            ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
            AddHandler btnAddExtra.Click, AddressOf HandleBtnExtraClick
            AddButtons.Add(intI, btnAddExtra)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtConcreteQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
            VP = VP + 20
            HP = 1
        Next
    End Sub
    Private Sub LoadControlsFromCraneArray(ByVal mitems As Integer)
        Dim intI As Integer
        HP = 1
        VP = 0
        Dim mcategoryname As String = "Cranes"
        Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
        CategoryTextBoxes.Clear()
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        QtyTextBoxes.Clear()
        HrsPermonthTextBoxes.Clear()
        concreteqtyTextboxes.Clear()
        Checkboxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        DepPercComboboxes.Clear()
        ShiftsComboboxes.Clear()
        AddButtons.Clear()
        mTabindex = 1
        For intI = 0 To mitems - 1
            Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
            Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
            Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
            Dim txtHrsPerMonth As New TextBox
            Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox
            Dim chkSelected As New CheckBox
            Dim txtConcreteQty As New TextBox   ', cmbDrive As New ComboBox
            Dim btnAddExtra As New Button

            chkSelected.Checked = False
            chkSelected.Height = 30
            chkSelected.Width = 24
            chkSelected.Left = HP
            chkSelected.Top = VP - 3
            HP = HP + chkSelected.Width + 1
            'txtMaintPerc.Font = New Font("verdana", 10)
            chkSelected.Tag = intI
            chkSelected.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler chkSelected.CheckedChanged, AddressOf HandleCheckboxStatus
            Checkboxes.Add(intI, chkSelected)

            txtCategory.Text = cranCategory(intI)
            txtCategory.WordWrap = True
            txtCategory.Height = 50
            txtCategory.Width = 65
            txtCategory.TextAlign = HorizontalAlignment.Center
            txtCategory.Left = HP
            txtCategory.Top = VP
            HP = HP + txtCategory.Width + 1
            txtCategory.Font = New Font("arial", 8)
            txtCategory.Tag = intI
            txtCategory.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCategory.Enabled = False
            CategoryTextBoxes.Add(intI, txtCategory)

            txtEquipname.Text = cranEName(intI)
            txtEquipname.WordWrap = True
            txtEquipname.Height = 60
            txtEquipname.Width = 150
            txtEquipname.TextAlign = HorizontalAlignment.Center
            txtEquipname.Left = HP
            txtEquipname.Top = VP
            HP = HP + txtEquipname.Width + 1
            txtEquipname.Font = New Font("arial", 8)
            txtEquipname.Tag = intI
            txtEquipname.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtEquipname.Enabled = False
            EquipNameTextBoxes.Add(intI, txtEquipname)

            txtCapacity.Text = cranCapacity(intI)
            txtCapacity.WordWrap = True
            txtCapacity.Height = 30
            txtCapacity.Width = 70
            txtCapacity.TextAlign = HorizontalAlignment.Center
            txtCapacity.Left = HP
            txtCapacity.Top = VP
            HP = HP + txtCapacity.Width + 1
            txtCapacity.Font = New Font("arial", 8)
            txtCapacity.Tag = intI
            txtCapacity.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCapacity.Enabled = False
            CapacityTextBoxes.Add(intI, txtCapacity)

            txtMakeModel.Text = cranMakeModel(intI)
            txtMakeModel.WordWrap = True
            txtMakeModel.Height = 50
            txtMakeModel.Width = 160
            txtMakeModel.TextAlign = HorizontalAlignment.Center
            txtMakeModel.Left = HP
            txtMakeModel.Top = VP
            HP = HP + txtMakeModel.Width + 1
            txtMakeModel.Font = New Font("arial", 8)
            txtMakeModel.Tag = intI
            txtMakeModel.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtMakeModel.Enabled = False
            MakeModelTextBoxes.Add(intI, txtMakeModel)

            dpMobDate.Value = cranMDate(intI).Date
            dpMobDate.Name = "MobDate"
            dpMobDate.Format = DateTimePickerFormat.Custom
            dpMobDate.CustomFormat = "dd-MMM-yyyy"
            dpMobDate.Height = 30
            dpMobDate.Width = 100
            dpMobDate.Left = HP
            dpMobDate.Top = VP
            HP = HP + dpMobDate.Width + 1
            dpMobDate.Font = New Font("arial", 8)
            dpMobDate.Enabled = False
            dpMobDate.Tag = intI
            dpMobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
            MobdatePickers.Add(intI, dpMobDate)

            dpDemobDate.Value = cranDMdate(intI)
            dpDemobDate.Name = "DemobDate"
            dpDemobDate.Format = DateTimePickerFormat.Custom
            dpDemobDate.CustomFormat = "dd-MMM-yyyy"
            dpDemobDate.Height = 30
            dpDemobDate.Width = 100
            dpDemobDate.Left = HP
            dpDemobDate.Top = VP
            HP = HP + dpDemobDate.Width + 1
            dpDemobDate.Font = New Font("arial", 8)
            dpDemobDate.Enabled = False
            dpDemobDate.Tag = intI
            dpDemobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
            DemobDatePickers.Add(intI, dpDemobDate)

            txtQty.Text = cranQuantity(intI)
            txtQty.Height = 30
            txtQty.Width = 30
            txtQty.Left = HP
            txtQty.Top = VP
            HP = HP + txtQty.Width + 1
            txtQty.Font = New Font("arial", 8)
            txtQty.Enabled = False
            txtQty.Tag = intI
            txtQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler txtQty.Validated, AddressOf HandleQty
            QtyTextBoxes.Add(intI, txtQty)


            txtHrsPerMonth.Text = cranHrsPerMonth(intI)
            txtHrsPerMonth.Height = 30
            txtHrsPerMonth.Width = 60
            txtHrsPerMonth.Left = HP
            txtHrsPerMonth.Top = VP
            HP = HP + txtHrsPerMonth.Width + 1
            txtHrsPerMonth.Font = New Font("arial", 8)
            txtHrsPerMonth.Enabled = False
            txtHrsPerMonth.Tag = intI
            txtHrsPerMonth.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            HrsPermonthTextBoxes.Add(intI, txtHrsPerMonth)

            cmbDepPerc.Items.Add(2.75)
            cmbDepPerc.Items.Add(1.25)
            cmbDepPerc.Items.Add(0.5)
            cmbDepPerc.Text = cranDep(intI)
            cmbDepPerc.Height = 30
            cmbDepPerc.Width = 60
            cmbDepPerc.Left = HP
            cmbDepPerc.Top = VP
            HP = HP + cmbDepPerc.Width + 1
            cmbDepPerc.Font = New Font("arial", 8)
            cmbDepPerc.Enabled = False
            cmbDepPerc.Tag = intI
            cmbDepPerc.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            DepPercComboboxes.Add(intI, cmbDepPerc)

            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(1.5)
            cmbShifts.Items.Add(2)
            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(0)
            cmbShifts.Text = cranShift(intI)
            cmbShifts.Height = 30
            cmbShifts.Width = 45
            cmbShifts.Left = HP
            cmbShifts.Top = VP
            HP = HP + cmbShifts.Width + 1
            cmbShifts.Font = New Font("arial", 8)
            cmbShifts.Enabled = False
            cmbShifts.Tag = intI
            cmbShifts.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            ShiftsComboboxes.Add(intI, cmbShifts)

            txtConcreteQty.Text = mConcreteQty
            txtConcreteQty.Height = 30
            txtConcreteQty.Width = 60
            txtConcreteQty.Left = HP
            txtConcreteQty.Top = VP
            txtConcreteQty.Font = New Font("arial", 8)
            txtConcreteQty.Enabled = False
            txtConcreteQty.Tag = intI
            txtConcreteQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            concreteqtyTextboxes.Add(intI, txtConcreteQty)
            If CategoryTextBoxes(intI).Text <> "Concreting" Then concreteqtyTextboxes(intI).Visible = False

            btnAddExtra.Text = "Add"
            btnAddExtra.Height = 20
            btnAddExtra.Width = 40
            btnAddExtra.Left = HP
            btnAddExtra.Top = VP
            btnAddExtra.Font = New Font("arial", 8)
            btnAddExtra.Enabled = False
            btnAddExtra.Tag = intI
            btnAddExtra.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
            ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
            AddHandler btnAddExtra.Click, AddressOf HandleBtnExtraClick
            AddButtons.Add(intI, btnAddExtra)
            If CategoryTextBoxes(intI).Text = "Conveyance" Or _
                 CategoryTextBoxes(intI).Text = "Major Others" Then btnAddExtra.Visible = False

            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtConcreteQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
            VP = VP + 20
            HP = 1
        Next
    End Sub
    Private Sub LoadControlsFromDGSetsArray(ByVal mitems As Integer)
        Dim intI As Integer
        HP = 1
        VP = 0
        Dim mcategoryname As String = "Dg Sets"
        Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
        CategoryTextBoxes.Clear()
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        QtyTextBoxes.Clear()
        HrsPermonthTextBoxes.Clear()
        concreteqtyTextboxes.Clear()
        Checkboxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        DepPercComboboxes.Clear()
        ShiftsComboboxes.Clear()
        AddButtons.Clear()
        mTabindex = 1
        For intI = 0 To mitems - 1
            Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
            Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
            Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
            Dim txtHrsPerMonth As New TextBox
            Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox
            Dim chkSelected As New CheckBox
            Dim txtConcreteQty As New TextBox   ', cmbDrive As New ComboBox
            Dim btnAddExtra As New Button

            chkSelected.Checked = False
            chkSelected.Height = 30
            chkSelected.Width = 24
            chkSelected.Left = HP
            chkSelected.Top = VP - 3
            HP = HP + chkSelected.Width + 1
            'txtMaintPerc.Font = New Font("verdana", 10)
            chkSelected.Tag = intI
            chkSelected.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler chkSelected.CheckedChanged, AddressOf HandleCheckboxStatus
            Checkboxes.Add(intI, chkSelected)

            txtCategory.Text = dgsetCategory(intI)
            'txtCategory.Multiline = True
            txtCategory.WordWrap = True
            txtCategory.Height = 50
            txtCategory.Width = 65
            txtCategory.TextAlign = HorizontalAlignment.Center
            txtCategory.Left = HP
            txtCategory.Top = VP
            HP = HP + txtCategory.Width + 1
            txtCategory.Font = New Font("arial", 8)
            txtCategory.Tag = intI
            txtCategory.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCategory.Enabled = False
            CategoryTextBoxes.Add(intI, txtCategory)

            txtEquipname.Text = dgsetEName(intI)
            txtEquipname.WordWrap = True
            txtEquipname.Height = 60
            txtEquipname.Width = 150
            txtEquipname.TextAlign = HorizontalAlignment.Center
            txtEquipname.Left = HP
            txtEquipname.Top = VP
            HP = HP + txtEquipname.Width + 1
            txtEquipname.Font = New Font("arial", 8)
            txtEquipname.Tag = intI
            txtEquipname.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtEquipname.Enabled = False
            EquipNameTextBoxes.Add(intI, txtEquipname)

            txtCapacity.Text = dgsetCapacity(intI)
            txtCapacity.WordWrap = True
            txtCapacity.Height = 30
            txtCapacity.Width = 70
            txtCapacity.TextAlign = HorizontalAlignment.Center
            txtCapacity.Left = HP
            txtCapacity.Top = VP
            HP = HP + txtCapacity.Width + 1
            txtCapacity.Font = New Font("arial", 8)
            txtCapacity.Tag = intI
            txtCapacity.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCapacity.Enabled = False
            CapacityTextBoxes.Add(intI, txtCapacity)

            txtMakeModel.Text = dgsetMakeModel(intI)
            txtMakeModel.WordWrap = True
            txtMakeModel.Height = 50
            txtMakeModel.Width = 160
            txtMakeModel.TextAlign = HorizontalAlignment.Center
            txtMakeModel.Left = HP
            txtMakeModel.Top = VP
            HP = HP + txtMakeModel.Width + 1
            txtMakeModel.Font = New Font("arial", 8)
            txtMakeModel.Tag = intI
            txtMakeModel.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtMakeModel.Enabled = False
            MakeModelTextBoxes.Add(intI, txtMakeModel)

            dpMobDate.Value = dgsetMDate(intI).Date
            dpMobDate.Name = "MobDate"
            dpMobDate.Format = DateTimePickerFormat.Custom
            dpMobDate.CustomFormat = "dd-MMM-yyyy"
            dpMobDate.Height = 30
            dpMobDate.Width = 100
            dpMobDate.Left = HP
            dpMobDate.Top = VP
            HP = HP + dpMobDate.Width + 1
            dpMobDate.Font = New Font("arial", 8)
            dpMobDate.Enabled = False
            dpMobDate.Tag = intI
            dpMobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
            MobdatePickers.Add(intI, dpMobDate)

            dpDemobDate.Value = dgsetDMdate(intI)
            dpDemobDate.Name = "DemobDate"
            dpDemobDate.Format = DateTimePickerFormat.Custom
            dpDemobDate.CustomFormat = "dd-MMM-yyyy"
            dpDemobDate.Height = 30
            dpDemobDate.Width = 100
            dpDemobDate.Left = HP
            dpDemobDate.Top = VP
            HP = HP + dpDemobDate.Width + 1
            dpDemobDate.Font = New Font("arial", 8)
            dpDemobDate.Enabled = False
            dpDemobDate.Tag = intI
            dpDemobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
            DemobDatePickers.Add(intI, dpDemobDate)

            txtQty.Text = dgsetQuantity(intI)
            txtQty.Height = 30
            txtQty.Width = 30
            txtQty.Left = HP
            txtQty.Top = VP
            HP = HP + txtQty.Width + 1
            txtQty.Font = New Font("arial", 8)
            txtQty.Enabled = False
            txtQty.Tag = intI
            txtQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler txtQty.Validated, AddressOf HandleQty
            QtyTextBoxes.Add(intI, txtQty)


            txtHrsPerMonth.Text = dgsetHrsPerMonth(intI)
            txtHrsPerMonth.Height = 30
            txtHrsPerMonth.Width = 60
            txtHrsPerMonth.Left = HP
            txtHrsPerMonth.Top = VP
            HP = HP + txtHrsPerMonth.Width + 1
            txtHrsPerMonth.Font = New Font("arial", 8)
            txtHrsPerMonth.Enabled = False
            txtHrsPerMonth.Tag = intI
            txtHrsPerMonth.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            HrsPermonthTextBoxes.Add(intI, txtHrsPerMonth)

            cmbDepPerc.Items.Add(2.75)
            cmbDepPerc.Items.Add(1.25)
            cmbDepPerc.Items.Add(0.5)
            cmbDepPerc.Text = dgsetDep(intI)
            cmbDepPerc.Height = 30
            cmbDepPerc.Width = 60
            cmbDepPerc.Left = HP
            cmbDepPerc.Top = VP
            HP = HP + cmbDepPerc.Width + 1
            cmbDepPerc.Font = New Font("arial", 8)
            cmbDepPerc.Enabled = False
            cmbDepPerc.Tag = intI
            cmbDepPerc.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            DepPercComboboxes.Add(intI, cmbDepPerc)

            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(1.5)
            cmbShifts.Items.Add(2)
            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(0)
            cmbShifts.Text = dgsetShift(intI)
            cmbShifts.Height = 30
            cmbShifts.Width = 45
            cmbShifts.Left = HP
            cmbShifts.Top = VP
            HP = HP + cmbShifts.Width + 1
            cmbShifts.Font = New Font("arial", 8)
            cmbShifts.Enabled = False
            cmbShifts.Tag = intI
            cmbShifts.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            ShiftsComboboxes.Add(intI, cmbShifts)

            txtConcreteQty.Text = mConcreteQty
            txtConcreteQty.Height = 30
            txtConcreteQty.Width = 60
            txtConcreteQty.Left = HP
            txtConcreteQty.Top = VP
            txtConcreteQty.Font = New Font("arial", 8)
            txtConcreteQty.Enabled = False
            txtConcreteQty.Tag = intI
            txtConcreteQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            concreteqtyTextboxes.Add(intI, txtConcreteQty)
            If CategoryTextBoxes(intI).Text <> "Concreting" Then concreteqtyTextboxes(intI).Visible = False

            btnAddExtra.Text = "Add"
            btnAddExtra.Height = 20
            btnAddExtra.Width = 40
            btnAddExtra.Left = HP
            btnAddExtra.Top = VP
            btnAddExtra.Font = New Font("arial", 8)
            btnAddExtra.Enabled = False
            btnAddExtra.Tag = intI
            btnAddExtra.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
            ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
            AddHandler btnAddExtra.Click, AddressOf HandleBtnExtraClick
            AddButtons.Add(intI, btnAddExtra)
            If CategoryTextBoxes(intI).Text = "Conveyance" Or _
                 CategoryTextBoxes(intI).Text = "Major Others" Then btnAddExtra.Visible = False

            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtConcreteQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
            VP = VP + 20
            HP = 1
        Next
    End Sub
    Private Sub LoadControlsFromMHArray(ByVal mitems As Integer)
        Dim intI As Integer
        HP = 1
        VP = 0
        Dim mcategoryname As String = "Material Handling"
        Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
        CategoryTextBoxes.Clear()
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        QtyTextBoxes.Clear()
        HrsPermonthTextBoxes.Clear()
        concreteqtyTextboxes.Clear()
        Checkboxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        DepPercComboboxes.Clear()
        ShiftsComboboxes.Clear()
        AddButtons.Clear()
        mTabindex = 1
        For intI = 0 To mitems - 1
            Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
            Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
            Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
            Dim txtHrsPerMonth As New TextBox
            Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox
            Dim chkSelected As New CheckBox
            Dim txtConcreteQty As New TextBox   ', cmbDrive As New ComboBox
            Dim btnAddExtra As New Button

            chkSelected.Checked = False
            chkSelected.Height = 30
            chkSelected.Width = 24
            chkSelected.Left = HP
            chkSelected.Top = VP - 3
            HP = HP + chkSelected.Width + 1
            'txtMaintPerc.Font = New Font("verdana", 10)
            chkSelected.Tag = intI
            chkSelected.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler chkSelected.CheckedChanged, AddressOf HandleCheckboxStatus
            Checkboxes.Add(intI, chkSelected)

            txtCategory.Text = mathCategory(intI)
            txtCategory.WordWrap = True
            txtCategory.Height = 50
            txtCategory.Width = 65
            txtCategory.TextAlign = HorizontalAlignment.Center
            txtCategory.Left = HP
            txtCategory.Top = VP
            HP = HP + txtCategory.Width + 1
            txtCategory.Font = New Font("arial", 8)
            txtCategory.Tag = intI
            txtCategory.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCategory.Enabled = False
            CategoryTextBoxes.Add(intI, txtCategory)

            txtEquipname.Text = mathEName(intI)
            txtEquipname.WordWrap = True
            txtEquipname.Height = 60
            txtEquipname.Width = 150
            txtEquipname.TextAlign = HorizontalAlignment.Center
            txtEquipname.Left = HP
            txtEquipname.Top = VP
            HP = HP + txtEquipname.Width + 1
            txtEquipname.Font = New Font("arial", 8)
            txtEquipname.Tag = intI
            txtEquipname.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtEquipname.Enabled = False
            EquipNameTextBoxes.Add(intI, txtEquipname)

            txtCapacity.Text = mathCapacity(intI)
            txtCapacity.WordWrap = True
            txtCapacity.Height = 30
            txtCapacity.Width = 70
            txtCapacity.TextAlign = HorizontalAlignment.Center
            txtCapacity.Left = HP
            txtCapacity.Top = VP
            HP = HP + txtCapacity.Width + 1
            txtCapacity.Font = New Font("arial", 8)
            txtCapacity.Tag = intI
            txtCapacity.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCapacity.Enabled = False
            CapacityTextBoxes.Add(intI, txtCapacity)

            txtMakeModel.Text = mathMakeModel(intI)
            txtMakeModel.WordWrap = True
            txtMakeModel.Height = 50
            txtMakeModel.Width = 160
            txtMakeModel.TextAlign = HorizontalAlignment.Center
            txtMakeModel.Left = HP
            txtMakeModel.Top = VP
            HP = HP + txtMakeModel.Width + 1
            txtMakeModel.Font = New Font("arial", 8)
            txtMakeModel.Tag = intI
            txtMakeModel.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtMakeModel.Enabled = False
            MakeModelTextBoxes.Add(intI, txtMakeModel)

            dpMobDate.Value = mathMDate(intI).Date
            dpMobDate.Name = "MobDate"
            dpMobDate.Format = DateTimePickerFormat.Custom
            dpMobDate.CustomFormat = "dd-MMM-yyyy"
            dpMobDate.Height = 30
            dpMobDate.Width = 100
            dpMobDate.Left = HP
            dpMobDate.Top = VP
            HP = HP + dpMobDate.Width + 1
            dpMobDate.Font = New Font("arial", 8)
            dpMobDate.Enabled = False
            dpMobDate.Tag = intI
            dpMobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
            MobdatePickers.Add(intI, dpMobDate)

            dpDemobDate.Value = mathDMdate(intI)
            dpDemobDate.Name = "DemobDate"
            dpDemobDate.Format = DateTimePickerFormat.Custom
            dpDemobDate.CustomFormat = "dd-MMM-yyyy"
            dpDemobDate.Height = 30
            dpDemobDate.Width = 100
            dpDemobDate.Left = HP
            dpDemobDate.Top = VP
            HP = HP + dpDemobDate.Width + 1
            dpDemobDate.Font = New Font("arial", 8)
            dpDemobDate.Enabled = False
            dpDemobDate.Tag = intI
            dpDemobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
            DemobDatePickers.Add(intI, dpDemobDate)

            txtQty.Text = mathQuantity(intI)
            txtQty.Height = 30
            txtQty.Width = 30
            txtQty.Left = HP
            txtQty.Top = VP
            HP = HP + txtQty.Width + 1
            txtQty.Font = New Font("arial", 8)
            txtQty.Enabled = False
            txtQty.Tag = intI
            txtQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler txtQty.Validated, AddressOf HandleQty
            QtyTextBoxes.Add(intI, txtQty)

            txtHrsPerMonth.Text = mathHrsPerMonth(intI)
            txtHrsPerMonth.Height = 30
            txtHrsPerMonth.Width = 60
            txtHrsPerMonth.Left = HP
            txtHrsPerMonth.Top = VP
            HP = HP + txtHrsPerMonth.Width + 1
            txtHrsPerMonth.Font = New Font("arial", 8)
            txtHrsPerMonth.Enabled = False
            txtHrsPerMonth.Tag = intI
            txtHrsPerMonth.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            HrsPermonthTextBoxes.Add(intI, txtHrsPerMonth)

            cmbDepPerc.Items.Add(2.75)
            cmbDepPerc.Items.Add(1.25)
            cmbDepPerc.Items.Add(0.5)
            cmbDepPerc.Text = mathDep(intI)
            cmbDepPerc.Height = 30
            cmbDepPerc.Width = 60
            cmbDepPerc.Left = HP
            cmbDepPerc.Top = VP
            HP = HP + cmbDepPerc.Width + 1
            cmbDepPerc.Font = New Font("arial", 8)
            cmbDepPerc.Enabled = False
            cmbDepPerc.Tag = intI
            cmbDepPerc.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            DepPercComboboxes.Add(intI, cmbDepPerc)

            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(1.5)
            cmbShifts.Items.Add(2)
            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(0)
            cmbShifts.Text = mathShift(intI)
            cmbShifts.Height = 30
            cmbShifts.Width = 45
            cmbShifts.Left = HP
            cmbShifts.Top = VP
            HP = HP + cmbShifts.Width + 1
            cmbShifts.Font = New Font("arial", 8)
            cmbShifts.Enabled = False
            cmbShifts.Tag = intI
            cmbShifts.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            ShiftsComboboxes.Add(intI, cmbShifts)

            txtConcreteQty.Text = mConcreteQty
            txtConcreteQty.Height = 30
            txtConcreteQty.Width = 60
            txtConcreteQty.Left = HP
            txtConcreteQty.Top = VP
            txtConcreteQty.Font = New Font("arial", 8)
            txtConcreteQty.Enabled = False
            txtConcreteQty.Tag = intI
            txtConcreteQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            concreteqtyTextboxes.Add(intI, txtConcreteQty)
            If CategoryTextBoxes(intI).Text <> "Concreting" Then concreteqtyTextboxes(intI).Visible = False

            btnAddExtra.Text = "Add"
            btnAddExtra.Height = 20
            btnAddExtra.Width = 40
            btnAddExtra.Left = HP
            btnAddExtra.Top = VP
            btnAddExtra.Font = New Font("arial", 8)
            btnAddExtra.Enabled = False
            btnAddExtra.Tag = intI
            btnAddExtra.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
            ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
            AddHandler btnAddExtra.Click, AddressOf HandleBtnExtraClick
            AddButtons.Add(intI, btnAddExtra)
            If CategoryTextBoxes(intI).Text = "Conveyance" Or _
                 CategoryTextBoxes(intI).Text = "Major Others" Then btnAddExtra.Visible = False

            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtConcreteQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
            VP = VP + 20
            HP = 1
        Next
    End Sub
    Private Sub LoadControlsFromNCArray(ByVal mitems As Integer)
        Dim intI As Integer
        HP = 1
        VP = 0
        Dim mcategoryname As String = "Non Concreting"
        Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
        'Me.Refresh()
        CategoryTextBoxes.Clear()
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        QtyTextBoxes.Clear()
        HrsPermonthTextBoxes.Clear()
        concreteqtyTextboxes.Clear()
        Checkboxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        DepPercComboboxes.Clear()
        ShiftsComboboxes.Clear()
        AddButtons.Clear()
        mTabindex = 1
        For intI = 0 To mitems - 1
            Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
            Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
            Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
            Dim txtHrsPerMonth As New TextBox
            Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox
            Dim chkSelected As New CheckBox
            Dim txtConcreteQty As New TextBox   ', cmbDrive As New ComboBox
            Dim btnAddExtra As New Button

            chkSelected.Checked = False
            chkSelected.Height = 30
            chkSelected.Width = 24
            chkSelected.Left = HP
            chkSelected.Top = VP - 3
            HP = HP + chkSelected.Width + 1
            'txtMaintPerc.Font = New Font("verdana", 10)
            chkSelected.Tag = intI
            chkSelected.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler chkSelected.CheckedChanged, AddressOf HandleCheckboxStatus
            Checkboxes.Add(intI, chkSelected)

            txtCategory.Text = noncCategory(intI)
            txtCategory.WordWrap = True
            txtCategory.Height = 50
            txtCategory.Width = 65
            txtCategory.TextAlign = HorizontalAlignment.Center
            txtCategory.Left = HP
            txtCategory.Top = VP
            HP = HP + txtCategory.Width + 1
            txtCategory.Font = New Font("arial", 8)
            txtCategory.Tag = intI
            txtCategory.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCategory.Enabled = False
            CategoryTextBoxes.Add(intI, txtCategory)

            txtEquipname.Text = noncEName(intI)
            txtEquipname.WordWrap = True
            txtEquipname.Height = 60
            txtEquipname.Width = 150
            txtEquipname.TextAlign = HorizontalAlignment.Center
            txtEquipname.Left = HP
            txtEquipname.Top = VP
            HP = HP + txtEquipname.Width + 1
            txtEquipname.Font = New Font("arial", 8)
            txtEquipname.Tag = intI
            txtEquipname.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtEquipname.Enabled = False
            EquipNameTextBoxes.Add(intI, txtEquipname)

            txtCapacity.Text = noncCapacity(intI)
            txtCapacity.WordWrap = True
            txtCapacity.Height = 30
            txtCapacity.Width = 70
            txtCapacity.TextAlign = HorizontalAlignment.Center
            txtCapacity.Left = HP
            txtCapacity.Top = VP
            HP = HP + txtCapacity.Width + 1
            txtCapacity.Font = New Font("arial", 8)
            txtCapacity.Tag = intI
            txtCapacity.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCapacity.Enabled = False
            CapacityTextBoxes.Add(intI, txtCapacity)

            txtMakeModel.Text = noncMakeModel(intI)
            txtMakeModel.WordWrap = True
            txtMakeModel.Height = 50
            txtMakeModel.Width = 160
            txtMakeModel.TextAlign = HorizontalAlignment.Center
            txtMakeModel.Left = HP
            txtMakeModel.Top = VP
            HP = HP + txtMakeModel.Width + 1
            txtMakeModel.Font = New Font("arial", 8)
            txtMakeModel.Tag = intI
            txtMakeModel.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtMakeModel.Enabled = False
            MakeModelTextBoxes.Add(intI, txtMakeModel)

            dpMobDate.Value = noncMDate(intI).Date
            dpMobDate.Name = "MobDate"
            dpMobDate.Format = DateTimePickerFormat.Custom
            dpMobDate.CustomFormat = "dd-MMM-yyyy"
            dpMobDate.Height = 30
            dpMobDate.Width = 100
            dpMobDate.Left = HP
            dpMobDate.Top = VP
            HP = HP + dpMobDate.Width + 1
            dpMobDate.Font = New Font("arial", 8)
            dpMobDate.Enabled = False
            dpMobDate.Tag = intI
            dpMobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
            MobdatePickers.Add(intI, dpMobDate)

            dpDemobDate.Value = noncDMdate(intI)
            dpDemobDate.Name = "DemobDate"
            dpDemobDate.Format = DateTimePickerFormat.Custom
            dpDemobDate.CustomFormat = "dd-MMM-yyyy"
            dpDemobDate.Height = 30
            dpDemobDate.Width = 100
            dpDemobDate.Left = HP
            dpDemobDate.Top = VP
            HP = HP + dpDemobDate.Width + 1
            dpDemobDate.Font = New Font("arial", 8)
            dpDemobDate.Enabled = False
            dpDemobDate.Tag = intI
            dpDemobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
            DemobDatePickers.Add(intI, dpDemobDate)

            txtQty.Text = noncQuantity(intI)
            txtQty.Height = 30
            txtQty.Width = 30
            txtQty.Left = HP
            txtQty.Top = VP
            HP = HP + txtQty.Width + 1
            txtQty.Font = New Font("arial", 8)
            txtQty.Enabled = False
            txtQty.Tag = intI
            txtQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler txtQty.Validated, AddressOf HandleQty
            QtyTextBoxes.Add(intI, txtQty)

            txtHrsPerMonth.Text = noncHrsPerMonth(intI)
            txtHrsPerMonth.Height = 30
            txtHrsPerMonth.Width = 60
            txtHrsPerMonth.Left = HP
            txtHrsPerMonth.Top = VP
            HP = HP + txtHrsPerMonth.Width + 1
            txtHrsPerMonth.Font = New Font("arial", 8)
            txtHrsPerMonth.Enabled = False
            txtHrsPerMonth.Tag = intI
            txtHrsPerMonth.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            HrsPermonthTextBoxes.Add(intI, txtHrsPerMonth)

            cmbDepPerc.Items.Add(2.75)
            cmbDepPerc.Items.Add(1.25)
            cmbDepPerc.Items.Add(0.5)
            cmbDepPerc.Text = noncDep(intI)
            cmbDepPerc.Height = 30
            cmbDepPerc.Width = 60
            cmbDepPerc.Left = HP
            cmbDepPerc.Top = VP
            HP = HP + cmbDepPerc.Width + 1
            cmbDepPerc.Font = New Font("arial", 8)
            cmbDepPerc.Enabled = False
            cmbDepPerc.Tag = intI
            cmbDepPerc.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            DepPercComboboxes.Add(intI, cmbDepPerc)

            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(1.5)
            cmbShifts.Items.Add(2)
            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(0)
            cmbShifts.Text = noncShift(intI)
            cmbShifts.Height = 30
            cmbShifts.Width = 45
            cmbShifts.Left = HP
            cmbShifts.Top = VP
            HP = HP + cmbShifts.Width + 1
            cmbShifts.Font = New Font("arial", 8)
            cmbShifts.Enabled = False
            cmbShifts.Tag = intI
            cmbShifts.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            ShiftsComboboxes.Add(intI, cmbShifts)

            txtConcreteQty.Text = mConcreteQty
            txtConcreteQty.Height = 30
            txtConcreteQty.Width = 60
            txtConcreteQty.Left = HP
            txtConcreteQty.Top = VP
            txtConcreteQty.Font = New Font("arial", 8)
            txtConcreteQty.Enabled = False
            txtConcreteQty.Tag = intI
            txtConcreteQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            concreteqtyTextboxes.Add(intI, txtConcreteQty)
            If CategoryTextBoxes(intI).Text <> "Concreting" Then concreteqtyTextboxes(intI).Visible = False

            btnAddExtra.Text = "Add"
            btnAddExtra.Height = 20
            btnAddExtra.Width = 40
            btnAddExtra.Left = HP
            btnAddExtra.Top = VP
            btnAddExtra.Font = New Font("arial", 8)
            btnAddExtra.Enabled = False
            btnAddExtra.Tag = intI
            btnAddExtra.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
            ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
            AddHandler btnAddExtra.Click, AddressOf HandleBtnExtraClick
            AddButtons.Add(intI, btnAddExtra)
            If CategoryTextBoxes(intI).Text = "Conveyance" Or _
                 CategoryTextBoxes(intI).Text = "Major Others" Then btnAddExtra.Visible = False

            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtConcreteQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
            VP = VP + 20
            HP = 1
        Next
    End Sub
    Private Sub LoadControlsFromMajOtherArray(ByVal mitems As Integer)
        Dim intI As Integer
        HP = 1
        VP = 0
        Dim mcategoryname As String = "Others"
        Me.tbcBdgetHeads.TabPages(0).Text = "Major Equipments - " & mcategoryname
        CategoryTextBoxes.Clear()
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        QtyTextBoxes.Clear()
        HrsPermonthTextBoxes.Clear()
        concreteqtyTextboxes.Clear()
        Checkboxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        DepPercComboboxes.Clear()
        ShiftsComboboxes.Clear()
        AddButtons.Clear()
        mTabindex = 1
        For intI = 0 To mitems - 1
            Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
            Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
            Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
            Dim txtHrsPerMonth As New TextBox
            Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox
            Dim chkSelected As New CheckBox
            Dim txtConcreteQty As New TextBox   ', cmbDrive As New ComboBox
            Dim btnAddExtra As New Button

            chkSelected.Checked = False
            chkSelected.Height = 30
            chkSelected.Width = 24
            chkSelected.Left = HP
            chkSelected.Top = VP - 3
            HP = HP + chkSelected.Width + 1
            'txtMaintPerc.Font = New Font("verdana", 10)
            chkSelected.Tag = intI
            chkSelected.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler chkSelected.CheckedChanged, AddressOf HandleCheckboxStatus
            Checkboxes.Add(intI, chkSelected)

            txtCategory.Text = majoCategory(intI)
            txtCategory.WordWrap = True
            txtCategory.Height = 50
            txtCategory.Width = 65
            txtCategory.TextAlign = HorizontalAlignment.Center
            txtCategory.Left = HP
            txtCategory.Top = VP
            HP = HP + txtCategory.Width + 1
            txtCategory.Font = New Font("arial", 8)
            txtCategory.Tag = intI
            txtCategory.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCategory.Enabled = False
            CategoryTextBoxes.Add(intI, txtCategory)

            txtEquipname.Text = majoEName(intI)
            txtEquipname.WordWrap = True
            txtEquipname.Height = 60
            txtEquipname.Width = 150
            txtEquipname.TextAlign = HorizontalAlignment.Center
            txtEquipname.Left = HP
            txtEquipname.Top = VP
            HP = HP + txtEquipname.Width + 1
            txtEquipname.Font = New Font("arial", 8)
            txtEquipname.Tag = intI
            txtEquipname.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtEquipname.Enabled = False
            EquipNameTextBoxes.Add(intI, txtEquipname)

            txtCapacity.Text = majoCapacity(intI)
            txtCapacity.WordWrap = True
            txtCapacity.Height = 30
            txtCapacity.Width = 70
            txtCapacity.TextAlign = HorizontalAlignment.Center
            txtCapacity.Left = HP
            txtCapacity.Top = VP
            HP = HP + txtCapacity.Width + 1
            txtCapacity.Font = New Font("arial", 8)
            txtCapacity.Tag = intI
            txtCapacity.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCapacity.Enabled = False
            CapacityTextBoxes.Add(intI, txtCapacity)

            txtMakeModel.Text = majoMakeModel(intI)
            txtMakeModel.WordWrap = True
            txtMakeModel.Height = 50
            txtMakeModel.Width = 160
            txtMakeModel.TextAlign = HorizontalAlignment.Center
            txtMakeModel.Left = HP
            txtMakeModel.Top = VP
            HP = HP + txtMakeModel.Width + 1
            txtMakeModel.Font = New Font("arial", 8)
            txtMakeModel.Tag = intI
            txtMakeModel.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtMakeModel.Enabled = False
            MakeModelTextBoxes.Add(intI, txtMakeModel)

            dpMobDate.Value = majoMDate(intI).Date
            dpMobDate.Name = "MobDate"
            dpMobDate.Format = DateTimePickerFormat.Custom
            dpMobDate.CustomFormat = "dd-MMM-yyyy"
            dpMobDate.Height = 30
            dpMobDate.Width = 100
            dpMobDate.Left = HP
            dpMobDate.Top = VP
            HP = HP + dpMobDate.Width + 1
            dpMobDate.Font = New Font("arial", 8)
            dpMobDate.Enabled = False
            dpMobDate.Tag = intI
            dpMobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
            MobdatePickers.Add(intI, dpMobDate)

            dpDemobDate.Value = majoDMdate(intI)
            dpDemobDate.Name = "DemobDate"
            dpDemobDate.Format = DateTimePickerFormat.Custom
            dpDemobDate.CustomFormat = "dd-MMM-yyyy"
            dpDemobDate.Height = 30
            dpDemobDate.Width = 100
            dpDemobDate.Left = HP
            dpDemobDate.Top = VP
            HP = HP + dpDemobDate.Width + 1
            dpDemobDate.Font = New Font("arial", 8)
            dpDemobDate.Enabled = False
            dpDemobDate.Tag = intI
            dpDemobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
            DemobDatePickers.Add(intI, dpDemobDate)

            txtQty.Text = majoQuantity(intI)
            txtQty.Height = 30
            txtQty.Width = 30
            txtQty.Left = HP
            txtQty.Top = VP
            HP = HP + txtQty.Width + 1
            txtQty.Font = New Font("arial", 8)
            txtQty.Enabled = False
            txtQty.Tag = intI
            txtQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler txtQty.Validated, AddressOf HandleQty
            QtyTextBoxes.Add(intI, txtQty)


            txtHrsPerMonth.Text = majoHrsPerMonth(intI)
            txtHrsPerMonth.Height = 30
            txtHrsPerMonth.Width = 60
            txtHrsPerMonth.Left = HP
            txtHrsPerMonth.Top = VP
            HP = HP + txtHrsPerMonth.Width + 1
            txtHrsPerMonth.Font = New Font("arial", 8)
            txtHrsPerMonth.Enabled = False
            txtHrsPerMonth.Tag = intI
            txtHrsPerMonth.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            HrsPermonthTextBoxes.Add(intI, txtHrsPerMonth)

            cmbDepPerc.Items.Add(2.75)
            cmbDepPerc.Items.Add(1.25)
            cmbDepPerc.Items.Add(0.5)
            cmbDepPerc.Text = majoDep(intI)
            cmbDepPerc.Height = 30
            cmbDepPerc.Width = 60
            cmbDepPerc.Left = HP
            cmbDepPerc.Top = VP
            HP = HP + cmbDepPerc.Width + 1
            cmbDepPerc.Font = New Font("arial", 8)
            cmbDepPerc.Enabled = False
            cmbDepPerc.Tag = intI
            cmbDepPerc.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            DepPercComboboxes.Add(intI, cmbDepPerc)

            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(1.5)
            cmbShifts.Items.Add(2)
            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(0)
            cmbShifts.Text = majoShift(intI)
            cmbShifts.Height = 30
            cmbShifts.Width = 45
            cmbShifts.Left = HP
            cmbShifts.Top = VP
            HP = HP + cmbShifts.Width + 1
            cmbShifts.Font = New Font("arial", 8)
            cmbShifts.Enabled = False
            cmbShifts.Tag = intI
            cmbShifts.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            ShiftsComboboxes.Add(intI, cmbShifts)

            txtConcreteQty.Text = mConcreteQty
            txtConcreteQty.Height = 30
            txtConcreteQty.Width = 60
            txtConcreteQty.Left = HP
            txtConcreteQty.Top = VP
            txtConcreteQty.Font = New Font("arial", 8)
            txtConcreteQty.Enabled = False
            txtConcreteQty.Tag = intI
            txtConcreteQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            concreteqtyTextboxes.Add(intI, txtConcreteQty)
            If CategoryTextBoxes(intI).Text <> "Concreting" Then concreteqtyTextboxes(intI).Visible = False

            btnAddExtra.Text = "Add"
            btnAddExtra.Height = 20
            btnAddExtra.Width = 40
            btnAddExtra.Left = HP
            btnAddExtra.Top = VP
            btnAddExtra.Font = New Font("arial", 8)
            btnAddExtra.Enabled = False

            btnAddExtra.Tag = intI
            btnAddExtra.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
            ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
            AddHandler btnAddExtra.Click, AddressOf HandleBtnExtraClick
            AddButtons.Add(intI, btnAddExtra)

            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtConcreteQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
            VP = VP + 20
            HP = 1
        Next
    End Sub
    Private Sub LoadControlsFromFexpArray(ByVal fexpItems As Integer)
        Dim intI As Integer
        HP = 1
        VP = 0
        mcategory = "FixedE Expenses"
        Me.tbcBdgetHeads.TabPages(0).Text = "Fixed Expenses"
        CategoryTextBoxes.Clear()
        QtyTextBoxes.Clear()
        CostTextBoxes.Clear()
        RemarksTextBoxes.Clear()
        Checkboxes.Clear()
        AmountTextBoxes.Clear()
        ClientBillingTextBoxes.Clear()
        CostPercTextBoxes.Clear()
        mTabindex = 1
        For intI = 0 To fexpItems - 1
            Dim txtCategory As New TextBox, txtQty As New TextBox, txtCost As New TextBox
            Dim txtRemarks As New TextBox, chkSelected As New CheckBox
            Dim txtClientBilling As New TextBox, txtCostperc As New TextBox, txtAmount As New TextBox

            chkSelected.Checked = False     'False
            chkSelected.Height = 30
            chkSelected.Width = 24
            chkSelected.Left = HP
            chkSelected.Top = VP - 3
            HP = HP + chkSelected.Width + 1
            'txtMaintPerc.Font = New Font("verdana", 10)
            chkSelected.Tag = intI
            chkSelected.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler chkSelected.CheckedChanged, AddressOf HandleFexpCheckboxStatus
            Checkboxes.Add(intI, chkSelected)

            txtCategory.Text = fexpCategory(intI)
            txtCategory.WordWrap = True
            txtCategory.Height = 50
            txtCategory.Width = 220
            txtCategory.TextAlign = HorizontalAlignment.Center
            txtCategory.Left = HP
            txtCategory.Top = VP
            HP = HP + txtCategory.Width + 1
            txtCategory.Font = New Font("arial", 8)
            txtCategory.Tag = intI
            txtCategory.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCategory.Enabled = False
            CategoryTextBoxes.Add(intI, txtCategory)

            txtQty.Text = FExpQty(intI)
            txtQty.Height = 60
            txtQty.Width = 40
            txtQty.TextAlign = HorizontalAlignment.Center
            txtQty.Left = HP
            txtQty.Top = VP
            HP = HP + txtQty.Width + 1
            txtQty.Font = New Font("arial", 8)
            txtQty.Tag = intI
            txtQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtQty.Enabled = False
            AddHandler txtQty.Validated, AddressOf HandleFexpAmtCalc
            QtyTextBoxes.Add(intI, txtQty)

            txtCost.Text = fexpCost(intI)
            txtCost.WordWrap = True
            txtCost.Height = 30
            txtCost.Width = 50
            txtCost.TextAlign = HorizontalAlignment.Center
            txtCost.Left = HP
            txtCost.Top = VP
            HP = HP + txtCost.Width + 1
            txtCost.Font = New Font("arial", 8)
            txtCost.Tag = intI
            txtCost.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCost.Enabled = False
            AddHandler txtCost.Validated, AddressOf HandleFexpAmtCalc
            CostTextBoxes.Add(intI, txtCost)

            txtAmount.Text = fexpAmount(intI)
            txtAmount.WordWrap = True
            txtAmount.Height = 30
            txtAmount.Width = 80
            txtAmount.TextAlign = HorizontalAlignment.Center
            txtAmount.Left = HP
            txtAmount.Top = VP
            HP = HP + txtAmount.Width + 1
            txtAmount.Font = New Font("arial", 8)
            txtAmount.Tag = intI
            txtAmount.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtAmount.Enabled = False
            AmountTextBoxes.Add(intI, txtAmount)

            txtClientBilling.Text = fexpClientBilling(intI)
            txtClientBilling.WordWrap = True
            txtClientBilling.Height = 30
            txtClientBilling.Width = 80
            txtClientBilling.TextAlign = HorizontalAlignment.Center
            txtClientBilling.Left = HP
            txtClientBilling.Top = VP
            HP = HP + txtClientBilling.Width + 1
            txtClientBilling.Font = New Font("arial", 8)
            txtClientBilling.Tag = intI
            txtClientBilling.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtClientBilling.Enabled = False
            ClientBillingTextBoxes.Add(intI, txtClientBilling)

            txtCostperc.Text = fexpCostPerc(intI)
            txtCostperc.WordWrap = True
            txtCostperc.Height = 30
            txtCostperc.Width = 80
            txtCostperc.TextAlign = HorizontalAlignment.Center
            txtCostperc.Left = HP
            txtCostperc.Top = VP
            HP = HP + txtCostperc.Width + 1
            txtCostperc.Font = New Font("arial", 8)
            txtCostperc.Tag = intI
            txtCostperc.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCostperc.Enabled = False
            CostPercTextBoxes.Add(intI, txtCostperc)

            txtRemarks.Text = FExpRemarks(intI)
            txtRemarks.WordWrap = True
            txtRemarks.Height = 50
            txtRemarks.Width = 300
            txtRemarks.TextAlign = HorizontalAlignment.Center
            txtRemarks.Left = HP
            txtRemarks.Top = VP
            HP = HP + txtRemarks.Width + 1
            txtRemarks.Font = New Font("arial", 8)
            txtRemarks.Tag = intI
            txtRemarks.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtRemarks.Enabled = False
            RemarksTextBoxes.Add(intI, txtRemarks)

            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCost)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtRemarks)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtAmount)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtClientBilling)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCostperc)
            VP = VP + 20
            HP = 1
        Next
    End Sub
    Private Sub LoadControlsFromBPFexpArray(ByVal BPfexpItems As Integer)
        Dim intI As Integer
        HP = 1
        VP = 0
        mcategory = "BP FixedE Expenses"
        Me.tbcBdgetHeads.TabPages(0).Text = "BP Fixed Expenses"
        CategoryTextBoxes.Clear()
        QtyTextBoxes.Clear()
        CostTextBoxes.Clear()
        RemarksTextBoxes.Clear()
        Checkboxes.Clear()
        AmountTextBoxes.Clear()
        ClientBillingTextBoxes.Clear()
        CostPercTextBoxes.Clear()
        mTabindex = 1
        For intI = 0 To BPfexpItems - 1
            Dim txtCategory As New TextBox, txtQty As New TextBox, txtCost As New TextBox
            Dim txtRemarks As New TextBox, chkSelected As New CheckBox
            Dim txtClientBilling As New TextBox, txtCostperc As New TextBox, txtAmount As New TextBox

            chkSelected.Checked = False
            chkSelected.Height = 30
            chkSelected.Width = 24
            chkSelected.Left = HP
            chkSelected.Top = VP - 3
            HP = HP + chkSelected.Width + 1
            chkSelected.Tag = intI
            chkSelected.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler chkSelected.CheckedChanged, AddressOf HandleFexpCheckboxStatus
            Checkboxes.Add(intI, chkSelected)

            txtCategory.Text = BPfexpCategory(intI)
            txtCategory.WordWrap = True
            txtCategory.Height = 50
            txtCategory.Width = 220
            txtCategory.TextAlign = HorizontalAlignment.Center
            txtCategory.Left = HP
            txtCategory.Top = VP
            HP = HP + txtCategory.Width + 1
            txtCategory.Font = New Font("arial", 8)
            txtCategory.Tag = intI
            txtCategory.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCategory.Enabled = False
            CategoryTextBoxes.Add(intI, txtCategory)

            txtQty.Text = BPFExpQty(intI)
            txtQty.Height = 60
            txtQty.Width = 40
            txtQty.TextAlign = HorizontalAlignment.Center
            txtQty.Left = HP
            txtQty.Top = VP
            HP = HP + txtQty.Width + 1
            txtQty.Font = New Font("arial", 8)
            txtQty.Tag = intI
            txtQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtQty.Enabled = False
            AddHandler txtQty.Validated, AddressOf HandleFexpAmtCalc
            QtyTextBoxes.Add(intI, txtQty)

            txtCost.Text = BPfexpCost(intI)
            txtCost.WordWrap = True
            txtCost.Height = 30
            txtCost.Width = 50
            txtCost.TextAlign = HorizontalAlignment.Center
            txtCost.Left = HP
            txtCost.Top = VP
            HP = HP + txtCost.Width + 1
            txtCost.Font = New Font("arial", 8)
            txtCost.Tag = intI
            txtCost.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCost.Enabled = False
            AddHandler txtCost.Validated, AddressOf HandleFexpAmtCalc
            CostTextBoxes.Add(intI, txtCost)

            txtAmount.Text = BPfexpAmount(intI)
            txtAmount.WordWrap = True
            txtAmount.Height = 30
            txtAmount.Width = 80
            txtAmount.TextAlign = HorizontalAlignment.Center
            txtAmount.Left = HP
            txtAmount.Top = VP
            HP = HP + txtAmount.Width + 1
            txtAmount.Font = New Font("arial", 8)
            txtAmount.Tag = intI
            txtAmount.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtAmount.Enabled = False
            AmountTextBoxes.Add(intI, txtAmount)

            txtClientBilling.Text = BPfexpClientBilling(intI)
            txtClientBilling.WordWrap = True
            txtClientBilling.Height = 30
            txtClientBilling.Width = 80
            txtClientBilling.TextAlign = HorizontalAlignment.Center
            txtClientBilling.Left = HP
            txtClientBilling.Top = VP
            HP = HP + txtClientBilling.Width + 1
            txtClientBilling.Font = New Font("arial", 8)
            txtClientBilling.Tag = intI
            txtClientBilling.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtClientBilling.Enabled = False
            ClientBillingTextBoxes.Add(intI, txtClientBilling)

            txtCostperc.Text = BPfexpCostPerc(intI)
            txtCostperc.WordWrap = True
            txtCostperc.Height = 30
            txtCostperc.Width = 80
            txtCostperc.TextAlign = HorizontalAlignment.Center
            txtCostperc.Left = HP
            txtCostperc.Top = VP
            HP = HP + txtCostperc.Width + 1
            txtCostperc.Font = New Font("arial", 8)
            txtCostperc.Tag = intI
            txtCostperc.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCostperc.Enabled = False
            CostPercTextBoxes.Add(intI, txtCostperc)

            txtRemarks.Text = BPFExpRemarks(intI)
            txtRemarks.WordWrap = True
            txtRemarks.Height = 50
            txtRemarks.Width = 300
            txtRemarks.TextAlign = HorizontalAlignment.Center
            txtRemarks.Left = HP
            txtRemarks.Top = VP
            HP = HP + txtRemarks.Width + 1
            txtRemarks.Font = New Font("arial", 8)
            txtRemarks.Tag = intI
            txtRemarks.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtRemarks.Enabled = False
            RemarksTextBoxes.Add(intI, txtRemarks)

            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCost)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtRemarks)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtAmount)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtClientBilling)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCostperc)
            VP = VP + 20
            HP = 1
        Next
    End Sub
    Private Sub LoadControlsFromMinorArray(ByVal MinorItems As Integer)
        Dim intI As Integer
        HP = 1
        VP = 0
        mcategory = "Minor Equipments"
        Me.tbcBdgetHeads.TabPages(0).Text = "Minor Equipments"
        CategoryTextBoxes.Clear()
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        QtyTextBoxes.Clear()
        HrsPermonthTextBoxes.Clear()
        PurchvalTextBoxes.Clear()
        Checkboxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        DepPercComboboxes.Clear()
        ShiftsComboboxes.Clear()
        AddButtons.Clear()
        IsNewMC.Clear()
        'Checkboxes.Clear()
        mTabindex = 1
        For intI = 0 To MinorItems - 1
            Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
            Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
            Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
            Dim txtHrsPerMonth As New TextBox
            Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox, txtPurchVal As New TextBox
            Dim chkSelected As New CheckBox, IsNew As New CheckBox
            'Dim txtConcreteQty As New TextBox   ', cmbDrive As New ComboBox
            Dim btnAddExtra As New Button

            chkSelected.Checked = False
            chkSelected.Height = 30
            chkSelected.Width = 24
            chkSelected.Left = HP
            chkSelected.Top = VP - 3
            HP = HP + chkSelected.Width + 1
            'txtMaintPerc.Font = New Font("verdana", 10)
            chkSelected.Tag = intI
            chkSelected.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler chkSelected.CheckedChanged, AddressOf HandleMinorCheckboxStatus
            Checkboxes.Add(intI, chkSelected)

            txtCategory.Text = MinorCategory(intI)
            txtCategory.WordWrap = True
            txtCategory.Height = 50
            txtCategory.Width = 80
            txtCategory.TextAlign = HorizontalAlignment.Center
            txtCategory.Left = HP
            txtCategory.Top = VP
            HP = HP + txtCategory.Width + 1
            txtCategory.Font = New Font("arial", 8)
            txtCategory.Tag = intI
            txtCategory.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCategory.Enabled = False
            CategoryTextBoxes.Add(intI, txtCategory)

            txtEquipname.Text = MinorEName(intI)
            txtEquipname.WordWrap = True
            txtEquipname.Height = 60
            txtEquipname.Width = 150
            txtEquipname.TextAlign = HorizontalAlignment.Center
            txtEquipname.Left = HP
            txtEquipname.Top = VP
            HP = HP + txtEquipname.Width + 1
            txtEquipname.Font = New Font("arial", 8)
            txtEquipname.Tag = intI
            txtEquipname.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtEquipname.Enabled = False
            EquipNameTextBoxes.Add(intI, txtEquipname)

            txtCapacity.Text = MinorCapacity(intI)
            txtCapacity.WordWrap = True
            txtCapacity.Height = 30
            txtCapacity.Width = 70
            txtCapacity.TextAlign = HorizontalAlignment.Center
            txtCapacity.Left = HP
            txtCapacity.Top = VP
            HP = HP + txtCapacity.Width + 1
            txtCapacity.Font = New Font("arial", 8)
            txtCapacity.Tag = intI
            txtCapacity.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCapacity.Enabled = False
            CapacityTextBoxes.Add(intI, txtCapacity)

            txtMakeModel.Text = MinorMakeModel(intI)
            txtMakeModel.WordWrap = True
            txtMakeModel.Height = 50
            txtMakeModel.Width = 160
            txtMakeModel.TextAlign = HorizontalAlignment.Center
            txtMakeModel.Left = HP
            txtMakeModel.Top = VP
            HP = HP + txtMakeModel.Width + 1
            txtMakeModel.Font = New Font("arial", 8)
            txtMakeModel.Tag = intI
            txtMakeModel.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtMakeModel.Enabled = False
            MakeModelTextBoxes.Add(intI, txtMakeModel)

            dpMobDate.Value = MinorMDate(intI).Date
            dpMobDate.Name = "MobDate"
            dpMobDate.Format = DateTimePickerFormat.Custom
            dpMobDate.CustomFormat = "dd-MMM-yyyy"
            dpMobDate.Height = 30
            dpMobDate.Width = 100
            dpMobDate.Left = HP
            dpMobDate.Top = VP
            HP = HP + dpMobDate.Width + 1
            dpMobDate.Font = New Font("arial", 8)
            dpMobDate.Enabled = False
            dpMobDate.Tag = intI
            dpMobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
            MobdatePickers.Add(intI, dpMobDate)

            dpDemobDate.Value = MinorDMdate(intI)
            dpDemobDate.Name = "DemobDate"
            dpDemobDate.Format = DateTimePickerFormat.Custom
            dpDemobDate.CustomFormat = "dd-MMM-yyyy"
            dpDemobDate.Height = 30
            dpDemobDate.Width = 100
            dpDemobDate.Left = HP
            dpDemobDate.Top = VP
            HP = HP + dpDemobDate.Width + 1
            dpDemobDate.Font = New Font("arial", 8)
            dpDemobDate.Enabled = False
            dpDemobDate.Tag = intI
            dpDemobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
            DemobDatePickers.Add(intI, dpDemobDate)

            txtQty.Text = MinorQuantity(intI)
            txtQty.Height = 30
            txtQty.Width = 30
            txtQty.Left = HP
            txtQty.Top = VP
            HP = HP + txtQty.Width + 1
            txtQty.Font = New Font("arial", 8)
            txtQty.Enabled = False
            txtQty.Tag = intI
            txtQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler txtQty.Validated, AddressOf HandleQty
            QtyTextBoxes.Add(intI, txtQty)


            txtHrsPerMonth.Text = MinorHrsPerMonth(intI)
            txtHrsPerMonth.Height = 30
            txtHrsPerMonth.Width = 60
            txtHrsPerMonth.Left = HP
            txtHrsPerMonth.Top = VP
            HP = HP + txtHrsPerMonth.Width + 5
            txtHrsPerMonth.Font = New Font("arial", 8)
            txtHrsPerMonth.Enabled = False
            txtHrsPerMonth.Tag = intI
            txtHrsPerMonth.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            HrsPermonthTextBoxes.Add(intI, txtHrsPerMonth)

            IsNew.Checked = MinorIsNew(intI).ToString()
            IsNew.Height = 30
            IsNew.Width = 40
            IsNew.Left = HP
            IsNew.Top = VP - 3
            HP = HP + IsNew.Width + 1
            'txtPurchValue.Font = New Font("arial", 8)
            IsNew.Tag = intI
            IsNew.TabIndex = TabIndex
            TabIndex = TabIndex + 1
            IsNew.Visible = True
            IsNew.Enabled = False
            AddHandler IsNew.CheckedChanged, AddressOf HandleNewMCStatus
            IsNewMC.Add(intI, IsNew)

            txtPurchVal.Text = MinorNewPurchVal(intI)
            txtPurchVal.Height = 30
            txtPurchVal.Width = 80
            txtPurchVal.Left = HP
            txtPurchVal.Top = VP
            HP = HP + txtPurchVal.Width + 1
            txtPurchVal.Font = New Font("arial", 8)
            txtPurchVal.Tag = intI
            txtPurchVal.TabIndex = TabIndex
            TabIndex = TabIndex + 1
            txtPurchVal.Visible = True
            txtPurchVal.Enabled = False
            PurchvalTextBoxes.Add(intI, txtPurchVal)

            cmbDepPerc.Items.Add(2.75)
            cmbDepPerc.Items.Add(1.25)
            cmbDepPerc.Items.Add(0.5)
            cmbDepPerc.Text = MinorDep(intI)
            cmbDepPerc.Height = 30
            cmbDepPerc.Width = 45
            cmbDepPerc.Left = HP
            cmbDepPerc.Top = VP
            cmbDepPerc.Font = New Font("arial", 8)
            cmbDepPerc.Visible = False
            cmbDepPerc.Enabled = False
            cmbDepPerc.Tag = intI
            cmbDepPerc.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            DepPercComboboxes.Add(intI, cmbDepPerc)

            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(1.5)
            cmbShifts.Items.Add(2)
            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(0)
            cmbShifts.SelectedIndex = 0
            cmbShifts.Height = 30
            cmbShifts.Width = 45
            cmbShifts.Left = HP
            cmbShifts.Top = VP
            cmbShifts.Font = New Font("arial", 8)
            cmbShifts.Enabled = False
            cmbShifts.Visible = False
            cmbShifts.Tag = intI
            cmbShifts.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            ShiftsComboboxes.Add(intI, cmbShifts)

            btnAddExtra.Text = "Add"
            btnAddExtra.Height = 20
            btnAddExtra.Width = 40
            btnAddExtra.Left = HP
            btnAddExtra.Top = VP
            btnAddExtra.Font = New Font("arial", 8)
            btnAddExtra.Enabled = False
            btnAddExtra.Tag = intI
            btnAddExtra.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
            ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
            AddHandler btnAddExtra.Click, AddressOf HandleMinorBtnExtraClick
            AddButtons.Add(intI, btnAddExtra)

            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(IsNew)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtPurchVal)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
            VP = VP + 20
            HP = 1
        Next
    End Sub
    Private Sub LoadControlsFromHiredArray(ByVal Hireditems As Integer)
        Dim intI As Integer
        HP = 1
        VP = 0
        mcategory = "Minor Equipments"
        Me.tbcBdgetHeads.TabPages(0).Text = "Hired Equipments"
        CategoryTextBoxes.Clear()
        EquipNameTextBoxes.Clear()
        CapacityTextBoxes.Clear()
        MakeModelTextBoxes.Clear()
        QtyTextBoxes.Clear()
        HrsPermonthTextBoxes.Clear()
        HireChargesTextBoxes.Clear()
        Checkboxes.Clear()
        MobdatePickers.Clear()
        DemobDatePickers.Clear()
        DepPercComboboxes.Clear()
        ShiftsComboboxes.Clear()
        AddButtons.Clear()
        mTabindex = 1
        For intI = 0 To Hireditems - 1
            Dim txtCategory As New TextBox, txtEquipname As New TextBox, txtCapacity As New TextBox
            Dim txtMakeModel As New TextBox, txtQty As New TextBox  ', txtmodel As New TextBox
            Dim dpMobDate As New DateTimePicker, dpDemobDate As New DateTimePicker
            Dim txtHrsPerMonth As New TextBox
            Dim cmbDepPerc As New ComboBox, cmbShifts As New ComboBox, txtHireVal As New TextBox
            Dim chkSelected As New CheckBox   ', IsNew As New CheckBox
            'Dim txtConcreteQty As New TextBox   ', cmbDrive As New ComboBox
            Dim btnAddExtra As New Button

            chkSelected.Checked = False
            chkSelected.Height = 30
            chkSelected.Width = 24
            chkSelected.Left = HP
            chkSelected.Top = VP - 3
            HP = HP + chkSelected.Width + 1
            chkSelected.Tag = intI
            chkSelected.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler chkSelected.CheckedChanged, AddressOf HandleHireCheckboxStatus
            Checkboxes.Add(intI, chkSelected)

            txtCategory.Text = HiredCategory(intI)
            txtCategory.WordWrap = True
            txtCategory.Height = 50
            txtCategory.Width = 80
            txtCategory.TextAlign = HorizontalAlignment.Center
            txtCategory.Left = HP
            txtCategory.Top = VP
            HP = HP + txtCategory.Width + 1
            txtCategory.Font = New Font("arial", 8)
            txtCategory.Tag = intI
            txtCategory.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCategory.Enabled = False
            CategoryTextBoxes.Add(intI, txtCategory)

            txtEquipname.Text = HiredEName(intI)
            txtEquipname.WordWrap = True
            txtEquipname.Height = 60
            txtEquipname.Width = 150
            txtEquipname.TextAlign = HorizontalAlignment.Center
            txtEquipname.Left = HP
            txtEquipname.Top = VP
            HP = HP + txtEquipname.Width + 1
            txtEquipname.Font = New Font("arial", 8)
            txtEquipname.Tag = intI
            txtEquipname.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtEquipname.Enabled = False
            EquipNameTextBoxes.Add(intI, txtEquipname)

            txtCapacity.Text = HiredCapacity(intI)
            txtCapacity.WordWrap = True
            txtCapacity.Height = 30
            txtCapacity.Width = 70
            txtCapacity.TextAlign = HorizontalAlignment.Center
            txtCapacity.Left = HP
            txtCapacity.Top = VP
            HP = HP + txtCapacity.Width + 1
            txtCapacity.Font = New Font("arial", 8)
            txtCapacity.Tag = intI
            txtCapacity.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtCapacity.Enabled = False
            CapacityTextBoxes.Add(intI, txtCapacity)

            txtMakeModel.Text = HiredMakeModel(intI)
            txtMakeModel.WordWrap = True
            txtMakeModel.Height = 50
            txtMakeModel.Width = 160
            txtMakeModel.TextAlign = HorizontalAlignment.Center
            txtMakeModel.Left = HP
            txtMakeModel.Top = VP
            HP = HP + txtMakeModel.Width + 1
            txtMakeModel.Font = New Font("arial", 8)
            txtMakeModel.Tag = intI
            txtMakeModel.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            txtMakeModel.Enabled = False
            MakeModelTextBoxes.Add(intI, txtMakeModel)

            dpMobDate.Value = HiredMDate(intI).Date
            dpMobDate.Name = "MobDate"
            dpMobDate.Format = DateTimePickerFormat.Custom
            dpMobDate.CustomFormat = "dd-MMM-yyyy"
            dpMobDate.Height = 30
            dpMobDate.Width = 100
            dpMobDate.Left = HP
            dpMobDate.Top = VP
            HP = HP + dpMobDate.Width + 1
            dpMobDate.Font = New Font("arial", 8)
            dpMobDate.Enabled = False
            dpMobDate.Tag = intI
            dpMobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpMobDate.Validated, AddressOf HandleCheckDates
            MobdatePickers.Add(intI, dpMobDate)

            dpDemobDate.Value = HiredDMdate(intI)
            dpDemobDate.Name = "DemobDate"
            dpDemobDate.Format = DateTimePickerFormat.Custom
            dpDemobDate.CustomFormat = "dd-MMM-yyyy"
            dpDemobDate.Height = 30
            dpDemobDate.Width = 100
            dpDemobDate.Left = HP
            dpDemobDate.Top = VP
            HP = HP + dpDemobDate.Width + 1
            dpDemobDate.Font = New Font("arial", 8)
            dpDemobDate.Enabled = False
            dpDemobDate.Tag = intI
            dpDemobDate.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler dpDemobDate.Validated, AddressOf HandleCheckDates
            DemobDatePickers.Add(intI, dpDemobDate)

            txtQty.Text = HiredQuantity(intI)
            txtQty.Height = 30
            txtQty.Width = 30
            txtQty.Left = HP
            txtQty.Top = VP
            HP = HP + txtQty.Width + 1
            txtQty.Font = New Font("arial", 8)
            txtQty.Enabled = False
            txtQty.Tag = intI
            txtQty.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            AddHandler txtQty.Validated, AddressOf HandleQty
            QtyTextBoxes.Add(intI, txtQty)


            txtHrsPerMonth.Text = HiredHrsPerMonth(intI)
            txtHrsPerMonth.Height = 30
            txtHrsPerMonth.Width = 60
            txtHrsPerMonth.Left = HP
            txtHrsPerMonth.Top = VP
            HP = HP + txtHrsPerMonth.Width + 5
            txtHrsPerMonth.Font = New Font("arial", 8)
            txtHrsPerMonth.Enabled = False
            txtHrsPerMonth.Tag = intI
            txtHrsPerMonth.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            HrsPermonthTextBoxes.Add(intI, txtHrsPerMonth)

            txtHireVal.Text = HiredHireCharges(intI)
            txtHireVal.Height = 30
            txtHireVal.Width = 80
            txtHireVal.Left = HP
            txtHireVal.Top = VP
            HP = HP + txtHireVal.Width + 1
            txtHireVal.Font = New Font("arial", 8)
            txtHireVal.Tag = intI
            txtHireVal.TabIndex = TabIndex
            TabIndex = TabIndex + 2
            txtHireVal.Visible = True
            txtHireVal.Enabled = False
            HireChargesTextBoxes.Add(intI, txtHireVal)

            cmbDepPerc.Items.Add(2.75)
            cmbDepPerc.Items.Add(1.25)
            cmbDepPerc.Items.Add(0.5)
            cmbDepPerc.Text = HiredDep(intI)
            cmbDepPerc.Height = 30
            cmbDepPerc.Width = 45
            cmbDepPerc.Left = HP
            cmbDepPerc.Top = VP
            cmbDepPerc.Font = New Font("arial", 8)
            cmbDepPerc.Visible = False
            cmbDepPerc.Enabled = False
            cmbDepPerc.Tag = intI
            cmbDepPerc.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            DepPercComboboxes.Add(intI, cmbDepPerc)

            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(1.5)
            cmbShifts.Items.Add(2)
            cmbShifts.Items.Add(1)
            cmbShifts.Items.Add(0)
            cmbShifts.SelectedIndex = 0
            cmbShifts.Height = 30
            cmbShifts.Width = 45
            cmbShifts.Left = HP
            cmbShifts.Top = VP
            cmbShifts.Font = New Font("arial", 8)
            cmbShifts.Enabled = False
            cmbShifts.Visible = False
            cmbShifts.Tag = intI
            cmbShifts.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            ShiftsComboboxes.Add(intI, cmbShifts)

            btnAddExtra.Text = "Add"
            btnAddExtra.Height = 20
            btnAddExtra.Width = 40
            btnAddExtra.Left = HP
            btnAddExtra.Top = VP
            btnAddExtra.Font = New Font("arial", 8)
            btnAddExtra.Enabled = False
            btnAddExtra.Tag = intI
            btnAddExtra.TabIndex = mTabindex
            mTabindex = mTabindex + 1
            Dim ToolTip1 As System.Windows.Forms.ToolTip = New System.Windows.Forms.ToolTip()
            ToolTip1.SetToolTip(btnAddExtra, "Click to Add more entries for this Equiment")
            AddHandler btnAddExtra.Click, AddressOf HandleHireBtnExtraClick
            AddButtons.Add(intI, btnAddExtra)

            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCategory)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtEquipname)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtCapacity)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtMakeModel)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtQty)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpMobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(dpDemobDate)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHrsPerMonth)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbDepPerc)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(cmbShifts)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(txtHireVal)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(chkSelected)
            Me.tbcBdgetHeads.TabPages(0).Controls.Add(btnAddExtra)
            VP = VP + 20
            HP = 1
        Next
    End Sub
    Private Sub optMajConvyance_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMajConvyance.CheckedChanged
        Dim intL As Integer
        If optMinorEquips.Checked = True Then
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optHireEquipments.Checked Then
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optFixedExp.Checked Or optBPFixedExp.Checked Then
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        ElseIf Me.optLightingEquips.Checked Then
            Panel7.Left = 1
            Panel7.Top = 20
            Panel7.Height = 30
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        Else
            Panel2.Left = 1
            Panel2.Top = 20
            Panel2.Height = 30
            Panel2.Visible = True
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If
        If Not optMajConvyance.Checked Then
            Dim intK As Integer, L As Integer
            ConveyanceItems = CategoryTextBoxes.Count
            L = ConveyanceItems - 1  'CategoryTextBoxes.Count - 1
            ReDim Preserve ConvChecked(L)
            ReDim Preserve ConvMobdate(L)
            ReDim Preserve ConvDemobdate(L)
            ReDim Preserve ConvQty(L)
            ReDim Preserve ConvHrs(L)
            ReDim Preserve ConvDepPerc(L)
            ReDim Preserve ConvShifts(L)
            ReDim convCategory(L)
            ReDim convEName(L)
            ReDim convCapacity(L)
            ReDim convMakeModel(L)
            ReDim convMDate(L)
            ReDim convDMdate(L)
            ReDim convQuantity(L)
            ReDim convHrsPerMonth(L)
            ReDim convDep(L)
            ReDim convShift(L)

            For intK = 0 To L
                convCategory(intK) = CategoryTextBoxes(intK).Text
                convEName(intK) = EquipNameTextBoxes(intK).Text
                convCapacity(intK) = CapacityTextBoxes(intK).Text
                convMakeModel(intK) = MakeModelTextBoxes(intK).Text
                convMDate(intK) = MobdatePickers(intK).Value.Date
                convDMdate(intK) = DemobDatePickers(intK).Value.Date
                convQuantity(intK) = Val(QtyTextBoxes(intK).Text)
                convHrsPerMonth(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                convDep(intK) = Val(DepPercComboboxes(intK).Text)
                convShift(intK) = Val(ShiftsComboboxes(intK).Text)
            Next
            For intK = 0 To L
                If Checkboxes(intK).Checked Then
                    ConvChecked(intK) = 1
                    ConvMobdate(intK) = MobdatePickers(intK).Value.Date
                    ConvDemobdate(intK) = DemobDatePickers(intK).Value.Date
                    ConvQty(intK) = Val(QtyTextBoxes(intK).Text)
                    ConvHrs(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                    ConvDepPerc(intK) = Val(DepPercComboboxes(intK).Text)
                    ConvShifts(intK) = Val(ShiftsComboboxes(intK).Text)
                Else
                    ConvChecked(intK) = 0
                End If
            Next
            ConveyanceItems = CategoryTextBoxes.Count
            WriteToConveyanceEquipsArray(ConveyanceItems)
            Exit Sub
        End If
        Button1.Enabled = True
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        mcategory = "Conveyance"
        SelectedCategory = mcategory
        If Not convFirsttime Then
            LoadControlsFromConvArray(ConveyanceItems)
            For intL = 0 To ConveyanceItems - 1
                If ConvChecked(intL) = 1 Then
                    Checkboxes(intL).Checked = True
                    MobdatePickers(intL).Value = ConvMobdate(intL)
                    DemobDatePickers(intL).Value = ConvDemobdate(intL)
                    QtyTextBoxes(intL).Text = ConvQty(intL)
                    HrsPermonthTextBoxes(intL).Text = ConvHrs(intL)
                    DepPercComboboxes(intL).Text = ConvDepPerc(intL)
                    ShiftsComboboxes(intL).Text = ConvShifts(intL)
                Else
                    Checkboxes(intL).Checked = False
                End If
            Next
            Me.Refresh()
        Else
            mcategory = "Conveyance"
            SelectedCategory = mcategory
            LoadControlsInpage(mcategory)
            convFirsttime = False
        End If
        ConveyanceItems = CategoryTextBoxes.Count
        'WriteToConveyanceEquipsArray(ConveyanceItems)
    End Sub

    Private Sub optMajCrane_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMajCrane.CheckedChanged
        Dim intL As Integer
        If optMinorEquips.Checked = True Then
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optHireEquipments.Checked Then
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optFixedExp.Checked Or optBPFixedExp.Checked Then
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        ElseIf Me.optLightingEquips.Checked Then
            Panel7.Left = 1
            Panel7.Top = 20
            Panel7.Height = 30
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        Else
            Panel2.Left = 1
            Panel2.Top = 20
            Panel2.Height = 30
            Panel2.Visible = True
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If
        If Not optMajCrane.Checked Then
            Dim intK As Integer, L As Integer
            CraneItems = CategoryTextBoxes.Count
            L = CraneItems - 1 'CategoryTextBoxes.Count - 1
            ReDim Preserve CraneChecked(L)
            ReDim Preserve CraneMobdate(L)
            ReDim Preserve CraneDemobdate(L)
            ReDim Preserve CraneQty(L)
            ReDim Preserve CraneHrs(L)
            ReDim Preserve CraneDepPerc(L)
            ReDim Preserve CraneShifts(L)
            ReDim cranCategory(L)
            ReDim cranEName(L)
            ReDim cranCapacity(L)
            ReDim cranMakeModel(L)
            ReDim cranMDate(L)
            ReDim cranDMdate(L)
            ReDim cranQuantity(L)
            ReDim cranHrsPerMonth(L)
            ReDim cranDep(L)
            ReDim cranShift(L)

            For intK = 0 To L
                cranCategory(intK) = CategoryTextBoxes(intK).Text
                cranEName(intK) = EquipNameTextBoxes(intK).Text
                cranCapacity(intK) = CapacityTextBoxes(intK).Text
                cranMakeModel(intK) = MakeModelTextBoxes(intK).Text
                cranMDate(intK) = MobdatePickers(intK).Value.Date
                cranDMdate(intK) = DemobDatePickers(intK).Value.Date
                cranQuantity(intK) = Val(QtyTextBoxes(intK).Text)
                cranHrsPerMonth(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                cranDep(intK) = Val(DepPercComboboxes(intK).Text)
                cranShift(intK) = Val(ShiftsComboboxes(intK).Text)
            Next
            For intK = 0 To L
                If Checkboxes(intK).Checked Then
                    CraneChecked(intK) = 1
                    CraneMobdate(intK) = MobdatePickers(intK).Value.Date
                    CraneDemobdate(intK) = DemobDatePickers(intK).Value.Date
                    CraneQty(intK) = Val(QtyTextBoxes(intK).Text)
                    CraneHrs(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                    CraneDepPerc(intK) = Val(DepPercComboboxes(intK).Text)
                    CraneShifts(intK) = Val(ShiftsComboboxes(intK).Text)
                Else
                    CraneChecked(intK) = 0
                End If
            Next
            CraneItems = CategoryTextBoxes.Count
            WriteToCraneEquipsArray(CraneItems)
            Exit Sub
        End If
        Button1.Enabled = True
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        mcategory = "Cranes"
        SelectedCategory = mcategory
        If Not cranFirstTime Then
            LoadControlsFromCraneArray(CraneItems)
            For intL = 0 To CraneItems - 1
                If CraneChecked(intL) = 1 Then
                    Checkboxes(intL).Checked = True
                    MobdatePickers(intL).Value = CraneMobdate(intL)
                    DemobDatePickers(intL).Value = CraneDemobdate(intL)
                    QtyTextBoxes(intL).Text = CraneQty(intL)
                    HrsPermonthTextBoxes(intL).Text = CraneHrs(intL)
                    DepPercComboboxes(intL).Text = CraneDepPerc(intL)
                    ShiftsComboboxes(intL).Text = CraneShifts(intL)
                Else
                    Checkboxes(intL).Checked = False
                End If
            Next
            Me.Refresh()
        Else
            mcategory = "Cranes"
            SelectedCategory = mcategory
            LoadControlsInpage(mcategory)
            For intL = 0 To Checkboxes.Count - 1
                If Checkboxes(intL).Checked Then
                    changeEnabled(True, intL)
                End If
            Next
            cranFirstTime = False
        End If
        CraneItems = CategoryTextBoxes.Count
    End Sub

    Private Sub optMajDGSets_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMajDGSets.CheckedChanged
        Dim intL As Integer
        If optMinorEquips.Checked = True Then
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optHireEquipments.Checked Then
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optFixedExp.Checked Or optBPFixedExp.Checked Then
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        ElseIf Me.optLightingEquips.Checked Then
            Panel7.Left = 1
            Panel7.Top = 20
            Panel7.Height = 30
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        Else
            Panel2.Left = 1
            Panel2.Top = 20
            Panel2.Height = 30
            Panel2.Visible = True
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If
        If Not optMajDGSets.Checked Then
            Dim intK As Integer, L As Integer
            DGSetItems = CategoryTextBoxes.Count
            L = DGSetItems - 1  'CategoryTextBoxes.Count
            ReDim Preserve DGSetsChecked(L)
            ReDim Preserve DGSetsMobdate(L)
            ReDim Preserve DGSetsDemobdate(L)
            ReDim Preserve DGSetsQty(L)
            ReDim Preserve DGSetsHrs(L)
            ReDim Preserve DGSetsDepPerc(L)
            ReDim Preserve DGSetsShifts(L)
            ReDim dgsetCategory(L)
            ReDim dgsetEName(L)
            ReDim dgsetCapacity(L)
            ReDim dgsetMakeModel(L)
            ReDim dgsetMDate(L)
            ReDim dgsetDMdate(L)
            ReDim dgsetQuantity(L)
            ReDim dgsetHrsPerMonth(L)
            ReDim dgsetDep(L)
            ReDim dgsetShift(L)

            For intK = 0 To L
                dgsetCategory(intK) = CategoryTextBoxes(intK).Text
                dgsetEName(intK) = EquipNameTextBoxes(intK).Text
                dgsetCapacity(intK) = CapacityTextBoxes(intK).Text
                dgsetMakeModel(intK) = MakeModelTextBoxes(intK).Text
                dgsetMDate(intK) = MobdatePickers(intK).Value.Date
                dgsetDMdate(intK) = DemobDatePickers(intK).Value.Date
                dgsetQuantity(intK) = Val(QtyTextBoxes(intK).Text)
                dgsetHrsPerMonth(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                dgsetDep(intK) = Val(DepPercComboboxes(intK).Text)
                dgsetShift(intK) = Val(ShiftsComboboxes(intK).Text)
            Next
            For intK = 0 To L
                If Checkboxes(intK).Checked Then
                    DGSetsChecked(intK) = 1
                    DGSetsMobdate(intK) = MobdatePickers(intK).Value.Date
                    DGSetsDemobdate(intK) = DemobDatePickers(intK).Value.Date
                    DGSetsQty(intK) = Val(QtyTextBoxes(intK).Text)
                    DGSetsHrs(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                    DGSetsDepPerc(intK) = Val(DepPercComboboxes(intK).Text)
                    DGSetsShifts(intK) = Val(ShiftsComboboxes(intK).Text)
                Else
                    DGSetsChecked(intK) = 0
                End If
            Next
            DGSetItems = CategoryTextBoxes.Count
            WriteToDgsetsEquipsArray(DGSetItems)
            Exit Sub
        End If
        Button1.Enabled = True
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        mcategory = "DG Sets"
        SelectedCategory = mcategory
        If Not dgsetFirstTime Then
            LoadControlsFromDGSetsArray(DGSetItems)
            For intL = 0 To DGSetItems - 1
                If DGSetsChecked(intL) = 1 Then
                    Checkboxes(intL).Checked = True
                    MobdatePickers(intL).Value = DGSetsMobdate(intL)
                    DemobDatePickers(intL).Value = DGSetsDemobdate(intL)
                    QtyTextBoxes(intL).Text = DGSetsQty(intL)
                    HrsPermonthTextBoxes(intL).Text = DGSetsHrs(intL)
                    DepPercComboboxes(intL).Text = DGSetsDepPerc(intL)
                    ShiftsComboboxes(intL).Text = DGSetsShifts(intL)
                Else
                    Checkboxes(intL).Checked = False
                End If
            Next
            Me.Refresh()
        Else
            mcategory = "DG sets"
            SelectedCategory = mcategory
            LoadControlsInpage(mcategory)
            For intL = 0 To Checkboxes.Count - 1
                If Checkboxes(intL).Checked Then
                    changeEnabled(True, intL)
                End If
            Next
            dgsetFirstTime = False
        End If
        DGSetItems = CategoryTextBoxes.Count
        'WriteToDgsetsEquipsArray(DGSetItems)
    End Sub

    Private Sub optMajMH_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMajMH.CheckedChanged
        Dim intL As Integer
        If optMinorEquips.Checked = True Then
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optHireEquipments.Checked Then
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optFixedExp.Checked Or optBPFixedExp.Checked Then
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        ElseIf Me.optLightingEquips.Checked Then
            Panel7.Left = 1
            Panel7.Top = 20
            Panel7.Height = 30
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        Else
            Panel2.Left = 1
            Panel2.Top = 20
            Panel2.Height = 30
            Panel2.Visible = True
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If

        If Not optMajMH.Checked Then
            Dim intK As Integer, L As Integer
            MHItems = CategoryTextBoxes.Count
            L = MHItems - 1   'CategoryTextBoxes.Count - 1
            ReDim Preserve MHChecked(L)
            ReDim Preserve MHMobdate(L)
            ReDim Preserve MHDemobdate(L)
            ReDim Preserve MHQty(L)
            ReDim Preserve MHHrs(L)
            ReDim Preserve MHDepPerc(L)
            ReDim Preserve MHShifts(L)
            ReDim mathCategory(L)
            ReDim mathEName(L)
            ReDim mathCapacity(L)
            ReDim mathMakeModel(L)
            ReDim mathMDate(L)
            ReDim mathDMdate(L)
            ReDim mathQuantity(L)
            ReDim mathHrsPerMonth(L)
            ReDim mathDep(L)
            ReDim mathShift(L)

            For intK = 0 To L
                mathCategory(intK) = CategoryTextBoxes(intK).Text
                mathEName(intK) = EquipNameTextBoxes(intK).Text
                mathCapacity(intK) = CapacityTextBoxes(intK).Text
                mathMakeModel(intK) = MakeModelTextBoxes(intK).Text
                mathMDate(intK) = MobdatePickers(intK).Value.Date
                mathDMdate(intK) = DemobDatePickers(intK).Value.Date
                mathQuantity(intK) = Val(QtyTextBoxes(intK).Text)
                mathHrsPerMonth(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                mathDep(intK) = Val(DepPercComboboxes(intK).Text)
                mathShift(intK) = Val(ShiftsComboboxes(intK).Text)
            Next
            For intK = 0 To L
                If Checkboxes(intK).Checked Then
                    MHChecked(intK) = 1
                    MHMobdate(intK) = MobdatePickers(intK).Value.Date
                    MHDemobdate(intK) = DemobDatePickers(intK).Value.Date
                    MHQty(intK) = Val(QtyTextBoxes(intK).Text)
                    MHHrs(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                    MHDepPerc(intK) = Val(DepPercComboboxes(intK).Text)
                    MHShifts(intK) = Val(ShiftsComboboxes(intK).Text)
                Else
                    MHChecked(intK) = 0
                End If
            Next
            MHItems = CategoryTextBoxes.Count
            WriteToMHEquipsArray(MHItems)
            Exit Sub
        End If
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        mcategory = "Material Handling"
        SelectedCategory = mcategory
        If Not mathFirstTime Then
            LoadControlsFromMHArray(MHItems)
            For intL = 0 To MHItems - 1
                If MHChecked(intL) = 1 Then
                    Checkboxes(intL).Checked = True
                    MobdatePickers(intL).Value = MHMobdate(intL)
                    DemobDatePickers(intL).Value = MHDemobdate(intL)
                    QtyTextBoxes(intL).Text = MHQty(intL)
                    HrsPermonthTextBoxes(intL).Text = MHHrs(intL)
                    DepPercComboboxes(intL).Text = MHDepPerc(intL)
                    ShiftsComboboxes(intL).Text = MHShifts(intL)
                Else
                    Checkboxes(intL).Checked = False
                End If
            Next
            Me.Refresh()
        Else
            mcategory = "Material Handling"
            SelectedCategory = mcategory
            LoadControlsInpage(mcategory)
            For intL = 0 To Checkboxes.Count - 1
                If Checkboxes(intL).Checked Then
                    changeEnabled(True, intL)
                End If
            Next
            mathFirstTime = False
        End If
        MHItems = CategoryTextBoxes.Count
        'WriteToMHEquipsArray(MHItems)
    End Sub

    Private Sub optMajNc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMajNc.CheckedChanged
        Dim intL As Integer
        If optMinorEquips.Checked = True Then
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optHireEquipments.Checked Then
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optFixedExp.Checked Or optBPFixedExp.Checked Then
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        ElseIf Me.optLightingEquips.Checked Then
            Panel7.Left = 1
            Panel7.Top = 20
            Panel7.Height = 30
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        Else
            Panel2.Left = 1
            Panel2.Top = 20
            Panel2.Height = 30
            Panel2.Visible = True
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If

        If Not optMajNc.Checked Then
            Dim intK As Integer, L As Integer
            NCItems = CategoryTextBoxes.Count
            L = NCItems - 1  'CategoryTextBoxes.Count - 1
            ReDim Preserve nccHECKED(L)
            ReDim Preserve NCMobdate(L)
            ReDim Preserve NCDemobdate(L)
            ReDim Preserve NCQty(L)
            ReDim Preserve NCHrs(L)
            ReDim Preserve NCDepPerc(L)
            ReDim Preserve NCShifts(L)
            ReDim noncCategory(L)
            ReDim noncEName(L)
            ReDim noncCapacity(L)
            ReDim noncMakeModel(L)
            ReDim noncMDate(L)
            ReDim noncDMdate(L)
            ReDim noncQuantity(L)
            ReDim noncHrsPerMonth(L)
            ReDim noncDep(L)
            ReDim noncShift(L)

            For intK = 0 To L
                noncCategory(intK) = CategoryTextBoxes(intK).Text
                noncEName(intK) = EquipNameTextBoxes(intK).Text
                noncCapacity(intK) = CapacityTextBoxes(intK).Text
                noncMakeModel(intK) = MakeModelTextBoxes(intK).Text
                noncMDate(intK) = MobdatePickers(intK).Value.Date
                noncDMdate(intK) = DemobDatePickers(intK).Value.Date
                noncQuantity(intK) = Val(QtyTextBoxes(intK).Text)
                noncHrsPerMonth(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                noncDep(intK) = Val(DepPercComboboxes(intK).Text)
                noncShift(intK) = Val(ShiftsComboboxes(intK).Text)
            Next
            For intK = 0 To L
                If Checkboxes(intK).Checked Then
                    nccHECKED(intK) = 1
                    NCMobdate(intK) = MobdatePickers(intK).Value.Date
                    NCDemobdate(intK) = DemobDatePickers(intK).Value.Date
                    NCQty(intK) = Val(QtyTextBoxes(intK).Text)
                    NCHrs(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                    NCDepPerc(intK) = Val(DepPercComboboxes(intK).Text)
                    NCShifts(intK) = Val(ShiftsComboboxes(intK).Text)
                Else
                    nccHECKED(intK) = 0
                End If
            Next
            NCItems = CategoryTextBoxes.Count
            WriteToNCEquipsArray(NCItems)
            Exit Sub
        End If
        Button1.Enabled = True
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        mcategory = "Non Concreting"
        SelectedCategory = mcategory
        If Not noncFirstTime Then
            LoadControlsFromNCArray(NCItems)
            For intL = 0 To NCItems - 1
                If nccHECKED(intL) = 1 Then
                    Checkboxes(intL).Checked = True
                    MobdatePickers(intL).Value = NCMobdate(intL)
                    DemobDatePickers(intL).Value = NCDemobdate(intL)
                    QtyTextBoxes(intL).Text = NCQty(intL)
                    HrsPermonthTextBoxes(intL).Text = NCHrs(intL)
                    DepPercComboboxes(intL).Text = NCDepPerc(intL)
                    ShiftsComboboxes(intL).Text = NCShifts(intL)
                Else
                    Checkboxes(intL).Checked = False
                End If
            Next
            Me.Refresh()
        Else
            mcategory = "Non Concreting"
            SelectedCategory = mcategory
            LoadControlsInpage(mcategory)
            For intL = 0 To Checkboxes.Count - 1
                If Checkboxes(intL).Checked Then
                    changeEnabled(True, intL)
                End If
            Next
            noncFirstTime = False
        End If
        NCItems = CategoryTextBoxes.Count
    End Sub

    Private Sub optMajOthers_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMajOthers.CheckedChanged
        Dim intL As Integer
        If optMinorEquips.Checked = True Then
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optHireEquipments.Checked Then
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optFixedExp.Checked Or optBPFixedExp.Checked Then
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        ElseIf Me.optLightingEquips.Checked Then
            Panel7.Left = 1
            Panel7.Top = 20
            Panel7.Height = 30
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        Else
            Panel2.Left = 1
            Panel2.Top = 20
            Panel2.Height = 30
            Panel2.Visible = True
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If

        If Not optMajOthers.Checked Then
            Dim intK As Integer, L As Integer
            MajorOtherItems = CategoryTextBoxes.Count
            L = MajorOtherItems - 1  'CategoryTextBoxes.Count - 1
            ReDim Preserve MajOthersChecked(L)
            ReDim Preserve MajOthersMobdate(L)
            ReDim Preserve MajOthersDemobdate(L)
            ReDim Preserve MajOthersQty(L)
            ReDim Preserve MajOthersHrs(L)
            ReDim Preserve MajOthersDepPerc(L)
            ReDim Preserve MajOthersShifts(L)
            ReDim majoCategory(L)
            ReDim majoEName(L)
            ReDim majoCapacity(L)
            ReDim majoMakeModel(L)
            ReDim majoMDate(L)
            ReDim majoDMdate(L)
            ReDim majoQuantity(L)
            ReDim majoHrsPerMonth(L)
            ReDim majoDep(L)
            ReDim majoShift(L)

            For intK = 0 To L
                majoCategory(intK) = CategoryTextBoxes(intK).Text
                majoEName(intK) = EquipNameTextBoxes(intK).Text
                majoCapacity(intK) = CapacityTextBoxes(intK).Text
                majoMakeModel(intK) = MakeModelTextBoxes(intK).Text
                majoMDate(intK) = MobdatePickers(intK).Value.Date
                majoDMdate(intK) = DemobDatePickers(intK).Value.Date
                majoQuantity(intK) = Val(QtyTextBoxes(intK).Text)
                majoHrsPerMonth(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                majoDep(intK) = Val(DepPercComboboxes(intK).Text)
                majoShift(intK) = Val(ShiftsComboboxes(intK).Text)
            Next
            For intK = 0 To L
                If Checkboxes(intK).Checked Then
                    MajOthersChecked(intK) = 1
                    MajOthersMobdate(intK) = MobdatePickers(intK).Value.Date
                    MajOthersDemobdate(intK) = DemobDatePickers(intK).Value.Date
                    MajOthersQty(intK) = Val(QtyTextBoxes(intK).Text)
                    MajOthersHrs(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                    MajOthersDepPerc(intK) = Val(DepPercComboboxes(intK).Text)
                    MajOthersShifts(intK) = Val(ShiftsComboboxes(intK).Text)
                Else
                    MajOthersChecked(intK) = 0
                End If
            Next
            MajorOtherItems = CategoryTextBoxes.Count
            WriteToMajOtherEquipsArray(MajorOtherItems)
            Exit Sub
        End If
        Button1.Enabled = True
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        mcategory = "Major Others"
        SelectedCategory = mcategory
        If Not majoFirstTime Then
            LoadControlsFromMajOtherArray(MajorOtherItems)
            For intL = 0 To MajorOtherItems - 1
                If MajOthersChecked(intL) = 1 Then
                    Checkboxes(intL).Checked = True
                    MobdatePickers(intL).Value = MHMobdate(intL)
                    DemobDatePickers(intL).Value = MHDemobdate(intL)
                    QtyTextBoxes(intL).Text = MHQty(intL)
                    HrsPermonthTextBoxes(intL).Text = MHHrs(intL)
                    DepPercComboboxes(intL).Text = MHDepPerc(intL)
                    ShiftsComboboxes(intL).Text = MHShifts(intL)
                Else
                    Checkboxes(intL).Checked = False
                End If
            Next
            Me.Refresh()
        Else
            mcategory = "Major Others"
            SelectedCategory = mcategory
            LoadControlsInpage(mcategory)
            For intL = 0 To Checkboxes.Count - 1
                If Checkboxes(intL).Checked Then
                    changeEnabled(True, intL)
                End If
            Next
            majoFirstTime = False
        End If
        MajorOtherItems = CategoryTextBoxes.Count
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        SelectedCategory = mcategory
        'Me.Hide()
        oForm = New AddedItems_new()
        oForm.ShowDialog()
        oForm = Nothing
    End Sub

    Private Sub optMinorEquips_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMinorEquips.CheckedChanged
        Dim intL As Integer
        If optMinorEquips.Checked = True Then
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optHireEquipments.Checked Then
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optFixedExp.Checked Or optBPFixedExp.Checked Then
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        ElseIf Me.optLightingEquips.Checked Then
            Panel7.Left = 1
            Panel7.Top = 20
            Panel7.Height = 30
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        Else
            Panel2.Left = 1
            Panel2.Top = 20
            Panel2.Height = 30
            Panel2.Visible = True
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If
        'ConcretingItems = CategoryTextBoxes.Count
        If Not optMinorEquips.Checked Then
            SelectedCategory = ""
            Dim intK As Integer, L As Integer
            MinorItems = CategoryTextBoxes.Count
            L = MinorItems - 1   'CategoryTextBoxes.Count
            ReDim Preserve MinorChecked(L)
            ReDim Preserve MinorMobdate(L)
            ReDim Preserve MinorDemobdate(L)
            ReDim Preserve MinorQty(L)
            ReDim Preserve MinorHrs(L)
            ReDim Preserve MinorNPV(L)
            ReDim Preserve MinorDepPerc(L)
            ReDim Preserve MinorShifts(L)
            ReDim MinorCategory(L)
            ReDim MinorEName(L)
            ReDim MinorCapacity(L)
            ReDim MinorMakeModel(L)
            ReDim MinorMDate(L)
            ReDim MinorDMdate(L)
            ReDim MinorQuantity(L)
            ReDim MinorHrsPerMonth(L)
            ReDim MinorNewPurchVal(L)
            ReDim MinorIsNew(L)
            ReDim MinorDep(L)
            ReDim MinorShift(L)
            ReDim MinorChecked(L)
            For intK = 0 To L
                MinorCategory(intK) = CategoryTextBoxes(intK).Text
                MinorEName(intK) = EquipNameTextBoxes(intK).Text
                MinorCapacity(intK) = CapacityTextBoxes(intK).Text
                MinorMakeModel(intK) = MakeModelTextBoxes(intK).Text
                MinorMDate(intK) = MobdatePickers(intK).Value.Date
                MinorDMdate(intK) = DemobDatePickers(intK).Value.Date
                MinorQuantity(intK) = Val(QtyTextBoxes(intK).Text)
                MinorHrsPerMonth(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                MinorNewPurchVal(intK) = Val(PurchvalTextBoxes(intK).Text)
                MinorDep(intK) = Val(DepPercComboboxes(intK).Text)
                MinorShift(intK) = Val(ShiftsComboboxes(intK).Text)
                MinorIsNew(intK) = IsNewMC(intK).Checked
            Next
            For intK = 0 To L
                If Checkboxes(intK).Checked Then
                    MinorChecked(intK) = 1
                    MinorMobdate(intK) = MobdatePickers(intK).Value.Date
                    MinorDemobdate(intK) = DemobDatePickers(intK).Value.Date
                    MinorQty(intK) = Val(QtyTextBoxes(intK).Text)
                    MinorHrs(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                    MinorNPV(intK) = Val(PurchvalTextBoxes(intK).Text)
                    MinorDepPerc(intK) = Val(DepPercComboboxes(intK).Text)
                    MinorShifts(intK) = Val(ShiftsComboboxes(intK).Text)
                    MinorIsNew(intK) = IsNewMC(intK).Checked
                Else
                    MinorChecked(intK) = 0
                End If
            Next
            MinorItems = CategoryTextBoxes.Count    ' - 1
            WriteToMinorEquipsArray(MinorItems)
            Exit Sub
        End If
        Button1.Enabled = True
        mcategory = "Minor Equipments"
        SelectedCategory = mcategory
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        'mcategory = "Minor Equipments"
        If Not MinorFirstTime Then
            LoadControlsFromMinorArray(MinorItems)
            For intL = 0 To MinorItems - 1
                If MinorChecked(intL) = 1 Then
                    Checkboxes(intL).Checked = True
                    MobdatePickers(intL).Value = MinorMobdate(intL)
                    DemobDatePickers(intL).Value = MinorDemobdate(intL)
                    QtyTextBoxes(intL).Text = MinorQty(intL)
                    HrsPermonthTextBoxes(intL).Text = MinorHrs(intL)
                    PurchvalTextBoxes(intL).Text = MinorNPV(intL)
                    DepPercComboboxes(intL).Text = MinorDepPerc(intL)
                    ShiftsComboboxes(intL).Text = MinorShifts(intL)
                    IsNewMC(intL).Checked = MinorIsNew(intL)
                Else
                    Checkboxes(intL).Checked = False
                End If
                'Checkboxes(intL + 1).Checked = IIf(ConcreteChecked(intL) = 1, True, False)
            Next
            Me.Refresh()
        Else
            mcategory = "Minor Equipments"
            SelectedCategory = mcategory
            LoadMinorControlsInpage()
            For intL = 0 To Checkboxes.Count - 1
                If Checkboxes(intL).Checked Then
                    MinorChangeEnabled(True, intL)
                End If
            Next
            MinorFirstTime = False
        End If
        MinorItems = CategoryTextBoxes.Count
    End Sub

    Private Sub optHireEquipments_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optHireEquipments.CheckedChanged
        Dim intL As Integer
        If optMinorEquips.Checked = True Then
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optHireEquipments.Checked Then
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optFixedExp.Checked Or optBPFixedExp.Checked Then
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        ElseIf Me.optLightingEquips.Checked Then
            Panel7.Left = 1
            Panel7.Top = 20
            Panel7.Height = 30
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        Else
            Panel2.Left = 1
            Panel2.Top = 20
            Panel2.Height = 30
            Panel2.Visible = True
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If
        'ConcretingItems = CategoryTextBoxes.Count
        If Not optHireEquipments.Checked Then
            SelectedCategory = ""
            Dim intK As Integer, L As Integer
            HireItems = CategoryTextBoxes.Count
            L = HireItems - 1   'CategoryTextBoxes.Count
            ReDim Preserve HiredChecked(L)
            ReDim Preserve HiredMobdate(L)
            ReDim Preserve HiredDemobdate(L)
            ReDim Preserve HiredQty(L)
            ReDim Preserve HiredHrs(L)
            ReDim Preserve HiredNPV(L)
            ReDim Preserve HiredDepPerc(L)
            ReDim Preserve HiredShifts(L)
            ReDim HiredCategory(L)
            ReDim HiredEName(L)
            ReDim HiredCapacity(L)
            ReDim HiredMakeModel(L)
            ReDim HiredMDate(L)
            ReDim HiredDMdate(L)
            ReDim HiredQuantity(L)
            ReDim HiredHrsPerMonth(L)
            ReDim HiredHireCharges(L)
            ReDim HiredDep(L)
            ReDim HiredShift(L)
            For intK = 0 To L
                HiredCategory(intK) = CategoryTextBoxes(intK).Text
                HiredEName(intK) = EquipNameTextBoxes(intK).Text
                HiredCapacity(intK) = CapacityTextBoxes(intK).Text
                HiredMakeModel(intK) = MakeModelTextBoxes(intK).Text
                HiredMDate(intK) = MobdatePickers(intK).Value.Date
                HiredDMdate(intK) = DemobDatePickers(intK).Value.Date
                HiredQuantity(intK) = Val(QtyTextBoxes(intK).Text)
                HiredHrsPerMonth(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                HiredHireCharges(intK) = Val(HireChargesTextBoxes(intK).Text)
                HiredDep(intK) = Val(DepPercComboboxes(intK).Text)
                HiredShift(intK) = Val(ShiftsComboboxes(intK).Text)
            Next
            For intK = 0 To L
                If Checkboxes(intK).Checked Then
                    HiredChecked(intK) = 1
                    HiredMobdate(intK) = MobdatePickers(intK).Value.Date
                    HiredDemobdate(intK) = DemobDatePickers(intK).Value.Date
                    HiredQty(intK) = Val(QtyTextBoxes(intK).Text)
                    HiredHrs(intK) = Val(HrsPermonthTextBoxes(intK).Text)
                    HiredNPV(intK) = Val(HireChargesTextBoxes(intK).Text)
                    HiredDepPerc(intK) = Val(DepPercComboboxes(intK).Text)
                    HiredShifts(intK) = Val(ShiftsComboboxes(intK).Text)
                Else
                    HiredChecked(intK) = 0
                End If
            Next
            HireItems = CategoryTextBoxes.Count
            WriteToHiredEquipsArray(HireItems)
            Exit Sub
        End If
        Button1.Enabled = True
        mcategory = "HiredEquipments"
        SelectedCategory = mcategory
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        mcategory = "HiredEquipments"
        If Not HiredFirstTime Then
            LoadControlsFromHiredArray(HireItems)
            For intL = 0 To HireItems - 1
                If HiredChecked(intL) = 1 Then
                    Checkboxes(intL).Checked = True
                    MobdatePickers(intL).Value = HiredMobdate(intL)
                    DemobDatePickers(intL).Value = HiredDemobdate(intL)
                    QtyTextBoxes(intL).Text = HiredQty(intL)
                    HrsPermonthTextBoxes(intL).Text = HiredHrs(intL)
                    'PurchvalTextBoxes(intL ).Text = HiredNPV(intL)
                    DepPercComboboxes(intL).Text = HiredDepPerc(intL)
                    ShiftsComboboxes(intL).Text = HiredShifts(intL)
                Else
                    Checkboxes(intL).Checked = False
                End If
                'Checkboxes(intL + 1).Checked = IIf(ConcreteChecked(intL) = 1, True, False)
            Next
            Me.Refresh()
        Else
            mcategory = "Hired Equipments"
            SelectedCategory = mcategory
            LoadHiredControlsInpage()
            For intL = 0 To Checkboxes.Count - 1
                If Checkboxes(intL).Checked Then
                    HireChangeEnabled(True, intL)
                End If
            Next
            HiredFirstTime = False
        End If
        HireItems = CategoryTextBoxes.Count
    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click, Label23.Click

    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click

    End Sub

    Private Sub optFixedExp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optFixedExp.CheckedChanged
        Dim intL As Integer
        If optMinorEquips.Checked = True Then
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optHireEquipments.Checked Then
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optFixedExp.Checked Or optBPFixedExp.Checked Then
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        ElseIf Me.optLightingEquips.Checked Then
            Panel7.Left = 1
            Panel7.Top = 20
            Panel7.Height = 30
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        Else
            Panel2.Left = 1
            Panel2.Top = 20
            Panel2.Height = 30
            Panel2.Visible = True
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If
        'ConcretingItems = CategoryTextBoxes.Count

        If Not optFixedExp.Checked Then
            SelectedCategory = ""
            Dim intK As Integer, L As Integer
            fexpItems = CategoryTextBoxes.Count
            L = fexpItems - 1 'CategoryTextBoxes.Count
            ReDim fexpCategory(L)
            ReDim FExpQty(L)
            ReDim fexpCost(L)
            ReDim fexpAmount(L)
            ReDim fexpClientBilling(L)
            ReDim fexpCostPerc(L)
            ReDim FExpRemarks(L)
            ReDim FexpChecked(L)

            For intK = 0 To L
                fexpCategory(intK) = CategoryTextBoxes(intK).Text
                'fexpQty(intK ) = QtyTextBoxes(intK).Text
                'fexpCost(intK ) = CostTextBoxes(intK).Text
                fexpAmount(intK) = AmountTextBoxes(intK).Text
                fexpClientBilling(intK) = ClientBillingTextBoxes(intK).Text
                fexpCostPerc(intK) = CostPercTextBoxes(intK).Text
                'fexpRemarks(intK ) = RemarksTextBoxes(intK).Text
            Next
            For intK = 0 To L
                If Checkboxes(intK).Checked Then
                    FexpChecked(intK) = 1
                    FExpQty(intK) = Val(QtyTextBoxes(intK).Text)
                    fexpCost(intK) = Val(CostTextBoxes(intK).Text)
                    FExpRemarks(intK) = RemarksTextBoxes(intK).Text
                Else
                    FexpChecked(intK) = 0
                End If
            Next
            fexpItems = CategoryTextBoxes.Count
            WriteToFixedExpArray(fexpItems)
            Exit Sub
        End If
        Button1.Enabled = True
        mcategory = "FixedExp"
        SelectedCategory = mcategory
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        mcategory = "Tr crane related exp"
        SelectedCategory = mcategory
        If Not fexpFirstTime Then
            LoadControlsFromFexpArray(fexpItems)
            For intL = 0 To fexpItems - 1
                If FexpChecked(intL) = 1 Then
                    Checkboxes(intL).Checked = True
                    QtyTextBoxes(intL).Text = FExpQty(intL)
                    CostTextBoxes(intL).Text = fexpCost(intL)
                    RemarksTextBoxes(intL).Text = FExpRemarks(intL)
                Else
                    Checkboxes(intL).Checked = False
                End If
            Next
            Me.Refresh()
        Else
            mcategory = "Tr crane related exp"
            SelectedCategory = mcategory
            LoadFexpControlsInpage(mcategory)
            For intL = 0 To Checkboxes.Count - 1
                If Checkboxes(intL).Checked Then
                    FexpChangeEnabled(True, intL)
                End If
            Next
            fexpFirstTime = False
        End If
        fexpItems = CategoryTextBoxes.Count
    End Sub

    Private Sub optBPFixedExp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optBPFixedExp.CheckedChanged
        Dim intL As Integer
        If optMinorEquips.Checked = True Then
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optHireEquipments.Checked Then
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optFixedExp.Checked Or optBPFixedExp.Checked Then
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        ElseIf Me.optLightingEquips.Checked Then
            Panel7.Left = 1
            Panel7.Top = 20
            Panel7.Height = 30
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        Else
            Panel2.Left = 1
            Panel2.Top = 20
            Panel2.Height = 30
            Panel2.Visible = True
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        End If
        If Not optBPFixedExp.Checked Then
            SelectedCategory = ""
            Dim intK As Integer, L As Integer
            BPFExpItems = CategoryTextBoxes.Count
            L = BPFExpItems - 1 'CategoryTextBoxes.Count
            ReDim BPfexpCategory(L)
            ReDim BPFExpQty(L)
            ReDim BPfexpCost(L)
            ReDim BPfexpAmount(L)
            ReDim BPfexpClientBilling(L)
            ReDim BPfexpCostPerc(L)
            ReDim BPFExpRemarks(L)
            ReDim BPFExpChecked(L)

            For intK = 0 To L
                BPfexpCategory(intK) = CategoryTextBoxes(intK).Text
                'BPfexpQty(intK ) = QtyTextBoxes(intK).Text
                'BPfexpCost(intK ) = CostTextBoxes(intK).Text
                BPfexpAmount(intK) = AmountTextBoxes(intK).Text
                BPfexpClientBilling(intK) = ClientBillingTextBoxes(intK).Text
                BPfexpCostPerc(intK) = CostPercTextBoxes(intK).Text
                'BPfexpRemarks(intK ) = RemarksTextBoxes(intK).Text
            Next
            For intK = 0 To L
                If Checkboxes(intK).Checked Then
                    BPFExpChecked(intK) = 1
                    BPFExpQty(intK) = Val(QtyTextBoxes(intK).Text)
                    BPfexpCost(intK) = Val(CostTextBoxes(intK).Text)
                    BPFExpRemarks(intK) = RemarksTextBoxes(intK).Text
                Else
                    BPFExpChecked(intK) = 0
                End If
            Next
            BPFExpItems = CategoryTextBoxes.Count
            WriteToFixedBPExpArray(BPFExpItems)
            Exit Sub
        End If
        Button1.Enabled = True
        mcategory = "FixedExp - BP"
        SelectedCategory = mcategory
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        mcategory = "BPlant related exp"
        SelectedCategory = mcategory
        If Not BPFExpFirstTime Then
            LoadControlsFromBPFexpArray(BPFExpItems)
            For intL = 0 To BPFExpItems - 1
                If BPFExpChecked(intL) = 1 Then
                    Checkboxes(intL).Checked = True
                    QtyTextBoxes(intL).Text = BPFExpQty(intL)
                    CostTextBoxes(intL).Text = BPfexpCost(intL)
                    RemarksTextBoxes(intL).Text = BPFExpRemarks(intL)
                Else
                    Checkboxes(intL).Checked = False
                End If
                'Checkboxes(intL + 1).Checked = IIf(ConcreteChecked(intL) = 1, True, False)
            Next
            Me.Refresh()
        Else
            mcategory = "BPlant related exp"
            SelectedCategory = mcategory
            LoadBPFExpControlsInpage(mcategory)
            For intL = 0 To Checkboxes.Count - 1
                If Checkboxes(intL).Checked Then
                    FexpChangeEnabled(True, intL)
                End If
            Next
            BPFExpFirstTime = False
        End If
        BPFExpItems = CategoryTextBoxes.Count
    End Sub

    Private Sub TotalElecExpensesPerc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TotalElecExpensesPerc.TextChanged

    End Sub

    Private Sub TotalElecExpensesPerc_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TotalElecExpensesPerc.Validating
        If Val(Me.TotalElecExpensesPerc.Text) < 0 Or Val(Me.TotalElecExpensesPerc.Text) > 1.5 Then
            MsgBox("total Electrical Expenses Percentage should be within the range of 0 and 1.5")
            e.Cancel = True
        End If
    End Sub

    Private Sub SaveElectricals_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveElectricals.Click
        Dim mTotalExp As Double, mCablesExp As Double, mPanelExp As Double, mLightsExp As Double, mEBDeposit As Double
        Dim mMiscExp As Double, msgstring As String = ""
        Dim currentsheetname2 As String = "Electrical"
        Dim currentsheetname1 As String, intI As Integer
        Dim Allvalid As Boolean = True
        Me.btnQuit.Enabled = True
        Me.btnClose.Enabled = False
        Me.lblError.Text = ""
        If Val(Me.TotalElecExpensesPerc.Text) >= 100 Then
            msgstring = msgstring & "Total Electrical expenses percentage should be less than 100" & vbNewLine
            Me.TotalElecExpensesPerc.Focus()
            Allvalid = False
        ElseIf Val(Me.CablesPerc.Text) >= 100 Then
            msgstring = msgstring & "Cables expenses percentage should be less than 100" & vbNewLine
            Me.CablesPerc.Focus()
            Allvalid = False
        ElseIf Val(Me.Panelsperc.Text) >= 100 Then
            msgstring = msgstring & "Total Electrical expenses percentage should be less than 100" & vbNewLine
            Me.Panelsperc.Focus()
            Allvalid = False
        ElseIf Val(Me.Depositperc.Text) >= 100 Then
            msgstring = msgstring & "EB Deposit and Other expenses percentage should be less than 100" & vbNewLine
            Me.Depositperc.Focus()
            Allvalid = False
        ElseIf Val(Me.Miscperc.Text) >= 100 Then
            msgstring = msgstring & "Miscellaneous expenses percentage should be less than 100" & vbNewLine
            Me.Miscperc.Focus()
            Allvalid = False
        ElseIf Val(Me.Lightsperc.Text) >= 100 Then
            msgstring = msgstring & "Lights and Accessories expenses percentage should be less than 100" & vbNewLine
            Me.Lightsperc.Focus()
            Allvalid = False
        End If
        If Not Allvalid Then
            Me.ErrorLabel.Text = msgstring
            Exit Sub
        End If
        mTotalExp = Val(Me.TotalElecExpensesPerc.Text)
        mCablesExp = Val(Me.CablesPerc.Text)
        mPanelExp = Val(Me.Panelsperc.Text)
        mLightsExp = Val(Me.Lightsperc.Text)
        mEBDeposit = Val(Me.Depositperc.Text)
        mMiscExp = Val(Me.Miscperc.Text)
        ErrorLabel.Visible = False
        If System.Math.Round((mCablesExp + mPanelExp + mLightsExp + mEBDeposit + mMiscExp), 2) <> System.Math.Round(mTotalExp, 2) Then
            Me.ErrorLabel.Text = "Percentage Break-up is not equal to the Total"
            Me.ErrorLabel.Visible = True
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
                'PrevSheetname = xlWorksheet.Name
                xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
                SheetNo = intI + 1
                Exit For
            End If
        Next
        xlWorksheet = xlWorkbook.Sheets.Item("Electrical")
        xlWorksheet.Activate()
        Category_Shortname = "Elec_"
        'xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle1")
        'xlRange.Value = mMainTitle1
        xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle2")
        If Len(Trim(xlRange.Value)) = 0 Then
            xlRange.Value = mMainTitle2
            'xlRange = xlWorksheet.Range(Category_Shortname & "MainTitle3")
            'xlRange.Value = mMainTitle3
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
        End If

        xlRange = xlWorksheet.Range("Elec_CablesExp")
        xlRange.Value = mProjectvalue * (Val(Me.CablesPerc.Text) / 100)
        xlRange = xlWorksheet.Range("Elec_PanelsExp")
        xlRange.Value = mProjectvalue * (Val(Me.Panelsperc.Text) / 100)
        xlRange = xlWorksheet.Range("Elec_LightsExp")
        xlRange.Value = mProjectvalue * (Val(Me.Lightsperc.Text) / 100)
        xlRange = xlWorksheet.Range("Elec_EBDeposit")
        xlRange.Value = mProjectvalue * (Val(Me.Depositperc.Text) / 100)
        xlRange = xlWorksheet.Range("Elec_MiscExp")
        xlRange.Value = mProjectvalue * (Val(Me.Miscperc.Text) / 100)
        xlRange = xlWorksheet.Range("Elec_reusablePanels")
        xlRange.Value = Me.ResusablePanels.Text
        xlRange = xlWorksheet.Range("Elec_ReusableLights")
        xlRange.Value = Me.ReusableLights.Text
        ErrorLabel.Text = "Electrical Expense Budget details added in " & xlWorksheet.Name
        ErrorLabel.Visible = True
        btnQuit.Enabled = True
    End Sub

    Private Sub Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Clear.Click
        Me.TotalElecExpensesPerc.Text = ""
        Me.CablesPerc.Text = ""
        Me.Panelsperc.Text = ""
        Me.Lightsperc.Text = ""
        Me.Depositperc.Text = ""
        Me.Miscperc.Text = ""
        Me.ErrorLabel.Text = ""
        Me.SaveElectricals.Enabled = True
    End Sub

    Private Sub NextTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextTab.Click
        Me.tbcBdgetHeads.SelectTab(2)
    End Sub

    Private Sub PrevTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrevTab.Click
        Me.tbcBdgetHeads.SelectTab(0)
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QtyBox1.TextChanged
        Me.AmountBox1.Text = Val(Me.QtyBox1.Text) * Val(Me.CostBox1.Text)
    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CostBox1.TextChanged
        Me.AmountBox1.Text = Val(Me.QtyBox1.Text) * Val(Me.CostBox1.Text)
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QtyBox2.TextChanged
        Me.AmonuntBox2.Text = Val(Me.QtyBox2.Text) * Val(Me.CostBox2.Text)
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CostBox2.TextChanged
        Me.AmonuntBox2.Text = Val(Me.QtyBox2.Text) * Val(Me.CostBox2.Text)
    End Sub

    Private Sub SavePipelineCost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SavePipelineCost.Click
        Dim currentsheetname2 As String, currentsheetname1 As String, intI As Integer
        currentsheetname2 = "Extra pipeline"
        Me.btnQuit.Enabled = False
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
                'PrevSheetname = xlWorksheet.Name
                'xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
                SheetNo = intI + 1
                Exit For
            End If
        Next
        intI = 1

        xlRange = xlWorksheet.Range("ExPipe_SlNo").Offset(intI, 0)
        'xlRange = xlRange.Offset(0, 2)
        'xlRange.Value = Me.PipelineExpHead1.Text
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.QtyBox1.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.CostBox1.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.RemarksBox1.Text
        PipelineCostError.Visible = False
        intI = intI + 1
        xlRange = xlWorksheet.Range("ExPipe_SlNo").Offset(intI, 0)
        'xlRange = xlRange.Offset(0, 2)
        'xlRange.Value = Me.PipelineExpHead2.Text
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Me.QtyBox2.Text
        xlRange = xlRange.Offset(0, 1)
        xlRange.Value = Me.CostBox2.Text
        xlRange = xlRange.Offset(0, 4)
        xlRange.Value = Me.RemarksBox2.Text
        PipelineCostError.Text = "Pipeline related expenses saved. "
        PipelineCostError.Visible = True
        Me.PipelineExpSave.Enabled = True
        btnQuit.Enabled = True
    End Sub

    Private Sub ClearPipeLineCost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearPipeLineCost.Click
        Dim ctl As Control
        For Each ctl In Me.tbcBdgetHeads.TabPages(3).Controls
            If ctl.GetType Is GetType(TextBox) Then
                ctl.Text = ""
            End If
        Next
    End Sub

    Private Sub PipelinePrevTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PipelinePrevTab.Click
        Me.tbcBdgetHeads.SelectTab(1)
    End Sub

    Private Sub PipeLineNextTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PipeLineNextTab.Click
        Me.tbcBdgetHeads.SelectTab(3)
    End Sub

    Private Sub EsalSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EsalSave.Click
        Dim currentsheetname2 As String, currentsheetname1 As String, intI As Integer
        currentsheetname2 = "Elect salary"
        Me.btnQuit.Enabled = False
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
                'PrevSheetname = xlWorksheet.Name
                'xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
                SheetNo = intI + 1
                Exit For
            End If
        Next
        intI = 1
        xlWorksheet.Range("Esal_ElecHelperMonths").Value = 0
        xlWorksheet.Range("Esal_ElectricianSalaryPM").Value = 0
        xlWorksheet.Range("Esal_ElectricianMonths").Value = 0
        xlWorksheet.Range("Esal_ElecHelperNos").Value = 0
        xlWorksheet.Range("Esal_ElecHelperSalaryPM").Value = 0
        xlWorksheet.Range("Esal_ElecHelperMonths").Value = 0
        xlWorksheet.Range("Esal_MechanicNos").Value = 0
        xlWorksheet.Range("Esal_MechanicSalaryPM").Value = 0
        xlWorksheet.Range("Esal_MechanicMonths").Value = 0
        xlWorksheet.Range("Esal_MechHelperNos").Value = 0
        xlWorksheet.Range("Esal_MechHelperSalaryPM").Value = 0
        xlWorksheet.Range("Esal_MechHelperMonths").Value = 0
        xlWorksheet.Range("Esal_FitterNos").Value = 0
        xlWorksheet.Range("Esal_FitterSalaryPM").Value = 0
        xlWorksheet.Range("Esal_FitterMonths").Value = 0
        xlWorksheet.Range("Esal_AutoElecNos").Value = 0
        xlWorksheet.Range("Esal_AutoElecSalaryPM").Value = 0
        xlWorksheet.Range("Esal_AutoElecMonths").Value = 0
        xlWorksheet.Range("Esal_WelderNos").Value = 0
        xlWorksheet.Range("Esal_WelderSalaryPM").Value = 0
        xlWorksheet.Range("Esal_WelderMonths").Value = 0

        xlRange = xlWorksheet.Range("Esal_ElectricianNos")
        xlRange.Value = Me.txtElectricianNos.Text
        xlRange = xlWorksheet.Range("Esal_ElectricianSalaryPM")
        xlRange.Value = Me.txtElectSalaryPM.Text
        xlRange = xlWorksheet.Range("Esal_ElectricianMonths")
        xlRange.Value = Me.txtElectMonths.Text

        xlRange = xlWorksheet.Range("Esal_ElecHelperNos")
        xlRange.Value = Me.txtEletHepNos.Text
        xlRange = xlWorksheet.Range("Esal_ElecHelperSalaryPM")
        xlRange.Value = Me.txtElectHelperSalaryPM.Text
        xlRange = xlWorksheet.Range("Esal_ElecHelperMonths")
        xlRange.Value = Me.txtElectHelperMonths.Text

        xlRange = xlWorksheet.Range("Esal_MechanicNos")
        xlRange.Value = Me.txtMechanicNos.Text
        xlRange = xlWorksheet.Range("Esal_MechanicSalaryPM")
        xlRange.Value = Me.txtMechanicSalaryPM.Text
        xlRange = xlWorksheet.Range("Esal_MechanicMonths")
        xlRange.Value = Me.txtMechanicMonths.Text

        xlRange = xlWorksheet.Range("Esal_MechHelperNos")
        xlRange.Value = Me.txtMechHelperNos.Text
        xlRange = xlWorksheet.Range("Esal_MechHelperSalaryPM")
        xlRange.Value = Me.txtMechHelperSalaryPM.Text
        xlRange = xlWorksheet.Range("Esal_MechHelperMonths")
        xlRange.Value = Me.txtMechHelperMonths.Text

        xlRange = xlWorksheet.Range("Esal_FitterNos")
        xlRange.Value = Me.txtTyreFitterNos.Text
        xlRange = xlWorksheet.Range("Esal_FitterSalaryPM")
        xlRange.Value = Me.txtTyreFitterSalaryPM.Text
        xlRange = xlWorksheet.Range("Esal_FitterMonths")
        xlRange.Value = Me.txtTyreFitterMonths.Text

        xlRange = xlWorksheet.Range("Esal_AutoElecNos")
        xlRange.Value = Me.txtAutoElecNos.Text
        xlRange = xlWorksheet.Range("Esal_AutoElecSalaryPM")
        xlRange.Value = Me.txtAutoElecSalaryPM.Text
        xlRange = xlWorksheet.Range("Esal_AutoElecMonths")
        xlRange.Value = Me.txtAutoElecMonths.Text

        xlRange = xlWorksheet.Range("Esal_WelderNos")
        xlRange.Value = Me.txtWelderNos.Text
        xlRange = xlWorksheet.Range("Esal_WelderSalaryPM")
        xlRange.Value = Me.txtWelderSalaryPM.Text
        xlRange = xlWorksheet.Range("Esal_WelderMonths")
        xlRange.Value = Me.txtWelderMonths.Text

        EsalMessage.Text = "Electricians, Mechanics and others' cost details Saved. "
        EsalMessage.Visible = True
        'Me.btnEsalSave.Enabled = False
        btnQuit.Enabled = True
    End Sub

    Private Sub EsalPrevTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EsalPrevTab.Click
        Me.tbcBdgetHeads.SelectTab(2)
    End Sub

    Private Sub EsalNextTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EsalNextTab.Click
        Me.tbcBdgetHeads.SelectTab(4)
    End Sub

    Private Sub MiscPrevTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MiscPrevTab.Click
        Me.tbcBdgetHeads.SelectTab(3)
    End Sub

    Private Sub MiscNextTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MiscNextTab.Click
        Me.tbcBdgetHeads.SelectTab(5)
    End Sub

    Private Sub MiscSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MiscSave.Click
        Dim currentsheetname2 As String, currentsheetname1 As String, intI As Integer
        currentsheetname2 = "Misc and Non ERP Purchases"
        Me.btnQuit.Enabled = False
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
                'PrevSheetname = xlWorksheet.Name
                'xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
                SheetNo = intI + 1
                Exit For
            End If
        Next
        intI = 1

        xlRange = xlWorksheet.Range("Misc_Slno").Offset(intI, 0)
        'xlRange = xlRange.Offset(0, 2)
        'xlRange.Value = Me.Misc_ExpHead1.Text
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Val(Me.MiscAmt1.Text) * 100000
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = Me.MiscRemarks1.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("Misc_Slno").Offset(intI, 0)
        'xlRange = xlRange.Offset(0, 2)
        'xlRange.Value = Me.Misc_ExpHead2.Text
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Val(Me.MiscAmt2.Text) * 100000
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = Me.MiscRemarks2.Text


        intI = intI + 1
        xlRange = xlWorksheet.Range("Misc_Slno").Offset(intI, 0)
        'xlRange = xlRange.Offset(0, 2)
        'xlRange.Value = Me.Misc_ExpHead3.Text
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Val(Me.MiscAmt3.Text) * 100000
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = Me.MiscRemarks3.Text

        intI = intI + 1
        xlRange = xlWorksheet.Range("Misc_Slno").Offset(intI, 0)
        'xlRange = xlRange.Offset(0, 2)
        'xlRange.Value = Me.Misc_ExpHead4.Text
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Val(Me.MiscAmt4.Text) * 100000
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = Me.MiscRemarks4.Text


        intI = intI + 1
        xlRange = xlWorksheet.Range("Misc_Slno").Offset(intI, 0)
        'xlRange = xlRange.Offset(0, 2)
        'xlRange.Value = Me.Misc_ExpHead5.Text
        xlRange = xlRange.Offset(0, 2)
        xlRange.Value = Val(Me.MiscAmt5.Text) * 100000
        xlRange = xlRange.Offset(0, 3)
        xlRange.Value = Me.MiscRemarks5.Text

        MiscMessage.Text = "Misc Expenses details Saved. "
        MiscMessage.Visible = True
        'Me.MiscSave.Enabled = False
        btnQuit.Enabled = True
    End Sub

    Private Sub MiscClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MiscClear.Click
        MiscAmt1.Text = ""
        MiscRemarks1.Text = ""
        MiscAmt2.Text = ""
        MiscRemarks2.Text = ""
        MiscAmt3.Text = ""
        MiscRemarks3.Text = ""
        MiscAmt4.Text = ""
        MiscRemarks4.Text = ""
        MiscAmt5.Text = ""
        MiscRemarks5.Text = ""
    End Sub

    Private Sub StaffCostPrevTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StaffCostPrevTab.Click
        Me.tbcBdgetHeads.SelectTab(4)
    End Sub

    Private Sub ManagerNos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ManagerNos.Validated
        Me.ManagerCost.Text = Val(Int(Me.ManagerNos.Text)) * Val(Me.TManagerSalary.Text) * Val(Me.ManagerMonths.Text)
    End Sub

    Private Sub ManagerNos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ManagerNos.Validating
        If Not IsNumeric(ManagerNos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub
    Private Sub TManagerSalary_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TManagerSalary.Validated
        Me.ManagerCost.Text = Val(Int(Me.ManagerNos.Text)) * Val(Me.TManagerSalary.Text) * Val(Me.ManagerMonths.Text)
    End Sub

    Private Sub TManagerSalary_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TManagerSalary.Validating
        If Not IsNumeric(TManagerSalary.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub
    Private Sub ManagerMonths_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ManagerMonths.Validated
        Me.ManagerCost.Text = Val(Int(Me.ManagerNos.Text)) * Val(Me.TManagerSalary.Text) * Val(Me.ManagerMonths.Text)
    End Sub

    Private Sub ManagerMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ManagerMonths.Validating
        If Not IsNumeric(ManagerMonths.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub
    Private Sub EngrNos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles EngrNos.Validated
        Me.EngrCost.Text = Val(Int(Me.EngrNos.Text)) * Val(Me.EngrSalary.Text) * Val(Me.EngrMonths.Text)
    End Sub

    Private Sub EngrNos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles EngrNos.Validating
        If Not IsNumeric(EngrNos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub EngrSalary_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles EngrSalary.Validated
        Me.EngrCost.Text = Val(Int(Me.EngrNos.Text)) * Val(Me.EngrSalary.Text) * Val(Me.EngrMonths.Text)
    End Sub

    Private Sub EngrSalary_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles EngrSalary.Validating
        If Not IsNumeric(EngrSalary.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub EngrMonths_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles EngrMonths.Validated
        Me.EngrCost.Text = Val(Int(Me.EngrNos.Text)) * Val(Me.EngrSalary.Text) * Val(Me.EngrMonths.Text)
    End Sub

    Private Sub EngrMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles EngrMonths.Validating
        If Not IsNumeric(EngrMonths.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub SupNos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles SupNos.Validated
        Me.SupCost.Text = Val(Int(Me.SupNos.Text)) * Val(Me.SupSalary.Text) * Val(Me.SupMonths.Text)
    End Sub

    Private Sub SupNos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SupNos.Validating
        If Not IsNumeric(SupNos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub SupSalary_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles SupSalary.Validated
        Me.SupCost.Text = Val(Int(Me.SupNos.Text)) * Val(Me.SupSalary.Text) * Val(Me.SupMonths.Text)
    End Sub

    Private Sub SupSalary_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SupSalary.Validating
        If Not IsNumeric(SupSalary.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub
    Private Sub SupMonths_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles SupMonths.Validated
        Me.SupCost.Text = Val(Int(Me.SupNos.Text)) * Val(Me.SupSalary.Text) * Val(Me.SupMonths.Text)
    End Sub

    Private Sub SupMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SupMonths.Validating
        If Not IsNumeric(SupMonths.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub
    Private Sub MechForemanNos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MechForemanNos.Validated
        Me.MechForemanCost.Text = Val(Int(Me.MechForemanNos.Text)) * Val(Me.MechForemanSalary.Text) * Val(Me.MechForemanMonths.Text)
    End Sub

    Private Sub MechForemanNos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MechForemanNos.Validating
        If Not IsNumeric(MechForemanNos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub MechForemanSalary_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MechForemanSalary.Validated
        Me.MechForemanCost.Text = Val(Int(Me.MechForemanNos.Text)) * Val(Me.MechForemanSalary.Text) * Val(Me.MechForemanMonths.Text)
    End Sub

    Private Sub MechForemanSalary_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MechForemanSalary.Validating
        If Not IsNumeric(MechForemanSalary.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub
    Private Sub MechForemanMonths_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MechForemanMonths.Validated
        Me.MechForemanCost.Text = Val(Int(Me.MechForemanNos.Text)) * Val(Me.MechForemanSalary.Text) * Val(Me.MechForemanMonths.Text)
    End Sub

    Private Sub MechForemanMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MechForemanMonths.Validating
        If Not IsNumeric(MechForemanMonths.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub ElecForemanNos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ElecForemanNos.Validated
        Me.ElecForemanCost.Text = Val(Int(Me.ElecForemanNos.Text)) * Val(Me.ElecForemanSalary.Text) * Val(Me.ElecForemanMonths.Text)
    End Sub

    Private Sub ElecForemanNos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ElecForemanNos.Validating
        If Not IsNumeric(ElecForemanNos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub ElecForemanMonths_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ElecForemanMonths.Validated
        Me.ElecForemanCost.Text = Val(Int(Me.ElecForemanNos.Text)) * Val(Me.ElecForemanSalary.Text) * Val(Me.ElecForemanMonths.Text)
    End Sub

    Private Sub ElecForemanMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ElecForemanMonths.Validating
        If Not IsNumeric(ElecForemanMonths.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub ElecForemanSalary_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ElecForemanSalary.Validated
        Me.ElecForemanCost.Text = Val(Int(Me.ElecForemanNos.Text)) * Val(Me.ElecForemanSalary.Text) * Val(Me.ElecForemanMonths.Text)
    End Sub

    Private Sub ElecForemanSalary_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ElecForemanSalary.Validating
        If Not IsNumeric(ElecForemanSalary.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub ManagerCost_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ManagerCost.TextChanged
        Me.TotStaffCost.Text = Val(ManagerCost.Text) + Val(EngrCost.Text) + Val(SupCost.Text) + Val(MechForemanCost.Text) + Val(ElecForemanCost.Text)
    End Sub

    Private Sub EngrCost_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles EngrCost.TextChanged
        Me.TotStaffCost.Text = Val(ManagerCost.Text) + Val(EngrCost.Text) + Val(SupCost.Text) + Val(MechForemanCost.Text) + Val(ElecForemanCost.Text)
    End Sub

    Private Sub SupCost_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles SupCost.TextChanged
        Me.TotStaffCost.Text = Val(ManagerCost.Text) + Val(EngrCost.Text) + Val(SupCost.Text) + Val(MechForemanCost.Text) + Val(ElecForemanCost.Text)
    End Sub

    Private Sub MechForemanCost_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MechForemanCost.TextChanged
        Me.TotStaffCost.Text = Val(ManagerCost.Text) + Val(EngrCost.Text) + Val(SupCost.Text) + Val(MechForemanCost.Text) + Val(ElecForemanCost.Text)
    End Sub

    Private Sub ElecForemanCost_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ElecForemanCost.TextChanged
        Me.TotStaffCost.Text = Val(ManagerCost.Text) + Val(EngrCost.Text) + Val(SupCost.Text) + Val(MechForemanCost.Text) + Val(ElecForemanCost.Text)
    End Sub

    Private Sub tbcBdgetHeads_Selected(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlEventArgs) Handles tbcBdgetHeads.Selected

    End Sub

    Private Sub tbcBdgetHeads_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbcBdgetHeads.SelectedIndexChanged
        If Me.tbcBdgetHeads.SelectedIndex <> 0 Then
            Me.Panel3.Enabled = False
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            If optBPFixedExp.Checked Then
                optBPFixedExp.Checked = False
            ElseIf optFixedExp.Checked Then
                optFixedExp.Checked = False
            ElseIf Me.optHireEquipments.Checked Then
                optHireEquipments.Checked = False
            ElseIf optLightingEquips.Checked Then
                optLightingEquips.Checked = False
            ElseIf optMajConcrete.Checked Then
                optMajConcrete.Checked = False
            ElseIf Me.optMajConvyance.Checked Then
                optMajConvyance.Checked = False
            ElseIf optMajCrane.Checked Then
                optMajCrane.Checked = False
            ElseIf optMajDGSets.Checked Then
                optMajDGSets.Checked = False
            ElseIf optMajMH.Checked Then
                optMajMH.Checked = False
            ElseIf optMajNc.Checked Then
                optMajNc.Checked = False
            ElseIf optMajOthers.Checked Then
                optMajOthers.Checked = False
            ElseIf optMinorEquips.Checked Then
                optMinorEquips.Checked = False
            End If
        Else
            Me.Panel3.Enabled = True
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = True
        End If
        If Me.tbcBdgetHeads.SelectedIndex = 1 Then
            ErrorLabel.Text = ""
            ErrorLabel.Visible = False
        End If
        If Me.tbcBdgetHeads.SelectedIndex = 2 Then
            PipelineCostError.Text = ""
            PipelineCostError.Visible = False
        End If
        If Me.tbcBdgetHeads.SelectedIndex = 3 Then
            EsalMessage.Text = ""
            EsalMessage.Visible = False
        End If
        If Me.tbcBdgetHeads.SelectedIndex = 4 Then
            MiscMessage.Text = ""
            MiscMessage.Visible = False
        End If
        If Me.tbcBdgetHeads.SelectedIndex = 5 Then
            StaffSalaryMessage.Text = ""
            StaffSalaryMessage.Visible = False
        End If
    End Sub

    Private Sub StaffCostSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StaffCostSave.Click
        Dim currentsheetname2 As String, currentsheetname1 As String, intI As Integer
        currentsheetname2 = "Staff Salary"
        Me.btnQuit.Enabled = False
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
                'PrevSheetname = xlWorksheet.Name
                'xlWorksheet = xlWorkbook.Sheets.Item(intI + 1)
                SheetNo = intI + 1
                Exit For
            End If
        Next
        intI = 1

        xlRange = xlWorksheet.Range("StaffSalary_ManagerNos")
        xlRange.Value = Me.ManagerNos.Text
        xlRange = xlWorksheet.Range("StaffSalary_EngNos")
        xlRange.Value = Me.EngrNos.Text
        xlRange = xlWorksheet.Range("StaffSalary_SupNos")
        xlRange.Value = Me.SupNos.Text
        xlRange = xlWorksheet.Range("StaffSalary_FormanNos")
        xlRange.Value = Me.MechForemanNos.Text
        xlRange = xlWorksheet.Range("StaffSalary_FormanElecNos")
        xlRange.Value = Me.ElecForemanNos.Text
        xlRange = xlWorksheet.Range("StaffSalary_ManagerSalaryPM")
        xlRange.Value = Me.TManagerSalary.Text
        xlRange = xlWorksheet.Range("StaffSalary_EngSalaryPM")
        xlRange.Value = Me.EngrSalary.Text
        xlRange = xlWorksheet.Range("StaffSalary_SupSalaryPM")
        xlRange.Value = Me.SupSalary.Text
        xlRange = xlWorksheet.Range("StaffSalary_FormanSalaryPM")
        xlRange.Value = Me.MechForemanSalary.Text
        xlRange = xlWorksheet.Range("StaffSalary_FormanElecSalaryPM")
        xlRange.Value = Me.ElecForemanSalary.Text
        xlRange = xlWorksheet.Range("StaffSalary_ManagerMonths")
        xlRange.Value = Me.ManagerMonths.Text
        xlRange = xlWorksheet.Range("StaffSalary_EngMonths")
        xlRange.Value = Me.EngrMonths.Text
        xlRange = xlWorksheet.Range("StaffSalary_SupMonths")
        xlRange.Value = Me.SupMonths.Text
        xlRange = xlWorksheet.Range("StaffSalary_FormanMonths")
        xlRange.Value = Me.MechForemanMonths.Text
        xlRange = xlWorksheet.Range("StaffSalary_FormanElecMonths")
        xlRange.Value = Me.ElecForemanMonths.Text

        StaffSalaryMessage.Text = "Salary Expense Saved. "
        StaffSalaryMessage.Visible = True
        btnQuit.Enabled = True
    End Sub


    Private Sub optLightingEquips_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optLightingEquips.CheckedChanged
        Dim intL As Integer
        If optMinorEquips.Checked = True Then
            Panel4.Left = 1
            Panel4.Top = 20
            Panel4.Height = 30
            Panel4.Visible = True
            Panel2.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optHireEquipments.Checked Then
            Panel5.Left = 1
            Panel5.Top = 20
            Panel5.Height = 30
            Panel5.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
        ElseIf Me.optFixedExp.Checked Or optBPFixedExp.Checked Then
            Panel6.Left = 1
            Panel6.Top = 20
            Panel6.Height = 30
            Panel6.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel7.Visible = False
        ElseIf Me.optLightingEquips.Checked Then
            Panel7.Left = 1
            Panel7.Top = 15
            Panel7.Height = 45
            Me.Label283.Text = "Conn" & vbNewLine & "Load"
            Me.Label282.Text = "Utility" & vbNewLine & "Factor"
            Panel7.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
        Else
            Panel2.Left = 1
            Panel2.Top = 20
            Panel2.Height = 30
            Panel2.Visible = True
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel2.BringToFront()
        End If

        If Not optLightingEquips.Checked Then
            mcategory = "Lighting"
            SelectedCategory = mcategory
            Dim intK As Integer, L As Integer
            LightingItems = EquipNameTextBoxes.Count
            L = LightingItems - 1
            ReDim Preserve LightingChecked(L)
            ReDim Preserve LightMDate(L)
            ReDim Preserve LightDMDate(L)
            ReDim Preserve LightQuantity(L)
            ReDim Preserve LightPowerPerUnit(L)
            ReDim Preserve LightConnectLoad(L)
            ReDim Preserve LightUtilityFactor(L)
            'ReDim LightCategory(L)
            ReDim LightEName(L)   'concEName(L)
            ReDim LightCapacity(L)      'concCapacity(L)
            ReDim LightMakeModel(L)     'concMakeModel(L)
            ReDim LightMDate(L)         'concMDate(L)
            ReDim LightDMDate(L)        'concDMdate(L)
            ReDim LightQuantity(L)           'concQuantity(L)
            ReDim LightConnectLoad(L)   'concHrsPerMonth(L)
            ReDim LightUtilityFactor(L) 'concDep(L)
            ReDim LightPowerPerUnit(L)
            ReDim LightingChecked(L)
            For intK = 0 To L
                LightEName(intK) = EquipNameTextBoxes(intK).Text
                LightCapacity(intK) = CapacityTextBoxes(intK).Text
                LightMakeModel(intK) = MakeModelTextBoxes(intK).Text
                LightMDate(intK) = MobdatePickers(intK).Value.Date
                LightDMDate(intK) = DemobDatePickers(intK).Value.Date
                LightQuantity(intK) = Val(QtyTextBoxes(intK).Text)
                LightPowerPerUnit(intK) = Val(PowerPerUnitTextBoxes(intK).Text)
                LightConnectLoad(intK) = Val(ConnectLoadTextBoxes(intK).Text)
                LightUtilityFactor(intK) = Val(UtilityFactorTextBoxes(intK).Text)
            Next
            For intK = 0 To L
                If Checkboxes(intK).Checked Then
                    LightingChecked(intK) = 1
                    LightMDate(intK) = MobdatePickers(intK).Value.Date
                    LightDMDate(intK) = DemobDatePickers(intK).Value.Date
                    LightQuantity(intK) = Val(QtyTextBoxes(intK).Text)
                    LightPowerPerUnit(intK) = Val(PowerPerUnitTextBoxes(intK).Text)
                    LightConnectLoad(intK) = Val(ConnectLoadTextBoxes(intK).Text)
                    LightUtilityFactor(intK) = Val(UtilityFactorTextBoxes(intK).Text)
                Else
                    LightingChecked(intK) = 0
                End If
            Next
            LightingItems = EquipNameTextBoxes.Count
            WriteToLightingEquipsArray(LightingItems)
            Exit Sub
        End If
        Button1.Enabled = True
        mcategory = "Lighting"
        SelectedCategory = mcategory
        Me.tbcBdgetHeads.TabPages(0).Controls.Clear()
        mcategory = "Lighting"
        If Not LightingFirstTime Then
            LoadControlsFromLightingArray(LightingItems)
            For intL = 0 To LightingItems - 1
                If LightingChecked(intL) = 1 Then
                    Checkboxes(intL).Checked = True
                    MobdatePickers(intL).Value = LightMDate(intL)
                    DemobDatePickers(intL).Value = LightDMDate(intL)
                    QtyTextBoxes(intL).Text = LightQuantity(intL)
                    HrsPermonthTextBoxes(intL).Text = LightPowerPerUnit(intL)
                    DepPercComboboxes(intL).Text = LightConnectLoad(intL)
                    ShiftsComboboxes(intL).Text = LightUtilityFactor(intL)
                Else
                    Checkboxes(intL).Checked = False
                End If
                'Checkboxes(intL + 1).Checked = IIf(ConcreteChecked(intL) = 1, True, False)
            Next
            Me.Refresh()
        Else
            mcategory = "Lighting"
            SelectedCategory = mcategory
            LoadLightingControlsInPage(mcategory)
            For intL = 0 To Checkboxes.Count - 1
                If Checkboxes(intL).Checked Then
                    LightingChangeEnabled(True, intL)
                End If
            Next
            LightingFirstTime = False
        End If
        LightingItems = EquipNameTextBoxes.Count
    End Sub
    Private Sub txtElectricianNos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtElectricianNos.Validated
        Me.txtElectricianSalary.Text = Val(Int(Me.txtElectricianNos.Text)) * Val(Me.txtElectSalaryPM.Text) * Val(Me.txtElectMonths.Text)
    End Sub

    Private Sub txtElectricianNos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtElectricianNos.Validating
        If Not IsNumeric(txtElectricianNos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtElectSalaryPM_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtElectSalaryPM.TextChanged

    End Sub

    Private Sub txtElectSalaryPM_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtElectSalaryPM.Validated
        Me.txtElectricianSalary.Text = Val(Int(Me.txtElectricianNos.Text)) * Val(Me.txtElectSalaryPM.Text) * Val(Me.txtElectMonths.Text)
    End Sub

    Private Sub txtElectSalaryPM_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtElectSalaryPM.Validating
        If Not IsNumeric(txtElectSalaryPM.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtElectMonths_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtElectMonths.TextChanged
        Me.txtElectricianSalary.Text = Val(Int(Me.txtElectricianNos.Text)) * Val(Me.txtElectSalaryPM.Text) * Val(Me.txtElectMonths.Text)
    End Sub

    Private Sub txtElectMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtElectMonths.Validating
        If Not IsNumeric(txtElectMonths.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtEletHepNos_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEletHepNos.TextChanged

    End Sub

    Private Sub txtEletHepNos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEletHepNos.Validated
        Me.txtElectHelperSalary.Text = Val(Int(Me.txtEletHepNos.Text)) * Val(Me.txtElectHelperSalaryPM.Text) * Val(Me.txtElectHelperMonths.Text)
    End Sub

    Private Sub txtEletHepNos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtEletHepNos.Validating
        If Not IsNumeric(txtEletHepNos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtElectHelperSalaryPM_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtElectHelperSalaryPM.Validated
        Me.txtElectHelperSalary.Text = Val(Int(Me.txtEletHepNos.Text)) * Val(Me.txtElectHelperSalaryPM.Text) * Val(Me.txtElectHelperMonths.Text)
    End Sub

    Private Sub txtElectHelperSalaryPM_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtElectHelperSalaryPM.Validating
        If Not IsNumeric(txtElectHelperSalaryPM.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtElectHelperMonths_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtElectHelperMonths.Validated
        Me.txtElectHelperSalary.Text = Val(Int(Me.txtEletHepNos.Text)) * Val(Me.txtElectHelperSalaryPM.Text) * Val(Me.txtElectHelperMonths.Text)
    End Sub

    Private Sub txtElectHelperMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtElectHelperMonths.Validating
        If Not IsNumeric(txtElectHelperMonths.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtMechanicNos_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMechanicNos.TextChanged

    End Sub

    Private Sub txtMechanicNos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMechanicNos.Validated
        Me.txtMechanicSalary.Text = Val(Int(Me.txtMechanicNos.Text)) * Val(Me.txtMechanicSalaryPM.Text) * Val(Me.txtMechanicMonths.Text)
    End Sub

    Private Sub txtMechanicNos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMechanicNos.Validating
        If Not IsNumeric(txtMechanicNos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtMechanicSalaryPM_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMechanicSalaryPM.Validated
        Me.txtMechanicSalary.Text = Val(Int(Me.txtMechanicNos.Text)) * Val(Me.txtMechanicSalaryPM.Text) * Val(Me.txtMechanicMonths.Text)
    End Sub

    Private Sub txtMechanicSalaryPM_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMechanicSalaryPM.Validating
        If Not IsNumeric(txtMechanicSalaryPM.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtMechanicMonths_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMechanicMonths.Validated
        Me.txtMechanicSalary.Text = Val(Int(Me.txtMechanicNos.Text)) * Val(Me.txtMechanicSalaryPM.Text) * Val(Me.txtMechanicMonths.Text)
    End Sub

    Private Sub txtMechanicMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMechanicMonths.Validating
        If Not IsNumeric(txtMechanicMonths.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub
    Private Sub txtMechHelperNos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMechHelperNos.Validated
        Me.txtMechHelpersalary.Text = Val(Int(Me.txtMechHelperNos.Text)) * Val(Me.txtMechHelperSalaryPM.Text) * Val(Me.txtMechHelperMonths.Text)
    End Sub

    Private Sub txtMechHelperNos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMechHelperNos.Validating
        If Not IsNumeric(txtMechHelperNos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtMechHelperSalaryPM_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMechHelperSalaryPM.Validated
        Me.txtMechHelpersalary.Text = Val(Int(Me.txtMechHelperNos.Text)) * Val(Me.txtMechHelperSalaryPM.Text) * Val(Me.txtMechHelperMonths.Text)
    End Sub

    Private Sub txtMechHelperSalaryPM_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMechHelperSalaryPM.Validating
        If Not IsNumeric(txtMechHelperSalaryPM.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtMechHelperMonths_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMechHelperMonths.Validated
        Me.txtMechHelpersalary.Text = Val(Int(Me.txtMechHelperNos.Text)) * Val(Me.txtMechHelperSalaryPM.Text) * Val(Me.txtMechHelperMonths.Text)
    End Sub

    Private Sub txtMechHelperMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMechHelperMonths.Validating
        If Not IsNumeric(txtMechHelperMonths.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtTyreFitterNos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTyreFitterNos.Validated
        Me.txtTyreFitterSalary.Text = Val(Int(Me.txtTyreFitterNos.Text)) * Val(Me.txtTyreFitterSalaryPM.Text) * Val(Me.txtTyreFitterMonths.Text)
    End Sub

    Private Sub txtTyreFitterNos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTyreFitterNos.Validating
        If Not IsNumeric(txtTyreFitterNos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub
    Private Sub txtAutoElecNos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAutoElecNos.Validated
        Me.txtAutoElecSalary.Text = Val(Int(Me.txtAutoElecNos.Text)) * Val(Me.txtAutoElecSalaryPM.Text) * Val(Me.txtAutoElecMonths.Text)
    End Sub

    Private Sub txtAutoElecNos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAutoElecNos.Validating
        If Not IsNumeric(txtAutoElecNos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtAutoElecSalaryPM_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAutoElecSalaryPM.Validated
        Me.txtAutoElecSalary.Text = Val(Int(Me.txtAutoElecNos.Text)) * Val(Me.txtAutoElecSalaryPM.Text) * Val(Me.txtAutoElecMonths.Text)
    End Sub

    Private Sub txtAutoElecSalaryPM_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAutoElecSalaryPM.Validating
        If Not IsNumeric(txtAutoElecSalaryPM.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtAutoElecMonths_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAutoElecMonths.Validated
        Me.txtAutoElecSalary.Text = Val(Int(Me.txtAutoElecNos.Text)) * Val(Me.txtAutoElecSalaryPM.Text) * Val(Me.txtAutoElecMonths.Text)
    End Sub

    Private Sub txtAutoElecMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAutoElecMonths.Validating
        If Not IsNumeric(txtAutoElecMonths.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtWelderNos_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWelderNos.TextChanged

    End Sub

    Private Sub txtWelderNos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWelderNos.Validated
        Me.txtWelderSalary.Text = Val(Int(Me.txtWelderNos.Text)) * Val(Me.txtWelderSalaryPM.Text) * Val(Me.txtWelderMonths.Text)
    End Sub

    Private Sub txtWelderNos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtWelderNos.Validating
        If Not IsNumeric(txtWelderNos.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtWelderSalaryPM_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWelderSalaryPM.Validated
        Me.txtWelderSalary.Text = Val(Int(Me.txtWelderNos.Text)) * Val(Me.txtWelderSalaryPM.Text) * Val(Me.txtWelderMonths.Text)
    End Sub

    Private Sub txtWelderSalaryPM_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtWelderSalaryPM.Validating
        If Not IsNumeric(txtWelderSalaryPM.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtWelderMonths_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWelderMonths.Validated
        Me.txtWelderSalary.Text = Val(Int(Me.txtWelderNos.Text)) * Val(Me.txtWelderSalaryPM.Text) * Val(Me.txtWelderMonths.Text)
    End Sub

    Private Sub txtWelderMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtWelderMonths.Validating
        If Not IsNumeric(txtWelderMonths.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub txtTyreFitterMonths_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTyreFitterMonths.TextChanged

    End Sub

    Private Sub txtTyreFitterMonths_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTyreFitterMonths.Validated
        Me.txtTyreFitterSalary.Text = Val(Int(Me.txtTyreFitterNos.Text)) * Val(Me.txtTyreFitterSalaryPM.Text) * Val(Me.txtTyreFitterMonths.Text)
    End Sub

    Private Sub txtTyreFitterMonths_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTyreFitterMonths.Validating
        If Not IsNumeric(txtTyreFitterMonths.Text) Then
            MsgBox("Only numeric values accepted")
            e.Cancel = True
        End If
    End Sub

    Private Sub EsalClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EsalClear.Click
        Dim ctl As Control
        For Each ctl In Me.tbcBdgetHeads.TabPages(3).Controls
            If ctl.GetType Is GetType(TextBox) Then
                ctl.Text = ""
            End If
        Next
    End Sub


    Private Sub optMajConcrete_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles optMajConcrete.Validating
        'If dovalidate Then
        '    Dim validated As Boolean = TestValidity(1, CategoryTextBoxes.Count)
        '    If Not validated Then
        '        dovalidate = False
        '        e.Cancel = True
        '    End If
        'End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        'Dim ans As Integer
        'ans = MsgBox("You have opted to close the application without saving any updated data." & vbNewLine & _
        '"Are you sure ?", MsgBoxStyle.YesNo)
        Dim frmclose As Form
        frmclose = New Form4
        frmclose.ShowDialog()
        frmclose = Nothing
        If answer = vbNo Then Exit Sub
        If Not moledbConnection.State.ToString().Equals("Closed") Then
            moledbConnection.Close()
            moledbConnection = Nothing
        End If
        If Not moledbConnection1.State.ToString().Equals("Closed") Then
            moledbConnection1.Close()
            moledbConnection1 = Nothing
        End If
        If Not xlWorkbook Is Nothing Then
            xlWorksheet = xlWorkbook.Sheets.Item(1)
            xlWorksheet.Select()
            xlApp.CalculateBeforeSave = True
            xlApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic
            xlApp.CalculateFull()
            xlWorkbook.Save()
            xlWorkbook.Close()
            xlWorkbook = Nothing
        End If
        xlApp.Quit()
        xlApp = Nothing
        System.GC.Collect()
        frmProjectDetails.Show()
    End Sub

    Private Sub MajorEquipments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MajorEquipments.Click

    End Sub
End Class
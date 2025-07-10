Imports System.Data
Imports System.Data.OleDb

Public Class frmProjectDetails
    Public moledbConnection As OleDbConnection
    Dim strStatement As String
    Dim moledbCommand As OleDbCommand
    Dim mOledbDataAdapter As OleDbDataAdapter
    Dim mReader As OleDbDataReader
    Dim mDataSet As DataSet
    Dim oForm As Form


    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub SaveFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        Me.txtWorkbookName.Text = Me.SaveFileDialog1.FileName
        Dim fi As System.IO.FileInfo
        fi = New System.IO.FileInfo(SaveFileDialog1.FileName)
        Me.txtWorkbookName.Text = SaveFileDialog1.FileName
        DestinationFolder = fi.Directory.ToString()
        mdbfilename = Replace(fi.Name, ".xls", ".mdb")
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        'xlApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        lo()
        Dim intI As Integer, moledbInsertCommand As New OleDbCommand
        Dim Records As Integer, strSql As String

        Dim InsertCommand As String   ', databaseCreated As Boolean
        Dim ShortName As String = ""
        FormLoaded = False

        If Len(Trim(Me.txtWorkbookName.Text)) = 0 Then
            MsgBox("Filename not specified to save")
            Exit Sub
        End If
        TemplatePath = (My.Application.Info.DirectoryPath)

        If (Me.dpStartDate.Value.Date > Me.dpEndDate.Value.Date And mWorkingMode = "New") Then
            MsgBox("Start date must be prior to the end date")
            Me.dpStartDate.Focus()
            Exit Sub
        End If
        If (Me.dpEndDate.Value.Date < Today()) Then
            MsgBox("Only furture dates are allowed for start and end dates")
            Me.dpEndDate.Focus()
            Exit Sub
        End If
        mMainTitle1 = "SHAPOORJI PALLONJI & CO. LTD"
        mMainTitle2 = "Project : " & Me.txtProjectdescription.Text
        'mMainTitle3 = "PMV Budget - Major Eqpts - Concreting Eqpts"
        mClient = Me.txtClient.Text
        mLocation = Me.txtLocation.Text
        mStartDate = Me.dpStartDate.Value.Date
        mStartDate = mStartDate.Date
        mEndDate = Me.dpEndDate.Value.Date
        mEndDate = mEndDate.Date
        mConcreteQty = Me.txtConcreteQuantity.Text
        FuelCostperLtr = Val(Me.txtFuelCostPerLtr.Text)
        PowerCostPerUnit = Val(Me.txtPowerCostPerUnit.Text)
        mProjectvalue = Val(Me.txtProjectValue.Text) * (10 ^ 7) ' to convert eneterd data into crores
        Currentfile = Me.txtWorkbookName.Text


        If mWorkingMode = "New" Then
            copyAddedItemsMdb(TemplatePath & "\AddedEquipments-ver20.mdb", DestinationFolder & "\" & mdbfilename)

            strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
            strconnection1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DestinationFolder & "\" & mdbfilename
            moledbConnection = New OleDbConnection(strConnection)
            If (moledbConnection.State.ToString().Equals("Closed")) Then
                moledbConnection.Open()
            End If
            moledbConnection1 = New OleDbConnection(strconnection1)
            If (moledbConnection1.State.ToString().Equals("Closed")) Then
                moledbConnection1.Open()
            End If

            mPassword = ""
            oForm = New fmgetpassword()
            oForm.ShowDialog()
            oForm = Nothing
            Me.lblMessage.Text = "Form Details being saved. Please wait..."
            Me.lblMessage.Visible = True
            Me.Refresh()
            If Len(Trim(mPassword)) = 0 Then Exit Sub
            InsertCommand = "INSERT INTO Projects Values ("
            InsertCommand = InsertCommand & "'" & Me.txtProjectCode.Text & "', "
            InsertCommand = InsertCommand & "'" & Me.txtProjectdescription.Text & "', "
            InsertCommand = InsertCommand & "'" & Me.txtLocation.Text & "', "
            InsertCommand = InsertCommand & Val(Me.txtProjectValue.Text) & ", "
            InsertCommand = InsertCommand & "'" & Me.dpStartDate.Text & "', "
            InsertCommand = InsertCommand & "'" & Me.dpEndDate.Text & "', "

            InsertCommand = InsertCommand & "'" & Me.txtClient.Text & "', "
            InsertCommand = InsertCommand & "'" & Me.txtWorkbookName.Text & "', "
            InsertCommand = InsertCommand & "'" & mPassword & "', "
            InsertCommand = InsertCommand & mConcreteQty & ", "
            InsertCommand = InsertCommand & "'" & DestinationFolder & "\" & mdbfilename & "', "
            InsertCommand = InsertCommand & Val(Me.txtFuelCostPerLtr.Text) & ", "
            InsertCommand = InsertCommand & Val(Me.txtPowerCostPerUnit.Text) & ")"

            Try
                moledbInsertCommand.CommandType = CommandType.Text
                moledbInsertCommand.CommandText = InsertCommand
                moledbInsertCommand.Connection = moledbConnection
                moledbInsertCommand.ExecuteNonQuery()
                moledbConnection.Close()
                moledbInsertCommand = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
                lblMessage.Visible = False
                Exit Sub
            End Try
            '----------creating correspondng excel 2orkbbok and sheets and saving project details for new project---------------------------
            With xlApp
                Try
                    xlWorkbook = .Workbooks.Open(TemplatePath & "\SP_ProjectBudgetTemplate-Ver2.xls")
                    xlWorkbook.SaveAs(Me.txtWorkbookName.Text)
                Catch ex As Exception
                    MsgBox("Could not open the Template file." & ex.Message)
                    Exit Sub
                End Try
            End With
            For intI = 1 To xlWorkbook.Sheets.Count
                xlWorksheet = xlWorkbook.Sheets(intI)
                If Not xlWorksheet.Visible Then Continue For
                With xlWorksheet
                    getCategoryShortname(xlWorksheet)
                    xlRange = .Range(Category_Shortname & "MainTitle1")
                    xlRange.Value = mMainTitle1
                    xlRange.Cells.Font.Size = 18
                    xlRange.Cells.Font.Bold = True
                    xlRange.Cells.RowHeight = 23
                    xlRange.Cells.WrapText = False
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = .Range(Category_Shortname & "MainTitle2")
                    xlRange.Value = mMainTitle2
                    xlRange.Cells.Font.Size = 18
                    xlRange.Cells.Font.Bold = True
                    xlRange.Cells.RowHeight = 23
                    xlRange.Cells.WrapText = False
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = .Range(Category_Shortname & "Client")
                    xlRange.Value = mClient
                    xlRange.Cells.Font.Size = 14
                    xlRange.Cells.Font.Bold = True
                    xlRange.Cells.RowHeight = 23
                    xlRange.Cells.WrapText = False
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = .Range(Category_Shortname & "Location")
                    xlRange.Value = mLocation
                    xlRange.Cells.Font.Size = 14
                    xlRange.Cells.Font.Bold = True
                    xlRange.Cells.RowHeight = 23
                    xlRange.Cells.WrapText = False
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = .Range(Category_Shortname & "StartDate")
                    xlRange.Value = mStartDate.Date.ToString()
                    xlRange.NumberFormat = "dd-mmm-yyyy"
                    xlRange.Cells.Font.Size = 14
                    xlRange.Cells.Font.Bold = True
                    xlRange.Cells.RowHeight = 23
                    xlRange.Cells.ColumnWidth = 21
                    xlRange.Cells.WrapText = False
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = .Range(Category_Shortname & "ProjectValue")
                    xlRange.Value = mProjectvalue
                    xlRange.Cells.Font.Size = 14
                    xlRange.Cells.Font.Bold = True
                    xlRange.Cells.RowHeight = 23
                    xlRange.Cells.WrapText = False
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange = .Range(Category_Shortname & "EndDate")
                    xlRange.Value = mEndDate.Date.ToString()
                    xlRange.Cells.Font.Size = 14
                    xlRange.Cells.Font.Bold = True
                    xlRange.Cells.RowHeight = 23
                    xlRange.Cells.ColumnWidth = 21
                    xlRange.Cells.WrapText = False
                    xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                    xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    xlRange.NumberFormat = "dd-mmm-yyyy"
                    If (intI > 1 And intI < 13) Then
                        RecordsInserted(intI) = 0
                        xlWorksheet.Range(Category_Shortname & "RecordsTotal").Value = 0
                    End If

                    If Category_Shortname = "Concrete_" Then
                        xlRange = .Range(Category_Shortname & "ConcreteQty")
                        xlRange.Value = mConcreteQty
                        xlRange.Cells.Font.Size = 14
                        xlRange.Cells.Font.Bold = True
                        xlRange.Cells.RowHeight = 23
                        xlRange.Cells.WrapText = False
                        xlRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                        xlRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    End If
                End With
            Next
            '--------------------------------------------------------------------------------------------------------------------
            xlFilename = Me.txtWorkbookName.Text
            CopyRecordFromMasterToAddedItemsArray("Concreting", "MajorConcreteEquips")
            CopyRecordFromMasterToAddedItemsArray("Conveyance", "MajorConveyanceEquips")
            CopyRecordFromMasterToAddedItemsArray("Cranes", "MajorCraneEquips")
            CopyRecordFromMasterToAddedItemsArray("DG Sets", "MajorDGSetsEquips")
            CopyRecordFromMasterToAddedItemsArray("Material Handling", "MajorMaterialhandlingEquips")
            CopyRecordFromMasterToAddedItemsArray("Non Concreting", "MajorNonConcreteEquips")
            CopyRecordFromMasterToAddedItemsArray("Major Others", "MajorOthers")
            CopyRecordFromMasterToAddedItemsArray("Minor Equipments", "MinorEquips")
            CopyRecordFromMasterToAddedItemsArray("HiredConveyance", "HiredEquips")
            CopyRecordFromMasterToAddedItemsArray("FixedExpenses", "FixedExpAdded")
            CopyRecordFromMasterToAddedItemsArray("BPFixedExpenses", "BPFixedExpAdded")
            CopyRecordFromMasterToAddedItemsArray("Lighting", "LightingEquipments")

            Records = concEquipsNames.Length
            ReDim ConcreteChecked(Records - 1)
            ReDim ConcreteMobdate(Records - 1)
            ReDim ConcreteDemobdate(Records - 1)
            ReDim ConcreteQty(Records - 1)
            ReDim ConcreteHrs(Records - 1)
            ReDim ConcreteDepPerc(Records - 1)
            ReDim ConcreteShifts(Records - 1)

            Records = convEquipsNames.Length
            ReDim ConvChecked(Records - 1)
            ReDim ConvMobdate(Records - 1)
            ReDim ConvDemobdate(Records - 1)
            ReDim ConvQty(Records - 1)
            ReDim ConvHrs(Records)
            ReDim ConvDepPerc(Records - 1)
            ReDim ConvShifts(Records - 1)

            Records = craneEquipsNames.Length
            ReDim CraneChecked(Records - 1)
            ReDim CraneMobdate(Records - 1)
            ReDim CraneDemobdate(Records - 1)
            ReDim CraneQty(Records - 1)
            ReDim CraneHrs(Records - 1)
            ReDim CraneDepPerc(Records - 1)
            ReDim CraneShifts(Records - 1)

            Records = dgsetsEquipsNames.Length
            ReDim DGSetsChecked(Records - 1)
            ReDim DGSetsMobdate(Records - 1)
            ReDim DGSetsDemobdate(Records - 1)
            ReDim DGSetsQty(Records - 1)
            ReDim DGSetsHrs(Records - 1)
            ReDim DGSetsDepPerc(Records - 1)
            ReDim DGSetsShifts(Records - 1)

            Records = MHEquipsNames.Length
            ReDim MHChecked(Records - 1)
            ReDim MHMobdate(Records - 1)
            ReDim MHDemobdate(Records - 1)
            ReDim MHQty(Records - 1)
            ReDim MHHrs(Records - 1)
            ReDim MHDepPerc(Records - 1)
            ReDim MHShifts(Records - 1)

            Records = NCEquipsNames.Length
            ReDim nccHECKED(Records - 1)
            ReDim NCMobdate(Records - 1)
            ReDim NCDemobdate(Records - 1)
            ReDim NCQty(Records - 1)
            ReDim NCHrs(Records - 1)
            ReDim NCDepPerc(Records - 1)
            ReDim NCShifts(Records - 1)

            Records = majOthersEquipsNames.Length
            ReDim MajOthersChecked(Records - 1)
            ReDim MajOthersMobdate(Records - 1)
            ReDim MajOthersDemobdate(Records - 1)
            ReDim MajOthersQty(Records - 1)
            ReDim MajOthersHrs(Records - 1)
            ReDim MajOthersDepPerc(Records - 1)
            ReDim MajOthersShifts(Records - 1)

            Records = minorEquipsNames.Length
            ReDim MinorChecked(Records - 1)
            ReDim MinorMobdate(Records - 1)
            ReDim MinorDemobdate(Records - 1)
            ReDim MinorQty(Records - 1)
            ReDim MinorHrs(Records - 1)
            ReDim MinorDepPerc(Records - 1)
            ReDim MinorShifts(Records - 1)
            ReDim MinorNPV(Records - 1)

            Records = hiredEquipsNames.Length
            ReDim HiredChecked(Records - 1)
            ReDim HiredMobdate(Records - 1)
            ReDim HiredDemobdate(Records - 1)
            ReDim HiredQty(Records - 1)
            ReDim HiredHrs(Records - 1)
            ReDim HiredDepPerc(Records - 1)
            ReDim HiredShifts(Records - 1)
            ReDim HiredNPV(Records - 1)

            Records = fixedExpCategoryNames.Length
            ReDim FexpChecked(Records - 1)
            ReDim FExpQty(Records - 1)
            ReDim FExpRemarks(Records - 1)
            Records = fixedBPExpCategoryNames.Length
            ReDim BPFExpChecked(Records - 1)
            ReDim BPFExpQty(Records - 1)
            ReDim BPFExpRemarks(Records - 1)
            mOledbDataAdapter = Nothing

            Records = lightingCategoryNames.Length
            ReDim LightingChecked(Records - 1)
            ReDim LightingQty(Records - 1)
            ReDim LightingMobDate(Records - 1)
            ReDim LightingDemobDate(Records - 1)
            ReDim LightingPowerPerUnit(Records - 1)
            ReDim LightingConnectLoad(Records - 1)
            ReDim LightingUtilityFactor(Records - 1)
            mOledbDataAdapter = Nothing
            mDataSet = Nothing
        Else  ' in case of exixting project
            CopyAddedConcreteItemsToAddedItemsArray()
            CopyAddedConveyanceItemsToAddedItemsArray()
            CopyAddedCraneItemsToAddedItemsArray()
            CopyAddedDgSetsItemsToAddedItemsArray()
            CopyAddedMHItemsToAddedItemsArray()
            CopyAddedNCItemsToAddedItemsArray()
            CopyAddedMajOtherItemsToAddedItemsArray()
            CopyAddedMinorItemsToAddedItemsArray()
            CopyAddedHiredItemsToAddedItemsArray()
            CopyAddedFixedExpItemsToAddedItemsArray()
            CopyAddedBPFixedExpItemsToAddedItemsArray()
            CopyAddedLightingItemsToAddedItemsArray()

            Me.lblMessage.Text = "Form Details are loading. Please wait..."
            Me.lblMessage.Visible = True
            Me.Refresh()
            xlFilename = Me.txtWorkbookName.Text
            If IsWorkbookOpen(xlFilename) Then
                MsgBox("The project Workbook " & xlFilename & " is open. " & vbNewLine & "Please Close and then Continue")
                Exit Sub
            End If
            Try
                xlWorkbook = xlApp.Workbooks.Open(xlFilename)
            Catch ex As Exception
                MsgBox("Error in opening project workbook. " & ex.ToString())
                Exit Sub
            End Try

            For intI = 2 To 12
                xlWorksheet = xlWorkbook.Sheets(intI)
                getCategoryShortname(xlWorksheet)
                RecordsInserted(intI) = xlWorksheet.Range(Category_Shortname & "RecordsTotal").Value
            Next
        End If
        xlApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual
        ' ----------------------redimensioning the control arrays for each category-------------------------------------
        Records = concEquipsNames.Length
        ReDim ConcreteChecked(Records - 1)
        ReDim ConcreteMobdate(Records - 1)
        ReDim ConcreteDemobdate(Records - 1)
        ReDim ConcreteQty(Records - 1)
        ReDim ConcreteHrs(Records)
        ReDim ConcreteDepPerc(Records - 1)
        ReDim ConcreteShifts(Records - 1)


        Records = convEquipsNames.Length
        ReDim ConvChecked(Records - 1)
        ReDim ConvMobdate(Records - 1)
        ReDim ConvDemobdate(Records - 1)
        ReDim ConvQty(Records - 1)
        ReDim ConvHrs(Records)
        ReDim ConvDepPerc(Records - 1)
        ReDim ConvShifts(Records - 1)

        Records = craneEquipsNames.Length
        ReDim CraneChecked(Records - 1)
        ReDim CraneMobdate(Records - 1)
        ReDim CraneDemobdate(Records - 1)
        ReDim CraneQty(Records - 1)
        ReDim CraneHrs(Records - 1)
        ReDim CraneDepPerc(Records - 1)
        ReDim CraneShifts(Records - 1)

        Records = dgsetsEquipsNames.Length
        ReDim DGSetsChecked(Records - 1)
        ReDim DGSetsMobdate(Records - 1)
        ReDim DGSetsDemobdate(Records - 1)
        ReDim DGSetsQty(Records - 1)
        ReDim DGSetsHrs(Records - 1)
        ReDim DGSetsDepPerc(Records - 1)
        ReDim DGSetsShifts(Records - 1)

        Records = MHEquipsNames.Length
        ReDim MHChecked(Records - 1)
        ReDim MHMobdate(Records - 1)
        ReDim MHDemobdate(Records - 1)
        ReDim MHQty(Records - 1)
        ReDim MHHrs(Records - 1)
        ReDim MHDepPerc(Records - 1)
        ReDim MHShifts(Records - 1)

        Records = NCEquipsNames.Length
        ReDim nccHECKED(Records - 1)
        ReDim NCMobdate(Records - 1)
        ReDim NCDemobdate(Records - 1)
        ReDim NCQty(Records - 1)
        ReDim NCHrs(Records - 1)
        ReDim NCDepPerc(Records - 1)
        ReDim NCShifts(Records - 1)

        Records = majOthersEquipsNames.Length
        ReDim MajOthersChecked(Records - 1)
        ReDim MajOthersMobdate(Records - 1)
        ReDim MajOthersDemobdate(Records - 1)
        ReDim MajOthersQty(Records - 1)
        ReDim MajOthersHrs(Records - 1)
        ReDim MajOthersDepPerc(Records - 1)
        ReDim MajOthersShifts(Records - 1)

        Records = minorEquipsNames.Length
        ReDim MinorChecked(Records - 1)
        ReDim MinorMobdate(Records - 1)
        ReDim MinorDemobdate(Records - 1)
        ReDim MinorQty(Records - 1)
        ReDim MinorHrs(Records - 1)
        ReDim MinorDepPerc(Records - 1)
        ReDim MinorShifts(Records - 1)
        ReDim MinorNPV(Records - 1)

        Records = hiredEquipsNames.Length
        ReDim HiredChecked(Records - 1)
        ReDim HiredMobdate(Records - 1)
        ReDim HiredDemobdate(Records - 1)
        ReDim HiredQty(Records - 1)
        ReDim HiredHrs(Records - 1)
        ReDim HiredDepPerc(Records - 1)
        ReDim HiredShifts(Records - 1)
        ReDim HiredNPV(Records - 1)
        ' ReDim FexpChecked(Records - 1)
        mOledbDataAdapter = Nothing

        Records = fixedExpCategoryNames.Length
        ReDim FexpChecked(Records - 1)
        ReDim FExpQty(Records - 1)
        ReDim FExpRemarks(Records - 1)
        mOledbDataAdapter = Nothing

        Records = fixedBPExpCategoryNames.Length
        ReDim BPFExpChecked(Records - 1)
        ReDim BPFExpQty(Records - 1)
        ReDim BPFExpRemarks(Records - 1)
        mOledbDataAdapter = Nothing

        Records = lightingCategoryNames.Length
        ReDim LightingChecked(Records - 1)
        ReDim LightingQty(Records - 1)
        ReDim LightingMobDate(Records - 1)
        ReDim LightingDemobDate(Records - 1)
        ReDim LightingPowerPerUnit(Records - 1)
        ReDim LightingConnectLoad(Records - 1)
        ReDim LightingUtilityFactor(Records - 1)
        ' -------------------------end of redimensioning---------------------------------------------------------------
        mOledbDataAdapter = Nothing
        mDataSet = Nothing


        If Not xlWorkbook Is Nothing Then
            xlApp.CalculateBeforeSave = True
            xlWorkbook.Save()
            xlWorkbook.Close()
        End If
        btnSave.Enabled = False
        xlWorkbook = Nothing
        System.GC.Collect()
        'Dim oForm As frmOptions
        Me.Hide()
        oForm = New frmOptionsNew()
        oForm.ShowDialog()
        oForm = Nothing
        Me.btnSave.Enabled = False
    End Sub
    Private Sub copyAddedItemsMdb(ByVal Source As String, ByVal Target As String)
        If System.IO.File.Exists(Source) = True Then
            System.IO.File.Copy(Source, Target)
        Else
            MsgBox("Sourc: " & Source & " is not found")
        End If
    End Sub
    Private Sub txtProjectValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtProjectValue_Validated(ByVal sender As Object, ByVal e As System.EventArgs)
        mProjectvalue = Val(Me.txtProjectValue.Text)
    End Sub

    Private Sub txtProjectValue_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        If Not IsNumeric(Me.txtProjectValue.Text) Then
            MsgBox("Project value should be a numeric value")
            Me.txtProjectValue.Focus()
            e.Cancel = True
        End If
    End Sub

    Private Sub dpStartDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        If Not IsDate(Me.dpStartDate.Value) Then
            MsgBox("Please Select a Date")
            Me.dpEndDate.Focus()
            e.Cancel = True
        End If
    End Sub

    Private Sub dpStartDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub dpEndDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        If Not IsDate(Me.dpEndDate.Value) Then
            MsgBox("Please Select a Date")
            Me.dpEndDate.Focus()
            e.Cancel = True
        End If
    End Sub

    Private Sub frmProjectDetails_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        lblMessage.Visible = False
    End Sub

    Private Sub frmProjectDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        apppath = Application.StartupPath
        strConnection = ""
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & apppath & "\EquipmentsMaster.mdb"
        moledbConnection = New OleDbConnection(strConnection)
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strStatement = "Select * from Projects"
        Dim mAdapter As OleDbDataAdapter
        Dim mdataset As DataSet = New DataSet
        mAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
        mAdapter.Fill(mdataset, "Projects")
        Dim mrow As DataRow
        Me.cmbProjects.Items.Clear()
        Me.cmbProjects.Items.Add(".New Project")
        If mdataset.Tables("Projects").Rows.Count > 0 Then
            For Each mrow In mdataset.Tables("Projects").Rows
                cmbProjects.Items.Add(mrow("ProjectTitle"))
            Next
        End If
        Me.cmbProjects.SelectedIndex = 0
        moledbCommand = Nothing
        frmprojectdetailsFirstTime = True
        Me.lblMessage.Visible = False
    End Sub
    Private Sub clearfields()
        Me.txtClient.Text = ""
        Me.txtLocation.Text = ""
        Me.txtProjectCode.Text = ""
        Me.txtProjectdescription.Text = ""
        Me.txtProjectValue.Text = 0
        Me.txtWorkbookName.Text = ""
        Me.dpStartDate.Value = Today()
        Me.dpEndDate.Value = Today()
    End Sub
    Private Sub cmbProjects_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbProjects.SelectedIndexChanged

    End Sub

    Private Sub btnBrowse_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        Me.SaveFileDialog1.ShowDialog()
    End Sub

    Private Sub cmbProjects_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbProjects.Validated
        Dim oForm As New Form
        If UCase(Me.cmbProjects.Text) = UCase(".New Project") Then
            clearfields()
            mWorkingMode = "New"
            Me.pnlDetails.Enabled = True
        Else
            Me.pnlDetails.Enabled = False
            If (moledbConnection.State.ToString().Equals("Closed")) Then
                moledbConnection.Open()
            End If
            strStatement = "Select *  from Projects where ProjectTitle = '" & Me.cmbProjects.Text & "'"
            mOledbDataAdapter = New OleDbDataAdapter(strStatement, moledbConnection)
            mDataSet = New DataSet()
            mOledbDataAdapter.Fill(mDataSet, "ExistingProjects")
            oForm = New fmgetpassword()
            oForm.ShowDialog()
            oForm = Nothing
            If mDataSet Is Nothing Or mDataSet.Tables("ExistingProjects").Rows.Count = 0 Then
                MsgBox("No entry found for project title " & cmbProjects.Text)
                Exit Sub
            End If
            For Each mRow In mDataSet.Tables("ExistingProjects").Rows
                If String.Compare(mPassword, mRow("Password").ToString(), False) <> 0 Then
                    MsgBox("Invalid Password. Cannot open the project file")
                    Exit Sub
                End If
            Next
            Me.txtProjectdescription.Text = mRow("ProjectTitle").ToString()
            Me.txtLocation.Text = mRow("Location").ToString()
            Me.txtClient.Text = mRow("Client").ToString()
            Me.dpStartDate.Value = mRow("StartDate").ToString()
            Me.dpEndDate.Value = mRow("EndDate").ToString()
            Me.txtProjectValue.Text = mRow("ProjectValue").ToString()
            Me.txtWorkbookName.Text = mRow("ProjectFileName").ToString()
            Me.txtProjectCode.Text = mRow("ProjectCode").ToString()
            Me.txtProjectCode.Enabled = False
            Me.txtConcreteQuantity.Text = mRow("ConcreteQty")
            mConcreteQty = Me.txtConcreteQuantity.Text
            mdbfilename = mRow("AddItemsMDB").ToString()
            Me.txtFuelCostPerLtr.Text = mRow("FuelCostPerLtr")
            Me.txtPowerCostPerUnit.Text = mRow("PowerCostPerUnit")
            mWorkingMode = "Edit"
            Me.pnlDetails.Enabled = False
        End If
        mDataSet = Nothing
        mOledbDataAdapter = Nothing
    End Sub

    'Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    'If Not FormLoaded Then
    '    '    If ProgressBar1.Value = ProgressBar1.Maximum Then ProgressBar1.Value = ProgressBar1.Minimum
    '    '    ProgressBar1.Value = ProgressBar1.Value + ProgressBar1.Step
    '    '    'If ProgressBar1.Value > ProgressBar1.Maximum Then ProgressBar1.Value = ProgressBar1.Minimum
    '    '    ProgressBar1.Refresh()
    '    'Else
    '    '    ProgressBar1.Visible = False
    '    '    Timer1.Stop()
    '    'End If
    'End Sub

    'Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
    '    ProgressBar1.Value = ProgressBar1.Value + 50
    '    ProgressBar1.Refresh()
    'End Sub

    'Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
    '    ProgressBar1.Visible = False
    'End Sub

    'Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
    '    oForm = New frmOptions()
    '    oForm.ShowDialog()
    'End Sub

    Private Sub BackgroundWorker1_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles BackgroundWorker1.Disposed

    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Dim msgtext As String
        'msgtext = "Form Details are loading. Please wait..."
        'If lblMessage.Visible Then
        '    If Len(Trim(lblMessage.Text)) = 0 Then
        '        lblMessage.Text = msgtext
        '    Else
        '        lblMessage.Text = ""
        '    End If
        '    Me.Refresh()
        'End If
    End Sub

    'Private Sub frmProjectDetails_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
    'Object, ByVal e As System.EventArgs) Handles lblMessage.Click

    'End Sub

    'Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click

    'End Sub
    '    lblMessage.Visible = False
    'End Sub

    'Private Sub lblMessage_Click(ByVal sender As System.
    'Private Sub pnlDetails_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlDetails.Paint

    'End Sub

    'Private Sub FuelCostPerLtr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFuelCostPerLtr.TextChanged

    'End Sub

    'Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click

    'End Sub

    Private Sub pnlDetails_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlDetails.Paint

    End Sub
End Class

Imports System.Data
Imports System.Data.OleDb
Imports System.IO.FileInfo


Module Module1
    'public variables and control array variable declaration
    Public xlApp As New Microsoft.Office.Interop.Excel.Application
    'Public xlApp As New Microsoft.Office.Interop.Excel.Application
    Public xlWorkbook As Microsoft.Office.Interop.Excel.Workbook
    Public xlWorksheet As Microsoft.Office.Interop.Excel.Worksheet
    Public xlRange As Microsoft.Office.Interop.Excel.Range
    Public strConnection As String, strconnection1 As String
    Public moledbConnection As OleDbConnection, moledbConnection1 As OleDbConnection
    Public mRow As DataRow
    Public SelectedCategory As String
    Public xlFilename As String
    Public mdbfilename As String
    Public mMainTitle1 As String
    Public mMainTitle2 As String
    Public mMainTitle3 As String
    Public mClient As String
    Public mLocation As String
    Public mStartDate As Date
    Public mEndDate As Date
    Public records_Majorequipments As Integer = 0
    Public records_MinorEquipments As Integer = 0
    Public records_Hiredequipments As Integer = 0
    Public records_electricalEquipments As Integer = 0
    Public Category_Shortname As String
    Public PrevSheetname As String
    Public RecordsInserted(20) As Integer
    Public SheetNo As Integer
    Public Startup As Boolean = True
    Public Sheetnames(20) As String, SheetIndices(20) As Integer, SheetsCount As Integer
    Public TemplatePath As String, Currentfile As String
    Public x1 As String, x2 As String
    Public Projectmonths As Integer
    Public DestinationFolder As String
    Public appPath As String = ""
    Public intX As Integer = Screen.PrimaryScreen.Bounds.Width
    Public intY As Integer = Screen.PrimaryScreen.Bounds.Height
    Public mProjectvalue As Long
    Public mConcreteQty As Double
    Public mWorkingMode As String
    Public mPassword As String
    Public mMinMaintperc As Single, mMaxMaintPerc As Single
    Public mcategory As String, Tablename As String
    Public FormLoaded As Boolean
    Public EditOperationSheet As String
    Public mDescription, mCapacity, mMake, mModel As String
    Public mMobDate, mDeMobDate As Date
    Public mQty As Integer, LightingItems As Integer
    Public mMonths As Integer
    Public BlankForm As Boolean = True
    Public frmprojectdetailsFirstTime As Boolean
    Public frmOptionsFirstTime As Boolean
    Public EditOrDelete As String
    Public Currenttab As Integer = 0
    Public NomoreDelete As Boolean = False
    Public HP As Integer = 10, VP As Integer = 20
    Public FuelCostperLtr As Single
    Public PowerCostPerUnit As Single
    Public txtmake As String, txtmodel As String
    Public ConcretingItems As Integer, CraneItems As Integer, MHItems As Integer
    Public DGSetItems As Integer, ConveyanceItems As Integer, NCItems As Integer, MajorOtherItems As Integer
    Public MinorItems As Integer, HireItems As Integer, fexpItems As Integer, BPFExpItems As Integer
    Public Counts As Integer = 0, Prevcategory As String, intj As Integer, RepeatFactor As Integer
    Public concretecheckeditems As Integer = 0, conveyancecheckeditems As Integer = 0, cranecheckeditems As Integer = 0
    Public mhcheckeditems As Integer = 0, nccheckeditems As Integer = 0, majorotherscheckeditems As Integer = 0
    Public dgsetscheckeditems As Integer = 0, MinorCheckedItems As Integer = 0, HiredCheckedItems As Integer = 0
    Public fexpCheckedItems As Integer, BPFexpCheckedItems As Integer, LightingCheckedItems

    'Declarin contol arrays
    Public Buttons As New Dictionary(Of String, Button)
    Public CategoryTextBoxes As New Dictionary(Of String, TextBox)
    Public EquipNameTextBoxes As New Dictionary(Of String, TextBox)
    Public CapacityTextBoxes As New Dictionary(Of String, TextBox)
    Public MakeModelTextBoxes As New Dictionary(Of String, TextBox)
    Public QtyTextBoxes As New Dictionary(Of String, TextBox)
    Public HrsPermonthTextBoxes As New Dictionary(Of String, TextBox)
    Public PurchvalTextBoxes As New Dictionary(Of String, TextBox)
    Public RepValueTextBoxes As New Dictionary(Of String, TextBox)
    Public MaintPercTextBoxes As New Dictionary(Of String, TextBox)
    Public concreteqtyTextboxes As New Dictionary(Of String, TextBox)
    Public Checkboxes As New Dictionary(Of Integer, CheckBox)
    Public IsNewMC As New Dictionary(Of Integer, CheckBox)
    Public AddButtons As New Dictionary(Of String, Button)
    Public HireChargesTextBoxes As New Dictionary(Of String, TextBox)
    Public CostTextBoxes As New Dictionary(Of String, TextBox)
    Public AmountTextBoxes As New Dictionary(Of String, TextBox)
    Public ClientBillingTextBoxes As New Dictionary(Of String, TextBox)
    Public CostPercTextBoxes As New Dictionary(Of String, TextBox)
    Public RemarksTextBoxes As New Dictionary(Of String, TextBox)
    Public PowerPerUnitTextBoxes As New Dictionary(Of Single, TextBox)
    Public ConnectLoadTextBoxes As New Dictionary(Of Single, TextBox)
    Public UtilityFactorTextBoxes As New Dictionary(Of Single, TextBox)



    Public MobdatePickers As New Dictionary(Of Integer, DateTimePicker)
    Public DemobDatePickers As New Dictionary(Of Integer, DateTimePicker)

    Public DepPercComboboxes As New Dictionary(Of Integer, ComboBox)
    Public ShiftsComboboxes As New Dictionary(Of Integer, ComboBox)
    'End of Control Arrays declaration

    Public ConcreteChecked() As Integer, ConvChecked() As Integer, CraneChecked() As Integer
    Public DGSetsChecked() As Integer, MHChecked() As Integer, nccHECKED() As Integer, MajOthersChecked() As Integer
    Public MinorChecked() As Integer, HiredChecked() As Integer, FexpChecked() As Integer, BPFExpChecked() As Integer
    Public LightingChecked() As Integer
    Public mTabindex As Integer

    Public concFirsttime As Boolean = True
    Public convFirsttime As Boolean = True
    Public cranFirstTime As Boolean = True, mathFirstTime As Boolean = True
    Public dgsetFirstTime As Boolean = True
    Public noncFirstTime As Boolean = True, majoFirstTime As Boolean = True
    Public MinorFirstTime As Boolean = True, HiredFirstTime As Boolean = True
    Public fexpFirstTime As Boolean = True, BPFExpFirstTime As Boolean = True
    Public LightingFirstTime As Boolean = True

    Public concreteitems As Integer ', conveyanceitems As Integer, craneitems As Integer, dgsetitems As Integer
    'Public mhitems As Integer, ncitems As Integer


    Public ConcreteMobdate() As Date, ConcreteDemobdate() As Date, ConcreteQty() As Integer, ConcreteHrs() As Integer
    Public ConcreteDepPerc() As Single, ConcreteShifts() As Single

    Public ConvMobdate() As Date, ConvDemobdate() As Date, ConvQty() As Integer, ConvHrs() As Integer
    Public ConvDepPerc() As Single, ConvShifts() As Single

    Public CraneMobdate() As Date, CraneDemobdate() As Date, CraneQty() As Integer, CraneHrs() As Integer
    Public CraneDepPerc() As Single, CraneShifts() As Single

    Public DGSetsMobdate() As Date, DGSetsDemobdate() As Date, DGSetsQty() As Integer, DGSetsHrs() As Integer
    Public DGSetsDepPerc() As Single, DGSetsShifts() As Single

    Public MHMobdate() As Date, MHDemobdate() As Date, MHQty() As Integer, MHHrs() As Integer
    Public MHDepPerc() As Single, MHShifts() As Single

    Public NCMobdate() As Date, NCDemobdate() As Date, NCQty() As Integer, NCHrs() As Integer
    Public NCDepPerc() As Single, NCShifts() As Single

    Public MajOthersMobdate() As Date, MajOthersDemobdate() As Date, MajOthersQty() As Integer, MajOthersHrs() As Integer
    Public MajOthersDepPerc() As Single, MajOthersShifts() As Single

    Public MinorMobdate() As Date, MinorDemobdate() As Date, MinorQty() As Integer, MinorHrs() As Integer
    Public MinorDepPerc() As Single, MinorShifts() As Single, MinorNPV() As Long

    Public HiredMobdate() As Date, HiredDemobdate() As Date, HiredQty() As Integer, HiredHrs() As Integer
    Public HiredDepPerc() As Single, HiredShifts() As Single, HiredNPV() As Long

    Public FExpQty() As Integer, FExpRemarks() As String
    Public BPFExpQty() As Integer, BPFExpRemarks() As String

    Public LightingMobDate() As Date, LightingDemobDate() As Date, LightingQty() As Integer
    Public LightingPowerPerUnit() As Single, LightingConnectLoad() As Single, LightingUtilityFactor() As Single

    Public concCategory() As String, concEName() As String, concCapacity() As String, concMakeModel() As String
    Public concMDate() As Date, concDMdate() As Date, concQuantity() As Integer, concHrsPerMonth() As Integer
    Public concDep() As Single, concShift() As Single

    Public convCategory() As String, convEName() As String, convCapacity() As String, convMakeModel() As String
    Public convMDate() As Date, convDMdate() As Date, convQuantity() As Integer, convHrsPerMonth() As Integer
    Public convDep() As Single, convShift() As Single

    Public cranCategory() As String, cranEName() As String, cranCapacity() As String, cranMakeModel() As String
    Public cranMDate() As Date, cranDMdate() As Date, cranQuantity() As Integer, cranHrsPerMonth() As Integer
    Public cranDep() As Single, cranShift() As Single

    Public dgsetCategory() As String, dgsetEName() As String, dgsetCapacity() As String, dgsetMakeModel() As String
    Public dgsetMDate() As Date, dgsetDMdate() As Date, dgsetQuantity() As Integer, dgsetHrsPerMonth() As Integer
    Public dgsetDep() As Single, dgsetShift() As Single

    Public mathCategory() As String, mathEName() As String, mathCapacity() As String, mathMakeModel() As String
    Public mathMDate() As Date, mathDMdate() As Date, mathQuantity() As Integer, mathHrsPerMonth() As Integer
    Public mathDep() As Single, mathShift() As Single

    Public noncCategory() As String, noncEName() As String, noncCapacity() As String, noncMakeModel() As String
    Public noncMDate() As Date, noncDMdate() As Date, noncQuantity() As Integer, noncHrsPerMonth() As Integer
    Public noncDep() As Single, noncShift() As Single

    Public majoCategory() As String, majoEName() As String, majoCapacity() As String, majoMakeModel() As String
    Public majoMDate() As Date, majoDMdate() As Date, majoQuantity() As Integer, majoHrsPerMonth() As Integer
    Public majoDep() As Single, majoShift() As Single

    Public MinorCategory() As String, MinorEName() As String, MinorCapacity() As String, MinorMakeModel() As String
    Public MinorMDate() As Date, MinorDMdate() As Date, MinorQuantity() As Integer, MinorHrsPerMonth() As Integer
    Public MinorDep() As Single, MinorShift() As Single, MinorNewPurchVal() As Long, MinorIsNew() As Boolean

    Public HiredCategory() As String, HiredEName() As String, HiredCapacity() As String, HiredMakeModel() As String
    Public HiredMDate() As Date, HiredDMdate() As Date, HiredQuantity() As Integer, HiredHrsPerMonth() As Integer
    Public HiredDep() As Single, HiredShift() As Single, HiredHireCharges() As Long

    Public fexpCategory() As String, fexpCost() As Long, fexpAmount() As Double
    Public fexpClientBilling() As Double, fexpCostPerc() As Single

    Public BPfexpCategory() As String, BPfexpCost() As Long, BPfexpAmount() As Double
    Public BPfexpClientBilling() As Double, BPfexpCostPerc() As Single

    Public LightEName() As String, LightCapacity() As String, LightMakeModel() As String, LightMDate() As Date
    Public LightDMDate() As Date, LightQuantity() As Integer, LightConnectLoad() As Single
    Public LightUtilityFactor() As Single, LightPowerPerUnit() As Single

    Public NewMCCost As Long, MinorEquipmentCost As Long
    Public RAndMPercentage As Single

    Public Ext_Conv_HirechargesTotal As Double = 0, Ext_Conv_FuelPermonthTotal As Double = 0, Ext_Conv_FuelProjectTotal As Double = 0

    Public concEquipsNames() As String, concEquipsCapacity() As String, concEquipsMake() As String, concEquipsModel() As String, concEquipsMobDate() As Date
    Public concEquipsDemobDate() As Date, concEquipsQty() As Integer, concEquipsChkd() As Integer, concEquipsHPM() As Integer, concEquipsDepPerc() As Single
    Public concEquipsRepValue() As Long, concEquipsShifts() As Single, concEquipsMaintPerc() As Single, concEquipsConcQty() As Long, concEquipsDrive() As String
    Public concEquipsPPU() As Single, concEquipsCLPerMc() As Single, concEquipsUF() As Single

    Public convEquipsNames() As String, convEquipsCapacity() As String, convEquipsMake() As String, convEquipsModel() As String, convEquipsMobDate() As String
    Public convEquipsDemobDate() As String, convEquipsQty() As Integer, convEquipsChkd() As Integer, convEquipsHPM() As Integer, convEquipsDepPerc() As Single
    Public convEquipsRepValue() As Long, convEquipsShifts() As Single, convEquipsMaintPerc() As Single, convEquipsConcQty() As Long, convEquipsDrive() As String
    Public convEquipsPPU() As Single, convEquipsCLPerMc() As Single, convEquipsUF() As Single

    Public craneEquipsNames() As String, craneEquipsCapacity() As String, craneEquipsMake() As String, craneEquipsModel() As String, craneEquipsMobDate() As String
    Public craneEquipsDemobDate() As String, craneEquipsQty() As Integer, craneEquipsChkd() As Integer, craneEquipsHPM() As Integer, craneEquipsDepPerc() As Single
    Public craneEquipsRepValue() As Long, craneEquipsShifts() As Single, craneEquipsMaintPerc() As Single, craneEquipsConcQty() As Long, craneEquipsDrive() As String
    Public craneEquipsPPU() As Single, craneEquipsCLPerMc() As Single, craneEquipsUF() As Single

    Public dgsetsEquipsNames() As String, dgsetsEquipsCapacity() As String, dgsetsEquipsMake() As String, dgsetsEquipsModel() As String, dgsetsEquipsMobDate() As String
    Public dgsetsEquipsDemobDate() As String, dgsetsEquipsQty() As Integer, dgsetsEquipsChkd() As Integer, dgsetsEquipsHPM() As Integer, dgsetsEquipsDepPerc() As Single
    Public dgsetsEquipsRepValue() As Long, dgsetsEquipsShifts() As Single, dgsetsEquipsMaintPerc() As Single, dgsetsEquipsConcQty() As Long, dgsetsEquipsDrive() As String
    Public dgsetsEquipsPPU() As Single, dgsetsEquipsCLPerMc() As Single, dgsetsEquipsUF() As Single

    Public MHEquipsNames() As String, MHEquipsCapacity() As String, MHEquipsMake() As String, MHEquipsModel() As String, MHEquipsMobDate() As String
    Public MHEquipsDemobDate() As String, MHEquipsQty() As Integer, MHEquipsChkd() As Integer, MHEquipsHPM() As Integer, MHEquipsDepPerc() As Single
    Public MHEquipsRepValue() As Long, MHEquipsShifts() As Single, MHEquipsMaintPerc() As Single, MHEquipsConcQty() As Long, MHEquipsDrive() As String
    Public MHEquipsPPU() As Single, MHEquipsCLPerMc() As Single, MHEquipsUF() As Single

    Public NCEquipsNames() As String, NCEquipsCapacity() As String, NCEquipsMake() As String, NCEquipsModel() As String, NCEquipsMobDate() As String
    Public NCEquipsDemobDate() As String, NCEquipsQty() As Integer, NCEquipsChkd() As Integer, NCEquipsHPM() As Integer, NCEquipsDepPerc() As Single
    Public NCEquipsRepValue() As Long, NCEquipsShifts() As Single, NCEquipsMaintPerc() As Single, NCEquipsConcQty() As Long, NCEquipsDrive() As String
    Public NCEquipsPPU() As Single, NCEquipsCLPerMc() As Single, NCEquipsUF() As Single

    Public majOthersEquipsNames() As String, majOthersEquipsCapacity() As String, majOthersEquipsMake() As String, majOthersEquipsModel() As String, majOthersEquipsMobDate() As String
    Public majOthersEquipsDemobDate() As String, majOthersEquipsQty() As Integer, majOthersEquipsChkd() As Integer, majOthersEquipsHPM() As Integer, majOthersEquipsDepPerc() As Single
    Public majOthersEquipsRepValue() As Long, majOthersEquipsShifts() As Single, majOthersEquipsMaintPerc() As Single, majOthersEquipsConcQty() As Long, majOthersEquipsDrive() As String
    Public majOthersEquipsPPU() As Single, majOthersEquipsCLPerMc() As Single, majOthersEquipsUF() As Single

    Public minorEquipsNames() As String, minorEquipsCapacity() As String, minorEquipsMake() As String, minorEquipsModel() As String, minorEquipsMobDate() As String
    Public minorEquipsDemobDate() As String, minorEquipsQty() As Integer, minorEquipsChkd() As Integer, minorEquipsHPM() As Integer, minorEquipsDepPerc() As Single
    Public minorEquipsNewCost() As Long, minorEquipsShifts() As Single, minorIsNewMC() As Boolean, minorEquipsDrive() As String
    Public minorEquipsPPU() As Single, minorEquipsCLPerMc() As Single, minorEquipsUF() As Single

    Public hiredCategoryNames() As String, hiredEquipsNames() As String, hiredEquipsCapacity() As String, hiredEquipsMake() As String, hiredEquipsModel() As String, hiredEquipsMobDate() As String
    Public hiredEquipsDemobDate() As String, hiredEquipsQty() As Integer, hiredEquipsChkd() As Integer, hiredEquipsHPM() As Integer, hiredEquipsDepPerc() As Single
    Public hiredEquipsHireCharges() As Long, hiredEquipsShifts() As Single

    Public fixedExpCategoryNames() As String, fixedExpEquipsQty() As Integer, fixedExpCost() As Long, fixedExpAmount() As Double, fixedExpRemarks() As String
    Public fixedExpEquipsChkd() As Integer, fixedExpProjValue() As Double, fixedExpEquipsCostPerc() As Single

    Public fixedBPExpCategoryNames() As String, fixedBPExpEquipsQty() As Integer, fixedBPExpCost() As Long, fixedBPExpAmount() As Double, fixedBPExpRemarks() As String
    Public fixedBPExpEquipsChkd() As Integer, fixedBPExpProjValue() As Double, fixedBPExpEquipsCostPerc() As Single

    Public lightingCategoryNames() As String, lightingEquipsNames() As String, lightingEquipsCapacity() As String, lightingEquipsMake() As String
    Public lightingEquipsModel() As String, lightingEquipsMobDate() As String
    Public lightingEquipsDemobDate() As String, lightingEquipsQty() As Integer, lightingEquipsChkd() As Integer, lightingEquipsPPU() As Single
    Public lightingEquipsCLPerMc() As Single, lightingEquipsUF() As Single

    Public DataSaved As Boolean, answer As Integer
    Public starttime As Date, endtime As Date
    'end of public variables and control array variable declaration


    Public Sub getCategoryShortname(ByVal xlsheet As Microsoft.Office.Interop.Excel.Worksheet)
        If UCase(Trim(xlsheet.Name)) = UCase("Concreting") Then
            Category_Shortname = "Concrete_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Conveyance") Then
            Category_Shortname = "Conv_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Cranes") Then
            Category_Shortname = "Cranes_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Material Handling") Then
            Category_Shortname = "MH_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Non concreting") Then
            Category_Shortname = "NC_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("DG Sets") Then
            Category_Shortname = "DG_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Major Others") Then
            Category_Shortname = "MajOthers_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Minor Eqpts") Then
            Category_Shortname = "Min_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("external Hire") Then
            Category_Shortname = "Ext_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("External Others") Then
            Category_Shortname = "ExtOthers_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Electrical") Then
            Category_Shortname = "Elec_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Top Sheet") Then
            Category_Shortname = "Out_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Power reqmnt Estimated Units") Then
            Category_Shortname = "PowerReq_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Load distribution") Then
            Category_Shortname = "LD_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Misc and Non ERP Purchases") Then
            Category_Shortname = "Misc_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Elect salary") Then
            Category_Shortname = "Esal_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Tr crane related exp") Then
            Category_Shortname = "CraneExp_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Bplant related exp") Then
            Category_Shortname = "BPlantExp_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Extra pipeline") Then
            Category_Shortname = "ExPipe_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("Staff Salary") Then
            Category_Shortname = "StaffSalary_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("PowerReqmt") Then
            Category_Shortname = "PowerReq_"
        ElseIf UCase(Trim(xlsheet.Name)) = UCase("PowerGen Cost") Then
            Category_Shortname = "Powergen_"
        End If
    End Sub
    Public Sub CopyRecordFromMasterToAddedItems(ByVal macategory As String, ByVal mtablename As String)
        Dim strSql As String, mOledbDataAdapter3 As OleDbDataAdapter
        Dim InsertCommand As String, Machine As DataRow
        Dim moleDbInsertComamnd As OleDbCommand
        Dim moledbDataSet3 As New DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
        If moledbConnection Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        If moledbConnection1 Is Nothing Then
            strconnection1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbf
            ilename()
            moledbConnection1 = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection1.Open()
        End If

        If (macategory = "Concreting" Or macategory = "Conveyance" Or macategory = "Cranes" Or macategory = "DG Sets" Or macategory = "Material Handling" Or _
           macategory = "Non Concreting" Or macategory = "Major Others") Then
            strSql = "Select * from MajorEquipments where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                InsertCommand = ""
                InsertCommand = "INSERT INTO " & mtablename & " Values ("
                InsertCommand = InsertCommand & "'" & Machine("EquipmentName").ToString() & "', "
                InsertCommand = InsertCommand & "'" & Machine("Capacity").ToString & "', "
                InsertCommand = InsertCommand & "'" & Machine("Make").ToString() & "', "
                InsertCommand = InsertCommand & "'" & Machine("Model").ToString() & "', "
                InsertCommand = InsertCommand & "'" & mStartDate.Date.ToString() & "', "
                InsertCommand = InsertCommand & "'" & mEndDate.Date.ToString() & "', "
                InsertCommand = InsertCommand & mqty & ", "
                InsertCommand = InsertCommand & chkd & ", "
                InsertCommand = InsertCommand & Val(Machine("Hrs_PerMonth").ToString()) & ", "
                InsertCommand = InsertCommand & DepPerc & ", "
                InsertCommand = InsertCommand & Val(Machine("RepValue").ToString()) & ", "
                InsertCommand = InsertCommand & mShifts & ", "
                InsertCommand = InsertCommand & mqty & ", "
                InsertCommand = InsertCommand & mConcreteQty & ", "
                InsertCommand = InsertCommand & "'" & Machine("Drive").ToString() & "', "
                InsertCommand = InsertCommand & Val(Machine("PowerPerUnit(HP)").ToString()) & ", "
                InsertCommand = InsertCommand & Val(Machine("ConnectedLoadPerMC").ToString()) & ", "
                InsertCommand = InsertCommand & Val(Machine("UtilityFactor").ToString()) & ")"
                'MsgBox(InsertCommand)

                Try
                    If (moledbConnection1.State.ToString().Equals("Closed")) Then
                        moledbConnection1.Open()
                    End If
                    moleDbInsertComamnd = New OleDbCommand
                    moleDbInsertComamnd.CommandType = CommandType.Text
                    moleDbInsertComamnd.CommandText = InsertCommand
                    moleDbInsertComamnd.Connection = moledbConnection1
                    moleDbInsertComamnd.ExecuteNonQuery()
                    'MsgBox(Me.txtProjectdescription.Text & "Record inserted in list of projects.")
                Catch ex As Exception
                    MsgBox(ex.ToString())
                Finally
                    moleDbInsertComamnd = Nothing
                    'moledbconnection3.Close()
                End Try
            Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        ElseIf macategory = "Minor Equipments" Then
            strSql = "Select * from MinorEquipments where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                InsertCommand = ""
                InsertCommand = "INSERT INTO " & mtablename & " Values ("
                InsertCommand = InsertCommand & "'" & Machine("EquipmentName").ToString() & "', "
                InsertCommand = InsertCommand & "'" & Machine("Capacity").ToString & "', "
                InsertCommand = InsertCommand & "'" & Machine("Make").ToString() & "', "
                InsertCommand = InsertCommand & "'" & Machine("Model").ToString() & "', "
                InsertCommand = InsertCommand & "'" & mStartDate.Date.ToString() & "', "
                InsertCommand = InsertCommand & "'" & mEndDate.Date.ToString() & "', "
                InsertCommand = InsertCommand & mqty & ", "
                InsertCommand = InsertCommand & chkd & ", "
                InsertCommand = InsertCommand & Val(Machine("Hrs_PerMonth").ToString()) & ", "
                InsertCommand = InsertCommand & DepPerc & ", "
                InsertCommand = InsertCommand & mShifts & ", "
                InsertCommand = InsertCommand & Val(Machine("CostOfnewEquipment").ToString()) & ", "
                InsertCommand = InsertCommand & ISNewMC & ", "
                InsertCommand = InsertCommand & "'" & Machine("Drive").ToString() & "', "
                InsertCommand = InsertCommand & Val(Machine("PowerPerUnit(HP)").ToString()) & ", "
                InsertCommand = InsertCommand & Val(Machine("ConnectedLoadPerMC").ToString()) & ", "
                InsertCommand = InsertCommand & Val(Machine("UtilityFactor").ToString()) & ")"

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
                End Try
            Next
        ElseIf (macategory = "HiredConveyance" Or macategory = "Excav / Earthwork" Or macategory = "GenSets" Or macategory = "Matl Handling" Or _
            macategory = "Matl Transport" Or macategory = "Others") Then
            strSql = "Select * from HiredEquipments where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                InsertCommand = ""
                InsertCommand = "INSERT INTO " & mtablename & " Values ("
                InsertCommand = InsertCommand & "'" & Machine("Categoryname").ToString() & "', "
                InsertCommand = InsertCommand & "'" & Machine("EquipmentName").ToString() & "', "
                InsertCommand = InsertCommand & "'" & Machine("Capacity").ToString & "', "
                InsertCommand = InsertCommand & "'" & Machine("Make").ToString() & "', "
                InsertCommand = InsertCommand & "'" & Machine("Model").ToString() & "', "
                InsertCommand = InsertCommand & "'" & mStartDate.Date.ToString() & "', "
                InsertCommand = InsertCommand & "'" & mEndDate.Date.ToString() & "', "
                InsertCommand = InsertCommand & mqty & ", "
                InsertCommand = InsertCommand & chkd & ", "
                InsertCommand = InsertCommand & Val(Machine("Hrs_PerMonth").ToString()) & ", "
                InsertCommand = InsertCommand & DepPerc & ", "
                InsertCommand = InsertCommand & mShifts & ", "
                InsertCommand = InsertCommand & Val(Machine("HireCharges").ToString()) & ")"

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
                End Try
            Next
        ElseIf macategory = "FixedExpenses" Then
            strSql = "Select * from FixedExpenses"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                InsertCommand = ""
                InsertCommand = "INSERT INTO " & mtablename & " Values ("
                InsertCommand = InsertCommand & "'" & Machine("Category").ToString() & "', "
                InsertCommand = InsertCommand & mqty & ","
                InsertCommand = InsertCommand & Val(Machine("Cost").ToString) & ", "
                InsertCommand = InsertCommand & Val(Machine("Cost").ToString) & ", "
                InsertCommand = InsertCommand & "'" & Machine("Remarks").ToString() & "', "
                InsertCommand = InsertCommand & chkd & ", "
                InsertCommand = InsertCommand & mProjectvalue & ", "
                InsertCommand = InsertCommand & Val(Machine("Cost").ToString) / mProjectvalue & ")"

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
                End Try
            Next
        ElseIf macategory = "BPFixedExpenses" Then
            strSql = "Select * from BPFixedExpenses"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                InsertCommand = ""
                InsertCommand = "INSERT INTO " & mtablename & " Values ("
                InsertCommand = InsertCommand & "'" & Machine("Category").ToString() & "', "
                InsertCommand = InsertCommand & mqty & ","
                InsertCommand = InsertCommand & Val(Machine("Cost").ToString) & ", "
                InsertCommand = InsertCommand & Val(Machine("Cost").ToString) & ", "
                InsertCommand = InsertCommand & "'" & Machine("Remarks").ToString() & "', "
                InsertCommand = InsertCommand & chkd & ", "
                InsertCommand = InsertCommand & mProjectvalue & ", "
                InsertCommand = InsertCommand & Val(Machine("Cost").ToString) / mProjectvalue & ")"

                Try
                    If (moledbConnection1.State.ToString().Equals("Closed")) Then
                        moledbConnection1.Open()
                    End If
                    moleDbInsertComamnd = New OleDbCommand
                    moleDbInsertComamnd.CommandType = CommandType.Text
                    moleDbInsertComamnd.CommandText = InsertCommand
                    moleDbInsertComamnd.Connection = moledbConnection1
                    moleDbInsertComamnd.ExecuteNonQuery()
                    'MsgBox(Me.txtProjectdescription.Text & "Record inserted in list of projects.")
                Catch ex As Exception
                    MsgBox(ex.ToString())
                Finally
                    moleDbInsertComamnd = Nothing
                    'moledbconnection3.Close()
                End Try
            Next
        ElseIf macategory = "Lighting" Then
            strSql = "Select * from LightingEquipments"   ' where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                InsertCommand = ""
                InsertCommand = "INSERT INTO " & mtablename & " Values ("
                InsertCommand = InsertCommand & "'" & Machine("Categoryname").ToString() & "', "
                InsertCommand = InsertCommand & "'" & Machine("EquipmentName").ToString() & "', "
                InsertCommand = InsertCommand & "'" & Machine("Capacity").ToString & "', "
                InsertCommand = InsertCommand & "'" & Machine("Make").ToString() & "', "
                InsertCommand = InsertCommand & "'" & Machine("Model").ToString() & "', "
                InsertCommand = InsertCommand & "'" & mStartDate.Date.ToString() & "', "
                InsertCommand = InsertCommand & "'" & mEndDate.Date.ToString() & "', "
                InsertCommand = InsertCommand & mqty & ", "
                InsertCommand = InsertCommand & chkd & ", "
                InsertCommand = InsertCommand & Val(Machine("PowerPerUnit").ToString()) & ", "
                InsertCommand = InsertCommand & Val(Machine("ConnectedLoad").ToString()) & ", "
                InsertCommand = InsertCommand & Val(Machine("UtilityFactor").ToString()) & ")"

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
                End Try
            Next
        End If
    End Sub
    Public Sub CopyRecordFromMasterToAddedItemsArray(ByVal macategory As String, ByVal mtablename As String)
        Dim strSql As String, mOledbDataAdapter3 As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet3 As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & "\EquipmentsMaster.mdb"
        If moledbConnection Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        If (macategory = "Concreting") Then
            moledbDataSet3 = New DataSet
            strSql = "Select * from MajorEquipments where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            counts = moledbDataSet3.Tables("MasterEquipments").Rows.Count - 1
            ReDim concEquipsNames(counts)
            ReDim concEquipsCapacity(counts)
            ReDim concEquipsMake(counts)
            ReDim concEquipsModel(counts)
            ReDim concEquipsMobDate(counts)
            ReDim concEquipsDemobDate(counts)
            ReDim concEquipsQty(counts)
            ReDim concEquipsChkd(counts)
            ReDim concEquipsHPM(counts)
            ReDim concEquipsDepPerc(counts)
            ReDim concEquipsRepValue(counts)
            ReDim concEquipsShifts(counts)
            ReDim concEquipsMaintPerc(counts)
            ReDim concEquipsConcQty(counts)
            ReDim concEquipsDrive(counts)
            ReDim concEquipsPPU(counts)
            ReDim concEquipsCLPerMc(counts)
            ReDim concEquipsUF(counts)

            pointer = 0
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                concEquipsNames(pointer) = Machine("EquipmentName").ToString()
                concEquipsCapacity(pointer) = Machine("Capacity").ToString
                concEquipsMake(pointer) = Machine("Make").ToString()
                concEquipsModel(pointer) = Machine("Model").ToString()
                concEquipsMobDate(pointer) = mStartDate.Date.ToString()
                concEquipsDemobDate(pointer) = mEndDate.Date.ToString()
                concEquipsQty(pointer) = mqty
                concEquipsChkd(pointer) = chkd
                concEquipsHPM(pointer) = Val(Machine("Hrs_PerMonth").ToString())
                concEquipsDepPerc(pointer) = DepPerc
                concEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
                concEquipsShifts(pointer) = mShifts
                concEquipsMaintPerc(pointer) = mqty
                concEquipsConcQty(pointer) = mConcreteQty
                concEquipsDrive(pointer) = Machine("Drive").ToString()
                concEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
                concEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
                concEquipsUF(pointer) = Val(Machine("UtilityFactor").ToString())
                pointer = pointer + 1
            Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
            mOledbDataAdapter3 = Nothing
            moledbDataSet3 = Nothing
        ElseIf (macategory = "Conveyance") Then
            strSql = "Select * from MajorEquipments where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            moledbDataSet3 = New DataSet
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            counts = moledbDataSet3.Tables("MasterEquipments").Rows.Count - 1
            'MsgBox(moledbDataSet3.Tables("MasterEquipments").Rows.Count)
            ReDim convEquipsNames(counts)
            ReDim convEquipsCapacity(counts)
            ReDim convEquipsMake(counts)
            ReDim convEquipsModel(counts)
            ReDim convEquipsMobDate(counts)
            ReDim convEquipsDemobDate(counts)
            ReDim convEquipsQty(counts)
            ReDim convEquipsChkd(counts)
            ReDim convEquipsHPM(counts)
            ReDim convEquipsDepPerc(counts)
            ReDim convEquipsRepValue(counts)
            ReDim convEquipsShifts(counts)
            ReDim convEquipsMaintPerc(counts)
            ReDim convEquipsConcQty(counts)
            ReDim convEquipsDrive(counts)
            ReDim convEquipsPPU(counts)
            ReDim convEquipsCLPerMc(counts)
            ReDim convEquipsUF(counts)
            pointer = 0
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                'If Not CategoryTextBoxes(intI).Text = "Others" Then mcategory = "external Hire" Else mcategory = "External Others"
                'InsertCommand = ""
                'InsertCommand = "INSERT INTO " & mtablename & " Values ("
                convEquipsNames(pointer) = Machine("EquipmentName").ToString()
                convEquipsCapacity(pointer) = Machine("Capacity").ToString
                convEquipsMake(pointer) = Machine("Make").ToString()
                convEquipsModel(pointer) = Machine("Model").ToString()
                convEquipsMobDate(pointer) = mStartDate.Date.ToString()
                convEquipsDemobDate(pointer) = mEndDate.Date.ToString()
                convEquipsQty(pointer) = mqty
                convEquipsChkd(pointer) = chkd
                convEquipsHPM(pointer) = Val(Machine("Hrs_PerMonth").ToString())
                convEquipsDepPerc(pointer) = DepPerc
                convEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
                convEquipsShifts(pointer) = mShifts
                convEquipsMaintPerc(pointer) = mqty
                convEquipsConcQty(pointer) = mConcreteQty
                convEquipsDrive(pointer) = Machine("Drive").ToString()
                convEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
                convEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
                convEquipsUF(pointer) = Val(Machine("UtilityFactor").ToString())
                pointer = pointer + 1
            Next '
            mOledbDataAdapter3 = Nothing
            moledbDataSet3 = Nothing
        ElseIf (macategory = "Cranes") Then
            strSql = "Select * from MajorEquipments where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            moledbDataSet3 = New DataSet
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            counts = moledbDataSet3.Tables("MasterEquipments").Rows.Count - 1
            ReDim craneEquipsNames(counts)
            ReDim craneEquipsCapacity(counts)
            ReDim craneEquipsMake(counts)
            ReDim craneEquipsModel(counts)
            ReDim craneEquipsMobDate(counts)
            ReDim craneEquipsDemobDate(counts)
            ReDim craneEquipsQty(counts)
            ReDim craneEquipsChkd(counts)
            ReDim craneEquipsHPM(counts)
            ReDim craneEquipsDepPerc(counts)
            ReDim craneEquipsRepValue(counts)
            ReDim craneEquipsShifts(counts)
            ReDim craneEquipsMaintPerc(counts)
            ReDim craneEquipsConcQty(counts)
            ReDim craneEquipsDrive(counts)
            ReDim craneEquipsPPU(counts)
            ReDim craneEquipsCLPerMc(counts)
            ReDim craneEquipsUF(counts)
            pointer = 0
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                craneEquipsNames(pointer) = Machine("EquipmentName").ToString()
                craneEquipsCapacity(pointer) = Machine("Capacity").ToString
                craneEquipsMake(pointer) = Machine("Make").ToString()
                craneEquipsModel(pointer) = Machine("Model").ToString()
                craneEquipsMobDate(pointer) = mStartDate.Date.ToString()
                craneEquipsDemobDate(pointer) = mEndDate.Date.ToString()
                craneEquipsQty(pointer) = mqty
                craneEquipsChkd(pointer) = chkd
                craneEquipsHPM(pointer) = Val(Machine("Hrs_PerMonth").ToString())
                craneEquipsDepPerc(pointer) = DepPerc
                craneEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
                craneEquipsShifts(pointer) = mShifts
                craneEquipsMaintPerc(pointer) = mqty
                craneEquipsConcQty(pointer) = mConcreteQty
                craneEquipsDrive(pointer) = Machine("Drive").ToString()
                craneEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
                craneEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
                craneEquipsUF(pointer) = Val(Machine("UtilityFactor").ToString())
                pointer = pointer + 1
            Next '
            mOledbDataAdapter3 = Nothing
            moledbDataSet3 = Nothing
        ElseIf (macategory = "DG Sets") Then
            strSql = "Select * from MajorEquipments where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            moledbDataSet3 = New DataSet
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            counts = moledbDataSet3.Tables("MasterEquipments").Rows.Count - 1
            ReDim dgsetsEquipsNames(counts)
            ReDim dgsetsEquipsCapacity(counts)
            ReDim dgsetsEquipsMake(counts)
            ReDim dgsetsEquipsModel(counts)
            ReDim dgsetsEquipsMobDate(counts)
            ReDim dgsetsEquipsDemobDate(counts)
            ReDim dgsetsEquipsQty(counts)
            ReDim dgsetsEquipsChkd(counts)
            ReDim dgsetsEquipsHPM(counts)
            ReDim dgsetsEquipsDepPerc(counts)
            ReDim dgsetsEquipsRepValue(counts)
            ReDim dgsetsEquipsShifts(counts)
            ReDim dgsetsEquipsMaintPerc(counts)
            ReDim dgsetsEquipsConcQty(counts)
            ReDim dgsetsEquipsDrive(counts)
            ReDim dgsetsEquipsPPU(counts)
            ReDim dgsetsEquipsCLPerMc(counts)
            ReDim dgsetsEquipsUF(counts)
            pointer = 0
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                dgsetsEquipsNames(pointer) = Machine("EquipmentName").ToString()
                dgsetsEquipsCapacity(pointer) = Machine("Capacity").ToString
                dgsetsEquipsMake(pointer) = Machine("Make").ToString()
                dgsetsEquipsModel(pointer) = Machine("Model").ToString()
                dgsetsEquipsMobDate(pointer) = mStartDate.Date.ToString()
                dgsetsEquipsDemobDate(pointer) = mEndDate.Date.ToString()
                dgsetsEquipsQty(pointer) = mqty
                dgsetsEquipsChkd(pointer) = chkd
                dgsetsEquipsHPM(pointer) = Val(Machine("Hrs_PerMonth").ToString())
                dgsetsEquipsDepPerc(pointer) = DepPerc
                dgsetsEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
                dgsetsEquipsShifts(pointer) = mShifts
                dgsetsEquipsMaintPerc(pointer) = mqty
                dgsetsEquipsConcQty(pointer) = mConcreteQty
                dgsetsEquipsDrive(pointer) = Machine("Drive").ToString()
                dgsetsEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
                dgsetsEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
                dgsetsEquipsUF(pointer) = Val(Machine("UtilityFactor").ToString())
                pointer = pointer + 1
            Next '
            mOledbDataAdapter3 = Nothing
            moledbDataSet3 = Nothing
        ElseIf (macategory = "Material Handling") Then
            strSql = "Select * from MajorEquipments where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            moledbDataSet3 = New DataSet
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            counts = moledbDataSet3.Tables("MasterEquipments").Rows.Count - 1
            ReDim MHEquipsNames(counts)
            ReDim MHEquipsCapacity(counts)
            ReDim MHEquipsMake(counts)
            ReDim MHEquipsModel(counts)
            ReDim MHEquipsMobDate(counts)
            ReDim MHEquipsDemobDate(counts)
            ReDim MHEquipsQty(counts)
            ReDim MHEquipsChkd(counts)
            ReDim MHEquipsHPM(counts)
            ReDim MHEquipsDepPerc(counts)
            ReDim MHEquipsRepValue(counts)
            ReDim MHEquipsShifts(counts)
            ReDim MHEquipsMaintPerc(counts)
            ReDim MHEquipsConcQty(counts)
            ReDim MHEquipsDrive(counts)
            ReDim MHEquipsPPU(counts)
            ReDim MHEquipsCLPerMc(counts)
            ReDim MHEquipsUF(counts)
            pointer = 0
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                MHEquipsNames(pointer) = Machine("EquipmentName").ToString()
                MHEquipsCapacity(pointer) = Machine("Capacity").ToString
                MHEquipsMake(pointer) = Machine("Make").ToString()
                MHEquipsModel(pointer) = Machine("Model").ToString()
                MHEquipsMobDate(pointer) = mStartDate.Date.ToString()
                MHEquipsDemobDate(pointer) = mEndDate.Date.ToString()
                MHEquipsQty(pointer) = mqty
                MHEquipsChkd(pointer) = chkd
                MHEquipsHPM(pointer) = Val(Machine("Hrs_PerMonth").ToString())
                MHEquipsDepPerc(pointer) = DepPerc
                MHEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
                MHEquipsShifts(pointer) = mShifts
                MHEquipsMaintPerc(pointer) = mqty
                MHEquipsConcQty(pointer) = mConcreteQty
                MHEquipsDrive(pointer) = Machine("Drive").ToString()
                MHEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
                MHEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
                MHEquipsUF(pointer) = Val(Machine("UtilityFactor").ToString())
                pointer = pointer + 1
            Next '
            mOledbDataAdapter3 = Nothing
            moledbDataSet3 = Nothing
        ElseIf (macategory = "Non Concreting") Then
            strSql = "Select * from MajorEquipments where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            moledbDataSet3 = New DataSet
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            counts = moledbDataSet3.Tables("MasterEquipments").Rows.Count - 1
            ReDim NCEquipsNames(counts)
            ReDim NCEquipsCapacity(counts)
            ReDim NCEquipsMake(counts)
            ReDim NCEquipsModel(counts)
            ReDim NCEquipsMobDate(counts)
            ReDim NCEquipsDemobDate(counts)
            ReDim NCEquipsQty(counts)
            ReDim NCEquipsChkd(counts)
            ReDim NCEquipsHPM(counts)
            ReDim NCEquipsDepPerc(counts)
            ReDim NCEquipsRepValue(counts)
            ReDim NCEquipsShifts(counts)
            ReDim NCEquipsMaintPerc(counts)
            ReDim NCEquipsConcQty(counts)
            ReDim NCEquipsDrive(counts)
            ReDim NCEquipsPPU(counts)
            ReDim NCEquipsCLPerMc(counts)
            ReDim NCEquipsUF(counts)
            pointer = 0
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                NCEquipsNames(pointer) = Machine("EquipmentName").ToString()
                NCEquipsCapacity(pointer) = Machine("Capacity").ToString
                NCEquipsMake(pointer) = Machine("Make").ToString()
                NCEquipsModel(pointer) = Machine("Model").ToString()
                NCEquipsMobDate(pointer) = mStartDate.Date.ToString()
                NCEquipsDemobDate(pointer) = mEndDate.Date.ToString()
                NCEquipsQty(pointer) = mqty
                NCEquipsChkd(pointer) = chkd
                NCEquipsHPM(pointer) = Val(Machine("Hrs_PerMonth").ToString())
                NCEquipsDepPerc(pointer) = DepPerc
                NCEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
                NCEquipsShifts(pointer) = mShifts
                NCEquipsMaintPerc(pointer) = mqty
                NCEquipsConcQty(pointer) = mConcreteQty
                NCEquipsDrive(pointer) = Machine("Drive").ToString()
                NCEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
                NCEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
                NCEquipsUF(pointer) = Val(Machine("UtilityFactor").ToString())
                pointer = pointer + 1
            Next '
            mOledbDataAdapter3 = Nothing
            moledbDataSet3 = Nothing
        ElseIf (macategory = "Major Others") Then
            strSql = "Select * from MajorEquipments where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            moledbDataSet3 = New DataSet
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            counts = moledbDataSet3.Tables("MasterEquipments").Rows.Count - 1
            ReDim majOthersEquipsNames(counts)
            ReDim majOthersEquipsCapacity(counts)
            ReDim majOthersEquipsMake(counts)
            ReDim majOthersEquipsModel(counts)
            ReDim majOthersEquipsMobDate(counts)
            ReDim majOthersEquipsDemobDate(counts)
            ReDim majOthersEquipsQty(counts)
            ReDim majOthersEquipsChkd(counts)
            ReDim majOthersEquipsHPM(counts)
            ReDim majOthersEquipsDepPerc(counts)
            ReDim majOthersEquipsRepValue(counts)
            ReDim majOthersEquipsShifts(counts)
            ReDim majOthersEquipsMaintPerc(counts)
            ReDim majOthersEquipsConcQty(counts)
            ReDim majOthersEquipsDrive(counts)
            ReDim majOthersEquipsPPU(counts)
            ReDim majOthersEquipsCLPerMc(counts)
            ReDim majOthersEquipsUF(counts)
            pointer = 0
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                majOthersEquipsNames(pointer) = Machine("EquipmentName").ToString()
                majOthersEquipsCapacity(pointer) = Machine("Capacity").ToString
                majOthersEquipsMake(pointer) = Machine("Make").ToString()
                majOthersEquipsModel(pointer) = Machine("Model").ToString()
                majOthersEquipsMobDate(pointer) = mStartDate.Date.ToString()
                majOthersEquipsDemobDate(pointer) = mEndDate.Date.ToString()
                majOthersEquipsQty(pointer) = mqty
                majOthersEquipsChkd(pointer) = chkd
                majOthersEquipsHPM(pointer) = Val(Machine("Hrs_PerMonth").ToString())
                majOthersEquipsDepPerc(pointer) = DepPerc
                majOthersEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
                majOthersEquipsShifts(pointer) = mShifts
                majOthersEquipsMaintPerc(pointer) = mqty
                majOthersEquipsConcQty(pointer) = mConcreteQty
                majOthersEquipsDrive(pointer) = Machine("Drive").ToString()
                majOthersEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
                majOthersEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
                majOthersEquipsUF(pointer) = Val(Machine("UtilityFactor").ToString())
                pointer = pointer + 1
            Next '
            mOledbDataAdapter3 = Nothing
            moledbDataSet3 = Nothing
        ElseIf macategory = "Minor Equipments" Then
            strSql = "Select * from MinorEquipments where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            moledbDataSet3 = New DataSet
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            counts = moledbDataSet3.Tables("MasterEquipments").Rows.Count - 1
            ReDim minorEquipsNames(counts)
            ReDim minorEquipsCapacity(counts)
            ReDim minorEquipsMake(counts)
            ReDim minorEquipsModel(counts)
            ReDim minorEquipsMobDate(counts)
            ReDim minorEquipsDemobDate(counts)
            ReDim minorEquipsQty(counts)
            ReDim minorEquipsChkd(counts)
            ReDim minorEquipsHPM(counts)
            ReDim minorEquipsDepPerc(counts)
            ReDim minorEquipsNewCost(counts)
            ReDim minorEquipsShifts(counts)
            ReDim minorIsNewMC(counts)
            ReDim minorEquipsDrive(counts)
            ReDim minorEquipsPPU(counts)
            ReDim minorEquipsCLPerMc(counts)
            ReDim minorEquipsUF(counts)
            pointer = 0
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                minorEquipsNames(pointer) = Machine("EquipmentName").ToString()
                minorEquipsCapacity(pointer) = Machine("Capacity").ToString
                minorEquipsMake(pointer) = Machine("Make").ToString()
                minorEquipsModel(pointer) = Machine("Model").ToString()
                minorEquipsMobDate(pointer) = mStartDate.Date.ToString()
                minorEquipsDemobDate(pointer) = mEndDate.Date.ToString()
                minorEquipsQty(pointer) = mqty
                minorEquipsChkd(pointer) = chkd
                minorEquipsHPM(pointer) = Val(Machine("Hrs_PerMonth").ToString())
                minorEquipsDepPerc(pointer) = DepPerc
                minorEquipsShifts(pointer) = mShifts
                minorEquipsNewCost(pointer) = Val(Machine("CostOfnewEquipment").ToString())
                minorIsNewMC(pointer) = ISNewMC
                minorEquipsDrive(pointer) = Machine("Drive").ToString()
                minorEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
                minorEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
                minorEquipsUF(pointer) = Val(Machine("UtilityFactor").ToString())
                pointer = pointer + 1
            Next
            mOledbDataAdapter3 = Nothing
            moledbDataSet3 = Nothing
        ElseIf (macategory = "HiredConveyance" Or macategory = "Excav / Earthwork" Or macategory = "GenSets" Or macategory = "Matl Handling" Or _
            macategory = "Matl Transport" Or macategory = "Others") Then
            strSql = "Select * from HiredEquipments" ' where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            moledbDataSet3 = New DataSet
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            counts = moledbDataSet3.Tables("MasterEquipments").Rows.Count - 1
            ReDim hiredCategoryNames(counts)
            ReDim hiredEquipsNames(counts)
            ReDim hiredEquipsCapacity(counts)
            ReDim hiredEquipsMake(counts)
            ReDim hiredEquipsModel(counts)
            ReDim hiredEquipsMobDate(counts)
            ReDim hiredEquipsDemobDate(counts)
            ReDim hiredEquipsQty(counts)
            ReDim hiredEquipsChkd(counts)
            ReDim hiredEquipsHPM(counts)
            ReDim hiredEquipsDepPerc(counts)
            ReDim hiredEquipsHireCharges(counts)
            ReDim hiredEquipsShifts(counts)
            pointer = 0
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                hiredCategoryNames(pointer) = Machine("Categoryname").ToString()
                hiredEquipsNames(pointer) = Machine("EquipmentName").ToString()
                hiredEquipsCapacity(pointer) = Machine("Capacity").ToString()
                hiredEquipsMake(pointer) = Machine("Make").ToString()
                hiredEquipsModel(pointer) = Machine("Model").ToString()
                hiredEquipsMobDate(pointer) = mStartDate.Date.ToString()
                hiredEquipsDemobDate(pointer) = mEndDate.Date.ToString()
                hiredEquipsQty(pointer) = mqty
                hiredEquipsChkd(pointer) = chkd
                hiredEquipsHPM(pointer) = Val(Machine("Hrs_PerMonth").ToString())
                hiredEquipsDepPerc(pointer) = DepPerc
                hiredEquipsShifts(pointer) = mShifts
                hiredEquipsHireCharges(pointer) = Val(Machine("HireCharges").ToString())
                pointer = pointer + 1
            Next
            mOledbDataAdapter3 = Nothing
            moledbDataSet3 = Nothing
        ElseIf macategory = "FixedExpenses" Then
            strSql = "Select * from FixedExpenses"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            moledbDataSet3 = New DataSet
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            counts = moledbDataSet3.Tables("MasterEquipments").Rows.Count - 1
            ReDim fixedExpCategoryNames(counts)
            ReDim fixedExpEquipsQty(counts)
            ReDim fixedExpCost(counts)
            ReDim fixedExpAmount(counts)
            ReDim fixedExpRemarks(counts)
            ReDim fixedExpEquipsChkd(counts)
            ReDim fixedExpProjValue(counts)
            ReDim fixedExpEquipsCostPerc(counts)
            pointer = 0
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                fixedExpCategoryNames(pointer) = Machine("Category").ToString()
                fixedExpEquipsQty(pointer) = mqty
                fixedExpCost(pointer) = Val(Machine("Cost").ToString)
                fixedExpAmount(pointer) = Val(Machine("Cost").ToString)
                fixedExpRemarks(pointer) = Machine("Remarks").ToString()
                fixedExpEquipsChkd(pointer) = chkd
                fixedExpProjValue(pointer) = mProjectvalue
                fixedExpEquipsCostPerc(pointer) = Val(Machine("Cost").ToString) / mProjectvalue
                pointer = pointer + 1
            Next
            mOledbDataAdapter3 = Nothing
            moledbDataSet3 = Nothing
        ElseIf macategory = "BPFixedExpenses" Then
            strSql = "Select * from BPFixedExpenses"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            moledbDataSet3 = New DataSet
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")

            counts = moledbDataSet3.Tables("MasterEquipments").Rows.Count - 1
            ReDim fixedBPExpCategoryNames(counts)
            ReDim fixedBPExpEquipsQty(counts)
            ReDim fixedBPExpCost(counts)
            ReDim fixedBPExpAmount(counts)
            ReDim fixedBPExpRemarks(counts)
            ReDim fixedBPExpEquipsChkd(counts)
            ReDim fixedBPExpProjValue(counts)
            ReDim fixedBPExpEquipsCostPerc(counts)
            pointer = 0
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                fixedBPExpCategoryNames(pointer) = Machine("Category").ToString()
                fixedBPExpEquipsQty(pointer) = mqty
                fixedBPExpCost(pointer) = Val(Machine("Cost").ToString)
                fixedBPExpAmount(pointer) = Val(Machine("Cost").ToString)
                fixedBPExpRemarks(pointer) = Machine("Remarks").ToString()
                fixedBPExpEquipsChkd(pointer) = chkd
                fixedBPExpProjValue(pointer) = mProjectvalue
                fixedBPExpEquipsCostPerc(pointer) = Val(Machine("Cost").ToString) / mProjectvalue
                pointer = pointer + 1
            Next
            mOledbDataAdapter3 = Nothing
            moledbDataSet3 = Nothing
        ElseIf macategory = "Lighting" Then
            strSql = "Select * from LightingEquipments"   ' where categoryname = '" & macategory & "'"
            mOledbDataAdapter3 = New OleDbDataAdapter(strSql, moledbConnection)
            moledbDataSet3 = New DataSet
            mOledbDataAdapter3.Fill(moledbDataSet3, "MasterEquipments")
            counts = moledbDataSet3.Tables("MasterEquipments").Rows.Count - 1
            ReDim lightingCategoryNames(counts)
            ReDim lightingEquipsNames(counts)
            ReDim lightingEquipsCapacity(counts)
            ReDim lightingEquipsMake(counts)
            ReDim lightingEquipsModel(counts)
            ReDim lightingEquipsMobDate(counts)
            ReDim lightingEquipsDemobDate(counts)
            ReDim lightingEquipsQty(counts)
            ReDim lightingEquipsChkd(counts)
            ReDim lightingEquipsPPU(counts)
            ReDim lightingEquipsCLPerMc(counts)
            ReDim lightingEquipsUF(counts)
            pointer = 0
            For Each Machine In moledbDataSet3.Tables("MasterEquipments").Rows
                lightingCategoryNames(pointer) = Machine("Categoryname").ToString()
                lightingEquipsNames(pointer) = Machine("EquipmentName").ToString()
                lightingEquipsCapacity(pointer) = Machine("Capacity").ToString()
                lightingEquipsMake(pointer) = Machine("Make").ToString()
                lightingEquipsModel(pointer) = Machine("Model").ToString()
                lightingEquipsMobDate(pointer) = mStartDate.Date.ToString()
                lightingEquipsDemobDate(pointer) = mEndDate.Date.ToString()
                lightingEquipsQty(pointer) = mqty
                lightingEquipsChkd(pointer) = chkd
                lightingEquipsPPU(pointer) = Val(Machine("PowerPerUnit").ToString())
                lightingEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoad").ToString())
                lightingEquipsUF(pointer) = Val(Machine("UtilityFactor").ToString())
                pointer = pointer + 1
            Next
            mOledbDataAdapter3 = Nothing
            moledbDataSet3 = Nothing
        End If
    End Sub
    Public Sub CopyAddedConcreteItemsToAddedItemsArray()
        Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        If moledbConnection1 Is Nothing Then
            moledbConnection1 = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection1.Open()
        End If
        strSql = "Select * from MajorConcreteEquips"
        mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        moledbDataSet = New DataSet
        mOledbDataAdapter.Fill(moledbDataSet, "ConcreteEquips")
        counts = moledbDataSet.Tables("ConcreteEquips").Rows.Count - 1
        ReDim Preserve concEquipsNames(counts)
        ReDim Preserve concEquipsCapacity(counts)
        ReDim Preserve concEquipsMake(counts)
        ReDim Preserve concEquipsModel(counts)
        ReDim Preserve concEquipsMobDate(counts)
        ReDim Preserve concEquipsDemobDate(counts)
        ReDim Preserve concEquipsQty(counts)
        ReDim Preserve concEquipsChkd(counts)
        ReDim Preserve concEquipsHPM(counts)
        ReDim Preserve concEquipsDepPerc(counts)
        ReDim Preserve concEquipsRepValue(counts)
        ReDim Preserve concEquipsShifts(counts)
        ReDim Preserve concEquipsMaintPerc(counts)
        ReDim Preserve concEquipsConcQty(counts)
        ReDim Preserve concEquipsDrive(counts)
        ReDim Preserve concEquipsPPU(counts)
        ReDim Preserve concEquipsCLPerMc(counts)
        ReDim Preserve concEquipsUF(counts)
        pointer = 0
        For Each Machine In moledbDataSet.Tables("ConcreteEquips").Rows
            concEquipsNames(pointer) = Machine("Description").ToString()
            concEquipsCapacity(pointer) = Machine("Capacity").ToString
            concEquipsMake(pointer) = Machine("Make").ToString()
            concEquipsModel(pointer) = Machine("Model").ToString()
            concEquipsMobDate(pointer) = Machine("MobDate").ToString()
            concEquipsDemobDate(pointer) = Machine("DeMobDate").ToString()
            concEquipsQty(pointer) = Machine("Qty").ToString()
            concEquipsChkd(pointer) = Machine("ChkBoxNo").ToString()
            concEquipsHPM(pointer) = Val(Machine("HrsperMonth").ToString())
            concEquipsDepPerc(pointer) = Val(Machine("Depperc").ToString())
            concEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
            concEquipsShifts(pointer) = Val(Machine("Shifts").ToString())
            concEquipsMaintPerc(pointer) = Val(Machine("MaintPerc").ToString())
            concEquipsConcQty(pointer) = Val(Machine("ConcreteQty").ToString())
            concEquipsDrive(pointer) = Machine("Drive").ToString()
            concEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
            concEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
            concEquipsUF(pointer) = Val(Machine("Utilityfactor").ToString())
            pointer = pointer + 1
        Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        mOledbDataAdapter = Nothing
        moledbDataSet = Nothing
    End Sub
    Public Sub CopyAddedConveyanceItemsToAddedItemsArray()
        Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        If moledbConnection1 Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection1.Open()
        End If
        strSql = "Select * from MajorConveyanceEquips"
        mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        moledbDataSet = New DataSet
        mOledbDataAdapter.Fill(moledbDataSet, "ConveyanceEquips")
        counts = moledbDataSet.Tables("ConveyanceEquips").Rows.Count - 1
        ReDim Preserve convEquipsNames(counts)
        ReDim Preserve convEquipsCapacity(counts)
        ReDim Preserve convEquipsMake(counts)
        ReDim Preserve convEquipsModel(counts)
        ReDim Preserve convEquipsMobDate(counts)
        ReDim Preserve convEquipsDemobDate(counts)
        ReDim Preserve convEquipsQty(counts)
        ReDim Preserve convEquipsChkd(counts)
        ReDim Preserve convEquipsHPM(counts)
        ReDim Preserve convEquipsDepPerc(counts)
        ReDim Preserve convEquipsRepValue(counts)
        ReDim Preserve convEquipsShifts(counts)
        ReDim Preserve convEquipsMaintPerc(counts)
        ReDim Preserve convEquipsConcQty(counts)
        ReDim Preserve convEquipsDrive(counts)
        ReDim Preserve convEquipsPPU(counts)
        ReDim Preserve convEquipsCLPerMc(counts)
        ReDim Preserve convEquipsUF(counts)
        pointer = 0
        For Each Machine In moledbDataSet.Tables("ConveyanceEquips").Rows
            convEquipsNames(pointer) = Machine("Description").ToString()
            convEquipsCapacity(pointer) = Machine("Capacity").ToString
            convEquipsMake(pointer) = Machine("Make").ToString()
            convEquipsModel(pointer) = Machine("Model").ToString()
            convEquipsMobDate(pointer) = Machine("MobDate").ToString()
            convEquipsDemobDate(pointer) = Machine("DeMobDate").ToString()
            convEquipsQty(pointer) = Machine("Qty").ToString()
            convEquipsChkd(pointer) = Machine("ChkBoxNo").ToString()
            convEquipsHPM(pointer) = Val(Machine("HrsperMonth").ToString())
            convEquipsDepPerc(pointer) = Val(Machine("Depperc").ToString())
            convEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
            convEquipsShifts(pointer) = Val(Machine("Shifts").ToString())
            convEquipsMaintPerc(pointer) = Val(Machine("MaintPerc").ToString())
            convEquipsConcQty(pointer) = Val(Machine("ConcreteQty").ToString())
            convEquipsDrive(pointer) = Machine("Drive").ToString()
            convEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
            convEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
            convEquipsUF(pointer) = Val(Machine("Utilityfactor").ToString())
            pointer = pointer + 1
        Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        mOledbDataAdapter = Nothing
        moledbDataSet = Nothing
    End Sub
    Public Sub CopyAddedCraneItemsToAddedItemsArray()
        Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        If moledbConnection1 Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strSql = "Select * from MajorCraneEquips"
        mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        moledbDataSet = New DataSet
        mOledbDataAdapter.Fill(moledbDataSet, "CraneEquips")
        counts = moledbDataSet.Tables("CraneEquips").Rows.Count - 1
        ReDim Preserve craneEquipsNames(counts)
        ReDim Preserve craneEquipsCapacity(counts)
        ReDim Preserve craneEquipsMake(counts)
        ReDim Preserve craneEquipsModel(counts)
        ReDim Preserve craneEquipsMobDate(counts)
        ReDim Preserve craneEquipsDemobDate(counts)
        ReDim Preserve craneEquipsQty(counts)
        ReDim Preserve craneEquipsChkd(counts)
        ReDim Preserve craneEquipsHPM(counts)
        ReDim Preserve craneEquipsDepPerc(counts)
        ReDim Preserve craneEquipsRepValue(counts)
        ReDim Preserve craneEquipsShifts(counts)
        ReDim Preserve craneEquipsMaintPerc(counts)
        ReDim Preserve craneEquipsConcQty(counts)
        ReDim Preserve craneEquipsDrive(counts)
        ReDim Preserve craneEquipsPPU(counts)
        ReDim Preserve craneEquipsCLPerMc(counts)
        ReDim Preserve craneEquipsUF(counts)
        pointer = 0
        For Each Machine In moledbDataSet.Tables("CraneEquips").Rows
            craneEquipsNames(pointer) = Machine("Description").ToString()
            craneEquipsCapacity(pointer) = Machine("Capacity").ToString
            craneEquipsMake(pointer) = Machine("Make").ToString()
            craneEquipsModel(pointer) = Machine("Model").ToString()
            craneEquipsMobDate(pointer) = Machine("MobDate").ToString()
            craneEquipsDemobDate(pointer) = Machine("DeMobDate").ToString()
            craneEquipsQty(pointer) = Machine("Qty").ToString()
            craneEquipsChkd(pointer) = Machine("ChkBoxNo").ToString()
            craneEquipsHPM(pointer) = Val(Machine("HrsperMonth").ToString())
            craneEquipsDepPerc(pointer) = Val(Machine("Depperc").ToString())
            craneEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
            craneEquipsShifts(pointer) = Val(Machine("Shifts").ToString())
            craneEquipsMaintPerc(pointer) = Val(Machine("MaintPerc").ToString())
            craneEquipsConcQty(pointer) = Val(Machine("ConcreteQty").ToString())
            craneEquipsDrive(pointer) = Machine("Drive").ToString()
            craneEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
            craneEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
            craneEquipsUF(pointer) = Val(Machine("Utilityfactor").ToString())
            pointer = pointer + 1
        Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        mOledbDataAdapter = Nothing
        moledbDataSet = Nothing
    End Sub
    Public Sub CopyAddedDgSetsItemsToAddedItemsArray()
        Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        If moledbConnection1 Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strSql = "Select * from MajorDGSetsEquips"
        mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        moledbDataSet = New DataSet
        mOledbDataAdapter.Fill(moledbDataSet, "DGSetsEquips")
        counts = moledbDataSet.Tables("DGSetsEquips").Rows.Count - 1
        ReDim Preserve dgsetsEquipsNames(counts)
        ReDim Preserve dgsetsEquipsCapacity(counts)
        ReDim Preserve dgsetsEquipsMake(counts)
        ReDim Preserve dgsetsEquipsModel(counts)
        ReDim Preserve dgsetsEquipsMobDate(counts)
        ReDim Preserve dgsetsEquipsDemobDate(counts)
        ReDim Preserve dgsetsEquipsQty(counts)
        ReDim Preserve dgsetsEquipsChkd(counts)
        ReDim Preserve dgsetsEquipsHPM(counts)
        ReDim Preserve dgsetsEquipsDepPerc(counts)
        ReDim Preserve dgsetsEquipsRepValue(counts)
        ReDim Preserve dgsetsEquipsShifts(counts)
        ReDim Preserve dgsetsEquipsMaintPerc(counts)
        ReDim Preserve dgsetsEquipsConcQty(counts)
        ReDim Preserve dgsetsEquipsDrive(counts)
        ReDim Preserve dgsetsEquipsPPU(counts)
        ReDim Preserve dgsetsEquipsCLPerMc(counts)
        ReDim Preserve dgsetsEquipsUF(counts)
        pointer = 0
        For Each Machine In moledbDataSet.Tables("DGSetsEquips").Rows
            dgsetsEquipsNames(pointer) = Machine("Description").ToString()
            dgsetsEquipsCapacity(pointer) = Machine("Capacity").ToString
            dgsetsEquipsMake(pointer) = Machine("Make").ToString()
            dgsetsEquipsModel(pointer) = Machine("Model").ToString()
            dgsetsEquipsMobDate(pointer) = Machine("MobDate").ToString()
            dgsetsEquipsDemobDate(pointer) = Machine("DeMobDate").ToString()
            dgsetsEquipsQty(pointer) = Machine("Qty").ToString()
            dgsetsEquipsChkd(pointer) = Machine("ChkBoxNo").ToString()
            dgsetsEquipsHPM(pointer) = Val(Machine("HrsperMonth").ToString())
            dgsetsEquipsDepPerc(pointer) = Val(Machine("Depperc").ToString())
            dgsetsEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
            dgsetsEquipsShifts(pointer) = Val(Machine("Shifts").ToString())
            dgsetsEquipsMaintPerc(pointer) = Val(Machine("MaintPerc").ToString())
            dgsetsEquipsConcQty(pointer) = Val(Machine("ConcreteQty").ToString())
            dgsetsEquipsDrive(pointer) = Machine("Drive").ToString()
            dgsetsEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
            dgsetsEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
            dgsetsEquipsUF(pointer) = Val(Machine("Utilityfactor").ToString())
            pointer = pointer + 1
        Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        mOledbDataAdapter = Nothing
        moledbDataSet = Nothing
    End Sub
    Public Sub CopyAddedMHItemsToAddedItemsArray()
        Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        If moledbConnection1 Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strSql = "Select * from MajorMaterialHandlingEquips"
        mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        moledbDataSet = New DataSet
        mOledbDataAdapter.Fill(moledbDataSet, "MHEquips")
        counts = moledbDataSet.Tables("MHEquips").Rows.Count - 1
        ReDim Preserve MHEquipsNames(counts)
        ReDim Preserve MHEquipsCapacity(counts)
        ReDim Preserve MHEquipsMake(counts)
        ReDim Preserve MHEquipsModel(counts)
        ReDim Preserve MHEquipsMobDate(counts)
        ReDim Preserve MHEquipsDemobDate(counts)
        ReDim Preserve MHEquipsQty(counts)
        ReDim Preserve MHEquipsChkd(counts)
        ReDim Preserve MHEquipsHPM(counts)
        ReDim Preserve MHEquipsDepPerc(counts)
        ReDim Preserve MHEquipsRepValue(counts)
        ReDim Preserve MHEquipsShifts(counts)
        ReDim Preserve MHEquipsMaintPerc(counts)
        ReDim Preserve MHEquipsConcQty(counts)
        ReDim Preserve MHEquipsDrive(counts)
        ReDim Preserve MHEquipsPPU(counts)
        ReDim Preserve MHEquipsCLPerMc(counts)
        ReDim Preserve MHEquipsUF(counts)
        pointer = 0
        For Each Machine In moledbDataSet.Tables("MHEquips").Rows
            MHEquipsNames(pointer) = Machine("Description").ToString()
            MHEquipsCapacity(pointer) = Machine("Capacity").ToString
            MHEquipsMake(pointer) = Machine("Make").ToString()
            MHEquipsModel(pointer) = Machine("Model").ToString()
            MHEquipsMobDate(pointer) = Machine("MobDate").ToString()
            MHEquipsDemobDate(pointer) = Machine("DeMobDate").ToString()
            MHEquipsQty(pointer) = Machine("Qty").ToString()
            MHEquipsChkd(pointer) = Machine("ChkBoxNo").ToString()
            MHEquipsHPM(pointer) = Val(Machine("HrsperMonth").ToString())
            MHEquipsDepPerc(pointer) = Val(Machine("Depperc").ToString())
            MHEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
            MHEquipsShifts(pointer) = Val(Machine("Shifts").ToString())
            MHEquipsMaintPerc(pointer) = Val(Machine("MaintPerc").ToString())
            MHEquipsConcQty(pointer) = Val(Machine("ConcreteQty").ToString())
            MHEquipsDrive(pointer) = Machine("Drive").ToString()
            MHEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
            MHEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
            MHEquipsUF(pointer) = Val(Machine("Utilityfactor").ToString())
            pointer = pointer + 1
        Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        mOledbDataAdapter = Nothing
        moledbDataSet = Nothing
    End Sub
    Public Sub CopyAddedNCItemsToAddedItemsArray()
        Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        If moledbConnection1 Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strSql = "Select * from MajorNonConcreteEquips"
        mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        moledbDataSet = New DataSet
        mOledbDataAdapter.Fill(moledbDataSet, "NCEquips")
        counts = moledbDataSet.Tables("NCEquips").Rows.Count - 1
        ReDim Preserve NCEquipsNames(counts)
        ReDim Preserve NCEquipsCapacity(counts)
        ReDim Preserve NCEquipsMake(counts)
        ReDim Preserve NCEquipsModel(counts)
        ReDim Preserve NCEquipsMobDate(counts)
        ReDim Preserve NCEquipsDemobDate(counts)
        ReDim Preserve NCEquipsQty(counts)
        ReDim Preserve NCEquipsChkd(counts)
        ReDim Preserve NCEquipsHPM(counts)
        ReDim Preserve NCEquipsDepPerc(counts)
        ReDim Preserve NCEquipsRepValue(counts)
        ReDim Preserve NCEquipsShifts(counts)
        ReDim Preserve NCEquipsMaintPerc(counts)
        ReDim Preserve NCEquipsConcQty(counts)
        ReDim Preserve NCEquipsDrive(counts)
        ReDim Preserve NCEquipsPPU(counts)
        ReDim Preserve NCEquipsCLPerMc(counts)
        ReDim Preserve NCEquipsUF(counts)
        pointer = 0
        For Each Machine In moledbDataSet.Tables("NCEquips").Rows
            NCEquipsNames(pointer) = Machine("Description").ToString()
            NCEquipsCapacity(pointer) = Machine("Capacity").ToString
            NCEquipsMake(pointer) = Machine("Make").ToString()
            NCEquipsModel(pointer) = Machine("Model").ToString()
            NCEquipsMobDate(pointer) = Machine("MobDate").ToString()
            NCEquipsDemobDate(pointer) = Machine("DeMobDate").ToString()
            NCEquipsQty(pointer) = Machine("Qty").ToString()
            NCEquipsChkd(pointer) = Machine("ChkBoxNo").ToString()
            NCEquipsHPM(pointer) = Val(Machine("HrsperMonth").ToString())
            NCEquipsDepPerc(pointer) = Val(Machine("Depperc").ToString())
            NCEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
            NCEquipsShifts(pointer) = Val(Machine("Shifts").ToString())
            NCEquipsMaintPerc(pointer) = Val(Machine("MaintPerc").ToString())
            NCEquipsConcQty(pointer) = Val(Machine("ConcreteQty").ToString())
            NCEquipsDrive(pointer) = Machine("Drive").ToString()
            NCEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
            NCEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
            NCEquipsUF(pointer) = Val(Machine("Utilityfactor").ToString())
            pointer = pointer + 1
        Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        mOledbDataAdapter = Nothing
        moledbDataSet = Nothing
    End Sub
    Public Sub CopyAddedMajOtherItemsToAddedItemsArray()
        Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        If moledbConnection1 Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strSql = "Select * from MajorOthers"
        mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        moledbDataSet = New DataSet
        mOledbDataAdapter.Fill(moledbDataSet, "MajOthers")
        counts = moledbDataSet.Tables("MajOthers").Rows.Count - 1
        ReDim Preserve majOthersEquipsNames(counts)
        ReDim Preserve majOthersEquipsCapacity(counts)
        ReDim Preserve majOthersEquipsMake(counts)
        ReDim Preserve majOthersEquipsModel(counts)
        ReDim Preserve majOthersEquipsMobDate(counts)
        ReDim Preserve majOthersEquipsDemobDate(counts)
        ReDim Preserve majOthersEquipsQty(counts)
        ReDim Preserve majOthersEquipsChkd(counts)
        ReDim Preserve majOthersEquipsHPM(counts)
        ReDim Preserve MajOthersDepPerc(counts)
        ReDim Preserve majOthersEquipsRepValue(counts)
        ReDim Preserve MajOthersShifts(counts)
        ReDim Preserve majOthersEquipsMaintPerc(counts)
        ReDim Preserve majOthersEquipsConcQty(counts)
        ReDim Preserve majOthersEquipsDrive(counts)
        ReDim Preserve majOthersEquipsPPU(counts)
        ReDim Preserve majOthersEquipsCLPerMc(counts)
        ReDim Preserve majOthersEquipsUF(counts)
        pointer = 0
        For Each Machine In moledbDataSet.Tables("MajOthers").Rows
            majOthersEquipsNames(pointer) = Machine("Description").ToString()
            majOthersEquipsCapacity(pointer) = Machine("Capacity").ToString
            majOthersEquipsMake(pointer) = Machine("Make").ToString()
            majOthersEquipsModel(pointer) = Machine("Model").ToString()
            majOthersEquipsMobDate(pointer) = Machine("MobDate").ToString()
            majOthersEquipsDemobDate(pointer) = Machine("DeMobDate").ToString()
            majOthersEquipsQty(pointer) = Machine("Qty").ToString()
            majOthersEquipsChkd(pointer) = Machine("ChkBoxNo").ToString()
            majOthersEquipsHPM(pointer) = Val(Machine("HrsperMonth").ToString())
            MajOthersDepPerc(pointer) = Val(Machine("Depperc").ToString())
            majOthersEquipsRepValue(pointer) = Val(Machine("RepValue").ToString())
            MajOthersShifts(pointer) = Val(Machine("Shifts").ToString())
            majOthersEquipsMaintPerc(pointer) = Val(Machine("MaintPerc").ToString())
            majOthersEquipsConcQty(pointer) = Val(Machine("ConcreteQty").ToString())
            majOthersEquipsDrive(pointer) = Machine("Drive").ToString()
            majOthersEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
            majOthersEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
            majOthersEquipsUF(pointer) = Val(Machine("Utilityfactor").ToString())
            pointer = pointer + 1
        Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        mOledbDataAdapter = Nothing
        moledbDataSet = Nothing
    End Sub
    Public Sub CopyAddedMinorItemsToAddedItemsArray()
        Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        If moledbConnection1 Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strSql = "Select * from MinorEquips"
        mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        moledbDataSet = New DataSet
        mOledbDataAdapter.Fill(moledbDataSet, "MinorEquips")
        counts = moledbDataSet.Tables("MinorEquips").Rows.Count - 1
        ReDim Preserve minorEquipsNames(counts)
        ReDim Preserve minorEquipsCapacity(counts)
        ReDim Preserve minorEquipsMake(counts)
        ReDim Preserve minorEquipsModel(counts)
        ReDim Preserve minorEquipsMobDate(counts)
        ReDim Preserve minorEquipsDemobDate(counts)
        ReDim Preserve minorEquipsQty(counts)
        ReDim Preserve minorEquipsChkd(counts)
        ReDim Preserve minorEquipsHPM(counts)
        ReDim Preserve minorEquipsDepPerc(counts)
        ReDim Preserve minorEquipsNewCost(counts)
        ReDim Preserve minorEquipsShifts(counts)
        ReDim Preserve minorIsNewMC(counts)
        ReDim Preserve minorEquipsDrive(counts)
        ReDim Preserve minorEquipsPPU(counts)
        ReDim Preserve minorEquipsCLPerMc(counts)
        ReDim Preserve minorEquipsUF(counts)
        pointer = 0
        For Each Machine In moledbDataSet.Tables("MinorEquips").Rows
            minorEquipsNames(pointer) = Machine("Description").ToString()
            minorEquipsCapacity(pointer) = Machine("Capacity").ToString
            minorEquipsMake(pointer) = Machine("Make").ToString()
            minorEquipsModel(pointer) = Machine("Model").ToString()
            minorEquipsMobDate(pointer) = Machine("MobDate").ToString()
            minorEquipsDemobDate(pointer) = Machine("DeMobDate").ToString()
            minorEquipsQty(pointer) = Machine("Qty").ToString()
            minorEquipsChkd(pointer) = Machine("ChkBoxNo").ToString()
            minorEquipsHPM(pointer) = Val(Machine("HrsperMonth").ToString())
            minorEquipsDepPerc(pointer) = Val(Machine("Depperc").ToString())
            minorEquipsNewCost(pointer) = Val(Machine("NewPurchVal").ToString())
            minorEquipsShifts(pointer) = Val(Machine("Shifts").ToString())
            minorIsNewMC(pointer) = Val(Machine("IsNewMc").ToString())
            minorEquipsDrive(pointer) = Machine("Drive").ToString()
            minorEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
            minorEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
            minorEquipsUF(pointer) = Val(Machine("Utilityfactor").ToString())
            pointer = pointer + 1
        Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        mOledbDataAdapter = Nothing
        moledbDataSet = Nothing
    End Sub
    Public Sub CopyAddedExcavItemsToAddedItemsArray()
        'Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        'Dim Machine As DataRow
        'Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        'Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        'strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        'If moledbConnection1 Is Nothing Then
        '    moledbConnection = New OleDbConnection(strConnection)
        'End If
        'If (moledbConnection1.State.ToString().Equals("Closed")) Then
        '    moledbConnection.Open()
        'End If
        'strSql = "Select * from MinorEquips"
        'mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        'moledbDataSet = New DataSet
        'mOledbDataAdapter.Fill(moledbDataSet, "MinorEquips")
        'counts = moledbDataSet.Tables("MinorEquips").Rows.Count - 1
        'ReDim Preserve minorEquipsNames(counts)
        'ReDim Preserve minorEquipsCapacity(counts)
        'ReDim Preserve minorEquipsMake(counts)
        'ReDim Preserve minorEquipsModel(counts)
        'ReDim Preserve minorEquipsMobDate(counts)
        'ReDim Preserve minorEquipsDemobDate(counts)
        'ReDim Preserve minorEquipsQty(counts)
        'ReDim Preserve minorEquipsChkd(counts)
        'ReDim Preserve minorEquipsHPM(counts)
        'ReDim Preserve MajOthersDepPerc(counts)
        'ReDim Preserve minorEquipsNewCost(counts)
        'ReDim Preserve MajOthersShifts(counts)
        'ReDim Preserve minorIsNewMC(counts)
        'ReDim Preserve minorEquipsDrive(counts)
        'ReDim Preserve minorEquipsPPU(counts)
        'ReDim Preserve minorEquipsCLPerMc(counts)
        'ReDim Preserve minorEquipsUF(counts)
        'pointer = 0
        'For Each Machine In moledbDataSet.Tables("MinorEquips").Rows
        '    minorEquipsNames(pointer) = Machine("Description").ToString()
        '    minorEquipsCapacity(pointer) = Machine("Capacity").ToString
        '    minorEquipsMake(pointer) = Machine("Make").ToString()
        '    minorEquipsModel(pointer) = Machine("Model").ToString()
        '    minorEquipsMobDate(pointer) = Machine("MobDate").ToString()
        '    minorEquipsDemobDate(pointer) = Machine("DeMobDate").ToString()
        '    minorEquipsQty(pointer) = Machine("Qty").ToString()
        '    minorEquipsChkd(pointer) = Machine("ChkBoxNo").ToString()
        '    minorEquipsHPM(pointer) = Val(Machine("HrsperMonth").ToString())
        '    minorEquipsDepPerc(pointer) = Val(Machine("Depperc").ToString())
        '    minorEquipsNewCost(pointer) = Val(Machine("RepValue").ToString())
        '    minorEquipsShifts(pointer) = Val(Machine("Shifts").ToString())
        '    minorIsNewMC(pointer) = Val(Machine("MaintPerc").ToString())
        '    minorEquipsDrive(pointer) = Machine("Drive").ToString()
        '    minorEquipsPPU(pointer) = Val(Machine("PowerPerUnit(HP)").ToString())
        '    minorEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
        '    minorEquipsUF(pointer) = Val(Machine("Utilityfactor").ToString())
        '    pointer = pointer + 1
        'Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        'mOledbDataAdapter = Nothing
        'moledbDataSet = Nothing
    End Sub
    Public Sub CopyAddedHiredItemsToAddedItemsArray()
        Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        If moledbConnection1 Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strSql = "Select * from HiredEquips"
        mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        moledbDataSet = New DataSet
        mOledbDataAdapter.Fill(moledbDataSet, "HiredEquips")
        counts = moledbDataSet.Tables("HiredEquips").Rows.Count - 1
        ReDim Preserve hiredCategoryNames(counts)
        ReDim Preserve hiredEquipsNames(counts)
        ReDim Preserve hiredEquipsCapacity(counts)
        ReDim Preserve hiredEquipsMake(counts)
        ReDim Preserve hiredEquipsModel(counts)
        ReDim Preserve hiredEquipsMobDate(counts)
        ReDim Preserve hiredEquipsDemobDate(counts)
        ReDim Preserve hiredEquipsQty(counts)
        ReDim Preserve hiredEquipsChkd(counts)
        ReDim Preserve hiredEquipsHPM(counts)
        ReDim Preserve hiredEquipsDepPerc(counts)
        ReDim Preserve hiredEquipsHireCharges(counts)
        ReDim Preserve hiredEquipsShifts(counts)

        pointer = 0
        For Each Machine In moledbDataSet.Tables("HiredEquips").Rows
            hiredCategoryNames(pointer) = Machine("Category").ToString()
            hiredEquipsNames(pointer) = Machine("Description").ToString()
            hiredEquipsCapacity(pointer) = Machine("Capacity").ToString
            hiredEquipsMake(pointer) = Machine("Make").ToString()
            hiredEquipsModel(pointer) = Machine("Model").ToString()
            hiredEquipsMobDate(pointer) = Machine("MobDate").ToString()
            hiredEquipsDemobDate(pointer) = Machine("DeMobDate").ToString()
            hiredEquipsQty(pointer) = Val(Machine("Qty").ToString())
            hiredEquipsChkd(pointer) = Val(Machine("ChkBoxNo").ToString())
            hiredEquipsHPM(pointer) = Val(Machine("HrsperMonth").ToString())
            hiredEquipsDepPerc(pointer) = Val(Machine("Depperc").ToString())
            hiredEquipsShifts(pointer) = Val(Machine("Shifts").ToString())
            hiredEquipsHireCharges(pointer) = Val(Machine("HireCharges").ToString())
            pointer = pointer + 1
        Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        mOledbDataAdapter = Nothing
        moledbDataSet = Nothing
    End Sub
    Public Sub CopyAddedFixedExpItemsToAddedItemsArray()
        Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        If moledbConnection1 Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strSql = "Select * from FixedExpAdded"
        mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        moledbDataSet = New DataSet
        mOledbDataAdapter.Fill(moledbDataSet, "fixedExp")
        counts = moledbDataSet.Tables("fixedExp").Rows.Count - 1
        ReDim Preserve fixedExpCategoryNames(counts)
        ReDim Preserve fixedExpEquipsQty(counts)
        ReDim Preserve fixedExpCost(counts)
        ReDim Preserve fixedExpAmount(counts)
        ReDim Preserve fixedExpRemarks(counts)
        ReDim Preserve fixedExpEquipsChkd(counts)
        ReDim Preserve fixedExpProjValue(counts)
        ReDim Preserve fixedExpEquipsCostPerc(counts)

        pointer = 0
        For Each Machine In moledbDataSet.Tables("fixedExp").Rows
            fixedExpCategoryNames(pointer) = Machine("Category").ToString()
            fixedExpEquipsQty(pointer) = Val(Machine("Qty").ToString())
            fixedExpEquipsChkd(pointer) = Val(Machine("chkBoxNo").ToString())
            fixedExpCost(pointer) = Val(Machine("Cost").ToString())
            fixedExpAmount(pointer) = Val(Machine("Amount").ToString())
            fixedExpRemarks(pointer) = Val(Machine("Remarks").ToString())
            fixedExpProjValue(pointer) = Val(Machine("ClientBilling").ToString())
            fixedExpEquipsCostPerc(pointer) = Machine("CostPerc").ToString()
            pointer = pointer + 1
        Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        mOledbDataAdapter = Nothing
        moledbDataSet = Nothing
    End Sub
    Public Sub CopyAddedBPFixedExpItemsToAddedItemsArray()
        Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        If moledbConnection1 Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strSql = "Select * from BPFixedExpAdded"
        mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        moledbDataSet = New DataSet
        mOledbDataAdapter.Fill(moledbDataSet, "BPfixedExp")
        counts = moledbDataSet.Tables("BPfixedExp").Rows.Count - 1
        ReDim Preserve fixedBPExpCategoryNames(counts)
        ReDim Preserve fixedBPExpEquipsQty(counts)
        ReDim Preserve fixedBPExpCost(counts)
        ReDim Preserve fixedBPExpAmount(counts)
        ReDim Preserve fixedBPExpRemarks(counts)
        ReDim Preserve fixedBPExpEquipsChkd(counts)
        ReDim Preserve fixedBPExpProjValue(counts)
        ReDim Preserve fixedBPExpEquipsCostPerc(counts)
        pointer = 0
        For Each Machine In moledbDataSet.Tables("BPfixedExp").Rows
            fixedBPExpCategoryNames(pointer) = Machine("Category").ToString()
            fixedBPExpEquipsQty(pointer) = Val(Machine("Qty").ToString())
            fixedBPExpEquipsChkd(pointer) = Val(Machine("chkBoxNo").ToString())
            fixedBPExpCost(pointer) = Val(Machine("Cost").ToString())
            fixedBPExpAmount(pointer) = Val(Machine("Amount").ToString())
            fixedBPExpRemarks(pointer) = Val(Machine("Remarks").ToString())
            fixedBPExpProjValue(pointer) = Val(Machine("ClientBilling").ToString())
            fixedBPExpEquipsCostPerc(pointer) = Machine("CostPerc").ToString()
            pointer = pointer + 1
        Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        mOledbDataAdapter = Nothing
        moledbDataSet = Nothing
    End Sub
    Public Sub CopyAddedLightingItemsToAddedItemsArray()
        Dim strSql As String, mOledbDataAdapter As OleDbDataAdapter
        Dim Machine As DataRow
        Dim moledbDataSet As DataSet, mqty As Integer = 1, chkd As Integer = 0, DepPerc As Single = 2.75, mShifts As Single = 1
        Dim ISNewMC As Boolean = False, pointer As Integer, counts As Integer
        strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfilename
        If moledbConnection1 Is Nothing Then
            moledbConnection = New OleDbConnection(strConnection)
        End If
        If (moledbConnection1.State.ToString().Equals("Closed")) Then
            moledbConnection.Open()
        End If
        strSql = "Select * from LightingEquipments"
        mOledbDataAdapter = New OleDbDataAdapter(strSql, moledbConnection1)
        moledbDataSet = New DataSet
        mOledbDataAdapter.Fill(moledbDataSet, "Lighting")
        counts = moledbDataSet.Tables("Lighting").Rows.Count - 1
        ReDim Preserve lightingCategoryNames(counts)
        ReDim Preserve lightingEquipsNames(counts)
        ReDim Preserve lightingEquipsCapacity(counts)
        ReDim Preserve lightingEquipsMake(counts)
        ReDim Preserve lightingEquipsModel(counts)
        ReDim Preserve lightingEquipsMobDate(counts)
        ReDim Preserve lightingEquipsDemobDate(counts)
        ReDim Preserve lightingEquipsQty(counts)
        ReDim Preserve lightingEquipsChkd(counts)
        ReDim Preserve lightingEquipsPPU(counts)
        ReDim Preserve lightingEquipsCLPerMc(counts)
        ReDim Preserve lightingEquipsUF(counts)
        pointer = 0
        For Each Machine In moledbDataSet.Tables("Lighting").Rows
            lightingCategoryNames(pointer) = Machine("Category").ToString()
            lightingEquipsNames(pointer) = Machine("Description").ToString()
            lightingEquipsCapacity(pointer) = Machine("Capacity").ToString()
            lightingEquipsMake(pointer) = Machine("Make").ToString()
            lightingEquipsModel(pointer) = Machine("Model").ToString()
            lightingEquipsMobDate(pointer) = Machine("MobDate").ToString()
            lightingEquipsDemobDate(pointer) = Machine("DeMobDate").ToString()
            lightingEquipsQty(pointer) = Val(Machine("Qty").ToString())
            lightingEquipsChkd(pointer) = Val(Machine("ChkBoxNo").ToString())
            lightingEquipsPPU(pointer) = Val(Machine("PowerPerUnit").ToString())
            lightingEquipsCLPerMc(pointer) = Val(Machine("ConnectedLoadPerMC").ToString())
            lightingEquipsUF(pointer) = Val(Machine("Utilityfactor").ToString())
            pointer = pointer + 1
        Next ' End of condition mcategory <> "Minor Equipments" And mcategory <> "Hired Equipments" 
        mOledbDataAdapter = Nothing
        moledbDataSet = Nothing
    End Sub


    Public Function GetTablename(ByVal catagory As String) As String
        GetTablename = ""
        If UCase(catagory) = UCase("Concreting") Then
            GetTablename = "MajorConcreteEquips"
        ElseIf UCase(catagory) = UCase("Cranes") Then
            GetTablename = "MajorCraneEquips"
        ElseIf UCase(catagory) = UCase("Material Handling") Then
            GetTablename = "MajorMaterialhandlingEquips"
        ElseIf UCase(catagory) = UCase("DG Sets") Then
            GetTablename = "MajorDGSetsEquips"
        ElseIf UCase(catagory) = UCase("Non Concreting") Then
            GetTablename = "MajorNonConcreteEquips"
        ElseIf UCase(catagory) = UCase("Conveyance") Then
            GetTablename = "MajorConveyanceEquips"
        ElseIf UCase(catagory) = UCase("Major Others") Then
            GetTablename = "MajorOthers"
        ElseIf UCase(catagory) = UCase("Minor Equipments") Then
            GetTablename = "MinorEquips"
        ElseIf UCase(catagory) = UCase("HiredConveyance") Or UCase(catagory) = UCase("Excav / Earthwork") Or _
              UCase(catagory) = UCase("Matl Handling") Or UCase(catagory) = UCase("Matl Transport") Or _
              UCase(catagory) = UCase("Gensets") Or UCase(catagory) = UCase("Others") Then
            GetTablename = "HiredEquips"
        ElseIf UCase(catagory) = UCase("FixedExp") Then
            GetTablename = "FixedExpAdded"
        ElseIf UCase(catagory) = UCase("BPFixed - Exp") Then
            GetTablename = "BPFixedExpAdded"
        ElseIf UCase(catagory) = UCase("Lighting") Then
            GetTablename = "LightingEquipments"
        ElseIf UCase(catagory) = UCase("PowerEquips") Then
            GetTablename = "MajorPowerEquipments"
        End If
    End Function
    Public Function IsWorkbookOpen(ByVal wbName As String) As Boolean
        Dim i As Long
        Dim XLAppFx As Microsoft.Office.Interop.Excel.Application
        Dim retval As Boolean = False
        Dim intI As Integer, lastPos As Integer = 0, newname As String = ""
        For intI = 0 To wbName.Length - 1
            If wbName.Substring(intI, 1) = "\" Then
                lastPos = intI
            End If
        Next
        wbName = wbName.Substring(lastPos + 1)

        XLAppFx = GetObject(, "Excel.Application")
        If Err.Number = 429 Then
            retval = False
        Else
            For i = XLAppFx.Workbooks.Count To 1 Step -1
                If XLAppFx.Workbooks(i).Name = wbName Then Exit For
            Next
            If i <> 0 Then retval = True
        End If
        Return retval
    End Function
    Public Function getSheetNo(ByVal wksheetname As String) As Integer
        If wksheetname = "Concreting" Then
            Return 2
        ElseIf wksheetname = "Cranes" Then
            Return 3
        ElseIf wksheetname = "Material Handling" Then
            Return 4
        ElseIf wksheetname = "Non Concreting" Then
            Return 5
        ElseIf wksheetname = "DG Sets" Then
            Return 6
        ElseIf wksheetname = "Conveyance" Then
            Return 7
        ElseIf wksheetname = "Major Others" Then
            Return 8
        ElseIf wksheetname = "external Hire" Then
            Return 9
        ElseIf wksheetname = "External Others" Then
            Return 10
        ElseIf wksheetname = "Minor Eqpts" Then
            Return 11
        ElseIf wksheetname = "PowerGen Cost" Then
            Return 19
        ElseIf wksheetname = "PowerReqmt" Then
            Return 20
        End If
    End Function
End Module

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProjectDetails
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Label1 = New System.Windows.Forms.Label
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.pnlDetails = New System.Windows.Forms.Panel
        Me.txtPowerCostPerUnit = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtFuelCostPerLtr = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.btnBrowse = New System.Windows.Forms.Button
        Me.dpEndDate = New System.Windows.Forms.DateTimePicker
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.dpStartDate = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtWorkbookName = New System.Windows.Forms.TextBox
        Me.txtProjectValue = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtConcreteQuantity = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtLocation = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtClient = New System.Windows.Forms.TextBox
        Me.txtProjectCode = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtProjectdescription = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.cmbProjects = New System.Windows.Forms.ComboBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.lblMessage = New System.Windows.Forms.Label
        Me.pnlDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte), True)
        Me.Label1.Location = New System.Drawing.Point(354, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(403, 26)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "SHAPOORJI PALLONJI & CO. LTD"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "xls"
        Me.SaveFileDialog1.InitialDirectory = ".\"
        Me.SaveFileDialog1.SupportMultiDottedExtensions = True
        Me.SaveFileDialog1.Title = "Select the Paht and name for the Budget Wokbook"
        '
        'btnClose
        '
        Me.btnClose.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(273, 506)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(152, 31)
        Me.btnClose.TabIndex = 16
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(537, 506)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(224, 31)
        Me.btnSave.TabIndex = 15
        Me.btnSave.Text = "Save && Proceed"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'pnlDetails
        '
        Me.pnlDetails.Controls.Add(Me.txtPowerCostPerUnit)
        Me.pnlDetails.Controls.Add(Me.Label13)
        Me.pnlDetails.Controls.Add(Me.txtFuelCostPerLtr)
        Me.pnlDetails.Controls.Add(Me.Label12)
        Me.pnlDetails.Controls.Add(Me.btnBrowse)
        Me.pnlDetails.Controls.Add(Me.dpEndDate)
        Me.pnlDetails.Controls.Add(Me.Label7)
        Me.pnlDetails.Controls.Add(Me.Label6)
        Me.pnlDetails.Controls.Add(Me.dpStartDate)
        Me.pnlDetails.Controls.Add(Me.Label5)
        Me.pnlDetails.Controls.Add(Me.txtWorkbookName)
        Me.pnlDetails.Controls.Add(Me.txtProjectValue)
        Me.pnlDetails.Controls.Add(Me.Label8)
        Me.pnlDetails.Controls.Add(Me.txtConcreteQuantity)
        Me.pnlDetails.Controls.Add(Me.Label11)
        Me.pnlDetails.Controls.Add(Me.txtLocation)
        Me.pnlDetails.Controls.Add(Me.Label4)
        Me.pnlDetails.Controls.Add(Me.Label3)
        Me.pnlDetails.Controls.Add(Me.txtClient)
        Me.pnlDetails.Controls.Add(Me.txtProjectCode)
        Me.pnlDetails.Controls.Add(Me.Label10)
        Me.pnlDetails.Controls.Add(Me.txtProjectdescription)
        Me.pnlDetails.Controls.Add(Me.Label2)
        Me.pnlDetails.Location = New System.Drawing.Point(40, 115)
        Me.pnlDetails.Name = "pnlDetails"
        Me.pnlDetails.Size = New System.Drawing.Size(990, 385)
        Me.pnlDetails.TabIndex = 2
        '
        'txtPowerCostPerUnit
        '
        Me.txtPowerCostPerUnit.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPowerCostPerUnit.Location = New System.Drawing.Point(409, 316)
        Me.txtPowerCostPerUnit.Name = "txtPowerCostPerUnit"
        Me.txtPowerCostPerUnit.Size = New System.Drawing.Size(424, 25)
        Me.txtPowerCostPerUnit.TabIndex = 12
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(208, 317)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(172, 18)
        Me.Label13.TabIndex = 34
        Me.Label13.Text = "Power Cost per Unit"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtFuelCostPerLtr
        '
        Me.txtFuelCostPerLtr.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFuelCostPerLtr.Location = New System.Drawing.Point(409, 285)
        Me.txtFuelCostPerLtr.Name = "txtFuelCostPerLtr"
        Me.txtFuelCostPerLtr.Size = New System.Drawing.Size(424, 25)
        Me.txtFuelCostPerLtr.TabIndex = 11
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(236, 288)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(145, 18)
        Me.Label12.TabIndex = 32
        Me.Label12.Text = "Fuel Cost Per Ltr"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'btnBrowse
        '
        Me.btnBrowse.Font = New System.Drawing.Font("Verdana", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowse.Location = New System.Drawing.Point(842, 345)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(102, 31)
        Me.btnBrowse.TabIndex = 14
        Me.btnBrowse.Text = "Browse..."
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'dpEndDate
        '
        Me.dpEndDate.CustomFormat = "dd-MMM-yyyy"
        Me.dpEndDate.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpEndDate.Location = New System.Drawing.Point(409, 220)
        Me.dpEndDate.Name = "dpEndDate"
        Me.dpEndDate.Size = New System.Drawing.Size(179, 25)
        Me.dpEndDate.TabIndex = 9
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(63, 347)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(334, 18)
        Me.Label7.TabIndex = 30
        Me.Label7.Text = "Name  && location to Save the workbook"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(302, 221)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(83, 18)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "End Date"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'dpStartDate
        '
        Me.dpStartDate.CustomFormat = "dd-MMM-yyyy"
        Me.dpStartDate.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpStartDate.Location = New System.Drawing.Point(409, 189)
        Me.dpStartDate.Name = "dpStartDate"
        Me.dpStartDate.Size = New System.Drawing.Size(177, 25)
        Me.dpStartDate.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(304, 190)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(93, 18)
        Me.Label5.TabIndex = 28
        Me.Label5.Text = "Start Date"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtWorkbookName
        '
        Me.txtWorkbookName.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWorkbookName.Location = New System.Drawing.Point(409, 345)
        Me.txtWorkbookName.Name = "txtWorkbookName"
        Me.txtWorkbookName.Size = New System.Drawing.Size(424, 25)
        Me.txtWorkbookName.TabIndex = 13
        '
        'txtProjectValue
        '
        Me.txtProjectValue.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProjectValue.Location = New System.Drawing.Point(409, 155)
        Me.txtProjectValue.Name = "txtProjectValue"
        Me.txtProjectValue.Size = New System.Drawing.Size(424, 25)
        Me.txtProjectValue.TabIndex = 7
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(198, 158)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(196, 18)
        Me.Label8.TabIndex = 26
        Me.Label8.Text = "Project Value in Crores"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtConcreteQuantity
        '
        Me.txtConcreteQuantity.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtConcreteQuantity.Location = New System.Drawing.Point(409, 251)
        Me.txtConcreteQuantity.Name = "txtConcreteQuantity"
        Me.txtConcreteQuantity.Size = New System.Drawing.Size(424, 25)
        Me.txtConcreteQuantity.TabIndex = 10
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(237, 254)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(157, 18)
        Me.Label11.TabIndex = 27
        Me.Label11.Text = "Concrete Quantity"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtLocation
        '
        Me.txtLocation.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocation.Location = New System.Drawing.Point(409, 120)
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.Size = New System.Drawing.Size(424, 25)
        Me.txtLocation.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(255, 125)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(140, 18)
        Me.Label4.TabIndex = 27
        Me.Label4.Text = "Project Location"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(177, 83)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(218, 18)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "Client name/Specification"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtClient
        '
        Me.txtClient.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtClient.Location = New System.Drawing.Point(409, 80)
        Me.txtClient.Name = "txtClient"
        Me.txtClient.Size = New System.Drawing.Size(424, 25)
        Me.txtClient.TabIndex = 5
        '
        'txtProjectCode
        '
        Me.txtProjectCode.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProjectCode.Location = New System.Drawing.Point(409, 8)
        Me.txtProjectCode.Name = "txtProjectCode"
        Me.txtProjectCode.Size = New System.Drawing.Size(424, 25)
        Me.txtProjectCode.TabIndex = 3
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(279, 11)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(106, 18)
        Me.Label10.TabIndex = 24
        Me.Label10.Text = "ProjectCode"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtProjectdescription
        '
        Me.txtProjectdescription.Font = New System.Drawing.Font("Verdana", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProjectdescription.Location = New System.Drawing.Point(409, 39)
        Me.txtProjectdescription.Name = "txtProjectdescription"
        Me.txtProjectdescription.Size = New System.Drawing.Size(424, 25)
        Me.txtProjectdescription.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(131, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(263, 18)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "Type the Description of Project"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Verdana", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(58, 64)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(381, 18)
        Me.Label9.TabIndex = 13
        Me.Label9.Text = "Which Project  you want to open to work with"
        '
        'cmbProjects
        '
        Me.cmbProjects.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbProjects.FormattingEnabled = True
        Me.cmbProjects.Items.AddRange(New Object() {".New Project"})
        Me.cmbProjects.Location = New System.Drawing.Point(452, 64)
        Me.cmbProjects.Name = "cmbProjects"
        Me.cmbProjects.Size = New System.Drawing.Size(463, 26)
        Me.cmbProjects.TabIndex = 1
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        '
        'Timer1
        '
        Me.Timer1.Interval = 500
        '
        'lblMessage
        '
        Me.lblMessage.AutoSize = True
        Me.lblMessage.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMessage.Location = New System.Drawing.Point(294, 548)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(437, 23)
        Me.lblMessage.TabIndex = 30
        Me.lblMessage.Text = "Form Details are loading. Please wait..."
        Me.lblMessage.Visible = False
        '
        'frmProjectDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1028, 580)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblMessage)
        Me.Controls.Add(Me.cmbProjects)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.pnlDetails)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmProjectDetails"
        Me.Text = "Form1"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlDetails.ResumeLayout(False)
        Me.pnlDetails.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents pnlDetails As System.Windows.Forms.Panel
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents dpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtWorkbookName As System.Windows.Forms.TextBox
    Friend WithEvents txtProjectValue As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtClient As System.Windows.Forms.TextBox
    Friend WithEvents txtProjectdescription As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cmbProjects As System.Windows.Forms.ComboBox
    Friend WithEvents txtProjectCode As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtConcreteQuantity As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents txtFuelCostPerLtr As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtPowerCostPerUnit As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
End Class

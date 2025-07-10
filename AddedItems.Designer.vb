<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AddedItems
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.btnMajorConcreteItems = New System.Windows.Forms.Button
        Me.btnMajorCraneItems = New System.Windows.Forms.Button
        Me.btnMajorMHItems = New System.Windows.Forms.Button
        Me.btnMajorNCItems = New System.Windows.Forms.Button
        Me.btnMajorDGItems = New System.Windows.Forms.Button
        Me.btnMajorConvItems = New System.Windows.Forms.Button
        Me.btnHiredItems = New System.Windows.Forms.Button
        Me.btnMinorItems = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.CausesValidation = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridView1.GridColor = System.Drawing.SystemColors.ControlLight
        Me.DataGridView1.Location = New System.Drawing.Point(12, 7)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(736, 426)
        Me.DataGridView1.TabIndex = 0
        '
        'btnMajorConcreteItems
        '
        Me.btnMajorConcreteItems.Location = New System.Drawing.Point(23, 437)
        Me.btnMajorConcreteItems.Name = "btnMajorConcreteItems"
        Me.btnMajorConcreteItems.Size = New System.Drawing.Size(163, 30)
        Me.btnMajorConcreteItems.TabIndex = 1
        Me.btnMajorConcreteItems.Text = "View Concrete items"
        Me.btnMajorConcreteItems.UseVisualStyleBackColor = True
        '
        'btnMajorCraneItems
        '
        Me.btnMajorCraneItems.Location = New System.Drawing.Point(220, 437)
        Me.btnMajorCraneItems.Name = "btnMajorCraneItems"
        Me.btnMajorCraneItems.Size = New System.Drawing.Size(163, 30)
        Me.btnMajorCraneItems.TabIndex = 1
        Me.btnMajorCraneItems.Text = "View Crane items"
        Me.btnMajorCraneItems.UseVisualStyleBackColor = True
        '
        'btnMajorMHItems
        '
        Me.btnMajorMHItems.Location = New System.Drawing.Point(406, 437)
        Me.btnMajorMHItems.Name = "btnMajorMHItems"
        Me.btnMajorMHItems.Size = New System.Drawing.Size(163, 30)
        Me.btnMajorMHItems.TabIndex = 1
        Me.btnMajorMHItems.Text = "View Materail handling items"
        Me.btnMajorMHItems.UseVisualStyleBackColor = True
        '
        'btnMajorNCItems
        '
        Me.btnMajorNCItems.Location = New System.Drawing.Point(589, 437)
        Me.btnMajorNCItems.Name = "btnMajorNCItems"
        Me.btnMajorNCItems.Size = New System.Drawing.Size(163, 30)
        Me.btnMajorNCItems.TabIndex = 1
        Me.btnMajorNCItems.Text = "View Non-Concrete items"
        Me.btnMajorNCItems.UseVisualStyleBackColor = True
        '
        'btnMajorDGItems
        '
        Me.btnMajorDGItems.Location = New System.Drawing.Point(23, 473)
        Me.btnMajorDGItems.Name = "btnMajorDGItems"
        Me.btnMajorDGItems.Size = New System.Drawing.Size(163, 30)
        Me.btnMajorDGItems.TabIndex = 1
        Me.btnMajorDGItems.Text = "View Generatorset  items"
        Me.btnMajorDGItems.UseVisualStyleBackColor = True
        '
        'btnMajorConvItems
        '
        Me.btnMajorConvItems.Location = New System.Drawing.Point(220, 473)
        Me.btnMajorConvItems.Name = "btnMajorConvItems"
        Me.btnMajorConvItems.Size = New System.Drawing.Size(163, 30)
        Me.btnMajorConvItems.TabIndex = 1
        Me.btnMajorConvItems.Text = "View Conveyance Vehicles"
        Me.btnMajorConvItems.UseVisualStyleBackColor = True
        '
        'btnHiredItems
        '
        Me.btnHiredItems.Location = New System.Drawing.Point(406, 473)
        Me.btnHiredItems.Name = "btnHiredItems"
        Me.btnHiredItems.Size = New System.Drawing.Size(163, 30)
        Me.btnHiredItems.TabIndex = 1
        Me.btnHiredItems.Text = "View Hired Items"
        Me.btnHiredItems.UseVisualStyleBackColor = True
        '
        'btnMinorItems
        '
        Me.btnMinorItems.Location = New System.Drawing.Point(589, 473)
        Me.btnMinorItems.Name = "btnMinorItems"
        Me.btnMinorItems.Size = New System.Drawing.Size(163, 30)
        Me.btnMinorItems.TabIndex = 1
        Me.btnMinorItems.Text = "View Minor equipments "
        Me.btnMinorItems.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(480, 510)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(140, 30)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close this Form"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEdit.Location = New System.Drawing.Point(362, 509)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(112, 31)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "Edit Entry"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(244, 510)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(112, 30)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "Delete Entry"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'AddedItems
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(755, 539)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnMajorNCItems)
        Me.Controls.Add(Me.btnMajorMHItems)
        Me.Controls.Add(Me.btnMajorCraneItems)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnMinorItems)
        Me.Controls.Add(Me.btnHiredItems)
        Me.Controls.Add(Me.btnMajorConvItems)
        Me.Controls.Add(Me.btnMajorDGItems)
        Me.Controls.Add(Me.btnMajorConcreteItems)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "AddedItems"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "AddedItems"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btnMajorConcreteItems As System.Windows.Forms.Button
    Friend WithEvents btnMajorCraneItems As System.Windows.Forms.Button
    Friend WithEvents btnMajorMHItems As System.Windows.Forms.Button
    Friend WithEvents btnMajorNCItems As System.Windows.Forms.Button
    Friend WithEvents btnMajorDGItems As System.Windows.Forms.Button
    Friend WithEvents btnMajorConvItems As System.Windows.Forms.Button
    Friend WithEvents btnHiredItems As System.Windows.Forms.Button
    Friend WithEvents btnMinorItems As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
End Class

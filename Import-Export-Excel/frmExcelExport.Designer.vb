<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExcelExport
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmExcelExport))
        Me.dgvDataToExport = New System.Windows.Forms.DataGridView()
        Me.btnImportToExcel = New System.Windows.Forms.Button()
        Me.btnRefreshExcel = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.txtFile = New System.Windows.Forms.TextBox()
        Me.lblFile = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.DataGridView4 = New System.Windows.Forms.DataGridView()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripSplitButton()
        Me.dgvChanges = New System.Windows.Forms.DataGridView()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.btnValidate = New System.Windows.Forms.Button()
        Me.btnOpenExcel = New System.Windows.Forms.Button()
        CType(Me.dgvDataToExport, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.dgvChanges, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvDataToExport
        '
        Me.dgvDataToExport.AllowUserToAddRows = False
        Me.dgvDataToExport.AllowUserToDeleteRows = False
        Me.dgvDataToExport.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgvDataToExport.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDataToExport.Location = New System.Drawing.Point(10, 128)
        Me.dgvDataToExport.Name = "dgvDataToExport"
        Me.dgvDataToExport.Size = New System.Drawing.Size(1097, 194)
        Me.dgvDataToExport.TabIndex = 0
        '
        'btnImportToExcel
        '
        Me.btnImportToExcel.Location = New System.Drawing.Point(233, 39)
        Me.btnImportToExcel.Name = "btnImportToExcel"
        Me.btnImportToExcel.Size = New System.Drawing.Size(75, 23)
        Me.btnImportToExcel.TabIndex = 1
        Me.btnImportToExcel.Text = "Import Excel"
        Me.btnImportToExcel.UseVisualStyleBackColor = True
        '
        'btnRefreshExcel
        '
        Me.btnRefreshExcel.Location = New System.Drawing.Point(111, 39)
        Me.btnRefreshExcel.Name = "btnRefreshExcel"
        Me.btnRefreshExcel.Size = New System.Drawing.Size(94, 23)
        Me.btnRefreshExcel.TabIndex = 2
        Me.btnRefreshExcel.Text = "Refresh Excel"
        Me.btnRefreshExcel.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'txtFile
        '
        Me.txtFile.Location = New System.Drawing.Point(105, 82)
        Me.txtFile.Name = "txtFile"
        Me.txtFile.Size = New System.Drawing.Size(455, 20)
        Me.txtFile.TabIndex = 3
        '
        'lblFile
        '
        Me.lblFile.AutoSize = True
        Me.lblFile.Location = New System.Drawing.Point(73, 88)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(26, 13)
        Me.lblFile.TabIndex = 4
        Me.lblFile.Text = "File:"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(11, 328)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1096, 64)
        Me.DataGridView1.TabIndex = 5
        '
        'DataGridView2
        '
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Location = New System.Drawing.Point(12, 409)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.Size = New System.Drawing.Size(1096, 77)
        Me.DataGridView2.TabIndex = 6
        '
        'DataGridView3
        '
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Location = New System.Drawing.Point(12, 504)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.Size = New System.Drawing.Size(1096, 127)
        Me.DataGridView3.TabIndex = 7
        '
        'DataGridView4
        '
        Me.DataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView4.Location = New System.Drawing.Point(12, 637)
        Me.DataGridView4.Name = "DataGridView4"
        Me.DataGridView4.Size = New System.Drawing.Size(1096, 135)
        Me.DataGridView4.TabIndex = 8
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 916)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1134, 22)
        Me.StatusStrip1.TabIndex = 9
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(121, 17)
        Me.ToolStripStatusLabel1.Text = "ToolStripStatusLabel1"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Image = CType(resources.GetObject("ToolStripStatusLabel2.Image"), System.Drawing.Image)
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(153, 20)
        Me.ToolStripStatusLabel2.Text = "ToolStripStatusLabel2"
        '
        'dgvChanges
        '
        Me.dgvChanges.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvChanges.Location = New System.Drawing.Point(11, 778)
        Me.dgvChanges.Name = "dgvChanges"
        Me.dgvChanges.Size = New System.Drawing.Size(1096, 135)
        Me.dgvChanges.TabIndex = 10
        '
        'Timer1
        '
        '
        'btnValidate
        '
        Me.btnValidate.Location = New System.Drawing.Point(336, 39)
        Me.btnValidate.Name = "btnValidate"
        Me.btnValidate.Size = New System.Drawing.Size(75, 23)
        Me.btnValidate.TabIndex = 11
        Me.btnValidate.Text = "Validate"
        Me.btnValidate.UseVisualStyleBackColor = True
        '
        'btnOpenExcel
        '
        Me.btnOpenExcel.Location = New System.Drawing.Point(476, 39)
        Me.btnOpenExcel.Name = "btnOpenExcel"
        Me.btnOpenExcel.Size = New System.Drawing.Size(75, 23)
        Me.btnOpenExcel.TabIndex = 12
        Me.btnOpenExcel.Text = "Open Excel File"
        Me.btnOpenExcel.UseVisualStyleBackColor = True
        '
        'frmExcelExport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1134, 938)
        Me.Controls.Add(Me.btnOpenExcel)
        Me.Controls.Add(Me.btnValidate)
        Me.Controls.Add(Me.dgvChanges)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.DataGridView4)
        Me.Controls.Add(Me.DataGridView3)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.lblFile)
        Me.Controls.Add(Me.txtFile)
        Me.Controls.Add(Me.btnRefreshExcel)
        Me.Controls.Add(Me.btnImportToExcel)
        Me.Controls.Add(Me.dgvDataToExport)
        Me.Name = "frmExcelExport"
        Me.Text = "Import-Export"
        CType(Me.dgvDataToExport, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.dgvChanges, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvDataToExport As System.Windows.Forms.DataGridView
    Friend WithEvents btnImportToExcel As System.Windows.Forms.Button
    Friend WithEvents btnRefreshExcel As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtFile As System.Windows.Forms.TextBox
    Friend WithEvents lblFile As System.Windows.Forms.Label
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView3 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView4 As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents dgvChanges As System.Windows.Forms.DataGridView
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripSplitButton
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents btnValidate As System.Windows.Forms.Button
    Friend WithEvents btnOpenExcel As System.Windows.Forms.Button

End Class

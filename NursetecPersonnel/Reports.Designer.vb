<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Reports
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Reports))
        Me.dgvDataBreaks = New System.Windows.Forms.DataGridView()
        Me.txtEmpID = New System.Windows.Forms.TextBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbBreakReport = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtpStart = New System.Windows.Forms.DateTimePicker()
        Me.dtpEnd = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnReset = New System.Windows.Forms.Button()
        Me.cmbEmployee = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.tspgbar = New System.Windows.Forms.ToolStripProgressBar()
        Me.lblProg = New System.Windows.Forms.ToolStripStatusLabel()
        Me.rdbtnExcel = New System.Windows.Forms.RadioButton()
        Me.rdbtnPDF = New System.Windows.Forms.RadioButton()
        Me.rdbtnPrinter = New System.Windows.Forms.RadioButton()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmbOffice = New System.Windows.Forms.ComboBox()
        Me.dgvPresent = New System.Windows.Forms.DataGridView()
        Me.Emp_Name = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ClockDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Time_In = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Time_Out = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Total_Hours = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Actual_Hours = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BreakNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BreakHours = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Notes = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgvDataBreaks, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.dgvPresent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvDataBreaks
        '
        Me.dgvDataBreaks.Anchor = System.Windows.Forms.AnchorStyles.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvDataBreaks.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvDataBreaks.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDataBreaks.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Emp_Name, Me.ClockDate, Me.Time_In, Me.Time_Out, Me.Total_Hours, Me.Actual_Hours, Me.BreakNo, Me.BreakHours, Me.Notes})
        Me.dgvDataBreaks.Location = New System.Drawing.Point(70, 266)
        Me.dgvDataBreaks.Name = "dgvDataBreaks"
        Me.dgvDataBreaks.RowHeadersWidth = 51
        Me.dgvDataBreaks.RowTemplate.Height = 24
        Me.dgvDataBreaks.Size = New System.Drawing.Size(1135, 296)
        Me.dgvDataBreaks.TabIndex = 0
        '
        'txtEmpID
        '
        Me.txtEmpID.Location = New System.Drawing.Point(125, 201)
        Me.txtEmpID.Name = "txtEmpID"
        Me.txtEmpID.Size = New System.Drawing.Size(121, 22)
        Me.txtEmpID.TabIndex = 1
        '
        'btnSearch
        '
        Me.btnSearch.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnSearch.Location = New System.Drawing.Point(1110, 568)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(95, 37)
        Me.btnSearch.TabIndex = 2
        Me.btnSearch.Text = "Generate"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(109, 172)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "EmployeeID"
        '
        'cmbBreakReport
        '
        Me.cmbBreakReport.FormattingEnabled = True
        Me.cmbBreakReport.Location = New System.Drawing.Point(127, 85)
        Me.cmbBreakReport.Name = "cmbBreakReport"
        Me.cmbBreakReport.Size = New System.Drawing.Size(121, 24)
        Me.cmbBreakReport.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(111, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(94, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Select Report"
        '
        'dtpStart
        '
        Me.dtpStart.Location = New System.Drawing.Point(333, 87)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.Size = New System.Drawing.Size(200, 22)
        Me.dtpStart.TabIndex = 7
        '
        'dtpEnd
        '
        Me.dtpEnd.Location = New System.Drawing.Point(596, 87)
        Me.dtpEnd.Name = "dtpEnd"
        Me.dtpEnd.Size = New System.Drawing.Size(200, 22)
        Me.dtpEnd.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(320, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 17)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Start Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(584, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 17)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "End Date"
        '
        'btnReset
        '
        Me.btnReset.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnReset.Location = New System.Drawing.Point(71, 568)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(134, 37)
        Me.btnReset.TabIndex = 11
        Me.btnReset.Text = "Reset Selection"
        Me.btnReset.UseVisualStyleBackColor = True
        '
        'cmbEmployee
        '
        Me.cmbEmployee.FormattingEnabled = True
        Me.cmbEmployee.Location = New System.Drawing.Point(331, 199)
        Me.cmbEmployee.Name = "cmbEmployee"
        Me.cmbEmployee.Size = New System.Drawing.Size(121, 24)
        Me.cmbEmployee.TabIndex = 12
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(318, 172)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(111, 17)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Employee Name"
        '
        'btnExport
        '
        Me.btnExport.Location = New System.Drawing.Point(1050, 201)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(74, 32)
        Me.btnExport.TabIndex = 14
        Me.btnExport.Text = "Print"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tspgbar, Me.lblProg})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 639)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1259, 22)
        Me.StatusStrip1.TabIndex = 15
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tspgbar
        '
        Me.tspgbar.Name = "tspgbar"
        Me.tspgbar.Size = New System.Drawing.Size(100, 16)
        Me.tspgbar.Visible = False
        '
        'lblProg
        '
        Me.lblProg.Name = "lblProg"
        Me.lblProg.Size = New System.Drawing.Size(0, 16)
        '
        'rdbtnExcel
        '
        Me.rdbtnExcel.AutoSize = True
        Me.rdbtnExcel.Location = New System.Drawing.Point(920, 163)
        Me.rdbtnExcel.Name = "rdbtnExcel"
        Me.rdbtnExcel.Size = New System.Drawing.Size(62, 21)
        Me.rdbtnExcel.TabIndex = 16
        Me.rdbtnExcel.TabStop = True
        Me.rdbtnExcel.Text = "Excel"
        Me.rdbtnExcel.UseVisualStyleBackColor = True
        '
        'rdbtnPDF
        '
        Me.rdbtnPDF.AutoSize = True
        Me.rdbtnPDF.Location = New System.Drawing.Point(988, 163)
        Me.rdbtnPDF.Name = "rdbtnPDF"
        Me.rdbtnPDF.Size = New System.Drawing.Size(56, 21)
        Me.rdbtnPDF.TabIndex = 17
        Me.rdbtnPDF.TabStop = True
        Me.rdbtnPDF.Text = "PDF"
        Me.rdbtnPDF.UseVisualStyleBackColor = True
        '
        'rdbtnPrinter
        '
        Me.rdbtnPrinter.AutoSize = True
        Me.rdbtnPrinter.Location = New System.Drawing.Point(1050, 163)
        Me.rdbtnPrinter.Name = "rdbtnPrinter"
        Me.rdbtnPrinter.Size = New System.Drawing.Size(71, 21)
        Me.rdbtnPrinter.TabIndex = 18
        Me.rdbtnPrinter.TabStop = True
        Me.rdbtnPrinter.Text = "Printer"
        Me.rdbtnPrinter.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(517, 174)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(86, 17)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = "Office Name"
        '
        'cmbOffice
        '
        Me.cmbOffice.FormattingEnabled = True
        Me.cmbOffice.Location = New System.Drawing.Point(530, 201)
        Me.cmbOffice.Name = "cmbOffice"
        Me.cmbOffice.Size = New System.Drawing.Size(121, 24)
        Me.cmbOffice.TabIndex = 19
        '
        'dgvPresent
        '
        Me.dgvPresent.Anchor = System.Windows.Forms.AnchorStyles.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvPresent.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvPresent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPresent.Location = New System.Drawing.Point(70, 266)
        Me.dgvPresent.Name = "dgvPresent"
        Me.dgvPresent.RowHeadersWidth = 51
        Me.dgvPresent.RowTemplate.Height = 24
        Me.dgvPresent.Size = New System.Drawing.Size(1135, 296)
        Me.dgvPresent.TabIndex = 21
        '
        'Emp_Name
        '
        Me.Emp_Name.DataPropertyName = "Emp_Name"
        Me.Emp_Name.HeaderText = "Employee Name"
        Me.Emp_Name.MinimumWidth = 6
        Me.Emp_Name.Name = "Emp_Name"
        Me.Emp_Name.Width = 125
        '
        'ClockDate
        '
        Me.ClockDate.DataPropertyName = "ClockDate"
        Me.ClockDate.HeaderText = "Date"
        Me.ClockDate.MinimumWidth = 6
        Me.ClockDate.Name = "ClockDate"
        Me.ClockDate.Width = 125
        '
        'Time_In
        '
        Me.Time_In.DataPropertyName = "Time_In"
        Me.Time_In.HeaderText = "Time In"
        Me.Time_In.MinimumWidth = 6
        Me.Time_In.Name = "Time_In"
        Me.Time_In.Width = 125
        '
        'Time_Out
        '
        Me.Time_Out.DataPropertyName = "Time_Out"
        Me.Time_Out.HeaderText = "Time Out"
        Me.Time_Out.MinimumWidth = 6
        Me.Time_Out.Name = "Time_Out"
        Me.Time_Out.Width = 125
        '
        'Total_Hours
        '
        Me.Total_Hours.DataPropertyName = "Total_Hours"
        Me.Total_Hours.HeaderText = "Total Hours"
        Me.Total_Hours.MinimumWidth = 6
        Me.Total_Hours.Name = "Total_Hours"
        Me.Total_Hours.Width = 125
        '
        'Actual_Hours
        '
        Me.Actual_Hours.DataPropertyName = "Actual_Hours"
        Me.Actual_Hours.HeaderText = "Actual_Hours"
        Me.Actual_Hours.MinimumWidth = 6
        Me.Actual_Hours.Name = "Actual_Hours"
        Me.Actual_Hours.Width = 125
        '
        'BreakNo
        '
        Me.BreakNo.DataPropertyName = "Breaks"
        Me.BreakNo.HeaderText = "Breaks"
        Me.BreakNo.MinimumWidth = 6
        Me.BreakNo.Name = "BreakNo"
        Me.BreakNo.Width = 125
        '
        'BreakHours
        '
        Me.BreakHours.DataPropertyName = "BreakHours"
        Me.BreakHours.HeaderText = "Break Hours"
        Me.BreakHours.MinimumWidth = 6
        Me.BreakHours.Name = "BreakHours"
        Me.BreakHours.Width = 125
        '
        'Notes
        '
        Me.Notes.DataPropertyName = "Notes"
        Me.Notes.HeaderText = "Notes"
        Me.Notes.MinimumWidth = 6
        Me.Notes.Name = "Notes"
        Me.Notes.Width = 125
        '
        'Reports
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1259, 661)
        Me.Controls.Add(Me.dgvPresent)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.cmbOffice)
        Me.Controls.Add(Me.rdbtnPrinter)
        Me.Controls.Add(Me.rdbtnPDF)
        Me.Controls.Add(Me.rdbtnExcel)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.txtEmpID)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.cmbEmployee)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtpEnd)
        Me.Controls.Add(Me.dtpStart)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmbBreakReport)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.dgvDataBreaks)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Reports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Breaks"
        Me.TopMost = True
        CType(Me.dgvDataBreaks, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.dgvPresent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgvDataBreaks As DataGridView
    Friend WithEvents txtEmpID As TextBox
    Friend WithEvents btnSearch As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents cmbBreakReport As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents dtpStart As DateTimePicker
    Friend WithEvents dtpEnd As DateTimePicker
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents btnReset As Button
    Friend WithEvents cmbEmployee As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents btnExport As Button
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents tspgbar As ToolStripProgressBar
    Friend WithEvents lblProg As ToolStripStatusLabel
    Friend WithEvents rdbtnExcel As RadioButton
    Friend WithEvents rdbtnPDF As RadioButton
    Friend WithEvents rdbtnPrinter As RadioButton
    Friend WithEvents Label6 As Label
    Friend WithEvents cmbOffice As ComboBox
    Friend WithEvents dgvPresent As DataGridView
    Friend WithEvents Emp_Name As DataGridViewTextBoxColumn
    Friend WithEvents ClockDate As DataGridViewTextBoxColumn
    Friend WithEvents Time_In As DataGridViewTextBoxColumn
    Friend WithEvents Time_Out As DataGridViewTextBoxColumn
    Friend WithEvents Total_Hours As DataGridViewTextBoxColumn
    Friend WithEvents Actual_Hours As DataGridViewTextBoxColumn
    Friend WithEvents BreakNo As DataGridViewTextBoxColumn
    Friend WithEvents BreakHours As DataGridViewTextBoxColumn
    Friend WithEvents Notes As DataGridViewTextBoxColumn
End Class

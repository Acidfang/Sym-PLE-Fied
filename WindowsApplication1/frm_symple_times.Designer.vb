<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_symple_times
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.dgv_symple = New System.Windows.Forms.DataGridView()
        Me.dgc_symple_section = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_symple_quantity = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_symple_unit = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_symple_type = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_symple_date = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.b_send_info = New System.Windows.Forms.Button()
        Me.dgv_summary = New System.Windows.Forms.DataGridView()
        Me.dgc_summary_setup = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_summary_run = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_summary_bs01 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_summary_cs01 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_summary_es01 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_summary_produced = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.tb_prod_num = New System.Windows.Forms.TextBox()
        Me.sb_job_select = New System.Windows.Forms.HScrollBar()
        Me.l_job_info = New System.Windows.Forms.Label()
        Me.b_final_confirm = New System.Windows.Forms.Button()
        Me.dgv_pending = New System.Windows.Forms.DataGridView()
        Me.dgc_pending_section = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_pending_quantity = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_pending_unit = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_pending_type = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_pending_send = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.dgc_pending_string = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cb_difference = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.b_refresh = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.nud_time_offset = New System.Windows.Forms.NumericUpDown()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        CType(Me.dgv_symple, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgv_summary, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgv_pending, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.nud_time_offset, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgv_symple
        '
        Me.dgv_symple.AllowUserToAddRows = False
        Me.dgv_symple.AllowUserToDeleteRows = False
        Me.dgv_symple.AllowUserToResizeColumns = False
        Me.dgv_symple.AllowUserToResizeRows = False
        Me.dgv_symple.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_symple.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgc_symple_section, Me.dgc_symple_quantity, Me.dgc_symple_unit, Me.dgc_symple_type, Me.dgc_symple_date})
        Me.dgv_symple.Location = New System.Drawing.Point(6, 22)
        Me.dgv_symple.MultiSelect = False
        Me.dgv_symple.Name = "dgv_symple"
        Me.dgv_symple.ReadOnly = True
        Me.dgv_symple.RowHeadersVisible = False
        Me.dgv_symple.Size = New System.Drawing.Size(392, 203)
        Me.dgv_symple.TabIndex = 0
        '
        'dgc_symple_section
        '
        Me.dgc_symple_section.HeaderText = "Section"
        Me.dgc_symple_section.Name = "dgc_symple_section"
        Me.dgc_symple_section.ReadOnly = True
        Me.dgc_symple_section.Width = 60
        '
        'dgc_symple_quantity
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.dgc_symple_quantity.DefaultCellStyle = DataGridViewCellStyle1
        Me.dgc_symple_quantity.HeaderText = "Quantity"
        Me.dgc_symple_quantity.Name = "dgc_symple_quantity"
        Me.dgc_symple_quantity.ReadOnly = True
        Me.dgc_symple_quantity.Width = 60
        '
        'dgc_symple_unit
        '
        Me.dgc_symple_unit.HeaderText = "Unit"
        Me.dgc_symple_unit.Name = "dgc_symple_unit"
        Me.dgc_symple_unit.ReadOnly = True
        Me.dgc_symple_unit.Width = 60
        '
        'dgc_symple_type
        '
        Me.dgc_symple_type.HeaderText = "Type"
        Me.dgc_symple_type.Name = "dgc_symple_type"
        Me.dgc_symple_type.ReadOnly = True
        Me.dgc_symple_type.Width = 60
        '
        'dgc_symple_date
        '
        Me.dgc_symple_date.HeaderText = "Date"
        Me.dgc_symple_date.Name = "dgc_symple_date"
        Me.dgc_symple_date.ReadOnly = True
        Me.dgc_symple_date.Width = 130
        '
        'b_send_info
        '
        Me.b_send_info.Location = New System.Drawing.Point(417, 240)
        Me.b_send_info.Name = "b_send_info"
        Me.b_send_info.Size = New System.Drawing.Size(170, 70)
        Me.b_send_info.TabIndex = 1
        Me.b_send_info.Text = "Send Info" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Send times to symple"
        Me.b_send_info.UseVisualStyleBackColor = True
        '
        'dgv_summary
        '
        Me.dgv_summary.AllowUserToAddRows = False
        Me.dgv_summary.AllowUserToDeleteRows = False
        Me.dgv_summary.AllowUserToResizeColumns = False
        Me.dgv_summary.AllowUserToResizeRows = False
        Me.dgv_summary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_summary.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgc_summary_setup, Me.dgc_summary_run, Me.dgc_summary_bs01, Me.dgc_summary_cs01, Me.dgc_summary_es01, Me.dgc_summary_produced})
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgv_summary.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgv_summary.Enabled = False
        Me.dgv_summary.Location = New System.Drawing.Point(6, 15)
        Me.dgv_summary.MultiSelect = False
        Me.dgv_summary.Name = "dgv_summary"
        Me.dgv_summary.ReadOnly = True
        Me.dgv_summary.RowHeadersVisible = False
        Me.dgv_summary.Size = New System.Drawing.Size(308, 43)
        Me.dgv_summary.TabIndex = 22
        '
        'dgc_summary_setup
        '
        Me.dgc_summary_setup.HeaderText = "Setup"
        Me.dgc_summary_setup.Name = "dgc_summary_setup"
        Me.dgc_summary_setup.ReadOnly = True
        Me.dgc_summary_setup.Width = 45
        '
        'dgc_summary_run
        '
        Me.dgc_summary_run.HeaderText = "Run"
        Me.dgc_summary_run.Name = "dgc_summary_run"
        Me.dgc_summary_run.ReadOnly = True
        Me.dgc_summary_run.Width = 45
        '
        'dgc_summary_bs01
        '
        Me.dgc_summary_bs01.HeaderText = "BS01"
        Me.dgc_summary_bs01.Name = "dgc_summary_bs01"
        Me.dgc_summary_bs01.ReadOnly = True
        Me.dgc_summary_bs01.Width = 45
        '
        'dgc_summary_cs01
        '
        Me.dgc_summary_cs01.HeaderText = "CS01"
        Me.dgc_summary_cs01.Name = "dgc_summary_cs01"
        Me.dgc_summary_cs01.ReadOnly = True
        Me.dgc_summary_cs01.Width = 45
        '
        'dgc_summary_es01
        '
        Me.dgc_summary_es01.HeaderText = "ES01"
        Me.dgc_summary_es01.Name = "dgc_summary_es01"
        Me.dgc_summary_es01.ReadOnly = True
        Me.dgc_summary_es01.Width = 45
        '
        'dgc_summary_produced
        '
        Me.dgc_summary_produced.HeaderText = "Produced"
        Me.dgc_summary_produced.Name = "dgc_summary_produced"
        Me.dgc_summary_produced.ReadOnly = True
        Me.dgc_summary_produced.Width = 80
        '
        'tb_prod_num
        '
        Me.tb_prod_num.Location = New System.Drawing.Point(320, 15)
        Me.tb_prod_num.Name = "tb_prod_num"
        Me.tb_prod_num.ReadOnly = True
        Me.tb_prod_num.Size = New System.Drawing.Size(78, 20)
        Me.tb_prod_num.TabIndex = 24
        Me.tb_prod_num.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'sb_job_select
        '
        Me.sb_job_select.LargeChange = 1
        Me.sb_job_select.Location = New System.Drawing.Point(320, 38)
        Me.sb_job_select.Name = "sb_job_select"
        Me.sb_job_select.Size = New System.Drawing.Size(78, 20)
        Me.sb_job_select.TabIndex = 25
        '
        'l_job_info
        '
        Me.l_job_info.AutoSize = True
        Me.l_job_info.Location = New System.Drawing.Point(251, 0)
        Me.l_job_info.Name = "l_job_info"
        Me.l_job_info.Size = New System.Drawing.Size(34, 13)
        Me.l_job_info.TabIndex = 26
        Me.l_job_info.Text = "1 of 1"
        '
        'b_final_confirm
        '
        Me.b_final_confirm.Location = New System.Drawing.Point(593, 240)
        Me.b_final_confirm.Name = "b_final_confirm"
        Me.b_final_confirm.Size = New System.Drawing.Size(148, 70)
        Me.b_final_confirm.TabIndex = 27
        Me.b_final_confirm.Text = "Final Confirm" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Close off Job)"
        Me.b_final_confirm.UseVisualStyleBackColor = True
        '
        'dgv_pending
        '
        Me.dgv_pending.AllowUserToAddRows = False
        Me.dgv_pending.AllowUserToDeleteRows = False
        Me.dgv_pending.AllowUserToResizeColumns = False
        Me.dgv_pending.AllowUserToResizeRows = False
        Me.dgv_pending.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_pending.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgc_pending_section, Me.dgc_pending_quantity, Me.dgc_pending_unit, Me.dgc_pending_type, Me.dgc_pending_send, Me.dgc_pending_string})
        Me.dgv_pending.Location = New System.Drawing.Point(6, 22)
        Me.dgv_pending.MultiSelect = False
        Me.dgv_pending.Name = "dgv_pending"
        Me.dgv_pending.ReadOnly = True
        Me.dgv_pending.RowHeadersVisible = False
        Me.dgv_pending.Size = New System.Drawing.Size(311, 203)
        Me.dgv_pending.TabIndex = 28
        '
        'dgc_pending_section
        '
        Me.dgc_pending_section.HeaderText = "Section"
        Me.dgc_pending_section.Name = "dgc_pending_section"
        Me.dgc_pending_section.ReadOnly = True
        Me.dgc_pending_section.Width = 60
        '
        'dgc_pending_quantity
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.dgc_pending_quantity.DefaultCellStyle = DataGridViewCellStyle3
        Me.dgc_pending_quantity.HeaderText = "Quantity"
        Me.dgc_pending_quantity.Name = "dgc_pending_quantity"
        Me.dgc_pending_quantity.ReadOnly = True
        Me.dgc_pending_quantity.Width = 60
        '
        'dgc_pending_unit
        '
        Me.dgc_pending_unit.HeaderText = "Unit"
        Me.dgc_pending_unit.Name = "dgc_pending_unit"
        Me.dgc_pending_unit.ReadOnly = True
        Me.dgc_pending_unit.Width = 60
        '
        'dgc_pending_type
        '
        Me.dgc_pending_type.HeaderText = "Type"
        Me.dgc_pending_type.Name = "dgc_pending_type"
        Me.dgc_pending_type.ReadOnly = True
        Me.dgc_pending_type.Width = 60
        '
        'dgc_pending_send
        '
        Me.dgc_pending_send.HeaderText = "Send"
        Me.dgc_pending_send.Name = "dgc_pending_send"
        Me.dgc_pending_send.ReadOnly = True
        Me.dgc_pending_send.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgc_pending_send.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.dgc_pending_send.Width = 50
        '
        'dgc_pending_string
        '
        Me.dgc_pending_string.HeaderText = "string"
        Me.dgc_pending_string.Name = "dgc_pending_string"
        Me.dgc_pending_string.ReadOnly = True
        Me.dgc_pending_string.Visible = False
        '
        'cb_difference
        '
        Me.cb_difference.AutoSize = True
        Me.cb_difference.Location = New System.Drawing.Point(156, 1)
        Me.cb_difference.Name = "cb_difference"
        Me.cb_difference.Size = New System.Drawing.Size(158, 17)
        Me.cb_difference.TabIndex = 29
        Me.cb_difference.Text = "Show Calculated Difference"
        Me.cb_difference.UseVisualStyleBackColor = True
        Me.cb_difference.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.b_refresh)
        Me.GroupBox1.Controls.Add(Me.dgv_symple)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.nud_time_offset)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(404, 232)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Times in Sym-PLE"
        '
        'b_refresh
        '
        Me.b_refresh.Location = New System.Drawing.Point(295, -1)
        Me.b_refresh.Name = "b_refresh"
        Me.b_refresh.Size = New System.Drawing.Size(96, 23)
        Me.b_refresh.TabIndex = 35
        Me.b_refresh.Text = "Refresh Times"
        Me.b_refresh.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(228, 4)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(61, 13)
        Me.Label2.TabIndex = 34
        Me.Label2.Text = "Time Offset"
        '
        'nud_time_offset
        '
        Me.nud_time_offset.Location = New System.Drawing.Point(178, 0)
        Me.nud_time_offset.Maximum = New Decimal(New Integer() {60, 0, 0, 0})
        Me.nud_time_offset.Minimum = New Decimal(New Integer() {60, 0, 0, -2147483648})
        Me.nud_time_offset.Name = "nud_time_offset"
        Me.nud_time_offset.Size = New System.Drawing.Size(44, 20)
        Me.nud_time_offset.TabIndex = 33
        Me.nud_time_offset.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.nud_time_offset.Value = New Decimal(New Integer() {5, 0, 0, -2147483648})
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.dgv_pending)
        Me.GroupBox2.Controls.Add(Me.cb_difference)
        Me.GroupBox2.Location = New System.Drawing.Point(416, 2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(325, 232)
        Me.GroupBox2.TabIndex = 31
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Times to enter"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.dgv_summary)
        Me.GroupBox3.Controls.Add(Me.sb_job_select)
        Me.GroupBox3.Controls.Add(Me.tb_prod_num)
        Me.GroupBox3.Controls.Add(Me.l_job_info)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 240)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(404, 70)
        Me.GroupBox3.TabIndex = 32
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Time on Summary"
        '
        'frm_symple_times
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(745, 317)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.b_final_confirm)
        Me.Controls.Add(Me.b_send_info)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_symple_times"
        Me.Text = "Sym-PLE Entry"
        CType(Me.dgv_symple, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgv_summary, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgv_pending, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.nud_time_offset, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgv_symple As System.Windows.Forms.DataGridView
    Friend WithEvents b_send_info As System.Windows.Forms.Button
    Friend WithEvents dgv_summary As System.Windows.Forms.DataGridView
    Friend WithEvents tb_prod_num As System.Windows.Forms.TextBox
    Friend WithEvents sb_job_select As System.Windows.Forms.HScrollBar
    Friend WithEvents l_job_info As System.Windows.Forms.Label
    Friend WithEvents b_final_confirm As System.Windows.Forms.Button
    Friend WithEvents dgv_pending As System.Windows.Forms.DataGridView
    Friend WithEvents cb_difference As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents nud_time_offset As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents b_refresh As System.Windows.Forms.Button
    Friend WithEvents dgc_symple_section As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_symple_quantity As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_symple_unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_symple_type As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_symple_date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_pending_section As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_pending_quantity As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_pending_unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_pending_type As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_pending_send As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents dgc_pending_string As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_summary_setup As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_summary_run As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_summary_bs01 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_summary_cs01 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_summary_es01 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_summary_produced As System.Windows.Forms.DataGridViewTextBoxColumn
End Class

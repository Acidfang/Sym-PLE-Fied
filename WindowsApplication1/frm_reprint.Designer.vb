<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_reprint
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.dgv_reprint = New System.Windows.Forms.DataGridView()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.dgc_reprint_num_out = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_reprint_kg_out = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_reprint_km_out = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_reprint_reprint = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        CType(Me.dgv_reprint, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgv_reprint
        '
        Me.dgv_reprint.AllowUserToAddRows = False
        Me.dgv_reprint.AllowUserToDeleteRows = False
        Me.dgv_reprint.AllowUserToResizeColumns = False
        Me.dgv_reprint.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgv_reprint.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgv_reprint.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_reprint.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgc_reprint_num_out, Me.dgc_reprint_kg_out, Me.dgc_reprint_km_out, Me.dgc_reprint_reprint})
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgv_reprint.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgv_reprint.Location = New System.Drawing.Point(8, 12)
        Me.dgv_reprint.Name = "dgv_reprint"
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgv_reprint.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgv_reprint.RowHeadersVisible = False
        Me.dgv_reprint.Size = New System.Drawing.Size(204, 256)
        Me.dgv_reprint.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(8, 274)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(62, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Check All"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(76, 274)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(136, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Reprint Selected"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'dgc_reprint_num_out
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.dgc_reprint_num_out.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgc_reprint_num_out.HeaderText = "#"
        Me.dgc_reprint_num_out.Name = "dgc_reprint_num_out"
        Me.dgc_reprint_num_out.ReadOnly = True
        Me.dgc_reprint_num_out.Width = 30
        '
        'dgc_reprint_kg_out
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.dgc_reprint_kg_out.DefaultCellStyle = DataGridViewCellStyle3
        Me.dgc_reprint_kg_out.HeaderText = "KG"
        Me.dgc_reprint_kg_out.Name = "dgc_reprint_kg_out"
        Me.dgc_reprint_kg_out.ReadOnly = True
        Me.dgc_reprint_kg_out.Width = 50
        '
        'dgc_reprint_km_out
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.dgc_reprint_km_out.DefaultCellStyle = DataGridViewCellStyle4
        Me.dgc_reprint_km_out.HeaderText = "KM"
        Me.dgc_reprint_km_out.Name = "dgc_reprint_km_out"
        Me.dgc_reprint_km_out.ReadOnly = True
        Me.dgc_reprint_km_out.Width = 50
        '
        'dgc_reprint_reprint
        '
        Me.dgc_reprint_reprint.HeaderText = "Reprint"
        Me.dgc_reprint_reprint.Name = "dgc_reprint_reprint"
        Me.dgc_reprint_reprint.ReadOnly = True
        Me.dgc_reprint_reprint.Width = 50
        '
        'frm_reprint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(219, 304)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.dgv_reprint)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_reprint"
        Me.Text = "Ticket Reprinting"
        CType(Me.dgv_reprint, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgv_reprint As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents dgc_reprint_num_out As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_reprint_kg_out As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_reprint_km_out As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_reprint_reprint As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class

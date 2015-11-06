<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_startup
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
        Me.cbo_machine = New System.Windows.Forms.ComboBox()
        Me.lb_shifts = New System.Windows.Forms.ListBox()
        Me.b_ok = New System.Windows.Forms.Button()
        Me.cbo_department = New System.Windows.Forms.ComboBox()
        Me.b_restore = New System.Windows.Forms.Button()
        Me.cb_ofline_mode = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'cbo_machine
        '
        Me.cbo_machine.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbo_machine.FormattingEnabled = True
        Me.cbo_machine.Location = New System.Drawing.Point(7, 35)
        Me.cbo_machine.Name = "cbo_machine"
        Me.cbo_machine.Size = New System.Drawing.Size(121, 21)
        Me.cbo_machine.TabIndex = 0
        '
        'lb_shifts
        '
        Me.lb_shifts.FormattingEnabled = True
        Me.lb_shifts.Location = New System.Drawing.Point(7, 62)
        Me.lb_shifts.Name = "lb_shifts"
        Me.lb_shifts.Size = New System.Drawing.Size(120, 17)
        Me.lb_shifts.TabIndex = 1
        '
        'b_ok
        '
        Me.b_ok.Location = New System.Drawing.Point(7, 85)
        Me.b_ok.Name = "b_ok"
        Me.b_ok.Size = New System.Drawing.Size(54, 23)
        Me.b_ok.TabIndex = 2
        Me.b_ok.Text = "&OK"
        Me.b_ok.UseVisualStyleBackColor = True
        '
        'cbo_department
        '
        Me.cbo_department.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbo_department.FormattingEnabled = True
        Me.cbo_department.Location = New System.Drawing.Point(7, 8)
        Me.cbo_department.Name = "cbo_department"
        Me.cbo_department.Size = New System.Drawing.Size(121, 21)
        Me.cbo_department.TabIndex = 3
        '
        'b_restore
        '
        Me.b_restore.Location = New System.Drawing.Point(67, 85)
        Me.b_restore.Name = "b_restore"
        Me.b_restore.Size = New System.Drawing.Size(60, 23)
        Me.b_restore.TabIndex = 4
        Me.b_restore.Text = "&Restore"
        Me.b_restore.UseVisualStyleBackColor = True
        '
        'cb_ofline_mode
        '
        Me.cb_ofline_mode.AutoSize = True
        Me.cb_ofline_mode.Location = New System.Drawing.Point(23, 114)
        Me.cb_ofline_mode.Name = "cb_ofline_mode"
        Me.cb_ofline_mode.Size = New System.Drawing.Size(86, 17)
        Me.cb_ofline_mode.TabIndex = 5
        Me.cb_ofline_mode.Text = "Offline Mode"
        Me.cb_ofline_mode.UseVisualStyleBackColor = True
        '
        'frm_startup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(135, 138)
        Me.ControlBox = False
        Me.Controls.Add(Me.cb_ofline_mode)
        Me.Controls.Add(Me.b_restore)
        Me.Controls.Add(Me.cbo_department)
        Me.Controls.Add(Me.b_ok)
        Me.Controls.Add(Me.lb_shifts)
        Me.Controls.Add(Me.cbo_machine)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_startup"
        Me.Text = "Startup"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbo_machine As System.Windows.Forms.ComboBox
    Friend WithEvents lb_shifts As System.Windows.Forms.ListBox
    Friend WithEvents b_ok As System.Windows.Forms.Button
    Friend WithEvents cbo_department As System.Windows.Forms.ComboBox
    Friend WithEvents b_restore As System.Windows.Forms.Button
    Friend WithEvents cb_ofline_mode As System.Windows.Forms.CheckBox
End Class

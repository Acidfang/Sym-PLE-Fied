<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_item_id
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
        Me.b_ok = New System.Windows.Forms.Button()
        Me.dgv_itemlist = New System.Windows.Forms.DataGridView()
        Me.dgc_itemlist_num_in = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_itemlist_item_id = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_itemlist_kg_in = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_itemlist_prod_num = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_itemlist_mat_num = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgv_itemlist, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'b_ok
        '
        Me.b_ok.Enabled = False
        Me.b_ok.Location = New System.Drawing.Point(5, 108)
        Me.b_ok.Name = "b_ok"
        Me.b_ok.Size = New System.Drawing.Size(152, 23)
        Me.b_ok.TabIndex = 1
        Me.b_ok.Text = "&OK"
        Me.b_ok.UseVisualStyleBackColor = True
        '
        'dgv_itemlist
        '
        Me.dgv_itemlist.AllowUserToAddRows = False
        Me.dgv_itemlist.AllowUserToDeleteRows = False
        Me.dgv_itemlist.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_itemlist.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgc_itemlist_num_in, Me.dgc_itemlist_item_id, Me.dgc_itemlist_kg_in, Me.dgc_itemlist_prod_num, Me.dgc_itemlist_mat_num})
        Me.dgv_itemlist.Location = New System.Drawing.Point(5, 7)
        Me.dgv_itemlist.MultiSelect = False
        Me.dgv_itemlist.Name = "dgv_itemlist"
        Me.dgv_itemlist.ReadOnly = True
        Me.dgv_itemlist.RowHeadersVisible = False
        Me.dgv_itemlist.Size = New System.Drawing.Size(152, 95)
        Me.dgv_itemlist.TabIndex = 3
        '
        'dgc_itemlist_num_in
        '
        Me.dgc_itemlist_num_in.HeaderText = "#"
        Me.dgc_itemlist_num_in.Name = "dgc_itemlist_num_in"
        Me.dgc_itemlist_num_in.ReadOnly = True
        Me.dgc_itemlist_num_in.Width = 30
        '
        'dgc_itemlist_item_id
        '
        Me.dgc_itemlist_item_id.HeaderText = "ID"
        Me.dgc_itemlist_item_id.Name = "dgc_itemlist_item_id"
        Me.dgc_itemlist_item_id.ReadOnly = True
        Me.dgc_itemlist_item_id.Width = 50
        '
        'dgc_itemlist_kg_in
        '
        Me.dgc_itemlist_kg_in.HeaderText = "KG"
        Me.dgc_itemlist_kg_in.Name = "dgc_itemlist_kg_in"
        Me.dgc_itemlist_kg_in.ReadOnly = True
        Me.dgc_itemlist_kg_in.Width = 50
        '
        'dgc_itemlist_prod_num
        '
        Me.dgc_itemlist_prod_num.HeaderText = "prod_num"
        Me.dgc_itemlist_prod_num.Name = "dgc_itemlist_prod_num"
        Me.dgc_itemlist_prod_num.ReadOnly = True
        Me.dgc_itemlist_prod_num.Visible = False
        '
        'dgc_itemlist_mat_num
        '
        Me.dgc_itemlist_mat_num.HeaderText = "mat_num"
        Me.dgc_itemlist_mat_num.Name = "dgc_itemlist_mat_num"
        Me.dgc_itemlist_mat_num.ReadOnly = True
        Me.dgc_itemlist_mat_num.Visible = False
        '
        'frm_item_id
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(162, 132)
        Me.ControlBox = False
        Me.Controls.Add(Me.dgv_itemlist)
        Me.Controls.Add(Me.b_ok)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frm_item_id"
        Me.Text = "Select Item"
        CType(Me.dgv_itemlist, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents b_ok As System.Windows.Forms.Button
    Friend WithEvents dgv_itemlist As System.Windows.Forms.DataGridView
    Friend WithEvents dgc_itemlist_num_in As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_itemlist_item_id As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_itemlist_kg_in As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_itemlist_prod_num As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_itemlist_mat_num As System.Windows.Forms.DataGridViewTextBoxColumn
End Class

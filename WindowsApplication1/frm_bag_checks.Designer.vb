<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_bag_checks
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
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Me.dgv_burst = New System.Windows.Forms.DataGridView()
        Me.nud_carton_num = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.CheckBox4 = New System.Windows.Forms.CheckBox()
        Me.l_burst_uap = New System.Windows.Forms.Label()
        Me.l_burst_aim = New System.Windows.Forms.Label()
        Me.l_burst_lap = New System.Windows.Forms.Label()
        Me.l_burst_1 = New System.Windows.Forms.Label()
        Me.l_burst_2 = New System.Windows.Forms.Label()
        Me.l_burst_3 = New System.Windows.Forms.Label()
        Me.l_burst_4 = New System.Windows.Forms.Label()
        Me.l_burst_5 = New System.Windows.Forms.Label()
        Me.l_burst_7 = New System.Windows.Forms.Label()
        Me.l_burst_6 = New System.Windows.Forms.Label()
        Me.l_burst_8 = New System.Windows.Forms.Label()
        Me.l_burst_9 = New System.Windows.Forms.Label()
        Me.l_burst_11 = New System.Windows.Forms.Label()
        Me.l_burst_10 = New System.Windows.Forms.Label()
        Me.l_burst_12 = New System.Windows.Forms.Label()
        Me.l_burst_fail = New System.Windows.Forms.Label()
        Me.l_burst_13 = New System.Windows.Forms.Label()
        Me.l_burst_14 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.nud_width = New System.Windows.Forms.NumericUpDown()
        Me.nud_length = New System.Windows.Forms.NumericUpDown()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CheckBox5 = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dgc_burst_test_num = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_burst_carton = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_burst_test_code = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_burst_burst = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_burst_length = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_burst_width = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_burst_date_ = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_burst_machine = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_burst_prod_num = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgc_burst_mat_num = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgv_burst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nud_carton_num, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nud_width, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nud_length, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgv_burst
        '
        Me.dgv_burst.AllowUserToAddRows = False
        Me.dgv_burst.AllowUserToDeleteRows = False
        Me.dgv_burst.AllowUserToResizeColumns = False
        Me.dgv_burst.AllowUserToResizeRows = False
        Me.dgv_burst.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_burst.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.dgc_burst_test_num, Me.dgc_burst_carton, Me.dgc_burst_test_code, Me.dgc_burst_burst, Me.dgc_burst_length, Me.dgc_burst_width, Me.dgc_burst_date_, Me.dgc_burst_machine, Me.dgc_burst_prod_num, Me.dgc_burst_mat_num})
        Me.dgv_burst.Location = New System.Drawing.Point(352, 294)
        Me.dgv_burst.Name = "dgv_burst"
        Me.dgv_burst.RowHeadersVisible = False
        Me.dgv_burst.Size = New System.Drawing.Size(376, 278)
        Me.dgv_burst.TabIndex = 0
        '
        'nud_carton_num
        '
        Me.nud_carton_num.Location = New System.Drawing.Point(6, 16)
        Me.nud_carton_num.Name = "nud_carton_num"
        Me.nud_carton_num.Size = New System.Drawing.Size(42, 20)
        Me.nud_carton_num.TabIndex = 5
        Me.nud_carton_num.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(49, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Carton #"
        '
        'Chart1
        '
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(8, 12)
        Me.Chart1.Name = "Chart1"
        Series1.ChartArea = "ChartArea1"
        Series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
        Series1.Legend = "Legend1"
        Series1.Name = "Series1"
        Me.Chart1.Series.Add(Series1)
        Me.Chart1.Size = New System.Drawing.Size(720, 276)
        Me.Chart1.TabIndex = 8
        Me.Chart1.Text = "Chart1"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(6, 102)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(55, 17)
        Me.CheckBox1.TabIndex = 9
        Me.CheckBox1.Text = "Splice"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(6, 125)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(85, 17)
        Me.CheckBox2.TabIndex = 9
        Me.CheckBox2.Text = "Jaw Change"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(6, 148)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(90, 17)
        Me.CheckBox3.TabIndex = 9
        Me.CheckBox3.Text = "Knife Change"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Location = New System.Drawing.Point(6, 171)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(82, 17)
        Me.CheckBox4.TabIndex = 9
        Me.CheckBox4.Text = "Maintenace"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'l_burst_uap
        '
        Me.l_burst_uap.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_uap.Location = New System.Drawing.Point(102, 32)
        Me.l_burst_uap.Name = "l_burst_uap"
        Me.l_burst_uap.Size = New System.Drawing.Size(100, 15)
        Me.l_burst_uap.TabIndex = 11
        Me.l_burst_uap.Text = "Upper Action Point"
        Me.l_burst_uap.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_aim
        '
        Me.l_burst_aim.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_aim.Location = New System.Drawing.Point(102, 112)
        Me.l_burst_aim.Name = "l_burst_aim"
        Me.l_burst_aim.Size = New System.Drawing.Size(100, 15)
        Me.l_burst_aim.TabIndex = 11
        Me.l_burst_aim.Text = "Aim"
        Me.l_burst_aim.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_lap
        '
        Me.l_burst_lap.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_lap.Location = New System.Drawing.Point(102, 160)
        Me.l_burst_lap.Name = "l_burst_lap"
        Me.l_burst_lap.Size = New System.Drawing.Size(100, 15)
        Me.l_burst_lap.TabIndex = 11
        Me.l_burst_lap.Text = "Lower Action Point"
        Me.l_burst_lap.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_1
        '
        Me.l_burst_1.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_1.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_1.Location = New System.Drawing.Point(203, 16)
        Me.l_burst_1.Name = "l_burst_1"
        Me.l_burst_1.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_1.TabIndex = 11
        Me.l_burst_1.Text = "58"
        Me.l_burst_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_2
        '
        Me.l_burst_2.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_2.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_2.Location = New System.Drawing.Point(203, 32)
        Me.l_burst_2.Name = "l_burst_2"
        Me.l_burst_2.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_2.TabIndex = 11
        Me.l_burst_2.Text = "57"
        Me.l_burst_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_3
        '
        Me.l_burst_3.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_3.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_3.Location = New System.Drawing.Point(203, 48)
        Me.l_burst_3.Name = "l_burst_3"
        Me.l_burst_3.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_3.TabIndex = 11
        Me.l_burst_3.Text = "56"
        Me.l_burst_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_4
        '
        Me.l_burst_4.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_4.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_4.Location = New System.Drawing.Point(203, 64)
        Me.l_burst_4.Name = "l_burst_4"
        Me.l_burst_4.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_4.TabIndex = 11
        Me.l_burst_4.Text = "55"
        Me.l_burst_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_5
        '
        Me.l_burst_5.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_5.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_5.Location = New System.Drawing.Point(203, 80)
        Me.l_burst_5.Name = "l_burst_5"
        Me.l_burst_5.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_5.TabIndex = 11
        Me.l_burst_5.Text = "54"
        Me.l_burst_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_7
        '
        Me.l_burst_7.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_7.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_7.Location = New System.Drawing.Point(203, 112)
        Me.l_burst_7.Name = "l_burst_7"
        Me.l_burst_7.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_7.TabIndex = 11
        Me.l_burst_7.Text = "52"
        Me.l_burst_7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_6
        '
        Me.l_burst_6.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_6.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_6.Location = New System.Drawing.Point(203, 96)
        Me.l_burst_6.Name = "l_burst_6"
        Me.l_burst_6.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_6.TabIndex = 11
        Me.l_burst_6.Text = "53"
        Me.l_burst_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_8
        '
        Me.l_burst_8.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_8.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_8.Location = New System.Drawing.Point(203, 128)
        Me.l_burst_8.Name = "l_burst_8"
        Me.l_burst_8.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_8.TabIndex = 11
        Me.l_burst_8.Text = "51"
        Me.l_burst_8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_9
        '
        Me.l_burst_9.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_9.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_9.Location = New System.Drawing.Point(203, 144)
        Me.l_burst_9.Name = "l_burst_9"
        Me.l_burst_9.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_9.TabIndex = 11
        Me.l_burst_9.Text = "50"
        Me.l_burst_9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_11
        '
        Me.l_burst_11.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_11.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_11.Location = New System.Drawing.Point(203, 176)
        Me.l_burst_11.Name = "l_burst_11"
        Me.l_burst_11.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_11.TabIndex = 11
        Me.l_burst_11.Text = "48"
        Me.l_burst_11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_10
        '
        Me.l_burst_10.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_10.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_10.Location = New System.Drawing.Point(203, 160)
        Me.l_burst_10.Name = "l_burst_10"
        Me.l_burst_10.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_10.TabIndex = 11
        Me.l_burst_10.Text = "49"
        Me.l_burst_10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_12
        '
        Me.l_burst_12.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_12.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_12.Location = New System.Drawing.Point(203, 192)
        Me.l_burst_12.Name = "l_burst_12"
        Me.l_burst_12.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_12.TabIndex = 11
        Me.l_burst_12.Text = "47"
        Me.l_burst_12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_fail
        '
        Me.l_burst_fail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_fail.Location = New System.Drawing.Point(102, 192)
        Me.l_burst_fail.Name = "l_burst_fail"
        Me.l_burst_fail.Size = New System.Drawing.Size(100, 15)
        Me.l_burst_fail.TabIndex = 11
        Me.l_burst_fail.Text = "Fail"
        Me.l_burst_fail.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_13
        '
        Me.l_burst_13.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_13.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_13.Location = New System.Drawing.Point(203, 208)
        Me.l_burst_13.Name = "l_burst_13"
        Me.l_burst_13.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_13.TabIndex = 11
        Me.l_burst_13.Text = "46"
        Me.l_burst_13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'l_burst_14
        '
        Me.l_burst_14.BackColor = System.Drawing.Color.LightGray
        Me.l_burst_14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.l_burst_14.Cursor = System.Windows.Forms.Cursors.Cross
        Me.l_burst_14.Location = New System.Drawing.Point(203, 224)
        Me.l_burst_14.Name = "l_burst_14"
        Me.l_burst_14.Size = New System.Drawing.Size(21, 15)
        Me.l_burst_14.TabIndex = 11
        Me.l_burst_14.Text = "45"
        Me.l_burst_14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(8, 545)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(338, 28)
        Me.Button1.TabIndex = 13
        Me.Button1.Text = "Add Test Data"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Location = New System.Drawing.Point(225, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(109, 184)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Bag Dimensions" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Target: " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "300mmx500mm " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Width x Length)"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'nud_width
        '
        Me.nud_width.Location = New System.Drawing.Point(225, 219)
        Me.nud_width.Name = "nud_width"
        Me.nud_width.Size = New System.Drawing.Size(50, 20)
        Me.nud_width.TabIndex = 18
        Me.nud_width.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'nud_length
        '
        Me.nud_length.Location = New System.Drawing.Point(284, 219)
        Me.nud_length.Name = "nud_length"
        Me.nud_length.Size = New System.Drawing.Size(50, 20)
        Me.nud_length.TabIndex = 18
        Me.nud_length.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(230, 203)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(35, 13)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Width"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(287, 203)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 13)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Length"
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = True
        Me.CheckBox5.Location = New System.Drawing.Point(106, 293)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(90, 17)
        Me.CheckBox5.TabIndex = 20
        Me.CheckBox5.Text = "Aussie Specs"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.nud_carton_num)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.CheckBox4)
        Me.GroupBox1.Controls.Add(Me.l_burst_uap)
        Me.GroupBox1.Controls.Add(Me.CheckBox3)
        Me.GroupBox1.Controls.Add(Me.l_burst_lap)
        Me.GroupBox1.Controls.Add(Me.CheckBox2)
        Me.GroupBox1.Controls.Add(Me.l_burst_fail)
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        Me.GroupBox1.Controls.Add(Me.l_burst_aim)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.l_burst_1)
        Me.GroupBox1.Controls.Add(Me.l_burst_5)
        Me.GroupBox1.Controls.Add(Me.l_burst_9)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.l_burst_3)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.l_burst_7)
        Me.GroupBox1.Controls.Add(Me.nud_length)
        Me.GroupBox1.Controls.Add(Me.l_burst_11)
        Me.GroupBox1.Controls.Add(Me.nud_width)
        Me.GroupBox1.Controls.Add(Me.l_burst_2)
        Me.GroupBox1.Controls.Add(Me.l_burst_6)
        Me.GroupBox1.Controls.Add(Me.l_burst_10)
        Me.GroupBox1.Controls.Add(Me.l_burst_14)
        Me.GroupBox1.Controls.Add(Me.l_burst_4)
        Me.GroupBox1.Controls.Add(Me.l_burst_13)
        Me.GroupBox1.Controls.Add(Me.l_burst_8)
        Me.GroupBox1.Controls.Add(Me.l_burst_12)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 294)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(340, 245)
        Me.GroupBox1.TabIndex = 21
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Test Data"
        '
        'Label5
        '
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Location = New System.Drawing.Point(98, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(1, 223)
        Me.Label5.TabIndex = 22
        '
        'dgc_burst_test_num
        '
        Me.dgc_burst_test_num.HeaderText = "Test #"
        Me.dgc_burst_test_num.Name = "dgc_burst_test_num"
        Me.dgc_burst_test_num.ReadOnly = True
        Me.dgc_burst_test_num.Width = 50
        '
        'dgc_burst_carton
        '
        Me.dgc_burst_carton.HeaderText = "Carton #"
        Me.dgc_burst_carton.Name = "dgc_burst_carton"
        Me.dgc_burst_carton.Width = 50
        '
        'dgc_burst_test_code
        '
        Me.dgc_burst_test_code.HeaderText = "Test Code"
        Me.dgc_burst_test_code.Name = "dgc_burst_test_code"
        '
        'dgc_burst_burst
        '
        Me.dgc_burst_burst.HeaderText = "Burst"
        Me.dgc_burst_burst.Name = "dgc_burst_burst"
        Me.dgc_burst_burst.Width = 50
        '
        'dgc_burst_length
        '
        Me.dgc_burst_length.HeaderText = "Length"
        Me.dgc_burst_length.Name = "dgc_burst_length"
        Me.dgc_burst_length.Width = 50
        '
        'dgc_burst_width
        '
        Me.dgc_burst_width.HeaderText = "Width"
        Me.dgc_burst_width.Name = "dgc_burst_width"
        Me.dgc_burst_width.Width = 50
        '
        'dgc_burst_date_
        '
        Me.dgc_burst_date_.HeaderText = "date"
        Me.dgc_burst_date_.Name = "dgc_burst_date_"
        Me.dgc_burst_date_.Visible = False
        '
        'dgc_burst_machine
        '
        Me.dgc_burst_machine.HeaderText = "machine"
        Me.dgc_burst_machine.Name = "dgc_burst_machine"
        Me.dgc_burst_machine.Visible = False
        '
        'dgc_burst_prod_num
        '
        Me.dgc_burst_prod_num.HeaderText = "prod_num"
        Me.dgc_burst_prod_num.Name = "dgc_burst_prod_num"
        Me.dgc_burst_prod_num.Visible = False
        '
        'dgc_burst_mat_num
        '
        Me.dgc_burst_mat_num.HeaderText = "mat_num"
        Me.dgc_burst_mat_num.Name = "dgc_burst_mat_num"
        Me.dgc_burst_mat_num.Visible = False
        '
        'frm_bag_checks
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(733, 580)
        Me.ControlBox = False
        Me.Controls.Add(Me.CheckBox5)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.dgv_burst)
        Me.Controls.Add(Me.Chart1)
        Me.Name = "frm_bag_checks"
        Me.Text = "Burst Checks"
        CType(Me.dgv_burst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nud_carton_num, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nud_width, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nud_length, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgv_burst As System.Windows.Forms.DataGridView
    Friend WithEvents nud_carton_num As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Chart1 As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents l_burst_uap As System.Windows.Forms.Label
    Friend WithEvents l_burst_aim As System.Windows.Forms.Label
    Friend WithEvents l_burst_lap As System.Windows.Forms.Label
    Friend WithEvents l_burst_1 As System.Windows.Forms.Label
    Friend WithEvents l_burst_2 As System.Windows.Forms.Label
    Friend WithEvents l_burst_3 As System.Windows.Forms.Label
    Friend WithEvents l_burst_4 As System.Windows.Forms.Label
    Friend WithEvents l_burst_5 As System.Windows.Forms.Label
    Friend WithEvents l_burst_7 As System.Windows.Forms.Label
    Friend WithEvents l_burst_6 As System.Windows.Forms.Label
    Friend WithEvents l_burst_8 As System.Windows.Forms.Label
    Friend WithEvents l_burst_9 As System.Windows.Forms.Label
    Friend WithEvents l_burst_11 As System.Windows.Forms.Label
    Friend WithEvents l_burst_10 As System.Windows.Forms.Label
    Friend WithEvents l_burst_12 As System.Windows.Forms.Label
    Friend WithEvents l_burst_fail As System.Windows.Forms.Label
    Friend WithEvents l_burst_13 As System.Windows.Forms.Label
    Friend WithEvents l_burst_14 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents nud_width As System.Windows.Forms.NumericUpDown
    Friend WithEvents nud_length As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dgc_burst_test_num As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_burst_carton As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_burst_test_code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_burst_burst As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_burst_length As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_burst_width As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_burst_date_ As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_burst_machine As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_burst_prod_num As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgc_burst_mat_num As System.Windows.Forms.DataGridViewTextBoxColumn
End Class

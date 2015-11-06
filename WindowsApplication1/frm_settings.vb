Public Class frm_settings

    Dim label_, printer_ As New DataTable

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Save settings to ini file
        WriteIni(Environ("USERPROFILE") & "\settings.ini", "Location", "Plant", ComboBox1.SelectedItem("plant"))
        WriteIni(environ("USERPROFILE") & "\settings.ini", "Location", "Department", ComboBox2.Text)
        'WriteIni(environ("USERPROFILE") & "\settings.ini", "ACSIS", "Label", ComboBox3.Text)
        WriteIni(Environ("USERPROFILE") & "\settings.ini", "ACSIS", "Printer", ComboBox4.Text)
        s_production = ComboBox2.SelectedItem("production")

        s_label = ComboBox3.Text
        s_printer = ComboBox4.Text
        Dim s As String = LCase(ComboBox2.Text)
        If Not s = s_department Then
            s_department = s
            s_production = ComboBox2.SelectedItem("production")
            If Not s_production Then
                s_shift = Nothing
                s_user = Nothing
                s_user_name = Nothing
                s_machine = Nothing
            End If

            frm_main.dgv_summary.Rows.Clear()
            frm_main.clear_job()
            frm_main.start_up()
            frm_main.set_department()
        End If
        Me.Close()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()


        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.DataSource = sql.get_table("db_departments", "plant", "=", ComboBox1.SelectedItem("plant"))
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Me.Visible = false
        With frm_settings_admin
            .StartPosition = FormStartPosition.CenterParent
            .Visible = False
            .ShowDialog()

        End With
        'Me.Visible = True
    End Sub

    Private Sub settings_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim dt, s As String
        dt = Format(Now, "dd/MM/yyyy")
        ComboBox1.DisplayMember = "name"
        ComboBox2.DisplayMember = "dept"
        ComboBox3.DisplayMember = "name"
        ComboBox4.DisplayMember = "name"
        ComboBox3.Enabled = CheckBox1.CheckState
        'bind SQL query to cbx

        ComboBox1.DataSource = sql.get_table("db_plants", "plant", ">", vbNull)

        s_plant = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Location", "Plant", 3203)
        select_cbo_item(ComboBox1, s_plant, "plant")

        ComboBox2.DataSource = dt_department 'sql.get_table("db_departments", "plant", "=", s_plant)

        s_department = LCase(ReadIni(Environ("USERPROFILE") & "\settings.ini", "Location", "Department", Nothing))
        select_cbo_item(ComboBox2, s_department, "dept")

        printer_ = sql.get_table("DP_Printers_Installed", "Printer_Location", "=", s_plant & " AND Label_Restriction='A5'", "symple")
        'label = sql.get_table("acsis_labels", "plant", "=", ComboBox1.SelectedItem("plant"))
        'printer = sql.get_table("acsis_printers", "plant", "=", ComboBox1.SelectedItem("plant"))

        If frm_main.tb_prod_num.Text = Nothing Then
            ComboBox3.Enabled = False
            CheckBox1.Enabled = False
        Else
            CheckBox1.Enabled = True
            If s_label = Nothing Then
                s = get_label(frm_main.tb_mat_num_out.Text, frm_main.tb_mat_desc_out.Text)
            Else
                s = s_label
            End If
            label_ = sql.get_table("DP_Print_Templates", "Label_Location", "=", "'Full' AND Label_Type='Item' ORDER BY List_Order", "symple")
            ComboBox3.DataSource = label_
            ComboBox3.DisplayMember = "Template_Code"

            If CheckBox1.Checked Then
                s = s_label
            End If

        select_cbo_item(ComboBox3, s, "Template_Code")
        End If

        ComboBox4.DataSource = printer_
        ComboBox4.DisplayMember = "Printer_Name"

        s = ReadIni(Environ("USERPROFILE") & "\settings.ini", "ACSIS", "Printer", Nothing)
        select_cbo_item(ComboBox4, s, "Printer_Name")

        Dim default_ As String = sql.read("DP_Logins", "Login_Id", "=", "'" & s_user & "' and Password='" & s_pass & "' AND Plant=" & s_plant, "Printer_Id", "symple")
        If default_ = Nothing Then
            Label6.Text = "Nothing"
        Else
            Label6.Text = default_
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim str As String = ComboBox4.SelectedItem("Printer_Name")
        sql.update("DP_Logins", "Printer_Id", "'" & str & "'", "Login_Id", "=", "'" & s_user & "' and Password='" & s_pass & "' AND Plant=" & s_plant, "symple")
        Label6.Text = str
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        ComboBox3.Enabled = CheckBox1.CheckState


        ComboBox3.DataSource = label_
        ComboBox3.DisplayMember = "Template_Code"
        ComboBox4.DataSource = printer_
        ComboBox4.DisplayMember = "Printer_Name"
        If s_label = Nothing Then
            ComboBox3.Text = get_label(frm_main.tb_mat_num_out.Text, frm_main.tb_mat_desc_out.Text)
        Else
            ComboBox3.Text = s_label
        End If


    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If s_loading Then

        End If
        'frm_main.dgv_summary.Rows.Clear()
    End Sub

End Class
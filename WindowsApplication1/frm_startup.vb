Public Class frm_startup

    'Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
    '    Dim dt As DataTable = schedule
    '    dt.DefaultView.RowFilter = "press = '" & ComboBox1.SelectedItem("machine") & "'"

    'End Sub


    Private Sub frm_startup_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        set_menu()
        SetForegroundWindow(Me.Handle)
        acsis.get_controls(0, True, False, True)
        If acsis.running(0) And s_user = Nothing Then acsis.get_logon_details(0)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles b_ok.Click
        If Not s_department = "graphics" And Not s_department = "planning" Then
            frm_main.tb_start_date.Text = Format(Now, "dd/MM/yyyy")
            s_shift = lb_shifts.Items(lb_shifts.SelectedIndex)("id")
            s_shift_name = UCase(lb_shifts.Items(lb_shifts.SelectedIndex)("name"))
            s_machine = cbo_machine.SelectedItem("name")
            s_extruder = cbo_machine.SelectedItem("extruder")

            WriteIni(Environ("USERPROFILE") & "\settings.ini", "Location", "Machine", s_machine)
            WriteIni(Environ("USERPROFILE") & "\settings.ini", "Location", "Department", cbo_department.Text)
            WriteIni(Environ("USERPROFILE") & "\settings.ini", "Machine", "Extruder", s_extruder)


            Dim tmp_dept As String = s_department
            If tmp_dept = "mounting" Then
                tmp_dept = "printing"
            End If
            Select Case s_department
                Case "printing"
                    frm_main.display_stations(cbo_machine.SelectedItem("stations"))
                    frm_main.b_main_personel.Text = "Add Helper"
                Case "bagplant"
                    frm_main.b_main_personel.Text = "Add Packer"
                Case "mounting"
                    frm_main.b_main_personel.Text = "Add Mounter"
            End Select

            s_loading = True
            frm_main.cbo_machine.SelectedIndex = cbo_machine.SelectedIndex
            s_loading = False
            frm_main.tb_shift.Text = s_shift
        Else
            s_shift = Nothing
        End If
        frm_main.set_department()
        Me.Close()
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.F4 Then
            ' Call your sub method here  .....

            Application.Exit()
            ' then avoid the key to reach the current control
            e.Handled = False
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_department.SelectedIndexChanged
        If s_loading Then Exit Sub
        s_department = LCase(cbo_department.Text)
        If s_loading Then Exit Sub
        s_production = cbo_department.SelectedItem("production")
        WriteIni(Environ("USERPROFILE") & "\settings.ini", "Location", "Department", cbo_department.Text)
        set_menu()


    End Sub
    Sub set_menu()
        s_loading = True
        'dt_machine = sql.get_table("db_machines", "plant", "=", s_plant & " AND dept='" & s_department & "' AND enabled=1")
        'cbo_machine.DataSource = dr_to_dt(dt_machines.Select("dept='" & s_department & "'"))
        'cbo_machine.DisplayMember = "name"
        cbo_department.DataSource = dt_department ' sql.get_table("db_departments", "plant", "=", s_plant)
        cbo_department.DisplayMember = "dept"


        s_department = LCase(ReadIni(Environ("USERPROFILE") & "\settings.ini", "Location", "Department", Nothing))

        If s_department = Nothing Then
            cbo_department.SelectedIndex = 0
            s_department = LCase(cbo_department.Text)
        Else
            cbo_department.Text = s_department
        End If


        Dim dr() As DataRow = Nothing
        Select Case s_department
            Case "graphics", "planning"
                lb_shifts.Visible = False
                cbo_machine.Visible = False
                b_ok.Top = cbo_machine.Top
                b_ok.Width = cbo_department.Width
                cb_ofline_mode.Top = b_ok.Top + cb_ofline_mode.Height + 7

                b_restore.Visible = False
                Me.Height = (cb_ofline_mode.Top + cb_ofline_mode.Height) + 39
                frm_main.set_department()
            Case Else
                cbo_machine.DataSource = dr_to_dt(dt_machines.Select("dept='" & s_department & "' AND enabled=1"))
                cbo_machine.DisplayMember = "name"
                frm_main.cbo_machine.DataSource = dr_to_dt(dt_machines.Select("dept='" & s_department & "' AND enabled=1"))
                frm_main.cbo_machine.DisplayMember = "name"
                If s_department = "filmsline" Then
                    Dim dt As Date = frm_main.tb_start_date.Text
                    Dim day As String = LCase(Format(dt, "ddd"))
                    lb_shifts.DataSource = dr_to_dt(dt_shifts.Select("dept='" & s_department & "' AND enabled=1 AND name LIKE '" & day & "%'", "id ASC"))
                Else
                    lb_shifts.DataSource = dr_to_dt(dt_shifts.Select("dept='" & s_department & "' AND enabled=1", "id ASC"))
                End If
                cbo_machine.Visible = True
                lb_shifts.Visible = True
                If Not s_machine = Nothing Then
                    select_cbo_item(cbo_machine, s_machine, "name")
                    If s_department = "printing" Then frm_main.display_stations(cbo_machine.SelectedItem("stations"))
                End If
                lb_shifts.DisplayMember = "name"
                lb_shifts.ValueMember = "id"

                lb_shifts.Height = (lb_shifts.Items.Count * 13) + 4
                Me.Height = (lb_shifts.Items.Count * 13) + 132 + cb_ofline_mode.Height
                b_ok.Top = lb_shifts.Top + lb_shifts.Height + 5
                b_restore.Top = b_ok.Top
                cb_ofline_mode.Top = b_ok.Top + cb_ofline_mode.Height + 7
                b_ok.Width = 54
                b_restore.Visible = True
        End Select
        If s_production Then
            lb_shifts.SelectedIndex = 0

        End If
        s_loading = False

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles b_restore.Click
        With frm_restore
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
        Me.Close()
    End Sub

    Private Sub cb_ofline_mode_CheckedChanged(sender As Object, e As EventArgs) Handles cb_ofline_mode.CheckedChanged
        If s_loading Then Exit Sub
        s_loading = True
        frm_main.cb_main_ofline_mode.Checked = cb_ofline_mode.Checked
        s_loading = False
    End Sub
End Class
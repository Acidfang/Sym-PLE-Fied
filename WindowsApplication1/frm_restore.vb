Public Class frm_restore
    Dim dt_names As DataTable

    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter And Button1.Enabled Then
            With DataGridView1

                dt_restore = sql.get_table("db_summary", "user_", "=", "'" & .CurrentRow.Cells(1).Value & "' AND date_='" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & _
                                  "' AND machine='" & .CurrentRow.Cells(2).Value & "' AND shift='" & .CurrentRow.Cells(3).Value & "'")
            End With
            If s_symple_access Then user_check()
            restore_it(0) 'Format(DateTimePicker1.Value, "dd/MM/yyyy"))
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        populate()
        show_jobs()
    End Sub
    Sub user_check()
        With DataGridView1
            If Not s_user = .CurrentRow.Cells(1).Value Then
                s_user = .CurrentRow.Cells(1).Value
                s_pass = sql.read("DP_Logins", "Login_Id", "=", "'" & s_user & "'", "Password", "symple")
                s_user_name = sql.read("DP_Logins", "Login_Id", "=", "'" & s_user & "'", "User_Name", "symple")
                If acsis.running(0) Then
                    w_sessions(0).process.Kill()
                    w_sessions(0).user = s_user
                    w_sessions(0).pass = s_pass
                    w_sessions(0).user_name = s_user_name
                    System.Threading.Thread.Sleep(1000)
                End If
                acsis.logon()
            End If
        End With
    End Sub
    Sub populate()
        Label1.Text = Nothing

        dt_restore = sql.get_table("db_summary", "plant", "=", s_plant & " AND dept='" & s_department & "' AND date_='" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "'")
        With DataGridView1
            .Rows.Clear()
            Dim found_name As Boolean = False
            If Not dt_restore Is Nothing Then
                For Each dr As DataRow In dt_restore.Rows
                    'If DataGridView1.RowCount > 0 Then
                    Dim add As Boolean = True
                    If .RowCount > 0 Then

                        For r As Integer = 0 To .RowCount - 1
                            If .Rows(r).Cells(1).Value = dr("user_") And _
                                .Rows(r).Cells(2).Value = dr("machine") And _
                                .Rows(r).Cells(3).Value = dr("shift") Then
                                add = False
                            End If
                        Next
                    End If
                    If add Then
                        .Rows.Add(0, dr("user_"), dr("machine"), dr("shift"))
                        If .RowCount = 1 Then .ClearSelection()
                        If dr("user_") = s_user And dr("machine") = s_machine Then
                            .CurrentCell = .Rows(.RowCount - 1).Cells(0)
                            .Rows(.RowCount - 1).Selected = True
                            show_jobs()
                            found_name = True

                        End If

                    End If

                    ' End If
                    '    If Not dt(i)("user_").ToString.Contains("PMNT") And Not dt(i)("user_") = Nothing Then
                    'For u As Integer = 0 To dt_names.Rows.Count - 1
                    '    If UCase(dt_names(u)("Login_Id")) = UCase(dt_(i)("user_")) And ComboBox1.Items.IndexOf(dt_names(u)("User_Name")) = -1 Then
                    '        ComboBox1.Items.Add(dt_names(u)("User_Name"))
                    '        Exit For
                    '    End If
                    'Next
                    '    End If
                    'End If
                Next
                If Not found_name Then
                    .CurrentCell = Nothing
                End If
                If s_symple_access Then
                    If .RowCount > 0 Then
                        For r As Integer = 0 To .RowCount - 1
                            For i As Integer = 0 To dt_names.Rows.Count - 1
                                If .Rows(r).Cells(1).Value = UCase(dt_names(i)("Login_Id")) Then
                                    .Rows(r).Cells(0).Value = dt_names(i)("User_Name")
                                    Exit For
                                End If
                            Next
                        Next
                    End If
                End If
            End If
            '  show_jobs()

        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.CurrentCell Is Nothing Then
            MsgBox("You must select an item to restore")
            Exit Sub
        End If
        With DataGridView1

            dt_restore = sql.get_table("db_summary", "user_", "=", "'" & .CurrentRow.Cells(1).Value & "' AND date_='" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & _
                              "' AND machine='" & .CurrentRow.Cells(2).Value & "' AND shift='" & .CurrentRow.Cells(3).Value & "'")
        End With
        frm_main.Refresh()
        If s_symple_access Then user_check()
        frm_main.set_department()
        restore_it(0) 'Format(DateTimePicker1.Value, "dd/MM/yyyy"))
    End Sub
    Sub show_jobs()
        Dim dt, str As String
        With DataGridView1
            Label1.Text = Nothing
            If .CurrentCell Is Nothing Then Exit Sub
            dt = "'" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "'"
            If .RowCount > 0 Then
                Dim dr() As DataRow = dt_restore.Select("user_='" & .CurrentRow.Cells(1).Value & "' AND date_=" & dt & _
                                           " AND machine='" & .CurrentRow.Cells(2).Value & "' AND shift='" & .CurrentRow.Cells(3).Value & "'")

                If Not .CurrentCell Is Nothing Then
                    If Not IsNothing(dt_restore) Then
                        str = "Saved jobs: " & dr.Length
                        Button1.Enabled = True
                    Else
                        str = "Nothing found for that date/user."
                        Button1.Enabled = False
                    End If
                    Label1.Text = str
                End If
            End If
        End With
    End Sub
    Sub restore_it(ByVal tab_page As Integer) 'user As String, restore_date As String, machine As String, shift As String)
        s_loading = True
        Dim row As Integer, col As String
        Me.Visible = False
        Me.Refresh()
        show_status("Restoring saved data")
        frm_main.tb_start_date.Text = dt_restore(0)("date_")
        If dt_restore Is Nothing Then Exit Sub
        frm_main.lb_multiple_job_list.Items.Clear()
        s_shift = dt_restore(0)("shift")
        Dim dr() As DataRow = dt_shifts.Select("id='" & s_shift & "' AND dept='" & s_department & "'")
        s_shift_name = UCase(dr(0)("name")) 'Replace(s_shift_name, UCase(Format(dt, "ddd")), UCase(Format(dt.AddDays(-1), "ddd")))

        's_shift_name = Replace(dt_restore(0)("user_"), "PMNT", Nothing)
        s_machine = dt_restore(0)("machine")
        s_assistant = dt_restore(0)("assistant").ToString
        If s_assistant = "" Then s_assistant = Nothing
        frm_main.tb_shift.Text = s_shift
        select_cbo_item(frm_main.cbo_machine, s_machine, "name")
        s_extruder = frm_main.cbo_machine.SelectedItem("extruder")
        frm_main.set_department()

        With frm_main.dgv_summary
            .Rows.Clear()
            For i As Integer = 0 To dt_restore.Rows.Count - 1 ' dt_restore.Compute("MAX(row)", "")
                .Rows.Add()
            Next
            For Each r In dt_restore.Rows
                row = r("row")
                For c As Integer = 0 To .Columns.Count - 1
                    Dim dp As Integer = 0
                    Dim s As String = Strings.Right(.Columns(c).DefaultCellStyle.Format, 1)
                    If IsNumeric(s) Then
                        dp = s
                    End If
                    If Not .Columns(c).ReadOnly Or c = 0 Then
                        If c = 0 Then
                            col = "prod_num"
                        Else
                            col = Replace(.Columns(c).Name, "dgc_summary_", Nothing)
                        End If

                        Select Case c
                            Case 1, 2
                                If Not IsNull(r(col)) Then
                                    .Rows(row).Cells(c).Value = r(col).ToString.PadLeft(4, "0"c)
                                End If
                            Case Else
                                Dim str As String
                                If IsNumeric(r(col)) Then
                                    str = FormatNumber(r(col), dp, , , TriState.False)
                                Else
                                    str = r(col).ToString
                                End If
                                .Rows(row).Cells(c).Value = str
                        End Select
                    End If

                Next
                '                .Rows(row).Cells(16).Value = r("schd_out").ToString

                .Rows(row).Cells("dgc_summary_setup").ToolTipText = .Rows(row).Cells("dgc_summary_st_setup").Value
                .Rows(row).Cells("dgc_summary_min_deck").ToolTipText = .Rows(row).Cells("dgc_summary_st_mindeck").Value
                .Rows(row).Cells("dgc_summary_mpm").ToolTipText = .Rows(row).Cells("dgc_summary_st_speed").Value

                .Rows(.RowCount - 1).Cells(0).Selected = True

                dt_produced = sql.get_table("db_produced", "prod_num", "=", r("prod_num"))
                s_selected_job = dt_schedules.Select("prod_num = " & .Rows(row).Cells("dgc_summary_prod_num").Value)
                frm_main.calc_times(r("row"))
                .CurrentCell = .Rows(row).Cells("dgc_summary_prod_num")
                'frm_main.calc_times()
                If s_department = "mounting" Then
                    .Rows(row).Cells("dgc_summary_desc").Value = s_selected_job(0)("mat_desc_fin")
                End If
            Next
            If s_department = "printing" Then

                If Not s_machine = Nothing Then
                    frm_main.display_stations(frm_main.cbo_machine.SelectedItem("stations"))
                End If
            End If
            s_production = True
            ' If Not IsNull(dt_restore(0)("tab_page")) Then
            frm_main.tc_main.SelectedIndex = tab_page
            'End If
            s_loading = False
            .ClearSelection()
            '.Rows(.RowCount - 1).Cells(0).Selected = True
            frm_main.lb_multiple_job_list.Items.Clear()
            .CurrentCell = .Rows(.RowCount - 1).Cells(0)
            frm_main.b_summary_multiple_job_add.Enabled = True
            s_selected_job = dt_schedules.Select("prod_num = " & .Rows(.RowCount - 1).Cells("dgc_summary_prod_num").Value)
            frm_main.tb_prod_num.Text = .Rows(.RowCount - 1).Cells("dgc_summary_prod_num").Value
            frm_main.lb_personel.Items.Clear()
            If Not s_assistant = Nothing Then frm_main.lb_personel.Items.AddRange(Split(s_assistant, "/"))
            frm_main.import_report()
            If Not IsNull(.Rows(.RowCount - 1).Cells("dgc_summary_multiple").Value) Then
                frm_main.lb_multiple_job_list.Items.AddRange(Split(.Rows(.RowCount - 1).Cells("dgc_summary_multiple").Value.ToString, ","))
                frm_main.calc_times_multiple(False)
                frm_main.gb_multiple_job.Visible = True
            End If
            'produced = sql.get_table("db_produced", "prod_num", "=", frm_main.TextBox2.Text)
            'frm_main.calc_times(.RowCount - 1)
        End With
        WriteIni(Environ("USERPROFILE") & "\settings.ini", "Location", "Machine", s_machine)
        WriteIni(Environ("USERPROFILE") & "\settings.ini", "Location", "Department", s_department)
        WriteIni(Environ("USERPROFILE") & "\settings.ini", "Machine", "Extruder", s_extruder)
        Me.Close()

        frm_status.Close()
        frm_main.display_job()
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)
        populate()
    End Sub

    Private Sub restore_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If s_symple_access Then dt_names = sql.get_table("DP_Logins", "plant", "=", s_plant, "symple")

        populate()
        DateTimePicker1.Value = FormatDateTime(frm_main.tb_start_date.Text)
        'If s_department = "mounting" Then
        '    Data.Text = "PMNT" & UCase(frm_startup.ListBox1.SelectedItem("name"))
        'End If
        populate()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DateTimePicker1.Value = DateTimePicker1.Value.AddDays(-1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DateTimePicker1.Value = DateTimePicker1.Value.AddDays(1)
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        With DataGridView1
            .Rows(.CurrentRow.Index).Selected = True
        End With

    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        With DataGridView1
            .Rows(.CurrentRow.Index).Selected = True
        End With
        show_jobs()

    End Sub


End Class
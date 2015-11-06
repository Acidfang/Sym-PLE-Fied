Public Class frm_symple_times

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles b_send_info.Click

        With dgv_pending
            If .RowCount = 0 Then Exit Sub
            Dim resp As MsgBoxResult = MsgBox("Send the selected information to Sym-PLE?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Send to Sym-PLE")
            If resp = MsgBoxResult.Yes Then
                acsis.open_job_screen(tb_prod_num.Text, False)

                show_status("Sending Run information to Sym-PLE")

                For r = 0 To .RowCount - 1
                    Dim send As Boolean, section, type, unit, quant As String

                    If IsNull(.Rows(r).Cells("dgc_pending_send").Value) Then
                        send = False
                    Else
                        send = .Rows(r).Cells("dgc_pending_send").Value
                    End If
                    If send Then
                        section = .Rows(r).Cells("dgc_pending_section").Value
                        If section = "Yield" Then
                            quant = .Rows(r).Cells("dgc_pending_quantity").Value
                        Else
                            quant = ConvertTime(CInt(.Rows(r).Cells("dgc_pending_quantity").Value), True)
                        End If
                        unit = .Rows(r).Cells("dgc_pending_unit").Value
                        type = .Rows(r).Cells("dgc_pending_type").Value
                        Select Case section
                            Case "Machine"

                                acsis.click_button(acsis.get_button_access("Run Confirmation", 0), False, 0)
                                show_status("Sending Run Time: " & Strings.Left(quant, 2) & " Hours " & Strings.Right(quant, 2) & " Mins")
                                acsis.insert_string("0000", w_sessions(0).winWnd(0).tbWnd(acsis.get_option("production_run_start")).handle)
                                acsis.insert_string(quant, w_sessions(0).winWnd(0).tbWnd(acsis.get_option("production_run_end")).handle)
                                acsis.click_button(acsis.get_option("production_run_confirm"), False, 0)
                            Case "Yield"
                                acsis.click_button(acsis.get_button_access("Run Confirmation", 0), False, 0)
                                show_status("Sending Yield: " & quant)
                                acsis.insert_string(quant, w_sessions(0).winWnd(0).tbWnd(acsis.get_option("production_run_quant")).handle)
                                acsis.click_button(acsis.get_option("production_run_confirm"), False, 0)
                            Case "Setup"

                                acsis.click_button(acsis.get_button_access("Setup Confirmation", 0), False, 0)
                                show_status("Sending Set-up Time: " & Strings.Left(quant, 2) & " Hours " & Strings.Right(quant, 2) & " Mins")
                                acsis.insert_string("0000", w_sessions(0).winWnd(0).tbWnd(acsis.get_option("production_setup_start")).handle)
                                acsis.insert_string(quant, w_sessions(0).winWnd(0).tbWnd(acsis.get_option("production_setup_end")).handle)
                                acsis.click_button(acsis.get_option("production_setup_confirm"), False, 0)
                            Case "Downtime"

                                acsis.click_button(acsis.get_button_access("DT & Scrap Recording", 0), False, 0)
                                show_status("Sending Downtime: " & Strings.Left(quant, 2) & " Hours " & Strings.Right(quant, 2) & " Mins" & "-" & type)

                                acsis.send_downtime(quant, type)
                        End Select
                    End If
                Next

                frm_main.dgv_summary.CurrentRow.Cells("dgc_summary_symple_times").Value = 1
                frm_main.export_summary()
                frm_main.display_job()
                frm_status.Close()
            End If
        End With

        get_times()
        SetForegroundWindow(Me.Handle)
    End Sub

    Private Sub symple_times_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        s_loading = True
        tb_prod_num.Text = frm_main.tb_prod_num.Text
        sb_job_select.Maximum = frm_main.dgv_summary.RowCount - 1
        sb_job_select.Value = frm_main.dgv_summary.CurrentRow.Index
        l_job_info.Text = sb_job_select.Value + 1 & " of " & sb_job_select.Maximum + 1
        b_send_info.Enabled = False
        cb_difference.CheckState = CheckState.Unchecked

        get_times()
        s_loading = False
    End Sub

    Sub get_times()


        Dim dts, dte As Date
        Dim s() As String = Split(frm_main.tb_start_date.Text, "/")
        dgv_symple.Rows.Clear()
        dgv_pending.Rows.Clear()
        dgv_summary.Rows.Clear()
        With frm_main.dgv_summary
            'dts = New Date(s(2), s(1), s(0), Strings.Left(.CurrentRow.Cells(1).Value, 2) - 1, 45, 0)
            Dim hour, min As Integer
            min = 60 + nud_time_offset.Value
            hour = Strings.Left(.Rows(0).Cells("dgc_summary_start").Value, 2)
            If min < 60 Then
                hour = hour - 1
            ElseIf min = 60 Then
                min = 0
            Else
                min = 0 + nud_time_offset.Value
            End If
            dts = New Date(s(2), s(1), s(0), hour, min, 0)
            For i As Integer = 0 To .RowCount - 1
                If Not IsNull(.Rows(i).Cells("dgc_summary_finish").Value) Then
                    hour = Strings.Left(.Rows(i).Cells("dgc_summary_finish").Value, 2)
                End If
            Next
            min = Math.Abs(nud_time_offset.Value)


            dte = New Date(s(2), s(1), s(0), hour, min, 0)
            If DateDiff(DateInterval.Hour, dts, dte) < 0 Then
                dte = dte.AddDays(1)

            End If
        End With
        Me.Text = "Loading data..."
        Me.Refresh()

        acsis.open_job_screen(frm_main.tb_prod_num.Text, False)
        If s_failed_job Then
            s_failed_job = False
            s_popup = 0
            frm_status.Close()
            Exit Sub
        End If
        show_status("Retrieving Times from Sym-PLE")
        With dgv_symple

            acsis.click_button(acsis.get_button_access("View Production Order", 0), False, 0)
            For i As Integer = 0 To w_sessions(0).winWnd(0).lbWnd(0).count
                s = Split(w_sessions(0).winWnd(0).lbWnd(0).text(i), "  ")
                Dim dn As Date = Now
                Dim ds As Date = s(0)
                Dim dd As Integer = Math.Abs(DateDiff(DateInterval.Minute, ds, dn))
                'dte = s(0)
                If dts < ds And dte > ds Or dd < 30 Then
                    Dim r As Integer = .RowCount
                    .Rows.Add()
                    .Rows(r).Cells(0).Value = s(1)
                    .Rows(r).Cells(1).Value = s(2)
                    .Rows(r).Cells(2).Value = s(3)
                    .Rows(r).Cells(3).Value = s(4)
                    .Rows(r).Cells(4).Value = s(0)
                    If dd > 30 Then

                        'End If
                        'If dd < 30 Then
                        '    dd = Math.Abs(DateDiff(DateInterval.Minute, dte, dn))
                        '    If dd > 30 Then
                        .Rows(r).DefaultCellStyle.BackColor = Color.Pink
                    End If
                    'End If
                End If
            Next
            acsis.click_button(acsis.get_button_back(0), False, 0)
            .ClearSelection()
        End With

        frm_status.Close()
        Me.Text = "Sym-PLE Entry"
        If dgv_symple.RowCount > 0 Then
            cb_difference.Visible = True
        Else
            cb_difference.Visible = False
        End If
        calc_times()
        SetForegroundWindow(Me.Handle)
    End Sub

    Private Sub HScrollBar1_ValueChanged(sender As Object, e As EventArgs) Handles sb_job_select.ValueChanged
        If s_loading Then Exit Sub
        Dim prod_num As Long
        With frm_main.dgv_summary
            .CurrentCell = .Rows(sb_job_select.Value).Cells("prod_num")
            .Rows(sb_job_select.Value).Cells("prod_num").Selected = True
            prod_num = .Rows(sb_job_select.Value).Cells("prod_num").Value
        End With
        tb_prod_num.Text = prod_num
        frm_main.tb_prod_num.Text = prod_num
        frm_main.display_job()
        l_job_info.Text = sb_job_select.Value + 1 & " of " & sb_job_select.Maximum + 1
        get_times()

    End Sub

    Sub calc_times()
        Dim uom As String = s_selected_job(0)("req_uom_1").ToString
        Dim tmp As DataTable = Nothing
        Dim r As Integer = -1
        With frm_main.dgv_summary
            If .RowCount > 0 Then
                r = .CurrentCell.RowIndex
                dgv_summary.Rows.Clear()
                dgv_summary.Rows.Add()
                For i As Integer = 5 To 9
                    If Not IsDBNull(.Rows(r).Cells(i).Value) And Not .Rows(r).Cells(i).Value = Nothing Then
                        dgv_summary.Rows(0).Cells(i - 5).Value = ConvertTime(.Rows(r).Cells(i).Value)
                    End If
                Next
                Select Case uom
                    Case "KM"
                        dgv_summary.Rows(0).Cells("dgc_summary_produced").Value = .Rows(r).Cells("dgc_summary_km_out").Value
                    Case "KG"
                        dgv_summary.Rows(0).Cells("dgc_summary_produced").Value = .Rows(r).Cells("dgc_summary_produced").Value
                    Case "ROL"
                        tmp = sql.get_table("db_produced", "prod_num", "=", tb_prod_num.Text & " AND user_='" & s_user & _
                                            "' AND date_='" & frm_main.tb_start_date.Text & "' AND NOT kg_out IS NULL AND NOT km_out IS NULL")

                        If Not IsNothing(tmp) Then
                            dgv_summary.Rows(0).Cells("dgc_summary_produced").Value = tmp.Rows.Count & " (Rolls)"
                        End If
                    Case "MU"
                        dgv_summary.Rows(0).Cells("dgc_summary_produced").Value = .Rows(r).Cells("dgc_summary_km_out").Value & " (Bags)"
                    Case Else
                        dgv_summary.Rows(0).Cells("dgc_summary_produced").Value = "N/A"
                End Select

            End If
        End With
        With dgv_pending
            .Rows.Clear()
            For i = 0 To dgv_summary.Columns.Count - 1
                If Not IsNull(dgv_summary.Rows(0).Cells(i).Value) Then
                    Dim s As String = Replace(dgv_summary.Rows(0).Cells(i).Value, " (Rolls)", Nothing)
                    s = Replace(s, " (Bags)", Nothing)
                    If Not CDec(s) = 0 Then
                        .Rows.Add()
                        Dim row As Integer = .RowCount - 1, section, type, unit, str As String, quant As Double
                        str = dgv_summary.Rows(0).Cells(i).Value
                        Select Case i
                            Case 0
                                section = "Setup"
                                type = Nothing
                                unit = "MIN"
                                quant = (Strings.Left(str, 2) * 60) + Strings.Right(str, 2)
                            Case 1
                                section = "Machine"
                                type = Nothing
                                unit = "MIN"
                                quant = (Strings.Left(str, 2) * 60) + Strings.Right(str, 2)
                            Case 2
                                section = "Downtime"
                                type = "BS01"
                                unit = "MIN"
                                quant = (Strings.Left(str, 2) * 60) + Strings.Right(str, 2)
                            Case 3
                                section = "Downtime"
                                type = "CS01"
                                unit = "MIN"
                                quant = (Strings.Left(str, 2) * 60) + Strings.Right(str, 2)
                            Case 4
                                section = "Downtime"
                                type = "ES01"
                                unit = "MIN"
                                quant = (Strings.Left(str, 2) * 60) + Strings.Right(str, 2)
                            Case 5
                                section = "Yield"
                                type = Nothing
                                unit = uom
                                Select Case uom
                                    Case "ROL"
                                        quant = Replace(str, " (Rolls)", Nothing)
                                    Case "MU"
                                        quant = Replace(str, " (Bags)", Nothing) / 1000
                                    Case Else
                                        quant = str
                                End Select
                            Case Else
                                section = Nothing
                                type = Nothing
                                unit = Nothing
                                quant = Nothing
                        End Select

                        .Rows(row).Cells("dgc_pending_section").Value = section
                        .Rows(row).Cells("dgc_pending_quantity").Value = FormatNumber(quant, 3, , , False)
                        .Rows(row).Cells("dgc_pending_unit").Value = unit
                        .Rows(row).Cells("dgc_pending_type").Value = type
                        .Rows(row).Cells("dgc_pending_string").Value = str
                    End If
                End If
            Next
            get_symple_values()
            If Not symple_values Is Nothing Then

                For r = 0 To .RowCount - 1
                    Dim v1 As String = .Rows(r).Cells("dgc_pending_section").Value & .Rows(r).Cells("dgc_pending_quantity").Value & _
                        .Rows(r).Cells("dgc_pending_unit").Value & .Rows(r).Cells("dgc_pending_type").Value
                    For i As Integer = 0 To symple_values.Length - 1
                        Dim v2 As String = symple_values(i).section & FormatNumber(symple_values(i).quantity, 3, , , TriState.False) & symple_values(i).unit & symple_values(i).type
                        If v1 = v2 Then
                            .Rows(r).DefaultCellStyle.BackColor = Color.LightGreen
                            Exit For
                        End If
                    Next
                Next
            End If
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles b_final_confirm.Click
        Dim do_prep As Boolean = False
        s_selected_job = dt_schedules.Select("prod_num=" & tb_prod_num.Text)

        Dim resp As MsgBoxResult = MsgBox("Final Confirm?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Final Confirm")
        If resp = MsgBoxResult.Yes Then
            acsis.open_job_screen(tb_prod_num.Text, False)
            do_prep = acsis.has_prep

            show_status("Processing Final Confirmation")

            acsis.click_button(acsis.get_button_access("Final Confirmation", 0), False, 0)
            If s_popup = 12 Then
                acsis.click_button(acsis.get_button_back(0), False, 0)
                s_popup = 0
            Else
                acsis.set_check(acsis.get_control_handle(acsis.get_acsis_option("production_final_check"), _
                                                         cls_ACSIS.eControlType.CheckBox, 0), True)

                acsis.click_button(acsis.get_acsis_option("production_final_confirm"), False, 0)
                acsis.click_button(acsis.get_button_back(0), False, 0)
            End If
            If do_prep Then
                acsis.open_job_screen(tb_prod_num.Text, True)

                show_status("Processing Final Confirmation - Pre Press")

                acsis.click_button(acsis.get_button_access("Final Confirmation", 0), False, 0)
                If s_popup = 12 Then
                    acsis.click_button(acsis.get_button_back(0), False, 0)
                    s_popup = 0
                Else
                    acsis.set_check(acsis.get_control_handle(acsis.get_acsis_option("production_final_check"), _
                                                             cls_ACSIS.eControlType.CheckBox, 0), True)

                    acsis.click_button(acsis.get_acsis_option("production_final_confirm"), False, 0)
                End If

            End If

            sql.update("db_schedule", "status", 10, "status", "=", 9 & " AND machine='" & s_selected_job(0)("machine") & "'")
            sql.update("db_schedule", "fin_date=" & sql.convert_value(Now) & ",status", 9, "prod_num", "=", tb_prod_num.Text)
            Me.Close()

            If frm_main.dgv_reject.RowCount > 0 Then
                MsgBox("There are items in the reject slip, you need to print this out")
                frm_main.tc_main.SelectedIndex = 4
            End If
            SetForegroundWindow(Me.Handle)
            frm_status.Close()
            frm_main.dgv_summary.CurrentRow.Cells("dgc_summary_final_conf").Value = 1
            frm_main.export_summary()
            frm_main.display_job()

            'frm_main.print_report()
            'MsgBox("The Production Report has been printed, please put it in the finish job packet.")
        End If
    End Sub

    Private Sub dgv_pending_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_pending.CellClick
        With dgv_pending
            .ClearSelection()
            If e.RowIndex = -1 Then Exit Sub
            If e.ColumnIndex = 4 Then
                If IsNull(.Rows(e.RowIndex).Cells("dgc_pending_send").Value) Then
                    .Rows(e.RowIndex).Cells("dgc_pending_send").Value = True
                Else
                    If .Rows(e.RowIndex).Cells("dgc_pending_send").Value = True Then
                        .Rows(e.RowIndex).Cells("dgc_pending_send").Value = False
                    Else
                        .Rows(e.RowIndex).Cells("dgc_pending_send").Value = True
                    End If
                End If

            End If
            b_send_info.Enabled = False
            If Not s_final_done Then
                For r As Integer = 0 To .RowCount - 1
                    If .Rows(r).Cells("dgc_pending_send").Value = True Then
                        b_send_info.Enabled = True
                        Exit For
                    End If
                Next
            End If

        End With
    End Sub
    Sub get_symple_values()
        Erase symple_values
        With dgv_symple
            If .RowCount > 0 Then
                Dim add As Boolean = False

                For r As Integer = 0 To dgv_symple.RowCount - 1
                    Dim i As Integer = -1
                    If symple_values Is Nothing Then
                        ReDim symple_values(0)
                        symple_values(0).section = .Rows(r).Cells("dgc_symple_section").Value
                        symple_values(0).type = .Rows(r).Cells("dgc_symple_type").Value.ToString
                        symple_values(0).unit = .Rows(r).Cells("dgc_symple_unit").Value
                        symple_values(0).quantity = .Rows(r).Cells("dgc_symple_quantity").Value
                    Else
                        add = True
                        For i = 0 To symple_values.Length - 1
                            If symple_values(i).section = .Rows(r).Cells("dgc_symple_section").Value And symple_values(i).type = .Rows(r).Cells("dgc_symple_type").Value.ToString Then
                                add = False
                                Exit For
                            End If
                        Next
                    End If
                    If add Then
                        ReDim Preserve symple_values(symple_values.Length)
                        i = symple_values.Length - 1
                        symple_values(i).section = .Rows(r).Cells("dgc_symple_section").Value
                        symple_values(i).type = .Rows(r).Cells("dgc_symple_type").Value.ToString
                        symple_values(i).unit = .Rows(r).Cells("dgc_symple_unit").Value
                        symple_values(i).quantity = .Rows(r).Cells("dgc_symple_quantity").Value
                    ElseIf i > -1 Then
                        symple_values(i).quantity = symple_values(i).quantity + .Rows(r).Cells("dgc_symple_quantity").Value
                    End If

                Next
            End If
        End With
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles cb_difference.CheckedChanged
        If cb_difference.Checked Then
            get_symple_values()

            For r As Integer = 0 To dgv_pending.RowCount - 1
                For i As Integer = 0 To symple_values.Length - 1
                    If dgv_pending.Rows(r).Cells("dgc_pending_section").Value = symple_values(i).section And _
                        dgv_pending.Rows(r).Cells("dgc_pending_type").Value = symple_values(i).type Then

                        If dgv_pending.Rows(r).Cells("dgc_pending_quantity").Value - symple_values(i).quantity < 0 Then
                            Dim msg As String = " has to be reversed - the time you are trying to enter is less that what is already in Sym-PLE"
                            Select Case dgv_pending.Rows(r).Cells("dgc_pending_section").Value
                                Case "Setup", "Yield", "Machine"
                                    msg = "The numbers for " & dgv_pending.Rows(r).Cells("dgc_pending_section").Value & msg
                                Case "Downtime"
                                    msg = "The times for " & dgv_pending.Rows(r).Cells("dgc_pending_section").Value & " - " & _
                                        dgv_pending.Rows(r).Cells("dgc_pending_type").Value & msg
                            End Select
                            MsgBox(msg, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Times have to be reversed!")
                        Else
                            If dgv_pending.Rows(r).Cells("dgc_pending_quantity").Value - symple_values(i).quantity = 0 Then
                                dgv_pending.Rows.Item(r).Visible = False
                                dgv_pending.Rows(r).Cells("dgc_pending_send").Value = False
                            Else
                                dgv_pending.Rows(r).Cells("dgc_pending_quantity").Value = _
                                    FormatNumber(dgv_pending.Rows(r).Cells("dgc_pending_quantity").Value - symple_values(i).quantity, 3)
                                dgv_pending.Rows(r).Cells("dgc_pending_string").Value = ConvertTime(dgv_pending.Rows(r).Cells("dgc_pending_quantity").Value, True)
                            End If
                        End If
                        Exit For
                    End If
                Next
            Next
        Else
            calc_times()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles b_refresh.Click
        get_times()
    End Sub
End Class
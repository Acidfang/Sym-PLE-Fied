Imports System.Runtime.InteropServices
Imports System.Deployment.Application
Imports System.IO
Imports System.IO.Ports
Imports System.Text

Module mod_scheduling
    Sub check_clipboard()
        frm_main.Timer1.Enabled = False
        s_loading = True
        Dim cb As String = Clipboard.GetText()                   'Get clipboard data as a string
        Dim rows() As String = cb.Split(ControlChars.NewLine)    'Split into rows
        Dim uom_count As Integer = 0
        Dim c1, c2, c3 As Integer
        Dim v1 As Boolean = False, v2 As Boolean = False, v3 As Boolean = False
        Dim num_up_count As Integer = 0
        Dim dr_previous() As DataRow
        If rows.Length > 10 Then
            Dim cells() As String = rows(0).Split(ControlChars.Tab)
            If cells.Length > 1 Then

                For c As Integer = 0 To cells.Length - 1
                    Select Case cells(c)
                        Case "Material"
                            If v1 Then Exit Select
                            uom_count = uom_count + 1
                            v1 = True
                            c1 = c
                        Case "Denom."
                            If v2 Then Exit Select
                            uom_count = uom_count + 1
                            v2 = True
                            c2 = c
                        Case "Numerat."
                            If v3 Then Exit Select
                            uom_count = uom_count + 1
                            v3 = True
                            c3 = c
                        Case "Document"
                            If v1 Then Exit Select
                            num_up_count = num_up_count + 1
                            v1 = True
                            c1 = c
                        Case "Cyld.Size"
                            If v2 Then Exit Select
                            num_up_count = num_up_count + 1
                            v2 = True
                            c2 = c
                        Case "Imps.Up"
                            If v3 Then Exit Select
                            num_up_count = num_up_count + 1
                            v3 = True
                            c3 = c
                    End Select
                Next
                If uom_count = 3 Then
                    ' Me.TopMost = True
                    frm_main.tssl_status.Text = "Status: Getting UOM data from clipboard..."
                    frm_main.Refresh()
                    rows = cb.Split(ControlChars.NewLine)    'Split into rows

                    uom_count = 0
                    For r As Integer = 1 To rows.Length - 1
                        frm_main.tspb_progress.Value = r / (rows.Length - 1) * 100
                        cells = rows(r).ToString.Replace(ControlChars.Lf, "").Split(ControlChars.Tab)
                        If cells.Length > 1 Then
                            If IsNumeric(cells(c1)) Then
                                Dim mat_num, numer, denom As Integer
                                mat_num = cells(c1)
                                numer = cells(c2)
                                denom = cells(c3)
                                If Not numer = denom Then

                                    Dim dr() As DataRow = dt_uom.Select("mat_num=" & mat_num)
                                    If dr.Length = 0 Then
                                        uom_count = uom_count + 1
                                        For i As Integer = 0 To frm_main.lb_scheduling_missing_uom.Items.Count - 1
                                            If mat_num = frm_main.lb_scheduling_missing_uom.Items(i) Then
                                                frm_main.lb_scheduling_missing_uom.Items.RemoveAt(i)
                                                Exit For
                                            End If
                                        Next
                                        sql.insert("db_uom", "mat_num,denom,numer", mat_num & "," & numer & "," & denom)
                                    End If
                                End If
                            End If
                        End If

                    Next
                    dt_uom = sql.get_table_full("db_uom", True)
                    frm_main.tssl_status.Text = "Status: UOM data updated (" & uom_count & ")"
                    Clipboard.Clear()
                ElseIf num_up_count = 3 Then
                    ' Me.TopMost = True

                    frm_main.tssl_status.Text = "Status: Getting Number Up data from clipboard..."
                    frm_main.Refresh()
                    rows = cb.Split(ControlChars.NewLine)    'Split into rows

                    num_up_count = 0
                    For r As Integer = 1 To rows.Length - 1
                        frm_main.tspb_progress.Value = r / (rows.Length - 1) * 100
                        cells = rows(r).ToString.Replace(ControlChars.Lf, "").Split(ControlChars.Tab)
                        If cells.Length > 1 Then
                            Dim dr() As DataRow = dt_num_up.Select("zaw='" & cells(c1) & "'")
                            If dr.Length = 0 Then
                                num_up_count = num_up_count + 1
                                For i As Integer = 0 To frm_main.lb_scheduling_missing_num_up.Items.Count - 1
                                    If cells(c1) = frm_main.lb_scheduling_missing_num_up.Items(i) Then
                                        frm_main.lb_scheduling_print_order.Items.RemoveAt(i)
                                        Exit For
                                    End If
                                Next
                                sql.insert("db_num_up", "zaw,cyl_size,num_up,plant", "'" & cells(c1) & "'," & _
                                           sql.convert_value(cells(c2)) & "," & CDec(cells(c3)) & "," & s_plant)
                            End If
                        End If

                    Next
                    dt_num_up = sql.get_table_full("db_num_up", True)
                    frm_main.tssl_status.Text = "Status: Number Up data updated (" & num_up_count & ")"
                    Clipboard.Clear()
                Else
                    frm_main.Show()
                    cells = rows(1).Split(ControlChars.Tab)

                    Dim col As Integer = 0
                    If cells.Length >= 20 Then

                        If cells(1) = "Ship Nr" Then
                            frm_main.tssl_status.Text = Nothing
                            frm_main.Timer1.Enabled = False
                            dt_schedules = sql.get_table("db_schedule", "plant", "=", s_plant)
                            dr_previous = dt_schedules.Select("last_schedule=1")
                            get_job_data(cb)
                            With frm_main.DataGridView1
                                .Rows.Clear()
                                For Each item In dr_previous
                                    Dim found As Boolean = False
                                    For i As Integer = 0 To jobs.Length - 1
                                        If jobs(i).prod_num = item("prod_num") Then
                                            found = True
                                            Exit For
                                        End If
                                    Next
                                    If Not found Then
                                        Dim r As Integer = .RowCount
                                        .Rows.Add()
                                        .Rows(r).Cells(0).Value = item("prod_num")
                                        .Rows(r).Cells(1).Value = item("machine")
                                        If item("dept") = "printing" Then
                                            .Rows(r).Cells(2).Value = item("label_desc")
                                        Else
                                            If IsNull(item("mat_desc_fin")) Then
                                                .Rows(r).Cells(2).Value = item("mat_desc_semi")
                                            Else
                                                .Rows(r).Cells(2).Value = item("mat_desc_fin")
                                            End If
                                        End If
                                    End If
                                Next

                            End With
                            Clipboard.Clear()
                            frm_main.Timer1.Enabled = True
                        End If
                    End If
                End If
            End If
        End If
        cb = Nothing
        If Process.GetProcessesByName("Excel").GetLength(0) > 0 Then
            Try
                xlApp = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application)

                If xlApp.Workbooks.Count = 0 Then
                    close_excel()
                Else
                    xlWorkBook = Nothing
                    For Each wb As Microsoft.Office.Interop.Excel.Workbook In xlApp.Workbooks
                        If wb.Name = "scheduling.xlsx" Then
                            xlWorkBook = wb
                        End If
                    Next
                    If xlWorkBook Is Nothing Then
                        frm_main.b_scheduling_print.Enabled = False
                    Else
                        frm_main.b_scheduling_print.Enabled = True
                        If frm_main.lb_scheduling_departments.Items.Count = 0 Then
                            Dim departments As String = Nothing
                            For Each ws As Microsoft.Office.Interop.Excel.Worksheet In xlWorkBook.Worksheets
                                If ws.Visible Then
                                    If Not ws.Name.Contains("Priority") Then
                                        departments = add_to_string(departments, get_department(ws.Name), ",", False)
                                    End If
                                End If
                            Next
                            frm_main.lb_scheduling_departments.Items.AddRange(Split(departments, ","))
                            frm_main.lb_scheduling_departments.SelectedIndex = 0
                        End If
                    End If
                End If

            Catch ex As Exception

            End Try
        Else
            frm_main.b_scheduling_print.Enabled = False

        End If
        frm_main.Timer1.Enabled = True
        s_loading = False
        If jobs Is Nothing Then
            frm_main.b_scheduling_build.Enabled = False
            frm_main.b_scheduling_remove_job.Enabled = False
        Else
            frm_main.b_scheduling_build.Enabled = True
            frm_main.b_scheduling_remove_job.Enabled = True
        End If

    End Sub



    Sub get_job_data(ByVal cb As String)
        ' cb = Clipboard.GetText

        Dim rows() As String = cb.Split(ControlChars.NewLine)    'Split into rows
        Dim s As String = Nothing
        Erase jobs

        Dim files() As String = Directory.GetFiles(get_file_path("pdf"))
        Erase s_pdf
        Dim str() As String
        For Each f In files
            Str = Split(f, "\")
            If s_pdf Is Nothing Then
                ReDim s_pdf(0)
                s_pdf(0) = Strings.Left(str(str.Length - 1), 5)
            Else
                ReDim Preserve s_pdf(s_pdf.Length)
                s_pdf(s_pdf.Length - 1) = Strings.Left(str(str.Length - 1), 5)

            End If
        Next        ' cells = rows(0).ToString.Replace(ControlChars.Lf, "").Split(ControlChars.Tab)
        s_machines = Nothing
        s_departments = Nothing
        Dim col As Integer = 0
        Dim add_to_list As Boolean = True
        Dim row_start As Integer = 0, row_finish As Integer = 0
        s_missing_num_up = Nothing
        s_missing_uom = Nothing
        frm_main.dgv_scheduling_missing_pdf.Rows.Clear()
        frm_main.DataGridView1.Rows.Clear()
        Dim skip As Boolean = False
        'With frm_main.dgv_scheduling
        '    .Rows.Clear()
        '    For c As Integer = 0 To .ColumnCount - 1
        '        .Columns(c).Visible = True
        '    Next
        '    .PerformLayout()
        'End With

        For r = 0 To rows.Length - 2
            frm_main.tspb_progress.Value = r / (rows.Length - 1) * 100
            Dim cells() As String = rows(r).ToString.Replace(ControlChars.Lf, "").Split(ControlChars.Tab)
            If r = rows.Length - 2 Then
                row_finish = r
                'add_job(row_start, row_finish)
            End If
            If cells.Length > 1 Then
                Select Case True
                    Case cells(0) = "LWCTOT"
                        'totals
                    Case cells(0) = "HEADER"
                    Case Not get_department(cells(1)) = Nothing
                        str = Split(cells(1), ":")
                        s = str(0)
                        str = Split(cells(1), "-")
                        s = s & " - " & str(1)
                        s_machines = add_to_string(s_machines, s, allow_double:=False)
                        s_departments = add_to_string(s_departments, get_department(s), allow_double:=False)
                    Case cells(11).Contains("    ")
                        If row_start = 0 Then
                            row_start = r
                        End If
                        If Not row_start = 0 And Not row_start = r Then
                            r = r - 1
                            row_finish = r - 1
                        End If

                End Select

                If r = row_start + 1 Then
                    If IsNumeric(cells(8)) And Not Strings.Left(cells(8), 1) = "5" Then
                        'skip = True
                        SetForegroundWindow(frm_main.Handle)
                        MsgBox("A job has been detected on " & s & " with no Job Number " & cells(8) & vbCr & cells(4) & vbCr & cells(6) & vbCr & cells(7))
                    End If
                End If

                If row_start > 0 And row_finish > 0 Then

                    add_to_list = True
                    If Not jobs Is Nothing And IsNumeric(cells(8)) Then
                        For i As Integer = 0 To jobs.Length - 1
                            If jobs(i).prod_num = cells(8) Then
                                add_to_list = False
                            End If
                        Next
                    End If


                    If add_to_list Then
                        If Not skip Then add_job(rows, row_start, row_finish)
                        row_start = 0
                        row_finish = 0
                        skip = False

                    End If
                End If
            End If


        Next
        With frm_main

            .lb_scheduling_departments.Items.Clear()
            .lb_scheduling_machines.Items.Clear()
            .dgv_scheduling.Rows.Clear()
            .lb_scheduling_missing_num_up.Items.Clear()
            .lb_scheduling_missing_uom.Items.Clear()
            .lb_scheduling_missing_num_up.Items.AddRange(Split(s_missing_num_up, ","))
            .lb_scheduling_missing_uom.Items.AddRange(Split(s_missing_uom, ","))
            .lb_scheduling_departments.Items.AddRange(Split(s_departments, ","))
            If .lb_scheduling_print_departments.Items.Count = 0 Then
                .lb_scheduling_print_departments.Items.AddRange(Split(s_departments, ","))
                .lb_scheduling_print_departments.SelectedIndex = 0
                WriteIni(Environ("USERPROFILE") & "\settings.ini", "Lists", "Departments", s_departments)
                str = Split(s_departments, ",")
                For i As Integer = 0 To UBound(str)
                    Dim m() As String = Split(s_machines, ",")
                    s = Nothing
                    For i1 As Integer = 0 To UBound(m)
                        If get_department(m(i1)) = str(i) Then
                            If s = Nothing Then
                                s = Strings.Left(m(i1), 6)
                            Else
                                s = s & "," & Strings.Left(m(i1), 6)
                            End If
                        End If
                    Next
                    WriteIni(Environ("USERPROFILE") & "\settings.ini", "Lists", str(i), s)
                Next
            End If
            .lb_scheduling_departments.SelectedIndex = 0
            If .lb_scheduling_departments.Items.Count > 1 Then .lb_scheduling_departments.Items.Add("All Departments")
            .lb_scheduling_machines.SelectedIndex = .lb_scheduling_machines.Items.Count - 1
        End With
        'display_jobs()
    End Sub
    Sub display_scheduled_jobs()
        Dim str() As String
        Dim mn, desc, uom1, uom2, req1, req2 As String

        frm_main.dgv_scheduling.Rows.Clear()
        For i As Integer = 0 To UBound(jobs)
            Dim add As Boolean = False
            mn = jobs(i).mat_num_fin
            If mn = 0 Then mn = jobs(i).mat_num_semi
            uom1 = jobs(i).req_uom_1
            req1 = FormatNumber(jobs(i).req_quant_1, 1, , , TriState.False)

            uom2 = jobs(i).req_uom_2
            If uom2 = Nothing Then
                uom2 = ""
                req2 = ""
            Else
                req2 = FormatNumber(jobs(i).req_quant_2, 1, , , TriState.False)
            End If
            If jobs(i).dept = "printing" Then
                desc = jobs(i).label_desc
            Else
                desc = jobs(i).customer
            End If
            If desc = Nothing Then desc = ""
            If frm_main.lb_scheduling_departments.SelectedItem = "All Departments" Then
                If frm_main.lb_scheduling_machines.SelectedItem.ToString.Contains("All Machines") Then
                    add = True
                Else
                    str = Split(frm_main.lb_scheduling_machines.SelectedItem, " ")
                    If str(0) = jobs(i).machine Then
                        add = True
                    End If
                End If
            Else
                If frm_main.lb_scheduling_machines.SelectedItem.ToString.Contains("All Machines") And _
                    frm_main.lb_scheduling_departments.SelectedItem = jobs(i).dept Then
                    add = True
                Else
                    str = Split(frm_main.lb_scheduling_machines.SelectedItem, " ")
                    If str(0) = jobs(i).machine Then
                        add = True
                    End If
                End If
            End If
            If add Then
                With frm_main.dgv_scheduling
                    Dim r As Integer = .RowCount
                    .Rows.Add()
                    With .Rows(r)
                        .Cells(0).Value = jobs(i).prod_num
                        .Cells(1).Value = jobs(i).mat_num_semi
                        .Cells(2).Value = jobs(i).mat_desc_semi
                        .Cells(3).Value = jobs(i).mat_num_fin
                        .Cells(4).Value = jobs(i).mat_desc_fin
                        .Cells(5).Value = jobs(i).mat_num_in
                        .Cells(6).Value = jobs(i).mat_desc_in
                        .Cells(7).Value = jobs(i).mat_num_in_1
                        .Cells(8).Value = jobs(i).mat_desc_in_1
                        .Cells(9).Value = jobs(i).mat_num_in_2
                        .Cells(10).Value = jobs(i).mat_desc_in_2
                        .Cells(11).Value = jobs(i).label_desc
                        .Cells(12).Value = jobs(i).cyl_size
                        .Cells(13).Value = jobs(i).inksys
                        .Cells(14).Value = jobs(i).inks
                        .Cells(15).Value = FormatNumber(jobs(i).issue_quant, 1, , , TriState.False).ToString
                        .Cells(16).Value = jobs(i).issue_uom
                        If jobs(i).dept = "printing" Then
                            .Cells(17).Value = FormatNumber(jobs(i).req_quant_1, 1, , , TriState.False).ToString
                        Else
                            .Cells(17).Value = FormatNumber(jobs(i).req_quant_1, 3, , , TriState.False).ToString
                        End If
                        .Cells(18).Value = jobs(i).req_uom_1
                        .Cells(19).Value = FormatNumber(jobs(i).req_quant_2, 1, , , TriState.False).ToString
                        .Cells(20).Value = jobs(i).req_uom_2
                        .Cells(21).Value = jobs(i).machine
                        .Cells(22).Value = jobs(i).date_due
                        .Cells(23).Value = jobs(i).time_setup
                        .Cells(24).Value = jobs(i).time_run
                        .Cells(25).Value = jobs(i).no_up
                        .Cells(26).Value = jobs(i).status
                        .Cells(27).Value = jobs(i).plant
                        .Cells(28).Value = jobs(i).dept
                        .Cells(29).Value = jobs(i).sales_doc
                        .Cells(30).Value = jobs(i).sales_line
                        .Cells(31).Value = jobs(i).customer
                        .Cells(32).Value = jobs(i).bag_info
                        .Cells(33).Value = jobs(i).issued
                        .Cells(34).Value = jobs(i).design
                        .Cells(35).Value = jobs(i).sched_text
                        .Cells(36).Value = jobs(i).logo_colour
                        .Cells(37).Value = jobs(i).date_added

                    End With
                    Dim BoldRow As New DataGridViewCellStyle With {.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)}
                    If jobs(i).status < 0 Then
                        .Rows(.RowCount - 1).DefaultCellStyle = BoldRow
                        .Rows(.RowCount - 1).DefaultCellStyle.ForeColor = Color.Blue
                    ElseIf jobs(i).status = 10 Then
                        .Rows(.RowCount - 1).DefaultCellStyle = BoldRow
                        .Rows(.RowCount - 1).DefaultCellStyle.ForeColor = Color.Red
                    ElseIf jobs(i).status = 2 Then
                        .Rows(.RowCount - 1).DefaultCellStyle = BoldRow
                        .Rows(.RowCount - 1).DefaultCellStyle.ForeColor = Color.Green
                    End If
                End With

            End If
        Next
        With frm_main.dgv_scheduling
            For c As Integer = 0 To .ColumnCount - 1
                .Columns(c).Visible = False
            Next
            frm_main.dgc_scheduling_prod_num.Visible = True
            frm_main.dgc_scheduling_req_quant_1.Visible = True
            frm_main.dgc_scheduling_req_uom_1.Visible = True
            frm_main.dgc_scheduling_issued.Visible = True
            frm_main.dgc_scheduling_date_due.Visible = True
            frm_main.dgc_scheduling_status.Visible = True
            frm_main.dgc_scheduling_issued.DefaultCellStyle.Format = "N1"
            frm_main.dgc_scheduling_req_quant_1.DefaultCellStyle.Format = "N1"
            frm_main.dgc_scheduling_req_quant_2.DefaultCellStyle.Format = "N1"
            Select Case frm_main.lb_scheduling_departments.SelectedItem
                Case "bagplant"
                    frm_main.dgc_scheduling_customer.Visible = True
                    frm_main.dgc_scheduling_logo_colour.Visible = True
                    frm_main.dgc_scheduling_mat_num_fin.Visible = True
                    frm_main.dgc_scheduling_mat_desc_fin.Visible = True
                    frm_main.dgc_scheduling_sales_doc.Visible = True
                    frm_main.dgc_scheduling_sales_line.Visible = True
                Case "filmsline"
                    frm_main.dgc_scheduling_mat_num_fin.Visible = True
                    frm_main.dgc_scheduling_mat_desc_fin.Visible = True
                    frm_main.dgc_scheduling_mat_num_semi.Visible = True
                    frm_main.dgc_scheduling_mat_desc_semi.Visible = True
                Case "printing"
                    frm_main.dgc_scheduling_req_quant_2.Visible = True
                    frm_main.dgc_scheduling_req_uom_2.Visible = True
                    frm_main.dgc_scheduling_label_desc.Visible = True
                    frm_main.dgc_scheduling_inks.Visible = True
                    frm_main.dgc_scheduling_mat_num_semi.Visible = True
                    frm_main.dgc_scheduling_mat_desc_semi.Visible = True
            End Select
        End With
    End Sub

    Sub export_jobs()
        dt_schedules = sql.get_table_full("db_schedule", True)
        sql.update("db_schedule", "status", 10 & ",last_schedule=0", "dept", "=", "'" & frm_main.lb_scheduling_departments.SelectedItem & "' AND plant=" & s_plant)
        With frm_main.dgv_scheduling
            Dim str, sql_string, c_name As String
            For r As Integer = 0 To .RowCount - 1
                If .Rows(r).Cells("dgc_scheduling_dept").Value = frm_main.lb_scheduling_departments.SelectedItem And Strings.Left(.Rows(r).Cells("dgc_scheduling_prod_num").Value, 1) = "5" Then
                    Dim prod_num As Long = .Rows(r).Cells("dgc_scheduling_prod_num").Value
                    Dim dr() As DataRow = dt_schedules.Select("prod_num=" & prod_num)

                    If dr Is Nothing Then
                        sql.insert("db_schedule", "prod_num,plant", prod_num & "," & s_plant)
                    ElseIf dr.Length > 1 Then
                        sql.delete("db_schedule", "prod_num", "=", prod_num & " AND plant=" & s_plant)
                        sql.insert("db_schedule", "prod_num,plant", prod_num & "," & s_plant)
                    ElseIf dr.Length = 0 Then
                        sql.insert("db_schedule", "prod_num,plant", prod_num & "," & s_plant)
                    End If

                    'If .Rows(r).Cells("dgc_scheduling_status").Value = -1 Then
                    '    .Rows(r).Cells("dgc_scheduling_status").Value = 0
                    'End If

                    str = Nothing
                    sql_string = Nothing

                    For c As Integer = 1 To .ColumnCount - 1
                        c_name = Replace(.Columns(c).Name, "dgc_scheduling_", Nothing)
                        If IsDBNull(.Rows(r).Cells(c).Value) Then
                            str = .Rows(r).Cells(c).Value.ToString
                        Else
                            str = .Rows(r).Cells(c).Value
                        End If

                        str = sql.convert_value(str)
                        If sql_string = Nothing Then
                            sql_string = c_name & "=" & str
                        Else
                            sql_string = sql_string & "," & c_name & "=" & str
                        End If

                    Next
                    sql_string = sql_string & ",last_schedule=1"
                    sql.update("db_schedule", "prod_num", .Rows(r).Cells("dgc_scheduling_prod_num").Value & "," & sql_string, "prod_num", "=", .Rows(r).Cells("dgc_scheduling_prod_num").Value & " AND plant=" & s_plant)

                End If

            Next
        End With
        dt_schedules = sql.get_table("db_schedule", "plant", "=", s_plant)
    End Sub

    Sub add_job(ByVal rows() As String, ByVal start As Integer, ByVal finish As Integer)


        Dim dr() As DataRow = Nothing
        Dim str() As String = Nothing
        Dim info() As String = Nothing
        Dim keep_values As Boolean = True
        Dim uom_found As Boolean = False
        Dim quant_found As Boolean = False
        Dim agreement_quant As Decimal = 0
        Dim order_quant As Decimal = 0
        If jobs Is Nothing Then
            ReDim jobs(0)
        Else
            ReDim Preserve jobs(jobs.Length)
        End If
        Dim j As Long = jobs.Length - 1
        If start = 158 Then
            keep_values = keep_values
        End If
        With frm_main.dgv_scheduling
            .Rows.Add()
            For r = start To finish
                Dim agreement As Boolean = True
                Dim cells() As String = rows(r).ToString.Replace(ControlChars.Lf, "").Split(ControlChars.Tab)
                If cells.Length > 1 Then

                    If IsNull(cells(6)) And IsNull(cells(7)) And IsNull(cells(8)) And IsNull(cells(9)) And IsNull(cells(10)) And Not IsNull(cells(11)) Then
                        str = Split(cells(11), "    ")
                        info = Split(str(0), ":")
                        jobs(j).time_setup = (info(0) * 60) + (info(1))
                        info = Split(str(1), ":")
                        jobs(j).time_run = (info(0) * 60) + (info(1))
                    ElseIf IsNumeric(cells(8)) And jobs(j).prod_num = 0 Then
                        dr = dt_schedules.Select("prod_num='" & cells(8) & "'")
                        jobs(j).prod_num = cells(8)
                        If jobs(j).prod_num = 0 Then
                            dr = dr
                        End If
                        jobs(j).machine = cells(11)
                        jobs(j).dept = get_department(cells(11))
                        jobs(j).date_due = cells(12)
                        jobs(j).plant = s_plant
                        If Not IsNull(cells(4)) Then
                            jobs(j).customer = cells(4)
                        End If

                    End If
                    'reqs
                    If Not IsNull(cells(2)) Then
                        If CLng(cells(2)) < 30000000 Then
                            agreement = False
                        End If
                        If jobs(j).sales_doc = Nothing Then
                            jobs(j).sales_doc = cells(2)
                            jobs(j).sales_line = cells(3)
                        Else

                            Dim sd() As String = Split(jobs(j).sales_doc, ":")
                            Dim sl() As String = Split(jobs(j).sales_line, ":")
                            Dim add As Boolean = True
                            For i As Integer = 0 To UBound(sd)
                                If sd(i) & sl(i) = cells(2) & cells(3) Then
                                    add = False
                                End If
                            Next
                            If add Then
                                jobs(j).sales_doc = add_to_string(jobs(j).sales_doc, cells(2), ":")
                                jobs(j).sales_line = add_to_string(jobs(j).sales_line, cells(3), ":")
                            End If

                        End If


                    End If
                    If IsNull(cells(2)) And Not IsNull(cells(3)) Then
                        jobs(j).customer = add_to_string(jobs(j).customer, cells(3), " ", False)
                    End If


                    If Not IsNull(cells(10)) Then
                        If IsNull(cells(5)) Then ' no BCOMP
                            If Not Trim(cells(8)) = "TOTAL" Then
                                If jobs(j).req_uom_1 = Nothing Or jobs(j).req_uom_1 = cells(10) Then
                                    jobs(j).req_uom_1 = cells(10)
                                    If IsNumeric(cells(9)) Then
                                        jobs(j).req_quant_1 = jobs(j).req_quant_1 + CDec(cells(9))
                                    End If
                                ElseIf jobs(j).req_uom_2 = Nothing Or jobs(j).req_uom_2 = cells(10) Then
                                    jobs(j).req_uom_2 = cells(10)
                                    jobs(j).req_quant_2 = jobs(j).req_quant_2 + CDec(cells(9))
                                Else
                                    j = j
                                End If
                                If IsNumeric(cells(9)) Then
                                    If agreement Then
                                        agreement_quant = agreement_quant + CDec(cells(9))
                                    Else
                                        order_quant = order_quant + CDec(cells(9))
                                    End If
                                End If

                            End If
                            If Strings.Left(cells(6), 1) = "1" Then
                                jobs(j).mat_num_fin = cells(6)
                                jobs(j).mat_desc_fin = cells(7)
                            ElseIf Strings.Left(cells(6), 1) = "2" Then
                                jobs(j).mat_num_semi = cells(6)
                                jobs(j).mat_desc_semi = cells(7)
                            End If
                        ElseIf Trim(cells(5)) = "BCOMP" Then
                            If Trim(cells(6)) = "ISSUED" Then
                                jobs(j).issued = True
                                If dr.Length > 0 Then

                                    If Not IsNull(dr(0)("mat_num_in")) And IsNull(jobs(j).mat_num_in) Then
                                        jobs(j).mat_num_in = dr(0)("mat_num_in")
                                        jobs(j).mat_desc_in = dr(0)("mat_desc_in")
                                    ElseIf Not IsNull(dr(0)("mat_num_in_1")) Then
                                        jobs(j).mat_num_in_1 = dr(0)("mat_num_in_1")
                                        jobs(j).mat_desc_in_1 = dr(0)("mat_desc_in_1")
                                    ElseIf Not IsNull(dr(0)("mat_num_in_2")) Then
                                        jobs(j).mat_num_in_2 = dr(0)("mat_num_in_2")
                                        jobs(j).mat_desc_in_2 = dr(0)("mat_desc_in_2")
                                    End If

                                    If Not IsNull(dr(0)("issue_uom")) Then
                                        jobs(j).issue_quant = dr(0)("issue_quant")
                                        jobs(j).issue_uom = dr(0)("issue_uom")
                                    End If

                                End If

                            Else

                                jobs(j).mat_num_in = cells(6)
                                jobs(j).mat_desc_in = cells(7)
                                jobs(j).issue_quant = cells(9)
                                jobs(j).issue_uom = cells(10)
                            End If

                        End If

                    End If

                    Select Case Trim(cells(6))
                        Case "INK DATA"
                            jobs(j).inks = Replace(cells(7), ";", ",")
                        Case "PRNT1", "PRNT2", "Bag Info"
                            str = Split(cells(7), " ")
                            For i = 0 To UBound(str)
                                info = Split(str(i), ":")
                                If UBound(info) > 0 Then
                                    Select Case info(0)
                                        Case "COLOR"
                                            jobs(j).logo_colour = info(1)
                                        Case "CYL"
                                            jobs(j).cyl_size = info(1)
                                        Case "DESIGN"
                                            jobs(j).design = info(1)
                                        Case "INKSYS"
                                            jobs(j).inksys = info(1)
                                        Case "MTL_AVL"
                                            jobs(j).MaterialAvail = info(1)
                                    End Select
                                End If
                            Next
                            If Trim(cells(6)) = "Bag Info" Then
                                jobs(j).bag_info = Trim(cells(7))
                            End If
                        Case "Lbl Dsc"
                            jobs(j).label_desc = add_to_string(jobs(j).label_desc, Trim(cells(7)), " ")
                        Case "Sch/Lng"
                            jobs(j).sched_text = add_to_string(jobs(j).sched_text, Trim(cells(7)), "|")
                            If cells(7).Contains("creased") Then
                                str = Split(Trim(cells(7)), " ")
                                Dim tmp_uom As String = Nothing, tmp_quant As Decimal
                                quant_found = False
                                uom_found = False
                                For Each item In str
                                    If Not quant_found Then tmp_quant = get_job_value(item, True)
                                    If Not uom_found Then tmp_uom = get_job_value(UCase(item), False)
                                    If tmp_quant > 0 Then quant_found = True
                                    If tmp_uom = "KM" Or tmp_uom = "KG" Or tmp_uom = "ROL" Or tmp_uom = "MU" Then uom_found = True

                                Next

                                If quant_found And uom_found Then
                                    jobs(j).req_quant_1 = tmp_quant
                                    jobs(j).req_uom_1 = tmp_uom
                                    jobs(j).req_uom_2 = Nothing
                                    keep_values = False
                                Else
                                    MsgBox(jobs(j).prod_num & " has invalid text for increasing/decreasing the production amount" & vbCr & vbCr & _
                                           "The text needs to contain the word ""creased"" and the quantity and uom" & vbCr & vbCr & _
                                           "e.g. increased 1500kg or decreased 1500 kg" & vbCr & vbCr & "The original values will remain until this is fixed, or you manually adjust it" & _
                                           vbCr & vbCr & "Wrong text: " & cells(7))
                                End If
                            End If
                    End Select

                End If
            Next
            dr = dt_schedules.Select("prod_num='" & jobs(j).prod_num & "'")

            If dr.Length > 0 Then

                If Not IsNull(dr(0)("status")) Then
                    jobs(j).status = dr(0)("status")
                    If Not IsNull(dr(0)("sched_date")) Then
                        jobs(j).date_added = dr(0)("sched_date")
                    Else
                        jobs(j).date_added = Format(Now.AddDays(-1), "dd/MM/yyyy HH:mm")
                    End If

                    If jobs(j).status = 11 Or jobs(j).status = -1 Then
                        jobs(j).status = 0
                    ElseIf jobs(j).status = -3 Then
                        Dim dt As Date = jobs(j).date_added
                        If DateDiff(DateInterval.Day, dt, Now) >= 1 Then
                            jobs(j).status = 0
                        Else
                            jobs(j).status = -1
                        End If
                    End If
                Else
                    jobs(j).status = 0
                    If Not IsNull(dr(0)("sched_date")) Then
                        jobs(j).date_added = dr(0)("sched_date")
                    Else
                        jobs(j).date_added = Format(Now.AddDays(-1), "dd/MM/yyyy HH:mm")
                    End If
                End If
                If Not jobs(j).machine = dr(0)("machine") And Not CInt(jobs(j).status) < 0 Then
                    jobs(j).sched_text = add_to_string(jobs(j).sched_text, "Was " & dr(0)("machine"), "|")
                End If
                If Not IsNull(dr(0)("index_")) Then
                    jobs(j).index = dr(0)("index_")
                End If
                If keep_values Then

                    If Not IsNull(dr(0)("req_quant_1")) Then
                        jobs(j).req_quant_1 = dr(0)("req_quant_1")
                    End If
                    If Not IsNull(dr(0)("req_quant_2")) Then
                        jobs(j).req_quant_2 = dr(0)("req_quant_2")
                    End If
                End If
            Else
                If jobs(j).dept = "bagplant" Then
                    jobs(j).status = -2
                Else
                    jobs(j).status = -1
                End If
                jobs(j).date_added = Format(Now, "dd/MM/yyyy HH:mm")
            End If
            Dim zaw As String = get_material_info(mat_desc:=jobs(j).mat_desc_semi).zaw
            If zaw = Nothing Then zaw = get_material_info(mat_desc:=jobs(j).mat_desc_fin).zaw

            If Not zaw = Nothing Then
                If jobs(j).machine = "PPGE09" Then
                    With frm_main.dgv_scheduling_missing_pdf
                        If Array.IndexOf(s_pdf, zaw) = -1 Then
                            Dim w As Integer
                            If jobs(j).mat_desc_in = Nothing Then
                                w = get_material_info(False, jobs(j).mat_desc_in).width
                            Else
                                w = get_material_info(False, jobs(j).mat_desc_fin).width
                            End If
                            .Rows.Add(zaw, jobs(j).cyl_size, w, jobs(j).label_desc)
                        End If
                    End With
                End If
                dr = dt_num_up.Select("zaw='" & zaw & "'")
                If dr.Length > 0 Then
                    jobs(j).no_up = dr(0)("num_up")
                Else
                    s_missing_num_up = add_to_string(s_missing_num_up, zaw, allow_double:=False)
                    zaw = get_material_info(mat_desc:=jobs(j).mat_desc_fin).zaw
                    If Not zaw = Nothing Then
                        dr = dt_num_up.Select("zaw='" & zaw & "'")
                        If dr.Length > 0 Then
                            jobs(j).no_up = dr(0)("num_up")
                        Else
                            s_missing_num_up = add_to_string(s_missing_num_up, zaw, allow_double:=False)
                        End If
                    End If
                End If

            End If
            If jobs(j).req_uom_2 = Nothing Then
                Select Case jobs(j).req_uom_1
                    Case "KG"
                        jobs(j).req_uom_2 = "M"
                        If jobs(j).mat_num_semi > 0 Then
                            dr = dt_uom.Select("mat_num=" & jobs(j).mat_num_semi)
                        Else
                            dr = dt_uom.Select("mat_num=" & jobs(j).mat_num_fin)
                        End If
                        If dr.Length > 0 Then
                            jobs(j).req_quant_2 = FormatNumber((dr(0)("denom") / dr(0)("numer")) * jobs(j).req_quant_1 * 1000, 2, , , TriState.False)
                        Else
                            If jobs(j).mat_num_fin > 0 Then
                                s_missing_uom = add_to_string(s_missing_uom, jobs(j).mat_num_fin, allow_double:=False)
                            Else
                                s_missing_uom = add_to_string(s_missing_uom, jobs(j).mat_num_semi, allow_double:=False)
                            End If

                            If jobs(j).mat_num_semi = 0 Then
                                jobs(j).req_quant_2 = get_roll_length(jobs(j).req_quant_1, jobs(j).mat_desc_fin)
                            Else
                                jobs(j).req_quant_2 = get_roll_length(jobs(j).req_quant_1, jobs(j).mat_desc_semi)
                            End If
                        End If
                    Case "KM"
                        If Not jobs(j).dept = "bagplant" Then

                            jobs(j).req_uom_2 = "KG"
                            Dim w, g As Integer
                            w = get_material_info(False, jobs(j).mat_desc_fin).width
                            g = get_material_info(True, jobs(j).mat_desc_fin).gauge
                            jobs(j).req_quant_2 = FormatNumber(w * jobs(j).req_quant_1 * g / 1000 * 2, 3, , , TriState.False)
                            j = j
                        End If

                    Case "MU"
                    Case "ROL"
                        jobs(j).req_uom_2 = "M"
                        jobs(j).req_quant_2 = jobs(j).req_quant_1 * get_material_info(mat_desc:=jobs(j).mat_desc_fin).length
                    Case "IMP"
                        jobs(j).req_uom_2 = "M"
                        jobs(j).req_quant_2 = (jobs(j).cyl_size / 1000) * jobs(j).req_quant_1
                    Case Else
                        j = j
                End Select
            End If

            Dim dtn As Date = Format(Now, "dd/MM/yyyy") & " 14:00"
            Dim d As String = Format(jobs(j).date_due, "ddd")
            Dim dtd As Date = Format(jobs(j).date_due, "dd/MM/yyyy") & " 14:00"
            If d = "Sun" Then
                jobs(j).date_due = jobs(j).date_due.AddDays(1)
            ElseIf d = "Sat" Then
                jobs(j).date_due = jobs(j).date_due.AddDays(2)
            End If
            Dim s As String = CInt(Format(jobs(j).date_due, "dd"))
            Select Case Right(s, 1)
                Case 1
                    s = s & "st"
                Case 2
                    s = s & "nd"
                Case 3
                    s = s & "rd"
                Case Else
                    s = s & "th"
            End Select
            If dtd.AddDays(-3) < dtn And Not jobs(j).dept = "printing" Then
                If Not agreement_quant = 0.001 Then
                    If Not order_quant = 0 Then
                        If jobs(j).sched_text = Nothing Then
                            If Not jobs(j).req_quant_1 = order_quant Then
                                jobs(j).sched_text = order_quant & " " & jobs(j).req_uom_1 & " Due " & Format(jobs(j).date_due, "ddd") & " " & s
                            Else
                                jobs(j).sched_text = "Due " & Format(jobs(j).date_due, "ddd") & " " & s
                            End If
                        ElseIf Not jobs(j).sched_text.Contains("Due ") Then
                            If Not jobs(j).req_quant_1 = order_quant Then
                                jobs(j).sched_text = add_to_string(jobs(j).sched_text, order_quant & " " & jobs(j).req_uom_1 & " Due " & Format(jobs(j).date_due, "ddd") & " " & s, "|")
                            Else
                                jobs(j).sched_text = add_to_string(jobs(j).sched_text, "Due " & Format(jobs(j).date_due, "ddd") & " " & s, "|")
                            End If

                        End If
                    End If
                Else
                    s = s
                End If
            End If

            'finalise job 
        End With
    End Sub

    Sub build_schedules(ByVal department As String)


        Dim str() As String
        Dim line As Boolean = False
        Dim count As Integer = 0
        Dim total As Integer = 0
        Dim lines As Integer = 0
        'str = Split(frm_main.lb_scheduling_machines.Items(frm_main.lb_scheduling_machines.Items.Count - 1), " (")
        'total = Replace(str(1), ")", Nothing)
        frm_main.tssl_status.Text = "Status: Opening Template workbook"
        Dim exeDir As New IO.FileInfo(Reflection.Assembly.GetExecutingAssembly.FullName)
        xlWorkBook = Nothing
        If xlApp Is Nothing Then
            frm_main.tssl_status.Text = "Status: loading Excel..."
            xlApp = New Microsoft.Office.Interop.Excel.Application
            With xlApp
                .UserControl = True
                .Visible = True
            End With
        Else
            For Each wb As Microsoft.Office.Interop.Excel.Workbook In xlApp.Workbooks
                If wb.Name = "scheduling.xlsx" Then
                    xlWorkBook = wb
                End If
            Next
        End If
        xlApp.ScreenUpdating = False
        xlApp.Application.DisplayAlerts = False
        If xlWorkBook Is Nothing Then xlWorkBook = xlApp.Workbooks.Open(s_resource_path & "scheduling.xlsx", [ReadOnly]:=False)
        xlWorkSheets = xlWorkBook.Worksheets
        For Each ws As Microsoft.Office.Interop.Excel.Worksheet In xlWorkBook.Worksheets
            If ws.Visible Then
                If Not ws.Name.Contains("Priority") And Not ws.Name = "Missing PDFs" Then
                    ws.Delete()
                End If
            End If
        Next
        If department = "printing" Then
            xlWorkSheet = xlWorkSheets("Missing PDFs")
            xlWorkSheet.Range("A3:D49").Value = Nothing
            With frm_main.dgv_scheduling_missing_pdf
                For i As Integer = 0 To .RowCount - 1
                    xlWorkSheet.Range("A" & i + 3).Value = .Rows(i).Cells(0).Value
                    xlWorkSheet.Range("B" & i + 3).Value = .Rows(i).Cells(1).Value
                    xlWorkSheet.Range("C" & i + 3).Value = .Rows(i).Cells(2).Value
                    xlWorkSheet.Range("D" & i + 3).Value = .Rows(i).Cells(3).Value
                Next
            End With
            For i As Integer = 0 To frm_main.lb_scheduling_machines.Items.Count - 1
                If Not frm_main.lb_scheduling_machines.Items(i).ToString.Contains("All Machines") Then
                    str = Split(frm_main.lb_scheduling_machines.Items(i), "(")
                    Dim machine As String = str(0)
                    frm_main.tssl_status.Text = "Status: Addding a page for " & machine
                    xlWorkSheet = xlWorkSheets("Template Printing")
                    xlWorkSheet.Copy(After:=xlWorkSheets("Printing Priority"))
                    xlWorkBook.Worksheets(xlWorkSheet.Name & " (2)").visible = True
                    xlWorkBook.Worksheets(xlWorkSheet.Name & " (2)").name = machine
                    xlWorkSheet = xlWorkSheets(machine)
                    line = False
                    lines = 0

                    For j As Integer = 0 To UBound(jobs)
                        If machine.Contains(jobs(j).machine) Then
                            lines = lines + 3
                            If DateDiff(DateInterval.Hour, Now, jobs(j).date_due) >= frm_main.nud_scheduling_hour_mark.Value And j > 0 And Not line Then
                                lines = lines + 2
                                line = True
                            End If

                            If Not jobs(j).sched_text = Nothing Then
                                str = Split(jobs(j).sched_text, "|")
                                lines = lines + str.Length
                            End If

                        End If
                    Next

                    lines = lines + 3
                    line = False
                    Dim xlarray(lines, 7) As Object
                    Dim r As Integer = 0
                    For j As Integer = 0 To UBound(jobs)
                        If machine.Contains(jobs(j).machine) Then
                            count = count + 1
                            frm_main.tssl_status.Text = "Status: Processing... " & count & " of " & total
                            'frm_main.tspb_progress.Value = count / total * 100
                            If DateDiff(DateInterval.Hour, Now, jobs(j).date_due) >= frm_main.nud_scheduling_hour_mark.Value And j > 0 And Not line Then
                                xlarray(r, 5) = frm_main.nud_scheduling_hour_mark.Value & " hr mark begins here"
                                r = r + 2
                                line = True
                            End If

                            Dim start_row As Integer = r + 4
                            Dim end_row As Integer = r + 5

                            If Not jobs(j).sched_text = Nothing Then
                                str = Split(jobs(j).sched_text, "|")
                            End If

                            xlarray(r, 0) = jobs(j).prod_num
                            xlarray(r + 1, 0) = jobs(j).inksys

                            xlarray(r, 1) = FormatNumber(jobs(j).req_quant_1, 1, , , TriState.False)
                            If Not IsNull(jobs(j).req_uom_2) Then xlarray(r + 1, 1) = FormatNumber(jobs(j).req_quant_2, 1, , , TriState.False)

                            xlarray(r, 2) = jobs(j).req_uom_1
                            If Not IsNull(jobs(j).req_uom_2) Then xlarray(r + 1, 2) = jobs(j).req_uom_2

                            xlarray(r, 3) = jobs(j).cyl_size
                            xlarray(r + 1, 3) = jobs(j).no_up

                            xlarray(r, 4) = Format(jobs(j).date_due, "dd/MM/yyyy")

                            If jobs(j).mat_num_semi > 0 Then
                                xlarray(r, 4) = jobs(j).mat_num_semi
                                xlarray(r + 1, 5) = jobs(j).mat_desc_semi
                            Else
                                xlarray(r, 4) = jobs(j).mat_num_fin
                                xlarray(r + 1, 5) = jobs(j).mat_desc_fin
                            End If

                            xlarray(r, 5) = jobs(j).label_desc
                            xlarray(r + 1, 6) = Replace(jobs(j).inks, ",", " ")
                            xlarray(r, 7) = get_status(jobs(j).status)


                            If Not jobs(j).sched_text = Nothing Then
                                str = Split(jobs(j).sched_text, "|")
                                For i1 = 0 To UBound(str)
                                    xlarray(r + 2 + i1, 4) = "Sch Txt"
                                    xlarray(r + 2 + i1, 5) = str(i1)
                                    'check_format(str(0), str(i), r)
                                Next

                            End If
                            For r = lines To 0 Step -1
                                If Not IsNull(xlarray(r, 5)) Then
                                    r = r + 2
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                    xlWorkSheet.Range("A" & 4 & ":H" & lines + 4).Value = xlarray

                    For r = 0 To lines
                        frm_main.tssl_status.Text = "Status: Formatting... " & r & " of " & lines & " lines"
                        frm_main.tspb_progress.Value = r / lines * 100
                        If Not IsNull(xlarray(r, 7)) Then

                            If xlarray(r, 7) = "New Job" Or xlarray(r, 7) = "Label Required" Then
                                Dim r_a As Integer = 0
                                If Not IsNull(xlarray(r + 2, 4)) Then
                                    r_a = 1
                                End If

                                xlWorkSheet.Range("A" & r + 4 & ":G" & r + 5 + r_a).Font.Color = Color.Blue

                                r = r
                            End If
                        End If

                        If Not IsNull(xlarray(r, 5)) Then
                            If IsNull(xlarray(r, 4)) Then
                                If xlarray(r, 5) = frm_main.nud_scheduling_hour_mark.Value & " hr mark begins here" Then
                                    xlWorkSheet.Range("A" & r + 4 & ":G" & r + 4).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                    xlWorkSheet.Range("A" & r + 4 & ":G" & r + 4).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
                                    xlWorkSheet.Range("A" & r + 4 & ":G" & r + 4).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                    xlWorkSheet.Range("A" & r + 4 & ":G" & r + 4).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
                                    xlWorkSheet.Range("F" & r + 4).Font.Bold = True
                                End If
                            ElseIf xlarray(r, 4).ToString = "Sch Txt" Then
                                xlWorkSheet.Range("E" & r + 4 & ":F" & r + 4).Font.Bold = True
                                xlWorkSheet.Range("E" & r + 4 & ":F" & r + 4).Font.Underline = True
                                With frm_main.dgv_scheduling_format

                                    For f As Integer = 0 To .RowCount - 2
                                        If .Rows(f).Cells(1).Value = False Then
                                            str = Split(.Rows(f).Cells(0).Value, ",")
                                            For Each find As String In str

                                                If xlarray(r, 5).ToString.Contains(find) Then

                                                    If .Rows(f).Cells(2).Value.ToString.Contains("Bold") Then
                                                        xlWorkSheet.Range("A" & r + 2 & ":G" & r + 3).Font.Bold = True
                                                    End If
                                                    If .Rows(f).Cells(2).Value.ToString.Contains("Underline") Then
                                                        xlWorkSheet.Range("A" & r + 2 & ":G" & r + 3).Font.Underline = True
                                                    End If
                                                    If .Rows(f).Cells(2).Value.ToString.Contains("Strikeout") Then
                                                        xlWorkSheet.Range("A" & r + 2 & ":G" & r + 3).Font.Strikethrough = True
                                                    End If
                                                    Dim col() As String = Split(.Rows(f).Cells(3).Value, ",")
                                                    xlWorkSheet.Range("A" & r + 2 & ":G" & r + 3).Interior.Color = Color.FromName(col(0))
                                                    xlWorkSheet.Range("A" & r + 2 & ":G" & r + 3).Font.Color = Color.FromName(col(1))
                                                End If
                                            Next

                                        End If
                                    Next

                                End With
                            End If
                        End If

                    Next
                    frm_main.tssl_status.Text = "Status: Setting Page Breaks"
                    For pb As Integer = 1 To 3
                        Dim row As Integer = 77 * pb
                        If pb > 1 Then row = row - 1
                        If Not xlWorkSheet.Range("F" & row).Value = Nothing And Not xlWorkSheet.Range("F" & row + 1).Value = Nothing Then
                            Dim c As Integer = 0
                            For start As Integer = row To 0 Step -1
                                If xlWorkSheet.Range("F" & start).Value = Nothing Then
                                    For e As Integer = 0 To c - 1
                                        xlWorkSheet.Rows(start).Insert()
                                    Next
                                    Exit For
                                Else
                                    c = c + 1
                                End If
                            Next
                        End If

                    Next
                End If
            Next

            xlWorkBook.Worksheets("Printing Priority").activate()
        Else
            For i As Integer = 0 To frm_main.lb_scheduling_machines.Items.Count - 1
                If Not frm_main.lb_scheduling_machines.Items(i).ToString.Contains("All Machines") Then
                    str = Split(frm_main.lb_scheduling_machines.Items(i), "(")
                    Dim machine As String = str(0)
                    frm_main.tssl_status.Text = "Status: Addding a page for " & machine
                    xlWorkSheet = xlWorkSheets("Template Bagplant")
                    xlWorkSheet.Copy(After:=xlWorkSheets("Printing Priority"))
                    xlWorkBook.Worksheets(xlWorkSheet.Name & " (2)").visible = True
                    xlWorkBook.Worksheets(xlWorkSheet.Name & " (2)").name = machine
                    xlWorkSheet = xlWorkSheets(machine)
                    Dim last_seal As String = Nothing
                    Dim add As Integer = 0
                    If frm_main.lb_scheduling_departments.SelectedItem = "filmsline" And Not machine.Contains("EXFL") Then
                        add = 2
                    End If
                    lines = 0 + add
                    For j As Integer = 0 To UBound(jobs)
                        If Not jobs(j).prod_num = 0 Then

                            If machine.Contains(jobs(j).machine) Then
                                If jobs(j).machine.Contains("EXFL") Then
                                    lines = lines + 2
                                Else
                                    lines = lines + 3
                                End If

                                If Not jobs(j).sched_text = Nothing Then
                                    str = Split(jobs(j).sched_text, "|")
                                    lines = lines + str.Length
                                End If

                                If Not jobs(j).bag_info = Nothing Then
                                    lines = lines + 1
                                End If

                                If seal_change(last_seal, get_material_info(mat_desc:=jobs(j).mat_desc_fin).seal_type, jobs(j).machine) Then
                                    lines = lines + 2
                                End If

                                last_seal = get_material_info(mat_desc:=jobs(j).mat_desc_fin).seal_type

                            End If
                        End If
                    Next

                    lines = lines + 3
                    line = False
                    last_seal = Nothing
                    Dim xlarray(lines, 8) As Object
                    Dim r As Integer = 0
                    For j As Integer = 0 To UBound(jobs)
                        If Not jobs(j).prod_num = 0 Then

                            If machine.Contains(jobs(j).machine) Then
                                count = count + 1
                                frm_main.tssl_status.Text = "Status: Processing... " & count & " of " & total
                                ' frm_main.tspb_progress.Value = count / total * 100

                                If jobs(j).dept = "bagplant" Then
                                    If seal_change(last_seal, get_material_info(mat_desc:=jobs(j).mat_desc_fin).seal_type, jobs(j).machine) Then
                                        xlarray(r, 4) = "Seal Change."
                                        r = r + 2
                                    End If
                                    last_seal = get_material_info(mat_desc:=jobs(j).mat_desc_fin).seal_type
                                End If

                                xlarray(r, 5) = jobs(j).prod_num
                                xlarray(r, 6) = FormatNumber(jobs(j).req_quant_1, 3, , , TriState.False)
                                xlarray(r, 7) = jobs(j).req_uom_1
                                xlarray(r, 8) = get_status(jobs(j).status)

                                If Not jobs(j).sales_doc = Nothing Then
                                    str = Split(jobs(j).sales_doc, ":")
                                    If str.Length > 1 Then
                                        xlarray(r, 0) = str(0)
                                        xlarray(r + 1, 0) = str(1)
                                        str = Split(jobs(j).sales_line, ":")
                                        xlarray(r, 1) = str(0)
                                        xlarray(r + 1, 1) = str(1)
                                    Else
                                        xlarray(r, 0) = jobs(j).sales_doc
                                        xlarray(r, 1) = jobs(j).sales_line
                                    End If
                                End If

                                xlarray(r, 2) = jobs(j).customer

                                If jobs(j).mat_num_fin > 0 Then
                                    xlarray(r, 3) = jobs(j).mat_num_fin
                                    xlarray(r, 4) = jobs(j).mat_desc_fin
                                Else
                                    xlarray(r, 3) = jobs(j).mat_num_semi
                                    xlarray(r, 4) = jobs(j).mat_desc_semi
                                End If

                                If jobs(j).issued Then
                                    xlarray(r + 1, 5) = "INVENTORY"
                                    xlarray(r + 1, 3) = "ISSUED"
                                ElseIf Not jobs(j).machine.Contains("EXFL") Then
                                    xlarray(r + 1, 3) = jobs(j).mat_num_in
                                    xlarray(r + 1, 4) = jobs(j).mat_desc_in
                                    xlarray(r + 1, 6) = FormatNumber(jobs(j).issue_quant, 1, , , TriState.False)
                                    xlarray(r + 1, 7) = jobs(j).issue_uom
                                End If

                                If Not jobs(j).bag_info = Nothing Then
                                    xlarray(r + 2, 3) = "Bag Info"
                                    xlarray(r + 2, 4) = jobs(j).bag_info
                                    r = r + 1
                                End If

                                If Not jobs(j).sched_text = Nothing Then
                                    str = Split(jobs(j).sched_text, "|")
                                    For i1 = 0 To UBound(str)
                                        xlarray(r + 2 + i1, 3) = "Sch Txt"
                                        xlarray(r + 2 + i1, 4) = str(i1)
                                        'check_format(str(0), str(i), r)
                                    Next

                                End If
                                For r = lines To 0 Step -1
                                    If Not IsNull(xlarray(r, 3)) Or Not IsNull(xlarray(r, 4)) Then
                                        r = r + 2
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    Next
                    xlWorkSheet.Range("A" & 3 + add & ":I" & lines - 2).Value = xlarray
                    xlWorkSheet.PageSetup.PrintArea = "$A$1:$H$" & lines - 2
                    If add > 0 Then
                        xlWorkSheet.Range("A3:I3").Merge()
                        xlWorkSheet.Range("A3:I3").Font.Size = 28
                        xlWorkSheet.Range("A3:I3").Font.Bold = True
                        xlWorkSheet.Range("A3:I3").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        xlWorkSheet.Range("A3:I3").Value = "Priority " & InputBox("What Priority is " & machine, "Set Priority")
                    End If
                    For r = 0 To lines
                        frm_main.tssl_status.Text = "Status: Formatting... " & r & " of " & lines & " lines"
                        frm_main.tspb_progress.Value = r / lines * 100
                        If Not IsNull(xlarray(r, 8)) Then

                            If xlarray(r, 8) = "New Job" Or xlarray(r, 8) = "Label Required" Then
                                Dim r_a As Integer = 0
                                If Not IsNull(xlarray(r + 2, 3)) Then
                                    r_a = 1
                                End If

                                xlWorkSheet.Range("F" & r + 3 + add & ":H" & r + 4 + r_a + add).Font.Color = Color.Blue

                                r = r
                            End If
                        End If
                        If Not IsNull(xlarray(r, 4)) Then
                            If IsNull(xlarray(r, 3)) Then
                                If xlarray(r, 4) = "Seal Change." Then
                                    xlWorkSheet.Range("A" & r + 3 & ":H" & r + 3).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                    xlWorkSheet.Range("A" & r + 3 & ":H" & r + 3).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
                                    xlWorkSheet.Range("A" & r + 3 & ":H" & r + 3).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                    xlWorkSheet.Range("A" & r + 3 & ":H" & r + 3).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
                                    xlWorkSheet.Range("E" & r + 3).Font.Bold = True
                                End If
                            ElseIf xlarray(r, 3).ToString = "Sch Txt" Then
                                xlWorkSheet.Range("D" & r + 3 + add & ":E" & r + 3 + add).Font.Bold = True
                                xlWorkSheet.Range("D" & r + 3 + add & ":E" & r + 3 + add).Font.Underline = True
                                With frm_main.dgv_scheduling_format

                                    For f As Integer = 0 To .RowCount - 1
                                        If .Rows(f).Cells(1).Value = False Then
                                            str = Split(.Rows(f).Cells(0).Value, ",")
                                            For Each find As String In str

                                                If xlarray(r, 4).ToString.Contains(find) Then
                                                    Dim r_a As Integer = 0
                                                    If xlarray(r + 3, 3) = "Bag Info" Then
                                                        r_a = 1
                                                    End If
                                                    If .Rows(f).Cells(2).Value.ToString.Contains("Bold") Then
                                                        xlWorkSheet.Range("A" & r + add & ":H" & r + 2 + r_a + add).Font.Bold = True
                                                    End If
                                                    If .Rows(f).Cells(2).Value.ToString.Contains("Underline") Then
                                                        xlWorkSheet.Range("A" & r + add & ":H" & r + 2 + r_a + add).Font.Underline = True
                                                    End If
                                                    If .Rows(f).Cells(2).Value.ToString.Contains("Strikeout") Then
                                                        xlWorkSheet.Range("A" & r + add & ":H" & r + 2 + r_a + add).Font.Strikethrough = True
                                                    End If
                                                    Dim col() As String = Split(.Rows(f).Cells(3).Value, ",")
                                                    xlWorkSheet.Range("A" & r + add & ":H" & r + 2 + r_a + add).Interior.Color = Color.FromName(col(0))
                                                    xlWorkSheet.Range("A" & r + add & ":H" & r + 2 + r_a + add).Font.Color = Color.FromName(col(1))
                                                End If
                                            Next
                                        End If
                                    Next
                                End With
                            ElseIf xlarray(r, 3).ToString = "Bag Info" Then
                                With frm_main.dgv_scheduling_format


                                    For f As Integer = 0 To .RowCount - 1
                                        If .Rows(f).Cells(1).Value = True Then
                                            str = Split(.Rows(f).Cells(0).Value, ",")
                                            For Each find As String In str
                                                If xlarray(r, 4).ToString.Contains(find) And ValidStamp(get_material_info(mat_desc:=xlarray(r - 2, 4)).print_type) Then

                                                    If .Rows(f).Cells(2).Value.ToString.Contains("Bold") Then
                                                        xlWorkSheet.Range("D" & r + 3 & ":E" & r + 3).Font.Bold = True
                                                    End If
                                                    If .Rows(f).Cells(2).Value.ToString.Contains("Underline") Then
                                                        xlWorkSheet.Range("D" & r + 3 & ":E" & r + 3).Font.Underline = True
                                                    End If
                                                    If .Rows(f).Cells(2).Value.ToString.Contains("Strikeout") Then
                                                        xlWorkSheet.Range("D" & r + 3 & ":E" & r + 3).Font.Strikethrough = True
                                                    End If
                                                    Dim col() As String = Split(.Rows(f).Cells(3).Value, ",")
                                                    xlWorkSheet.Range("D" & r + 3 & ":E" & r + 3).Interior.Color = Color.FromName(col(0))
                                                    xlWorkSheet.Range("D" & r + 3 & ":E" & r + 3).Font.Color = Color.FromName(col(1))
                                                End If
                                            Next
                                        End If
                                    Next

                                End With
                            End If

                        End If
                    Next
                    Dim inserted As Integer = 0
                    frm_main.tssl_status.Text = "Status: Setting Page Breaks"
                    For pb As Integer = 1 To 3
                        Dim row As Integer = 94 * pb
                        If Not xlWorkSheet.Range("D" & row).Value = Nothing And Not xlWorkSheet.Range("D" & row + 1).Value = Nothing Then
                            Dim c As Integer = 0
                            For start As Integer = row To 0 Step -1
                                If xlWorkSheet.Range("D" & start).Value = Nothing Then
                                    For e As Integer = 0 To c - 1
                                        xlWorkSheet.Rows(start).Insert()
                                        inserted = inserted + 1
                                    Next
                                    Exit For
                                Else
                                    c = c + 1
                                End If
                            Next
                        End If

                    Next
                    xlWorkSheet.PageSetup.PrintArea = "$A$1:$H$" & lines - 2 + inserted

                End If
            Next
            If frm_main.lb_scheduling_departments.SelectedItem = "bagplant" Then xlWorkBook.Worksheets("Bagplant Priority").activate()

        End If
        frm_main.tspb_progress.Value = 100
        frm_main.tssl_status.Text = "Status: All done!"
        xlWorkBook.Save()
        xlApp.ScreenUpdating = True
        SetForegroundWindow(get_handle(xlWorkBook.Name))

    End Sub

    Function ValidStamp(ByVal Stamp As String) As Boolean
        'Only require logo colour when stamped in Conversion, print Types, S, I, and X (and XA, XK, XN)
        Select Case Stamp
            Case "S", "I", "X", "XA", "XK", "XN"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Function xlLastRow() As Integer
        Return xlWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Microsoft.Office.Interop.Excel.XlSpecialCellsValue.xlTextValues).Row
    End Function


    Function seal_change(ByVal last_seal As String, ByVal seal As String, ByVal machine As String) As Boolean

        If last_seal = seal Or last_seal = Nothing Then
            Return False
        Else
            seal_change = True
            '         If Not Right(machine, 2) = 17 Then
            If seal = "TE" Or seal = "SE" Then
                If last_seal = "TE" Or last_seal = "SE" Then
                    Return False
                Else
                    Return True
                End If
            End If
            'End If
        End If
    End Function

    Sub scheduling_setting_save()
        Dim s As String = Nothing
        If frm_main.lb_scheduling_print_order.Items.Count > 0 And frm_main.lb_scheduling_print_departments.SelectedIndex > -1 Then
            For i As Integer = 0 To frm_main.lb_scheduling_print_order.Items.Count - 1
                If s = Nothing Then
                    s = frm_main.lb_scheduling_print_order.Items(i)
                Else
                    s = s & "," & frm_main.lb_scheduling_print_order.Items(i)
                End If
            Next
            WriteIni(Environ("USERPROFILE") & "\settings.ini", "Print Order", frm_main.lb_scheduling_print_departments.SelectedItem, s)
        End If
        WriteIni(Environ("USERPROFILE") & "\settings.ini", "Text Formatting", Nothing, Nothing)

        With frm_main.dgv_scheduling_format

            If .RowCount > 0 Then
                For r As Integer = 0 To .RowCount - 1
                    If Not IsNull(.Rows(r).Cells(0).Value) Then
                        s = Nothing
                        For c As Integer = 0 To .ColumnCount - 1
                            If c = 1 Then
                                s = add_to_string(s, Math.Abs(CInt(.Rows(r).Cells(c).Value)), "|", True)
                            Else
                                s = add_to_string(s, .Rows(r).Cells(c).Value, "|", True)
                            End If
                        Next
                        WriteIni(Environ("USERPROFILE") & "\settings.ini", "Text Formatting", r, s)
                    End If
                Next
            End If
        End With
        WriteIni(Environ("USERPROFILE") & "\settings.ini", "Label Selection", Nothing, Nothing)
        With frm_main.dgv_scheduling_label

            If .RowCount > 0 Then
                For r As Integer = 0 To .RowCount - 1
                    If Not IsNull(.Rows(r).Cells(0).Value) Then
                        s = Nothing
                        For c As Integer = 0 To .ColumnCount - 1
                            s = add_to_string(s, .Rows(r).Cells(c).Value, "|", True)
                        Next
                        WriteIni(Environ("USERPROFILE") & "\settings.ini", "Label Selection", r, s)
                    End If
                Next
            End If
        End With
    End Sub

    Sub scheduling_setting_load()
        s_loading = True
        Dim s As String = Nothing
        If frm_main.lb_scheduling_print_departments.Items.Count > 0 Then
            s = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Print Order", frm_main.lb_scheduling_print_departments.SelectedItem, Nothing)
            If Not s = Nothing Then
                frm_main.lb_scheduling_print_order.Items.Clear()
                frm_main.lb_scheduling_print_order.Items.AddRange(Split(s, ","))
            Else
                s = s
            End If
        End If
        Dim i As Integer = 0
        Dim str() As String
        s = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Text Formatting", i, Nothing)
        With frm_main.dgv_scheduling_format
            .Rows.Clear()
            While Not s = Nothing
                str = Split(s, "|")
                .Rows.Add()
                For c As Integer = 0 To UBound(str)
                    If c = 1 Then
                        .Rows(i).Cells(c).Value = CBool(str(c))
                    Else
                        .Rows(i).Cells(c).Value = str(c)
                    End If
                Next
                i = i + 1
                s = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Text Formatting", i, Nothing)
            End While
            .ClearSelection()
        End With
        i = 0
        s = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Label Selection", i, Nothing)
        With frm_main.dgv_scheduling_label
            .Rows.Clear()
            While Not s = Nothing
                str = Split(s, "|")
                .Rows.Add()
                For c As Integer = 0 To UBound(str)
                    .Rows(i).Cells(c).Value = str(c)
                Next
                i = i + 1
                s = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Label Selection", i, Nothing)
            End While
            .ClearSelection()
        End With
        s_loading = False
    End Sub

    Sub format_text(ByVal row As Integer)
        With frm_main

            .tb_scheduling_format.Font = get_text_format()
            .tb_scheduling_format.BackColor = .l_scheduling_colour_interior.BackColor
            .tb_scheduling_format.ForeColor = .l_scheduling_colour_font.BackColor

        End With

    End Sub
    Function get_text_format() As Font
        Dim style As FontStyle
        With frm_main
            If .cb_scheduling_format_strikeout.Checked Then style = style Or FontStyle.Strikeout
            If .cb_scheduling_format_bold.Checked Then style = style Or FontStyle.Bold
            If .cb_scheduling_format_underline.Checked Then style = style Or FontStyle.Underline
            Return New Font("Microsoft Sans Serif", 8.25, style)
        End With
    End Function
End Module

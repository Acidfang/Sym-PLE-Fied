Public Class frm_main

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        Me.SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        ' sql.delete("db_schedule", "prod_num", "=", 5009885842)
        'Add any initialization after the InitializeComponent() call.
        s_plant = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Location", "Plant", 3203)
        s_machine = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Location", "Machine", Nothing)
        s_department = LCase(ReadIni(Environ("USERPROFILE") & "\settings.ini", "Location", "department", "printing"))
        s_printer = ReadIni(Environ("USERPROFILE") & "\settings.ini", "ACSIS", "Printer", Nothing)
        s_extruder = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Machine", "Extruder", False)
        Dim file As String = get_file_path("burst_data")
        If System.IO.File.Exists(file) Then
            If Not System.IO.File.Exists(Environ("USERPROFILE") & "\specification master sheet with Aussie.xls") Or _
                Not System.IO.File.GetLastWriteTime(file) = _
                System.IO.File.GetLastWriteTime(Environ("USERPROFILE") & "\specification master sheet with Aussie.xls") Then
                System.IO.File.Copy(file, _
                                    Environ("USERPROFILE") & "\specification master sheet with Aussie.xls", True)
            End If
        End If
        l_main_sched_text.Text = Nothing
        With Me
            .StartPosition = FormStartPosition.CenterScreen
        End With
        With dgv_job_requirements
            For i As Integer = 0 To .ColumnCount - 1
                .Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
        End With
        dt_symple_session = sql.get_table("acsis_session", "plant", "=", s_plant)
        w_sessions(0).process = New Process
        w_sessions(1).process = New Process

        Timer1.Enabled = True
        s_loading = False
    End Sub
    Private Sub UpdateControls(ByVal myControlIN As Control)
        Dim myControl

        For Each myControl In myControlIN.Controls
            UpdateControls(myControl)
            Dim my_ctrl As Control

            If Not TypeOf myControl Is DataGridView And Not TypeOf myControl Is TextBox And Not TypeOf myControl Is ComboBox Then
                my_ctrl = myControl
                AddHandler my_ctrl.MouseClick, AddressOf dgv_focus
            End If
        Next

    End Sub

    Private Sub dgv_focus(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If tc_main.SelectedIndex = -1 Then Exit Sub
        Select Case tc_main.TabPages(tc_main.SelectedIndex).Name
            Case "tp_production"
                dgv_production.Focus()
            Case "tp_summary"
                dgv_summary.Focus()
        End Select
    End Sub
#Region "Buttons"

    Private Sub b_main_settings_Click(sender As Object, e As EventArgs) Handles b_main_settings.Click
        With frm_settings
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub

    Private Sub b_summary_print_Click(sender As Object, e As EventArgs) Handles b_summary_print.Click
        print_summary()
    End Sub

    Private Sub b_summary_remove_job_Click(sender As Object, e As EventArgs) Handles b_summary_remove_job.Click
        With dgv_summary
            If .RowCount = 0 Then Exit Sub
            Dim response As MsgBoxResult = MsgBox("Are you sure you want to delete this job?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Delete Job")
            If response = MsgBoxResult.Yes Then

                If .SelectedRows.Count >= 0 Then
                    s_loading = True
                    sql.delete("db_summary", "prod_num", "=", .CurrentRow.Cells(0).Value & " AND date_='" & tb_start_date.Text & "' AND user_='" & s_user & "' AND shift='" & s_shift & "' AND machine='" & s_machine & "'")
                    .Rows.RemoveAt(dgv_summary.CurrentCell.RowIndex)
                    export_summary()
                    If dgv_summary.RowCount > 0 Then
                        s_selected_job = dt_schedules.Select("prod_num='" & .CurrentRow.Cells(0).Value & "'")
                        display_job()
                    End If

                    s_loading = False
                End If
            End If

        End With
    End Sub

    Private Sub b_summary_symple_times_Click(sender As Object, e As EventArgs) Handles b_summary_symple_times.Click
        If dgv_summary.RowCount = 0 Then Exit Sub

        If Not IsNull(dgv_summary.Rows(0).Cells("dgc_summary_start").Value) Then
            With frm_symple_times
                .StartPosition = FormStartPosition.CenterParent
                .ShowDialog()
            End With
        End If
    End Sub

    Private Sub b_main_logoff_Click(sender As Object, e As EventArgs) Handles b_main_logoff.Click

        Dim resp As MsgBoxResult = MsgBoxResult.Yes
        With dgv_summary
            For r As Integer = 0 To .RowCount - 1
                If IsNull(.Rows(r).Cells("dgc_summary_symple_times").Value) And _
                    IsNull(.Rows(r).Cells("dgc_summary_symple_times").Value) Then
                    .CurrentCell = .Rows(r).Cells("dgc_summary_prod_num")
                    s_selected_job = dt_schedules.Select("prod_num=" & .Rows(r).Cells("dgc_summary_prod_num").Value)
                    display_job()
                    frm_status.Close()
                    resp = MsgBox("You have not entered Sym-PLE information for " & .Rows(r).Cells("dgc_summary_prod_num").Value & vbCr & _
                                   "Are you sure you want to log off?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Sym-PLE Input missing!")
                    If resp = MsgBoxResult.No Then
                        Exit Sub
                    End If
                    Exit For
                End If
            Next
        End With

        s_loading = True
        dgv_summary.Rows.Clear()
        clear_job()
        s_loading = False
        s_user = Nothing
        s_pass = Nothing
        s_user_name = Nothing
        s_assistant = Nothing
        s_shift = Nothing
        lb_personel.Items.Clear()
        tc_main.SelectedIndex = 0
        tb_start_date.Text = Format(Now, "dd/MM/yyyy")
        Dim tWnd As IntPtr = FindWindowByCaption(0, lb_symple_sesions.SelectedItem)
        w_sessions(0).process.Kill()
        lb_symple_sesions.Items(0) = " "
        lb_symple_sesions.SelectedIndex = -1
        lb_symple_users.Items(0) = " "
        lb_symple_users.SelectedIndex = -1
    End Sub

    Private Sub b_summary_open_schedule_Click(sender As Object, e As EventArgs) Handles b_summary_open_schedule.Click
        If s_department = "mounting" And lb_personel.Items.Count = 0 Then
            MsgBox("You need to add a mounter first", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "No mounter added")
        Else
            With frm_schedules
                .StartPosition = FormStartPosition.CenterParent
                .ShowDialog()
            End With
        End If

    End Sub

    Private Sub b_summary_restore_Click(sender As Object, e As EventArgs) Handles b_summary_restore.Click
        With frm_restore
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub

    Private Sub b_summary_clear_Click(sender As Object, e As EventArgs) Handles b_summary_clear.Click
        s_loading = True
        export_summary()
        dgv_summary.Rows.Clear()
        tb_start_date.Text = Format(Now, "dd/MM/yyyy")
        s_shift = Nothing
        s_loading = False
    End Sub

    Private Sub b_enter_item_Click(sender As Object, e As EventArgs) Handles b_report_enter_item.Click
        Dim num_in, first_roll, last_roll, count As Integer, kg_in, km_in As Decimal, _
            mat_num, prod_num, prod_num_orig As Long, _
            str, item_id As String, _
            exit_do As Boolean = False

        item_id = Nothing
        If s_assistant = Nothing And s_department = "printing" Then
            If s_preparing Then
                s_assistant = s_user_name
            Else
                If lb_personel.Items.Count = 0 Then
                    MsgBox("You need to add a helper")
                    With frm_helper
                        .StartPosition = FormStartPosition.CenterParent
                        .Button3.Enabled = True
                        .ShowDialog()
                    End With
                Else
                    For i As Integer = 0 To lb_personel.Items.Count - 1
                        If s_assistant = Nothing Then
                            s_assistant = lb_personel.Items(i)
                        Else
                            s_assistant = s_assistant & "/" & lb_personel.Items(i)
                        End If
                    Next
                End If
            End If

        End If
        If s_assistant = Nothing And s_department = "printing" Then Exit Sub
        With dgv_production
            For r As Integer = 0 To .RowCount - 1
                If .Rows(r).Cells("dgc_production_num_in").Value.ToString.Contains("&") Then
                    Dim nums_in() As String
                    nums_in = Split(dt_produced(r)("num_in"), " & ")
                    For i As Integer = 0 To UBound(nums_in)
                        If nums_in(i) > num_in Then num_in = nums_in(i)
                    Next
                ElseIf IsNumeric(.Rows(r).Cells("dgc_production_num_in").Value) Then
                    If .Rows(r).Cells("dgc_production_num_in").Value > num_in Then num_in = .Rows(r).Cells("dgc_production_num_in").Value
                End If
            Next
            num_in = num_in + 1
        End With
        dt_produced = sql.get_table("db_produced", "prod_num", "=", tb_prod_num.Text)
        If IsNull(dt_produced) Then
            num_in = 1
        Else
            For r As Integer = 0 To dt_produced.Rows.Count - 1
                If dt_produced(r)("num_in").ToString.Contains("&") Then
                Else
                    If dt_produced(r)("num_in") > num_in Then num_in = dt_produced(r)("num_in")
                End If
            Next
        End If
        first_roll = num_in
        prod_num = tb_prod_num.Text
        If cbo_mat_num_in.Items.Count = 0 Then
            With frm_material_info
                .StartPosition = FormStartPosition.CenterParent
                .ShowDialog()
            End With
            If cbo_mat_num_in.Items.Count = 0 Then Exit Sub
        End If
        Select Case s_department

            Case "printing", "bagplant", "filmsline"

                If cbo_mat_num_in.Items.Count = 0 Then
                    With frm_material_info
                        .StartPosition = FormStartPosition.CenterParent
                        .ShowDialog()
                    End With
                End If
                If cbo_mat_num_in.Items.Count = 0 And s_department = "printing" Then
                    If s_assistant.Contains(s_user_name) Then s_assistant = Nothing
                    Exit Sub
                End If
                s_loading = True
                SetForegroundWindow(Me.Handle)
                With dgv_production
                    If .RowCount > 0 Then

                        If s_department = "filmsline" And can_delete(.RowCount - 1) Then
                            Dim resp As MsgBoxResult = MsgBox("Do you want to clear empty lines?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Clear empty rows")
                            If resp = MsgBoxResult.Yes Then
                                For r As Integer = .RowCount - 1 To 0 Step -1
                                    If can_delete(r) And .Rows(r).Cells("dgc_production_num_in").Value = "-" Then
                                        .Rows.RemoveAt(r)
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    End If
                    Do While 1 = 1
                        'enter code here to pickup big barcode
                        If exit_do Then Exit Do
                        Do While 1 = 1
                            If exit_do Then Exit Do
                            frm_status.Close()
                            s_roll_info = Nothing
                            With s_roll_info
                                item_id = .item_id
                                mat_num = .mat_num
                                kg_in = .kg_in
                                km_in = .km_in
                            End With
                            str = InputBox("Enter the production number on the label" & vbCrLf & "Roll: " & num_in, "Enter Production Number", prod_num_orig)
                            If str = Nothing Then
                                exit_do = True
                                Exit Do
                            ElseIf IsNumeric(str) Then
                                prod_num_orig = str
                                Dim dt As DataTable = sql.get_table("db_schedule", "prod_num", "=", "'" & prod_num_orig & "'")
                                If Not dt Is Nothing Then
                                    Dim dr() As DataRow = dt.Select("mat_num_semi=" & s_selected_job(0)("mat_num_in"))
                                    If dr.Length > 0 Then
                                        dt = sql.get_table("db_produced", "prod_num", "=", "'" & prod_num_orig & "' AND kg_out IS NOT NULL")
                                        'Dim dt_used As DataTable = sql.get_table("db_produced", "prod_num_orig", "=", "'" & prod_num_orig & "' AND rts IS NOT NULL")
                                        ' used comparison code
                                        If Not dt Is Nothing Then
                                            With frm_item_info
                                                .DataGridView1.Rows.Clear()
                                                For row As Integer = 0 To dt.Rows.Count - 1
                                                    Dim add As Boolean = True
                                                    For r As Integer = 0 To dgv_production.RowCount - 1
                                                        If Not dgv_production.Rows(r).Cells("dgc_production_item_id").Value.ToString = "-" Then
                                                            If dgv_production.Rows(r).Cells("dgc_production_prod_num_orig").Value = prod_num_orig And _
                                                                dgv_production.Rows(r).Cells("dgc_production_item_id").Value = dt(row)("num_out") Then
                                                                add = False
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next
                                                    If add Then
                                                        .DataGridView1.Rows.Add()
                                                        .DataGridView1.Rows(.DataGridView1.RowCount - 1).Cells(0).Value = dt(row)("num_out")
                                                        .DataGridView1.Rows(.DataGridView1.RowCount - 1).Cells(1).Value = _
                                                            FormatNumber(dt(row)("kg_out"), 1, , , TriState.False)
                                                        If dt(row)("dept") = "filmsline" Then
                                                            .DataGridView1.Rows(.DataGridView1.RowCount - 1).Cells(2).Value = _
                                                                FormatNumber(dt(row)("km_out") / 1000, 2, , , TriState.False)
                                                        Else
                                                            .DataGridView1.Rows(.DataGridView1.RowCount - 1).Cells(2).Value = _
                                                                FormatNumber(dt(row)("km_out"), 0, , , TriState.False)
                                                        End If
                                                        .DataGridView1.Rows(.DataGridView1.RowCount - 1).Cells(3).Value = dr(0)("mat_num_semi")
                                                    End If
                                                Next
                                                .DataGridView1.Sort(.DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
                                                If .DataGridView1.RowCount > 0 Then
                                                    .StartPosition = FormStartPosition.CenterParent
                                                    .ShowDialog()
                                                Else
                                                    MsgBox("There are no other items to add from this job number")
                                                    Exit Do
                                                End If
                                            End With
                                            With s_roll_info
                                                item_id = .item_id
                                                mat_num = .mat_num
                                                kg_in = .kg_in
                                                km_in = .km_in
                                            End With
                                        End If
                                    End If
                                End If
                            End If
                            Do While item_id = Nothing

                                str = InputBox("Enter the material number on the label" & vbCrLf & "Roll: " & num_in)
                                If str = Nothing Then
                                    exit_do = True
                                    Exit Do
                                ElseIf IsNumeric(str) Then
                                    mat_num = str
                                End If
                                If cbo_mat_num_in.FindStringExact(mat_num) > -1 Then

                                    cbo_mat_num_in.SelectedIndex = cbo_mat_num_in.FindStringExact(mat_num)
                                    cbo_mat_desc_in.SelectedIndex = cbo_mat_num_in.SelectedIndex

                                    str = InputBox("Enter the roll number on the label" & vbCrLf & "Roll: " & num_in)
                                    If str = Nothing Then
                                        exit_do = True
                                        Exit Do
                                    Else
                                        item_id = str
                                        str = Nothing
                                    End If
                                    If exit_do Then Exit Do
                                    Do While str = Nothing
                                        str = InputBox("Enter the roll weight on the label" & vbCrLf & "Roll: " & num_in & vbCrLf & "Roll ID: " & item_id)
                                        If Not str = Nothing And IsNumeric(str) Then
                                            kg_in = str
                                        ElseIf str = Nothing Then
                                            kg_in = 0
                                            str = "0"
                                        Else
                                            MsgBox("The result of the scan is """ & str & """ and this is not a valid number, please try again", _
                                                   MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Incorrect item")
                                            str = Nothing
                                        End If
                                    Loop
                                Else
                                    MsgBox("That material number does not match!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error!")
                                End If
                            Loop
                            If Not item_id = Nothing Then

                                count = count + 1
                                Dim lines As Integer = nud_items_out.Value
                                If s_department = "bagplant" Then lines = 1
                                For i As Integer = 1 To lines
                                    .Rows.Add()
                                    .Rows(.RowCount - 1).Cells("dgc_production_prod_num").Value = prod_num
                                    If i = 1 Then
                                        .Rows(.RowCount - 1).Cells("dgc_production_num_in").Value = num_in
                                        .Rows(.RowCount - 1).Cells("dgc_production_item_id").Value = item_id
                                        .Rows(.RowCount - 1).Cells("dgc_production_kg_in").Value = FormatNumber(kg_in, 1)
                                        .Rows(.RowCount - 1).Cells("dgc_production_km_in").Value = km_in
                                        .Rows(.RowCount - 1).Cells("dgc_production_kg_in_orig").Value = FormatNumber(kg_in, 1)
                                        .Rows(.RowCount - 1).Cells("dgc_production_mat_num").Value = mat_num
                                        .Rows(.RowCount - 1).Cells("dgc_production_prod_num_orig").Value = prod_num_orig

                                    Else
                                        .Rows(.RowCount - 1).Cells("dgc_production_num_in").Value = "-"
                                        .Rows(.RowCount - 1).Cells("dgc_production_item_id").Value = "-"
                                        .Rows(.RowCount - 1).Cells("dgc_production_kg_in").Value = "-"
                                    End If
                                    .Rows(.RowCount - 1).Cells("dgc_production_assistant").Value = s_assistant

                                Next

                                last_roll = num_in
                                num_in = num_in + 1

                            End If
                        Loop

                    Loop
                    If last_roll = 0 Then
                        Exit Sub
                    End If
                    str = Strings.Mid(tb_mat_desc_out.Text, 10, 1)
                    If get_material_info.material_type = material_type.barrier And count > 0 And s_department = "printing" Then
                        If MsgBox("Print the tickets?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Print Tickets") = MsgBoxResult.Yes Then
                            xlApp = New Microsoft.Office.Interop.Excel.Application
                            xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
                            xlApp.ScreenUpdating = False
                            xlWorkBook = xlApp.Workbooks.Open("Z:\Production\Printing Dept\Printing (Prod logon)\Data\Roll Tickets.xls", [ReadOnly]:=True)
                            xlWorkSheets = xlWorkBook.Worksheets
                            xlWorkSheet = xlWorkSheets(1)
                            xlWorkSheet.Unprotect()
                            xlWorkSheet.Range("A5").Value = cbo_mat_desc_in.Text
                            xlWorkSheet.Range("H1").Value = cbo_mat_num_in.Text
                            xlWorkSheet.Range("F26").Value = Replace(s_assistant, "'", Nothing)
                            xlWorkSheet.Range("A31").Value = "IN-SPEC"
                            Dim response As MsgBoxResult
                            If first_roll = last_roll Then
                                response = MsgBox("Number the ticket?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Number Ticket")
                            Else
                                response = MsgBox("Number the tickets?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Number Tickets")
                            End If
                            xlApp.ScreenUpdating = True
                            For Each s As Microsoft.Office.Interop.Excel.Shape In xlWorkSheet.Shapes
                                If s.Name = "Oval 6" Then
                                    If response = MsgBoxResult.Yes Then
                                        s.Visible = True
                                    End If
                                    For i As Integer = first_roll To last_roll
                                        s.TextFrame.Characters.Text = i
                                        xlWorkSheet.PrintOutEx(1, 1, 1, False)
                                        'xlApp.ScreenUpdating = False
                                    Next
                                    Exit For
                                End If
                            Next
                            'End If
                            xlWorkBook.Close(False)
                            close_excel()

                        End If
                    End If
                End With
        End Select
        If s_department = "printing" Then
            If s_assistant.Contains(s_user_name) Then s_assistant = Nothing
        End If
        If Not l_prep_warning.Visible Then calc_totals()
        s_loading = False
        calc_totals()
        export_report()
    End Sub

    Private Sub b_setup_reset_Click(sender As Object, e As EventArgs) Handles b_setup_reset.Click
        lb_inks.Items.Clear()
        lb_inks.Items.AddRange(lb_inks_list.Items)
        lb_inks.SelectedIndex = 0
        clear_inks()
    End Sub

    Private Sub b_report_part_roll_Click(sender As Object, e As EventArgs) Handles b_report_part_roll.Click
        Select Case b_report_part_roll.Text

            Case "End Prep"
                mod_frm_main.end_prep()
            Case "Part Roll"
                mod_frm_main.part_roll()
            Case "Item(s) Out"
                s_selected_job = dt_schedules.Select("prod_num=" & tb_prod_num.Text)
                Dim kg As String = InputBox("How many KG?", "KG Out", 0)
                If IsNumeric(kg) And Not kg = "0" Then
                    Dim m As String = tb_roll_length.Text
                    If cb_report_use_estimate.Checked Then
                        m = FormatNumber(kg * get_conversion_factor(True, tb_mat_num_out.Text, False), 0, , , TriState.False)
                    End If
                    m = InputBox("How many Meters?", "KG Out", m)
                    If IsNumeric(m) Then
                        If s_extruder And cb_production_ply_separator.Checked Then
                            mod_frm_main.items_out(kg, m, False)
                            mod_frm_main.items_out(kg, m, True)
                        Else
                            mod_frm_main.items_out(kg, m, True)
                        End If

                    End If

                End If
            Case "Carton Out"
                mod_frm_main.carton_out()
            Case Else
                MsgBox("L")
        End Select
        s_loading = False

    End Sub

    Private Sub b_main_show_zaw_Click(sender As Object, e As EventArgs) Handles b_main_show_zaw.Click
        open_zaw(b_main_show_zaw.Text)
    End Sub

    Private Sub b_comments_Click(sender As Object, e As EventArgs) Handles b_main_comments.Click
        With Comments_
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()

        End With
    End Sub

    Private Sub b_rts_send_Click(sender As Object, e As EventArgs) Handles b_rts_send.Click
        Dim movement As Integer = 104
        Select Case s_department
            Case "filmsline"
                movement = 103
            Case "printing"
                movement = 104
            Case "bagplant"
                movement = 105
        End Select

        Dim response As MsgBoxResult = MsgBox("RTS these items?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Send RTS")
        Dim c, i As Integer
        If response = MsgBoxResult.Yes Then
            Dim mat_num As Integer = 0
            Dim quant As Decimal = 0
            Dim items As Integer = 0

            With dgv_rts_pending

                For r As Integer = 0 To .RowCount - 1

                    mat_num = .Rows(r).Cells(0).Value
                    quant = .Rows(r).Cells(1).Value
                    items = .Rows(r).Cells(2).Value

                    acsis.open_window(cls_ACSIS.eWindowTarget.WarehouseReceiptsOther, 0)

                    show_status("Processing RTS")

                    acsis.select_cb_item(acsis.get_acsis_option("rts_screen_movement"), 0)

                    w_sessions(0).process.WaitForInputIdle()

                    acsis.insert_string(tb_prod_num.Text, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("rts_screen_prod_num")).handle)
                    acsis.insert_string(mat_num, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("rts_screen_mat_num")).handle)

                    While acsis.get_text(w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("rts_screen_desc")).handle, False) = Nothing
                        If c > 10 Then
                            acsis.click_button(acsis.get_button_back(0), False, 0)
                            acsis.open_window(cls_ACSIS.eWindowTarget.WarehouseReceiptsOther, 0)
                            acsis.select_cb_item(acsis.get_acsis_option("rts_screen_movement"), 0)
                            acsis.insert_string(tb_prod_num.Text, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("rts_screen_prod_num")).handle)
                            acsis.insert_string(mat_num, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("rts_screen_mat_num")).handle)
                            c = 0
                        End If
                        acsis.popup_check(True, 0)
                        If i = acsis.get_acsis_option("rts_screen_prod_num") Then
                            i = acsis.get_acsis_option("rts_screen_mat_num")
                        Else
                            i = acsis.get_acsis_option("rts_screen_prod_num")
                        End If
                        c = c + 1
                        SendMessage(w_sessions(0).winWnd(0).tbWnd(i).handle, WM_SETFOCUS, 0, 0)
                        w_sessions(0).process.WaitForInputIdle()
                    End While
                    acsis.insert_string(quant, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("rts_screen_quant")).handle)
                    acsis.insert_string(items, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("rts_screen_items")).handle)
                    acsis.insert_string("RTS", w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("rts_screen_dest_bin")).handle)
                    acsis.insert_string(movement, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("rts_screen_dest_type")).handle)
                    For b As Integer = 0 To w_sessions(0).winWnd(0).btnWnd.Length - 1
                        If w_sessions(0).winWnd(0).btnWnd(b).text(0) = acsis.get_acsis_option("rts_screen_put_away") Then
                            While w_sessions(0).winWnd(0).btnWnd(b).enabled = False
                                acsis.popup_check(True, 0)
                                If i = acsis.get_acsis_option("rts_screen_dest_bin") Then
                                    i = acsis.get_acsis_option("rts_screen_source_bin")
                                Else
                                    i = acsis.get_acsis_option("rts_screen_dest_bin")
                                End If
                                SendMessage(w_sessions(0).winWnd(0).tbWnd(i).handle, WM_SETFOCUS, 0, 0)
                                acsis.get_controls(0, True, False, True)
                                w_sessions(0).process.WaitForInputIdle()
                            End While
                            Exit For
                        End If
                    Next

                    acsis.click_button(acsis.get_acsis_option("rts_screen_put_away"), False, 0)
                    acsis.click_button(acsis.get_acsis_option("rts_screen_put_away"), False, 0)
                    acsis.click_button(acsis.get_button_back(0), False, 0)
                    s_loading = True
                Next
            End With

            With dgv_rts
                Dim row As Integer = 9
                For r As Integer = 0 To .RowCount - 1
                    If .Rows(r).Cells("dgc_rts_send").Value = True Then
                        If Not get_original_weight(.Rows(r).Cells("dgc_rts_item_id").Value, _
                                            .Rows(r).Cells("dgc_rts_num_in").Value) = CDec(.Rows(r).Cells("dgc_rts_quantity").Value) Then

                            acsis.open_window(cls_ACSIS.eWindowTarget.ReLabeling, 0)

                            show_status("Reprinting Tickets")

                            acsis.set_check(acsis.get_control_handle(acsis.get_acsis_option("relabel_screen_material"), _
                                                                     cls_ACSIS.eControlType.OptionButton, 0), True)

                            acsis.insert_string(mat_num, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("relabel_screen_mat_num")).handle)
                            While acsis.get_text(w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("relabel_screen_date")).handle, False) = Nothing ' date
                                If i = acsis.get_acsis_option("relabel_screen_to") Then
                                    i = acsis.get_acsis_option("relabel_screen_from")
                                Else
                                    i = acsis.get_acsis_option("relabel_screen_to")
                                End If
                                SendMessage(w_sessions(0).winWnd(0).tbWnd(i).handle, WM_SETFOCUS, 0, 0)
                                w_sessions(0).process.WaitForInputIdle()
                            End While
                            acsis.get_controls(0, True, False, True)

                            acsis.select_cb_item(s_label_rts, 0)
                            acsis.select_cb_item(s_printer, 0)

                            acsis.insert_string(.Rows(r).Cells("dgc_rts_quantity").Value, _
                                                w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("relabel_screen_quant")).handle)

                            acsis.insert_string(.Rows(r).Cells("dgc_rts_item_id").Value, _
                                                w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("relabel_screen_from")).handle)

                            acsis.insert_string(.Rows(r).Cells("dgc_rts_item_id").Value, _
                                                w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("relabel_screen_to")).handle)

                            acsis.insert_string(.Rows(r).Cells("dgc_rts_item_id").Value, _
                                                w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("relabel_screen_of")).handle)

                            acsis.click_button(acsis.get_acsis_option("relabel_screen_print"), False, 0)
                            acsis.click_button(acsis.get_button_back(0), False, 0)
                        End If
                        .Rows(r).Cells("dgc_rts_send").Value = False
                        .Rows(r).Cells("dgc_rts_quantity_symple").Value = .Rows(r).Cells("dgc_rts_quantity").Value
                        row = row + 1
                    End If
                Next

            End With
            export_rts()
            s_loading = False

            Select Case s_department
                Case "printing"
                    MsgBox("RTS successfully processed, the old RTS Slip is no longer needed so it will not be printed.")
                Case "filmsline"
                    print_rts()
            End Select
            dgv_rts_pending.Rows.Clear()
            b_rts_send.Enabled = False
        End If

        frm_status.Close()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles b_report_pallet.Click
        With dgv_production

            If Not .RowCount = 0 And Not IsNull(dt_produced) Then
                With frm_pallet
                    .StartPosition = FormStartPosition.CenterParent
                    .ShowDialog()
                End With
            End If

        End With

    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles b_reject_print.Click
        print_reject()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles b_report_merge.Click
        With frm_merge_roll
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub

#End Region

#Region "Combobox"



    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)
        If s_loading Then Exit Sub
        s_loading = True
        import_downtime("BS01")
        import_downtime("CS01")
        calc_downtime(get_prod_num_row(tb_prod_num.Text))
        s_loading = False
    End Sub

#End Region

#Region "Datagridview"

    Private Sub dgv_rts_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_rts.CellClick
        With dgv_rts
            .ClearSelection()
            If e.RowIndex = -1 Then Exit Sub
            If .Columns(e.ColumnIndex).Name = "dgc_rts_send" Then

                If .CurrentRow.Cells("dgc_rts_send").Value = True Then
                    .CurrentRow.Cells("dgc_rts_send").Value = False
                Else
                    If .CurrentRow.Cells("dgc_rts_quantity").Value = _
                        .CurrentRow.Cells("dgc_rts_quantity_symple").Value And _
                        .CurrentRow.Cells("dgc_rts_send").Value = False Then

                        Dim response As MsgBoxResult = MsgBox("This item has been detected in RTS already, do you want to resend?", _
                                                              MsgBoxStyle.YesNo, "RTS already done!")
                        If response = MsgBoxResult.Yes Then
                            .CurrentRow.Cells("dgc_rts_send").Value = True
                        End If
                    ElseIf .CurrentRow.Cells("dgc_rts_send").Value = False Then
                        .CurrentRow.Cells("dgc_rts_send").Value = True
                    Else
                        .CurrentRow.Cells("dgc_rts_send").Value = False
                    End If
                End If
            End If

        End With
    End Sub

    Private Sub dgv_rts_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_rts.CellValueChanged
        If s_loading Then Exit Sub
        If e.ColumnIndex < 3 Then export_rts()

        calc_rts()
    End Sub

    Private Sub dgv_summary_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_summary.CellClick

        Select Case dgv_summary.Columns(e.ColumnIndex).Name
            Case "dgc_summary_cyl_in", "dgc_summary_ink_in"
                Select Case s_department
                    Case "printing"
                        With frm_setup
                            .StartPosition = FormStartPosition.CenterParent
                            .ShowDialog()
                        End With
                        dgv_summary.ClearSelection()

                End Select
            Case "dgc_summary_bs01", "dgc_summary_cs01"
                Select Case s_department
                    Case "printing", "filmsline"
                        Dim resp As MsgBoxResult = MsgBox("Do you want to add downtime?", MsgBoxStyle.YesNo, "Downtime Selected")
                        If resp = MsgBoxResult.Yes Then
                            tc_main.SelectedIndex = 2
                        Else
                            dgv_summary.ClearSelection()
                        End If
                End Select

        End Select
        If e.RowIndex > -1 Then
            With lb_multiple_job_list
                If .Items.IndexOf(dgv_summary.CurrentRow.Cells(0).Value) = -1 Then
                    b_summary_multiple_job_add.Enabled = True
                Else
                    .SelectedIndex = .Items.IndexOf(dgv_summary.CurrentRow.Cells(0).Value)
                End If
            End With
        Else
            b_summary_multiple_job_add.Enabled = False
        End If
    End Sub


    Private Sub dgv_summary_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_summary.CellEnter
        If s_loading Or e.RowIndex = -1 Then Exit Sub
        If dgv_summary.CurrentRow Is Nothing Then Exit Sub
        If dgv_summary.CurrentRow.Cells("dgc_summary_mounter").Selected Then dgv_summary.ClearSelection()
        If dt_schedules Is Nothing Then
            dt_schedules = sql.get_table("db_schedule", "plant", "=", s_plant)
        End If
        's_selected_job = dt_schedule.Select("prod_num='" & dgv_summary.CurrentRow.Cells(0).Value & "'")
        If cms_summary.Visible Then Exit Sub


        ' schedule = sql.get_table("db_schedule", "plant", "=", s_plant)

    End Sub

    Private Sub dgv_summary_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_summary.CellDoubleClick
        If s_department = "mounting" Then Exit Sub
        dgv_summary.EndEdit()
        tc_main.SelectedIndex = 1
        s_selected_job = dt_schedules.Select("prod_num='" & dgv_summary.CurrentRow.Cells(0).Value & "'")
        'display_job()
    End Sub

    Private Sub dgv_summary_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_summary.CellEndEdit
        If s_loading Then Exit Sub
        s_loading = True
        With dgv_summary
            Dim dp As Integer = 0
            If e.ColumnIndex > 4 And .Columns(e.ColumnIndex).Visible And Not .Columns(e.ColumnIndex).Name = "dgc_summary_mpm" Then
                Dim s As String = Strings.Right(.Columns(e.ColumnIndex).DefaultCellStyle.Format, 1)
                If IsNumeric(s) Then
                    dp = s
                End If
                ' .CurrentRow.Cells(e.ColumnIndex).ValueType = GetType(Decimal)
                If Not IsNull(.CurrentRow.Cells(e.ColumnIndex).Value) Then
                    .CurrentRow.Cells(e.ColumnIndex).Value = FormatNumber(.CurrentRow.Cells(e.ColumnIndex).Value, dp, , , TriState.False)
                End If
            End If
            Dim c_name As String = .Columns(e.ColumnIndex).Name
            Select Case c_name
                Case "dgc_summary_start", "dgc_summary_finish"
                    For i As Integer = 0 To lb_multiple_job_list.Items.Count - 1
                        For r As Integer = 0 To .RowCount - 1
                            If lb_multiple_job_list.Items(i).ToString = .Rows(r).Cells("dgc_summary_prod_num").Value.ToString Then
                                .Rows(r).Cells(.CurrentCell.ColumnIndex).Value = .CurrentCell.Value.ToString.PadLeft(4, "0"c)
                            End If
                        Next
                    Next
                    If Not .CurrentCell.Value = Nothing Then
                        .CurrentCell.Value = .CurrentCell.Value.ToString.PadLeft(4, "0"c)
                        If .CurrentRow.Index = 0 And c_name = "dgc_summary_start" Then
                            Dim d As Date = FormatDateTime(tb_start_date.Text & " " & Strings.Left(.CurrentCell.Value, 2) & ":" & Strings.Right(.CurrentCell.Value, 2) & ":" & "00")
                            ' Dim d1 As Date = FormatDateTime(Today.AddDays(1) & " 00:12:00")
                            If d > Now.AddHours(4) Then
                                sql.delete("db_summary", "user_", "=", "'" & s_user & "' AND date_='" & tb_start_date.Text & "' AND machine='" & s_machine & "'")
                                If s_department = "filmsline" Then
                                    Dim dt As Date = tb_start_date.Text
                                    Dim last() As String = Split(s_shift_name, " ")
                                    Dim day As String = Format(dt.AddDays(-1), "ddd")
                                    Dim dr() As DataRow = dt_shifts.Select("name LIKE '" & day & "%' AND name LIKE '%" & last(1) & "%'")
                                    s_shift_name = UCase(dr(0)("name")) 'Replace(s_shift_name, UCase(Format(dt, "ddd")), UCase(Format(dt.AddDays(-1), "ddd")))
                                    'Dim day As String = Format(dt.AddDays(-1), "ddd")
                                    s_shift = dr(0)("id") 'sql.read("db_shifts", "plant", "=", s_plant & " AND dept='" & s_department & "' AND enabled=1 AND name='" & s_shift_name & "'", "id")

                                    tb_shift.Text = s_shift
                                End If
                                tb_start_date.Text = Today.AddDays(-1)
                            End If
                        End If
                    End If
                    If c_name = "dgc_summary_finish" Then
                        If .RowCount - 1 > e.RowIndex And Not s_department = "mounting" And lb_multiple_job_list.Items.Count = 0 Then
                            .Rows(e.RowIndex + 1).Cells("dgc_summary_start").Value = .Rows(e.RowIndex).Cells("dgc_summary_finish").Value
                        End If
                    End If
                Case "dgc_summary_setup"
                    If lb_multiple_job_list.Items.Count > 0 Then
                        Dim new_value As Decimal = FormatNumber(.CurrentCell.Value / lb_multiple_job_list.Items.Count, 2, , , TriState.False)
                        For i As Integer = 0 To lb_multiple_job_list.Items.Count - 1
                            For r As Integer = 0 To .RowCount - 1
                                If lb_multiple_job_list.Items(i).ToString = .Rows(r).Cells("dgc_summary_prod_num").Value.ToString Then
                                    .Rows(r).Cells(.CurrentCell.ColumnIndex).Value = new_value
                                End If
                            Next
                        Next
                    End If
            End Select

            If lb_multiple_job_list.Items.Count = 0 Then
                calc_times(.CurrentRow.Index)
            Else
                calc_times_multiple(True)
            End If
        End With
        '  calc_times(e.RowIndex)

        export_summary()

        s_loading = False
    End Sub

    Private Sub dgv_summary_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgv_summary.EditingControlShowing
        With dgv_summary

            'Select Case .Columns(.CurrentCell.ColumnIndex).Name
            ' Case "dgc_summary_start", "dgc_summary_finish", "dgc_summary_cyl_in", "dgc_summary_ink_in"
            Dim dp As Integer = 0
            If .Columns(.CurrentCell.ColumnIndex).DefaultCellStyle.Format.Contains("N") Then
                dp = Strings.Right(.Columns(.CurrentCell.ColumnIndex).DefaultCellStyle.Format, 1)
            End If
            Select Case True
                Case dp = 0
                    allow_decimal = False
                Case Else
                    allow_decimal = True
            End Select
        End With
        RemoveHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf CellKeyPress
        AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf CellKeyPress
    End Sub

    Private Sub dgv_summary_KeyDown(sender As Object, e As KeyEventArgs) Handles dgv_summary.KeyDown
        With dgv_summary
            If .CurrentRow.Cells("dgc_summary_mounter").Selected Then
                .EndEdit()
            End If
            If e.KeyCode = Keys.Back Or e.KeyCode = Keys.Delete Then
                If Not .CurrentCell.ReadOnly Then
                    If lb_multiple_job_list.Items.Count = 0 Then
                        .CurrentCell.Value = Nothing
                        If .CurrentCell.ColumnIndex = 2 Then
                            If .RowCount - 1 > .CurrentCell.RowIndex Then
                                .Rows(.CurrentCell.RowIndex + 1).Cells(1).Value = Nothing
                            End If
                        End If
                    Else
                        For i As Integer = 0 To lb_multiple_job_list.Items.Count - 1
                            For r As Integer = 0 To .RowCount - 1
                                If lb_multiple_job_list.Items(i).ToString = .Rows(r).Cells(0).Value.ToString Then
                                    .Rows(r).Cells(.CurrentCell.ColumnIndex).Value = Nothing
                                    calc_times(r)

                                End If
                            Next
                        Next
                    End If

                    export_summary()
                End If
            End If
        End With
    End Sub

    Private Sub dgv_reject_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_reject.CellClick
        If e.RowIndex = -1 Then Exit Sub
        With dgv_reject

            Select Case Replace(.Columns(e.ColumnIndex).Name, "dgc_reject_", Nothing)
                Case "code", "reason"
                    With frm_reject
                        .StartPosition = FormStartPosition.CenterParent
                        .ShowDialog()
                    End With
            End Select
        End With
    End Sub

    Private Sub dgv_reject_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_reject.CellEndEdit

        export_reject()
    End Sub

    Private Sub dgv_reject_KeyDown(sender As Object, e As KeyEventArgs) Handles dgv_reject.KeyDown
        With dgv_reject
            Select Case .CurrentCell.ColumnIndex
                Case 1, 4, 5, 7, 8
                    If e.KeyCode = Keys.Back Or e.KeyCode = Keys.Delete Then
                        .CurrentCell.Value = Nothing
                    End If
                    export_reject()
            End Select

        End With

    End Sub


    Private Sub dgv_production_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_production.CellEndEdit

        If s_loading Then Exit Sub
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim rn As Integer, uom, unit As String, quant, kg, km As Decimal
        Dim edit As Boolean = True
        If get_finish_time() < Now Then
            MsgBox("You are editing the report after the finish time that is entered" & vbCr & _
                   "This may be because the previous user has not logged off")
        End If
        unit = Nothing
        s_report_export_all = False


        If dgv_summary.RowCount = 0 Then
            add_job(tb_prod_num.Text)
            export_summary()
        End If

        With dgv
            If .CurrentCell.Value = "." Then .CurrentCell.Value = Nothing
            If .CurrentCell.Value Is Nothing Then Exit Sub
            Dim dp As Integer = 0
            Dim s As String = Replace(.Columns(.CurrentCell.ColumnIndex).DefaultCellStyle.Format, "N", Nothing)
            If IsNumeric(s) And Not .CurrentCell.Value = Nothing Then
                dp = s
                .CurrentCell.Value = FormatNumber(.CurrentCell.Value, dp, , , False)
            End If

            Dim i_info As item_info = get_item_info(.CurrentRow.Index, True)

            If IsDBNull(.CurrentRow.Cells(.CurrentCell.ColumnIndex).Value) Then Exit Sub

            For r As Integer = 0 To .RowCount - 1
                Dim n_out As Integer = 0
                If IsNumeric(.Rows(r).Cells("dgc_production_num_out").Value) Then
                    n_out = .Rows(r).Cells("dgc_production_num_out").Value
                    If n_out > rn Then rn = n_out
                End If
            Next
            rn = rn + 1

            If IsNull(.CurrentRow.Cells("dgc_production_num_out").Value) And Not .Columns(e.ColumnIndex).Name = "dgc_production_comments" Then
                .CurrentRow.Cells("dgc_production_num_out").Value = "-"
            End If

            If e.ColumnIndex = 4 Then
                unit = "KM"
            ElseIf e.ColumnIndex = 5 Then
                unit = "KG"
            End If
            If Not unit = Nothing Then
                For r As Integer = 0 To .RowCount - 1
                    If Not r = .CurrentRow.Index Then
                        If Not IsNull(.Rows(r).Cells("dgc_production_kg_out").Value) And Not IsNull(.Rows(r).Cells("dgc_production_km_out").Value) Then
                            kg = kg + .Rows(r).Cells("dgc_production_kg_out").Value
                            km = km + .Rows(r).Cells("dgc_production_km_out").Value
                        End If
                    End If
                Next

                If Not kg = 0 And Not km = 0 And Not .CurrentCell.Value = Nothing Then
                    If unit = "KG" Then
                        quant = .CurrentRow.Cells("dgc_production_km_out").Value * (kg / km)
                    Else
                        quant = .CurrentRow.Cells("dgc_production_kg_out").Value * (km / kg)
                    End If
                End If

            End If

            If IsNull(.CurrentRow.Cells("dgc_production_kg_out").Value) And IsNull(.CurrentRow.Cells("dgc_production_km_out").Value) Then
                If Not IsNull(.CurrentRow.Cells("dgc_production_num_out").Value) Then
                    If Not .CurrentRow.Cells("dgc_production_num_out").Value.ToString = "-" Then
                        If .CurrentRow.Cells("dgc_production_num_out").Value >= rn - 1 Then
                            .CurrentRow.Cells("dgc_production_num_out").Value = Nothing

                        End If
                    End If
                End If
            End If


            If Not .Columns(e.ColumnIndex).Name = "dgc_production_comments" Then
                If IsNull(.CurrentRow.Cells("dgc_production_user_").Value) Then
                    .CurrentRow.Cells("dgc_production_date_").Value = tb_start_date.Text
                    .CurrentRow.Cells("dgc_production_shift").Value = s_shift
                    .CurrentRow.Cells("dgc_production_user_").Value = s_user
                    .CurrentRow.Cells("dgc_production_machine").Value = cbo_machine.SelectedItem("machine")
                End If
            End If

            Select Case .Columns(e.ColumnIndex).Name
                Case "dgc_production_kg_out", "dgc_production_km_out"
                    If IsNull(.Rows(i_info.row_start).Cells("dgc_production_kg_out").Value) And _
                        .Columns(e.ColumnIndex).Name = "dgc_production_kg_out" Then

                        .CurrentCell.Value = Nothing
                        MsgBox("You cant enter the first item here. Use the first row for the roll.")
                        Exit Sub
                    End If
                    If .Columns(e.ColumnIndex).Name = "dgc_production_kg_out" Then report_kg_out(sender)
                    If .Columns(e.ColumnIndex).Name = "dgc_production_km_out" Then report_km_out(sender, e)

                    'mod_frm_main.report_remaining_estimate(e)

                    uom = s_selected_job(0)("req_uom_1")

                    'If IsNull(.CurrentRow.Cells("dgc_production_num_out").Value) Then
                    '    str = Nothing
                    'Else
                    '    str = .CurrentRow.Cells("dgc_production_num_out").Value
                    'End If
                    'If str = "-" Then str = Nothing
                    'If str = Nothing And e.ColumnIndex > 3 And e.ColumnIndex < 8 Then
                    If .CurrentRow.Cells("dgc_production_num_out").Value.ToString = "-" And e.ColumnIndex > 3 And e.ColumnIndex < 8 And Not .CurrentCell.Value = Nothing Then
                        If Not .CurrentCell.Value = 0 Then
                            .CurrentRow.Cells("dgc_production_num_out").Value = rn
                        End If
                        For r As Integer = 0 To .CurrentRow.Index
                            If IsNull(.Rows(r).Cells("dgc_production_num_out").Value) Then
                                .Rows(r).Cells("dgc_production_num_out").Value = "-"
                            End If
                        Next

                    End If
                    item_entered(uom, quant)

                Case "dgc_production_rts"

                    If i_info.kg_in < .CurrentCell.Value Then
                        Dim resp As MsgBoxResult = MsgBox("You are trying to RTS more than what was issued on the item (" & _
                                                          FormatNumber(i_info.kg_in, 1, , , TriState.False) & " KG)" & vbCr & _
                                                          "Are you sure you want to RTS this amount (" & _
                                                          FormatNumber(.CurrentCell.Value, 1, , , TriState.False) & " KG)?", _
                                                          MsgBoxStyle.Critical + MsgBoxStyle.YesNo)
                        If resp = MsgBoxResult.No Then
                            .CurrentCell.Value = Nothing
                            Exit Sub
                        End If
                    End If
                    If i_info.total_rolls = 1 Then
                        For r As Integer = i_info.row_start To i_info.row_end
                            If Not IsNull(.Rows(r).Cells("dgc_production_rts").Value) And Not r = .CurrentRow.Index Then
                                .CurrentCell.Value = Nothing
                                MsgBox("You can't RTS the same roll from a different line, change the original RTS")
                                Exit Sub
                            End If
                        Next
                    ElseIf i_info.total_rolls = i_info.rts_count Then
                        .CurrentCell.Value = Nothing
                        MsgBox("You can't RTS any more items, there is nothing to RTS from.")
                        Exit Sub
                    End If

                    Dim last_rts As Decimal = 0

                    If IsDBNull(sql.read("db_rts", "prod_num", "=", tb_prod_num.Text & " AND item_id='" & i_info.item_id & _
                                         "' AND num_in='" & i_info.num_in & "'", "quantity_symple")) Then
                        last_rts = 0
                    Else
                        last_rts = sql.read("db_rts", "prod_num", "=", tb_prod_num.Text & " AND item_id='" & i_info.item_id & _
                                            "' AND num_in='" & i_info.num_in & "'", "quantity_symple")
                    End If

                    If .CurrentRow.Cells(e.ColumnIndex).Value < last_rts Then
                        MsgBox("The last RTS entered on this roll is MORE than you have entered, you will need to reverse the old RTS of " & last_rts, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "RTS Reversal needed!")
                    End If

                    s_loading = True
                    With dgv_rts
                        If .RowCount > 0 Then
                            For r As Integer = 0 To .RowCount - 1
                                If .Rows(r).Cells("dgc_rts_num_in").Value = i_info.num_in And .Rows(r).Cells("dgc_rts_item_id").Value = i_info.item_id Then
                                    If dgv_production.CurrentRow.Cells(e.ColumnIndex).Value = .Rows(r).Cells("dgc_rts_quantity").Value Then
                                        edit = False
                                    Else
                                        .Rows.RemoveAt(r)
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                        If edit Then
                            .Rows.Add()
                            add_rts(i_info.prod_num, i_info.mat_num, i_info.num_in, i_info.item_id, dgv_production.CurrentRow.Cells(e.ColumnIndex).Value)
                            dgv_production.CurrentRow.Cells("dgc_production_rts_info").Value = _
                                 i_info.prod_num & "," & i_info.item_id & "," & i_info.num_in & "," & i_info.mat_num
                        End If
                        s_loading = False
                        export_rts()
                    End With
                Case "dgc_production_reject"

                    Dim response As MsgBoxResult
                    response = MsgBox("Send to Reject Slip?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Reject")
                    If response = MsgBoxResult.Yes Then
                        s_loading = True
                        With dgv_reject
                            If .RowCount > 0 Then
                                For r As Integer = 0 To .RowCount - 1
                                    If Not IsNull(.Rows(r).Cells("dgc_reject_item_no").Value) And Not IsNull(.Rows(r).Cells("dgc_reject_num_in").Value) Then
                                        If .Rows(r).Cells("dgc_reject_item_no").Value = i_info.item_id And .Rows(r).Cells("dgc_reject_num_in").Value = i_info.num_in Then
                                            If dgv_production.CurrentRow.Cells(e.ColumnIndex).Value = .Rows(r).Cells("dgc_reject_quantity").Value Then
                                                edit = False
                                            Else
                                                .Rows.RemoveAt(r)
                                            End If
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                            If edit Then
                                .Rows.Add(i_info.prod_num.ToString)
                                add_reject(i_info.mat_num, i_info.item_id, i_info.num_in, dgv_production.CurrentRow.Cells(e.ColumnIndex).Value)
                                dgv_production.CurrentRow.Cells("dgc_production_reject_info").Value = _
                                     i_info.prod_num & "," & i_info.item_id & "," & i_info.num_in & "," & i_info.mat_num
                            End If
                        End With
                        s_loading = False
                        export_reject()
                    End If
                Case "dgc_production_width_"
                    If IsNumeric(.CurrentCell.Value) And s_extruder Then
                        Dim l, w As Integer, d, g As Decimal
                        d = get_item_desity(s_selected_job(0)("mat_desc_semi"))
                        l = .CurrentRow.Cells("dgc_production_km_out").Value
                        w = .CurrentRow.Cells("dgc_production_width_").Value
                        kg = .CurrentRow.Cells("dgc_production_kg_out").Value
                        g = FormatNumber(kg / d / l / w * 1000000)
                        .CurrentRow.Cells("dgc_production_gauge").Value = g
                    End If
            End Select
            If i_info.remaining <= 0 Then
                For r As Integer = i_info.row_start To i_info.row_end
                    If can_delete(r) Then
                        .Rows(r).Visible = False
                        s_report_export_all = True
                    End If
                Next

            End If
            sql.update("db_schedule", "status", 2, "prod_num", "=", tb_prod_num.Text)

            If s_report_export_all Then
                export_report()
            Else
                For i As Integer = i_info.row_start To i_info.row_end
                    export_report(i)
                Next
            End If
            ' End If
            calc_totals()
            .PerformLayout()

            s_loading = False
        End With


        ' get_production_info(e.RowIndex)

    End Sub

    Private Sub dgv_production_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv_production.CellMouseClick
        If e.ColumnIndex > -1 And e.RowIndex > -1 Then
            With dgv_production
                If e.Button = Windows.Forms.MouseButtons.Right Then

                    '.ClearSelection()
                    s_loading = False
                    Dim i_info As item_info = get_item_info(e.RowIndex, True)
                    s_loading = True
                    If i_info.row_end = .RowCount - 1 And i_info.used_kg = 0 Then
                        tsm_production_delete_row.Enabled = True
                        If s_department = "bagplant" Then
                            If Not can_delete(e.RowIndex, True) Then
                                tsm_production_delete_row.Enabled = False
                            End If
                        End If
                    ElseIf can_delete(e.RowIndex) Then
                        tsm_production_delete_row.Enabled = True
                    Else
                        tsm_production_delete_row.Enabled = False
                    End If
                    Dim cell As DataGridViewCell = .Rows(e.RowIndex).Cells(e.ColumnIndex)
                    .CurrentCell = cell
                    .CurrentCell.Selected = True
                    If e.RowIndex = 0 Then
                        tsm_production_insert_row_above.Enabled = False
                    Else
                        tsm_production_insert_row_above.Enabled = True
                    End If

                    If IsNull(.Rows(e.RowIndex).Cells("dgc_production_user_").Value) Then
                        If s_department = "printing" Then
                            If .Rows(e.RowIndex).Cells("dgc_production_num_in").Value.ToString = "-" And _
                            .Rows(e.RowIndex).Cells("dgc_production_item_id").Value.ToString = "-" And _
                            .Rows(e.RowIndex).Cells("dgc_production_kg_in").Value.ToString = "-" Then
                                tsm_production_insert_row_above.Enabled = True
                                If .RowCount - 1 = e.RowIndex Then
                                    tsm_production_insert_row_below.Enabled = True
                                Else
                                    If IsNull(.Rows(e.RowIndex + 1).Cells("dgc_production_user_").Value) Then
                                        tsm_production_insert_row_below.Enabled = True
                                    Else
                                        tsm_production_insert_row_below.Enabled = False
                                    End If
                                End If
                            Else
                                If .RowCount - 1 > .CurrentRow.Index Then
                                    For r As Integer = .CurrentRow.Index + 1 To .RowCount - 1
                                        If IsNumeric(.Rows(r).Cells("dgc_production_num_in").Value) Then
                                            tsm_production_delete_row.Enabled = False
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                        Else
                            tsm_production_insert_row_above.Enabled = True
                            If .RowCount - 1 = e.RowIndex Then
                                tsm_production_insert_row_below.Enabled = True
                            Else
                                If IsNull(.Rows(e.RowIndex + 1).Cells("dgc_production_user_").Value) Then
                                    tsm_production_insert_row_below.Enabled = True
                                Else
                                    tsm_production_insert_row_below.Enabled = False
                                End If
                            End If
                        End If

                    Else
                        tsm_production_insert_row_above.Enabled = False
                        If .RowCount - 1 = e.RowIndex Then
                            tsm_production_insert_row_below.Enabled = True
                        Else
                            If IsNull(.Rows(e.RowIndex + 1).Cells("dgc_production_user_").Value) Then
                                tsm_production_insert_row_below.Enabled = True
                            Else
                                tsm_production_insert_row_below.Enabled = False
                            End If
                        End If
                    End If
                    tsm_production_reprint.Enabled = False
                    Select Case get_material_info.material_type
                        Case material_type.barrier
                            If Not IsNull(.Rows(e.RowIndex).Cells("dgc_production_kg_out").Value) And IsNumeric(.Rows(e.RowIndex).Cells("dgc_production_kg_out").Value) Then
                                If .Rows(e.RowIndex).Cells("dgc_production_kg_out").Value > 0 Then
                                    tsm_production_reprint.Enabled = s_symple_access
                                End If
                            End If
                        Case Else
                            If Not IsNull(.Rows(e.RowIndex).Cells("dgc_production_kg_out").Value) And _
                                IsNumeric(.Rows(e.RowIndex).Cells("dgc_production_kg_out").Value) And _
                                Not IsNull(.Rows(e.RowIndex).Cells("dgc_production_km_out").Value) And _
                                IsNumeric(.Rows(e.RowIndex).Cells("dgc_production_km_out").Value) Then

                                If .Rows(e.RowIndex).Cells("dgc_production_kg_out").Value > 0 And .Rows(e.RowIndex).Cells("dgc_production_km_out").Value > 0 Then
                                    tsm_production_reprint.Enabled = True
                                End If
                            End If
                    End Select
                    If .RowCount - 1 = e.RowIndex And s_extruder Then
                        tsm_production_delete_row.Enabled = True
                    End If

                    tsm_production_clear.Enabled = True
                    For r As Integer = 0 To .RowCount - 1
                        If Not IsNull(.Rows(r).Cells("dgc_production_num_out").Value) Then
                            tsm_production_clear.Enabled = False
                            Exit For
                        End If
                    Next
    
                    cms_production.Show(Cursor.Position)
                    s_loading = False
                ElseIf e.Button = Windows.Forms.MouseButtons.Left And s_department = "bagplant" Then
                    If IsNull(.Rows(e.RowIndex).Cells("dgc_production_num_out").Value) Then
                        frm_bag_checks.nud_carton_num.Value = 1
                    Else
                        frm_bag_checks.nud_carton_num.Value = .Rows(e.RowIndex).Cells("dgc_production_num_out").Value
                    End If
                End If
            End With
        End If
    End Sub

    Private Sub dgv_production_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgv_production.EditingControlShowing
        Dim dgv As DataGridView = CType(sender, DataGridView)
        allow_decimal = True

        With dgv
            Select Case .Columns(.CurrentCell.ColumnIndex).Name
                Case "dgc_production_comments", "dgc_production_barcode"
                    RemoveHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf CellKeyPress
                Case Else
                    RemoveHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf CellKeyPress
                    AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf CellKeyPress
            End Select
        End With
    End Sub

    Private Sub dgv_production_KeyDown(sender As Object, e As KeyEventArgs) Handles dgv_production.KeyDown

        If e.KeyCode = Keys.Back Or e.KeyCode = Keys.Delete Then
            Dim rows As String = Nothing
            Dim dgv As DataGridView = CType(sender, DataGridView)
            Dim can_remove As Boolean = True
            Dim str() As String
            With dgv
                Dim i_info As item_info = get_item_info(.CurrentRow.Index, True)
                Dim row As Integer = .CurrentRow.Index
                Dim v = .CurrentCell.Value
                If .CurrentCell.ReadOnly Then Exit Sub
                If .RowCount = 0 Then Exit Sub
                .CurrentCell.Value = Nothing
                ' If nud_items_out.Value > 1 And .CurrentCell.ColumnIndex = 4 Then
                Select Case .Columns(.CurrentCell.ColumnIndex).Name
                    Case "dgc_production_kg_out", "dgc_production_km_out"
                        'find last roll to add value to
                        For r As Integer = .RowCount - 1 To .CurrentRow.Index Step -1
                            If r > .CurrentRow.Index And Not can_delete(r, True) And Not .CurrentRow.Index = 0 Then
                                can_remove = False
                                Exit For
                            ElseIf .CurrentRow.Index = 0 And .RowCount > 1 And .Columns(.CurrentCell.ColumnIndex).Name = "dgc_production_kg_out" Then
                                If Not can_delete(.CurrentRow.Index, True) Then
                                    can_remove = False
                                End If
                            End If
                        Next
                        If Not can_remove Then
                            MsgBox("You can't remove this entry because there is other information claimed after this." & vbCr & _
                                   "It could be a comment, RTS or Reject." & vbCr & "Please remove the conflicting items first.")

                        End If
                        If can_remove Then
                            i_info = get_item_info(.CurrentRow.Index, True)

                            If Not i_info.used_kg = 0 Then
                                For i As Integer = i_info.row_end To i_info.row_start Step -1

                                    If i < .CurrentRow.Index And .CurrentRow.Visible And IsNumeric(.Rows(i).Cells("dgc_production_kg_in").Value) And _
                                        IsNumeric(.CurrentRow.Cells("dgc_production_kg_in").Value) Then

                                        .Rows(i).Cells("dgc_production_kg_in").Value = _
                                            CDec(.Rows(i).Cells("dgc_production_kg_in").Value + CDec(.CurrentRow.Cells("dgc_production_kg_in").Value))
                                        rows = add_to_string(rows, i, ",", False)

                                        .CurrentRow.Cells("dgc_production_kg_in").Value = "-"
                                        rows = add_to_string(rows, .CurrentRow.Index, ",", False)
                                        row = -1
                                        Exit For
                                    End If
                                Next
                            End If
                        End If


                    Case "dgc_production_rts"
                        With dgv_rts

                            If .RowCount > 0 Then
                                If Not IsNull(dgv_production.CurrentRow.Cells("dgc_production_rts_info").Value) Then
                                    str = Split(dgv_production.CurrentRow.Cells("dgc_production_rts_info").Value, ",")
                                    If str.Length = 4 Then
                                        For r As Integer = 0 To .RowCount - 1
                                            If .Rows(r).Cells("dgc_rts_item_id").Value = str(1) And .Rows(r).Cells("dgc_rts_num_in").Value = str(2) Then
                                                dgv_production.CurrentRow.Cells("dgc_production_rts_info").Value = Nothing
                                                rows = add_to_string(rows, dgv_production.CurrentRow.Index, ",", False)

                                                .Rows.RemoveAt(r)
                                                calc_rts()
                                                export_rts()
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        End With

                    Case "dgc_production_reject"

                        With dgv_reject

                            If .RowCount > 0 Then
                                If Not IsNull(dgv_production.CurrentRow.Cells("dgc_production_reject_info").Value) Then
                                    str = Split(dgv_production.CurrentRow.Cells("dgc_production_reject_info").Value, ",")
                                    If str.Length = 3 Then
                                        For r As Integer = 0 To .RowCount - 1
                                            If .Rows(r).Cells(1).Value = str(1) Then
                                                dgv_production.CurrentRow.Cells("dgc_production_reject_info").Value = Nothing
                                                rows = add_to_string(rows, dgv_production.CurrentRow.Index, ",", False)
                                                .Rows.RemoveAt(r)
                                                export_reject()
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        End With
                End Select

                If can_remove Then
                    .CurrentCell.Value = Nothing
                    rows = add_to_string(rows, .CurrentRow.Index, ",", False)

                    If IsNull(.CurrentRow.Cells("dgc_production_kg_out").Value) And _
                        IsNull(.CurrentRow.Cells("dgc_production_km_out").Value) And Not _
                        IsNull(.CurrentRow.Cells("dgc_production_num_out").Value) Then

                        For r As Integer = .RowCount - 1 To .CurrentRow.Index Step -1
                            If IsNumeric(.Rows(r).Cells("dgc_production_num_out").Value) Then
                                If r = .CurrentRow.Index Then
                                    .CurrentRow.Cells("dgc_production_num_out").Value = Nothing
                                    rows = add_to_string(rows, .CurrentRow.Index, ",", False)
                                    Exit For
                                End If
                            ElseIf r = .CurrentRow.Index Then
                                If .Rows(r).Cells("dgc_production_num_out").Value.ToString = "-" Then
                                    .CurrentRow.Cells("dgc_production_num_out").Value = Nothing
                                    rows = add_to_string(rows, .CurrentRow.Index, ",", False)
                                End If
                            End If
                        Next
                    End If

                    If Not has_info(.CurrentRow.Index) Then
                        .CurrentRow.Cells("dgc_production_date_").Value = Nothing
                        .CurrentRow.Cells("dgc_production_shift").Value = Nothing
                        .CurrentRow.Cells("dgc_production_user_").Value = Nothing
                        rows = add_to_string(rows, .CurrentRow.Index, ",", False)
                    End If
                    'add delete numout if all nothing

                    dt_produced = dgv_to_dt(dgv_production) 'sql.get_table("db_produced", "prod_num", "=", tb_prod_num.Text)
                    export_report_rows(rows)
                    calc_times(get_prod_num_row(tb_prod_num.Text))
                    calc_totals()

                Else
                    .CurrentCell.Value = v
                    'other msgbox
                End If
                get_production_info(.CurrentRow.Index)
            End With
        End If
    End Sub



#End Region

    Sub set_item_spaces()

        'Dim count As Integer = sql.count("db_produced", "prod_num=" & tb_prod_num.Text & " AND NOT item_id", "=", "'0'")

        With dgv_production
            If .RowCount = 0 Then Exit Sub
            .CurrentCell = .CurrentRow.Cells("dgc_production_comments")
            'Dim r As Integer = -1
            'Dim a As Integer = 0
            'If current = 1 Then
            '    For r = 1 To nud_items_out.Value - 1
            '        .Rows.Add()
            '        .Rows(r).Cells("dgc_production_num_in").Value = "-"
            '        .Rows(r).Cells("dgc_production_item_id").Value = "-"
            '        .Rows(r).Cells("dgc_production_kg_in").Value = "-"
            '        .Rows(r).Cells("dgc_production_prod_num").Value = .Rows(r - 1).Cells("dgc_production_prod_num").Value
            '        .Rows(r).Cells("dgc_production_assistant").Value = .Rows(r - 1).Cells("dgc_production_assistant").Value
            '    Next

            ' If current < total Then
            If .RowCount = 0 Then Exit Sub
            Dim last_num As Integer = -1
            Dim x_f As Boolean = False
            For r = .RowCount - 1 To 0 Step -1
                Dim added As Boolean = False
                s_loading = False
                Dim i_info As item_info = get_item_info(r, True)
                Dim c_count As Integer = i_info.row_end - i_info.row_start + 1
                Dim n_count As Integer = nud_items_out.Value
                s_loading = True
                'If Not .Rows(r).Cells("dgc_production_num_in").Value.ToString = "-" Then
                If i_info.used_kg = 0 Then
                    If c_count < n_count Then
                        If i_info.row_end + 1 = .RowCount Then
                            .Rows.Insert(i_info.row_end + 1, 1)
                            added = True
                        ElseIf IsNull(.Rows(i_info.row_end + 1).Cells("dgc_production_num_out").Value) Then
                            .Rows.Insert(i_info.row_end + 1, 1)
                            added = True
                        End If
                    ElseIf c_count > n_count And can_delete(i_info.row_end) Then
                        .Rows.RemoveAt(i_info.row_end)
                        last_num = i_info.num_in
                    End If
                ElseIf i_info.num_in + 1 = last_num Or i_info.row_end = .RowCount - 1 Then
                    If c_count < n_count Then
                        If i_info.row_end + 1 = .RowCount Then
                            .Rows.Insert(i_info.row_end + 1, 1)
                            added = True
                        ElseIf IsNull(.Rows(i_info.row_end + 1).Cells("dgc_production_num_out").Value) Then
                            .Rows.Insert(i_info.row_end + 1, 1)
                            added = True
                        End If
                    ElseIf c_count > n_count And can_delete(i_info.row_end) Then
                        .Rows.RemoveAt(i_info.row_end)
                    End If
                    x_f = True

                End If
                If added Then
                    last_num = i_info.num_in
                    .Rows(i_info.row_end + 1).Cells("dgc_production_num_in").Value = "-"
                    .Rows(i_info.row_end + 1).Cells("dgc_production_item_id").Value = "-"
                    .Rows(i_info.row_end + 1).Cells("dgc_production_kg_in").Value = "-"
                    .Rows(i_info.row_end + 1).Cells("dgc_production_prod_num").Value = .Rows(i_info.row_start).Cells("dgc_production_prod_num").Value
                    .Rows(i_info.row_end + 1).Cells("dgc_production_assistant").Value = .Rows(i_info.row_start).Cells("dgc_production_assistant").Value
                End If
                If x_f Then Exit For
            Next
            .Focus()
        End With
        export_report()
        s_loading = False

    End Sub


    Private Sub nud_items_out_ValueChanged(sender As Object, e As EventArgs) Handles nud_items_out.ValueChanged
        If s_loading Then Exit Sub
        '       sql.delete("db_produced", "prod_num", "NULL", " IS ")
        sql.update("db_schedule", "items_out", nud_items_out.Value, "prod_num", "=", tb_prod_num.Text)

        s_selected_job(0)("items_out") = nud_items_out.Value
        ' dt_schedule = sql.get_table("db_schedule", "plant", "=", s_plant)

        Select Case s_department
            Case "printing", "filmsline"
                set_item_spaces()
        End Select

        s_loading = False

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        cb_main_loading.Checked = s_loading
        If System.Windows.Forms.Form.MouseButtons = Windows.Forms.MouseButtons.Left _
            Or System.Windows.Forms.Form.MouseButtons = Windows.Forms.MouseButtons.Right _
            Or System.Windows.Forms.Form.MouseButtons = Windows.Forms.MouseButtons.Middle Then
            Exit Sub
        End If
        If Not s_start_up Then
            Timer1.Enabled = False
            start_up()
        End If
        If s_department = "planning" And cb_scheduling_monitor_clipboard.Checked Then
            check_clipboard()
        End If

        Dim tWnd As IntPtr = get_handle("WFT - Port Reader")
        If Not tWnd = 0 And s_port_reader = Nothing Then
            With port_reader
                Dim p() As Process
                p = Process.GetProcessesByName("PortReader")
                If p.Length > 0 Then
                    port_reader = p(0)
                    s_port_reader = Replace(port_reader.MainModule.FileName, "PortReader.exe", Nothing)
                    For i As Integer = 1 To 19
                        Dim device As String = ReadIni(s_port_reader & "PortReader.ini", "COM" & i, "DeviceName", Nothing)
                        If Not device = Nothing Then
                            Dim clip As String = ReadIni(s_port_reader & "PortReader.ini", "COM" & i, "OutputClipboard", Nothing)
                            Dim reg As String = ReadIni(s_port_reader & "PortReader.ini", "COM" & i, "OutputRegApp", Nothing)
                            Dim kill As Boolean = False

                            If reg = Nothing Then
                                kill = True
                                WriteIni(s_port_reader & "PortReader.ini", "COM" & i, "OutputRegApp", "PortReader")
                                WriteIni(s_port_reader & "PortReader.ini", "COM" & i, "OutputRegSection", "Data")
                                WriteIni(s_port_reader & "PortReader.ini", "COM" & i, "OutputRegKey", "Weight")
                            End If

                            If LCase(clip) = "yes" Then
                                kill = True
                                WriteIni(s_port_reader & "PortReader.ini", "COM" & i, "OutputClipboard", "no")
                            End If

                            If kill Then
                                p(0).Kill()
                                Process.Start(s_port_reader & "PortReader.exe")
                                Exit For
                            End If
                        End If

                    Next
                End If

            End With
        End If

        If tWnd = 0 Then
            gb_scales.Visible = False
        Else
            gb_scales.Visible = True
            Dim rkey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\VB and VBA Program Settings\PortReader\Data", False)
            Dim w As Decimal
            Dim s As Decimal
            If Not rkey Is Nothing Then
                s = Replace(rkey.GetValue("Weight", 0), " ", Nothing) ' My.Computer.Registry.GetValue("Hkey_Current_User\Software\VBandVBAProgramSettings\PortReader\Data\Weight", "Weight", Nothing)
            Else
                s = 0
            End If

   
            If IsNumeric(tb_scale_tare.Text) Then w = tb_scale_tare.Text
            If IsNumeric(s) Then
                tb_scale_weight.Text = s
                tb_scale_tared.Text = FormatNumber(s - w, 1, , , TriState.False)
                b_main_use_weight.Enabled = True
                b_main_tare.Enabled = True
            Else
                tb_scale_weight.Text = "0.00"
                tb_scale_tared.Text = "0.00"
                b_main_use_weight.Enabled = False
                b_main_tare.Enabled = False
 
            End If
        End If

        Dim ts As TimeSpan
        For i As Integer = 0 To 1
            With w_sessions(i)
                Try
                    If .process.Responding Then
                        b_main_logoff.Enabled = True
                        If .user = Nothing Then
                            acsis.get_logon_details(i)
                        End If
                        If s_version = Nothing Then
                            s_version = FileVersionInfo.GetVersionInfo(.process.MainModule.FileName).FileVersion
                        End If
                    End If
                Catch ex As Exception
                    'symple not open 
                    If i = 0 Then b_main_logoff.Enabled = False
                End Try

            End With
        Next

        If Not s_loading And frm_status.Visible Then frm_status.Close()
        If frm_settings_admin.Visible Or frm_logon.Visible Then Exit Sub
        If frm_startup.Visible Then Exit Sub
        'If startup Then start_up()
        'get restore info
        If frm_settings_admin.Visible Then Exit Sub

        'check for updated files and export to sql

        If Not sw_schedules.IsRunning Then
            check_for_update()
            sw_schedules.Start()
        Else
            If sw_schedules.Elapsed.Minutes > 30 Then
                sw_schedules.Reset()
                'display_job()
                check_for_update()
                s_loading = True
                dt_schedules = sql.get_table("db_schedule", "plant", "=", s_plant)
                dt_uom = sql.get_table_full("db_uom", True)
                If s_department = "bagplant" Then

                    Dim file As String = get_file_path("burst_data")
                    If System.IO.File.Exists(file) Then
                        If Not System.IO.File.Exists(Environ("USERPROFILE") & "\specification master sheet with Aussie.xls") Or _
                            Not System.IO.File.GetLastWriteTime(file) = _
                            System.IO.File.GetLastWriteTime(Environ("USERPROFILE") & "\specification master sheet with Aussie.xls") Then
                            System.IO.File.Copy(file, _
                                                Environ("USERPROFILE") & "\specification master sheet with Aussie.xls", True)
                        End If
                    End If
                End If
                s_loading = False
            End If
        End If

        If sw_tables.IsRunning Then
            ts = sw_tables.Elapsed
            If ts.Minutes >= 30 Then
                WriteIni(Environ("USERPROFILE") & "\settings.ini", "Export", "Time", Now)
                sql.export_tables()
            End If
        Else
            Dim dt As Date = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Export", "Time", Now)
            Dim m As Integer = DateDiff(DateInterval.Minute, dt, Now)
            If m > 30 Then
                WriteIni(Environ("USERPROFILE") & "\settings.ini", "Export", "Time", Now)
                sql.export_tables()
            End If
            sw_tables.Start()
        End If

        If sw_sql.IsRunning Then
            ts = sw_sql.Elapsed
            If ts.Minutes >= 10 Then
                sql.connect("lean", True)
                If s_sql_access Then
                    update_files()
                    cb_sql_lean.Text = "SQL"
                    sw_sql.Stop()
                Else
                    sw_sql.Reset()
                End If
            Else
                cb_sql_lean.Text = "SQL: " & 9 - ts.Minutes & ":" & (60 - ts.Seconds).ToString.PadLeft(2, "0")
            End If

        End If

        If Not frm_startup.Visible And s_shift = Nothing And Not frm_logon.Visible And Not s_preparing And Not s_department = "graphics" And Not s_department = "planning" Then

            With frm_startup
                .StartPosition = FormStartPosition.CenterParent
                .ShowDialog()
            End With

        End If
        'If cb_sql_symple.Checked And Not s_shift = Nothing Then
        If s_symple_access And Not s_shift = Nothing Then
            If System.IO.File.Exists(get_file_path("symple")) Then
                If Not acsis.running(0) Then acsis.get_controls(0, True, False, True)
                If acsis.running(0) And Not s_department = "graphics" Then
                    'acsis.get_controls(0, True, False, True)
                    If Not acsis.running(0) Then
                        If Not s_user = Nothing And Not s_pass = Nothing Then
                            acsis.logon()
                        ElseIf Not frm_logon.Visible And Not frm_startup.Visible Then
                            With frm_logon
                                .StartPosition = FormStartPosition.CenterParent
                                .ShowDialog()
                            End With
                        End If
                    End If


                ElseIf s_pass = Nothing Or s_user = Nothing Then
                    If Not s_department = "graphics" Then
                        If acsis.running(0) Then acsis.get_logon_details(0)
                        If s_user = Nothing Then
                            With frm_logon
                                .StartPosition = FormStartPosition.CenterParent
                                .ShowDialog()
                            End With
                        End If
                    End If

                Else
                    If acsis.running(0) Then
                        If frm_logon.Visible Then frm_logon.Close()
                        If s_label = Nothing Or s_printer = Nothing Then
                            If dgv_summary.RowCount > 0 Then
                                With frm_settings
                                    If Not .Visible Then
                                        .StartPosition = FormStartPosition.CenterParent
                                        .ShowDialog()
                                    End If

                                End With
                            End If

                        End If
                        cb_sql_symple.Text = "Sym-PLE"
                        If s_check_pop Then acsis.popup_check(True, 0)
                    Else
                        acsis.logon()
                    End If

                End If

            End If
        Else

            ts = sw_symple.Elapsed
            If ts.Minutes >= 10 Then
                sql.connect("symple", True)
                If s_symple_access Then
                    Dim logins As New DataTable
                    logins = sql.get_table("DP_Logins", "Login_Id", "=", "'" & s_user & "' and Password='" & s_pass & "' AND Plant=" & s_plant, "symple")
                    If IsNothing(logins) Then
                        MsgBox("Wrong login information, please check your entry!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "invalid login/pass!")
                        s_user = Nothing
                        s_pass = Nothing
                    Else
                        acsis.logon()
                    End If
                    cb_sql_symple.Text = "Sym-PLE"
                    sw_symple.Stop()
                Else
                    sw_symple.Reset()
                End If
            Else
                cb_sql_symple.Text = "Sym-PLE: " & 9 - ts.Minutes & ":" & (60 - ts.Seconds).ToString.PadLeft(2, "0")
            End If
            If Not acsis.running(0) Then b_main_logoff.Enabled = False
        End If


        If tb_prod_num.Text = Nothing And Not s_department = "graphics" Then
            For i As Integer = 1 To tc_main.TabPages.Count - 1
                If i = 6 Then
                    If s_department = "mounting" Then
                        tc_main.TabPages(i).Enabled = True
                    End If
                Else
                    tc_main.TabPages(i).Enabled = False
                End If
            Next

            s_label = "No Job Selected"
        Else
            For i As Integer = 1 To tc_main.TabPages.Count - 1
                tc_main.TabPages(i).Enabled = True
            Next
        End If
        cb_sql_symple.Checked = s_symple_access
        cb_sql_lean.Checked = s_sql_access
        If s_user_name = Nothing Then
            tssl_symple_info_login.Text = Nothing
        Else
            tssl_symple_info_login.Text = "Sym-PLE Logged in as: " & s_user_name
            If lb_personel.Items.Count = 0 And s_department = "mounting" Then
                lb_personel.Items.Add(s_user_name)
            End If
        End If
        tssl_symple_info_label.Text = "Printer: " & s_printer & " Label: " & s_label
        If lb_inks.Items.Count > 0 And s_department = "printing" Then
            l_inks_done.Visible = True
        Else
            l_inks_done.Visible = False
        End If
        If l_inks_done.Visible And l_prep_warning.Visible = False Then
            tc_main.SelectedIndex = 5
        End If
        If dgv_summary.Rows.Count = 0 Then
            l_summary_warning.Visible = True
        Else
            l_summary_warning.Visible = False
        End If
        Select Case tc_main.SelectedIndex
            Case 0
            Case 1
            Case 2
            Case 3
            Case 4 ' reject
                Dim en As Boolean = True
                With dgv_reject
                    If .RowCount > 0 Then
                        For r As Integer = 0 To .RowCount - 1
                            For c As Integer = 0 To 5
                                If IsNull(dgv_reject.Rows(r).Cells(c).Value) Then
                                    en = False
                                    Exit For
                                End If
                            Next
                            If en = False Then Exit For
                        Next
                    Else
                        en = False
                    End If

                End With

                b_reject_print.Enabled = en
        End Select
        With dgv_job_requirements
            For i As Integer = 0 To .RowCount - 1
                If Replace(.Rows(i).Cells(4).Value, "%", Nothing) >= 100 Then
                    If .Rows(i).DefaultCellStyle.BackColor = Color.White Or .Rows(i).DefaultCellStyle.BackColor = Color.LightGreen Then
                        .Rows(i).DefaultCellStyle.BackColor = Color.LightPink
                    Else
                        .Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
                    End If
                Else
                    .Rows(i).DefaultCellStyle.BackColor = Color.White
                End If
            Next
        End With
        If Not s_department = "graphics" And s_shift = Nothing And Not s_department = "planning" Then
            With frm_startup
                If .Visible = False Then
                    .StartPosition = FormStartPosition.CenterParent
                    .ShowDialog()
                End If
            End With
        End If
        b_summary_symple_times.Enabled = s_symple_access
        If s_department = "planning" Then

        End If
    End Sub

    Function get_prod_num_row(ByVal prod_num As Long) As Integer
        With dgv_summary
            For i As Integer = 0 To .RowCount - 1
                If .Rows(i).Cells(0).Value = prod_num Then
                    Return i
                    Exit Function
                End If
            Next
        End With
        Return 0
    End Function


    Function get_original_weight(ByVal item_id As String, ByVal num_in As Integer) As Decimal

        get_original_weight = 0

        With dgv_production
            For r As Integer = 0 To .RowCount - 1
                If .Rows(r).Cells("dgc_production_num_in").Value.ToString.Contains("&") Then
                    Dim str_num_in() As String = Split(.Rows(r).Cells("dgc_production_num_in").Value, " & ")
                    Dim str_item_id() As String = Split(.Rows(r).Cells("dgc_production_item_id").Value, " & ")
                    For i As Integer = 0 To UBound(str_num_in)
                        If str_num_in(i) = num_in And str_item_id(i) = item_id Then
                            Dim kg() As String = Split(.Rows(r).Cells("dgc_production_kg_in_orig").Value, ",")
                            Return kg(i)
                        End If
                    Next
                ElseIf .Rows(r).Cells("dgc_production_num_in").Value.ToString = num_in.ToString And .Rows(r).Cells("dgc_production_item_id").Value.ToString = item_id.ToString Then
                    Return .Rows(r).Cells("dgc_production_kg_in_orig").Value
                End If

            Next
        End With
    End Function

    Sub add_job(ByVal prod_num As Long, Optional ByVal desc As String = Nothing, Optional ByVal mounter As String = Nothing, Optional ByVal stripped As Boolean = False, Optional ByVal machine_ As String = Nothing)
        Dim found As Boolean
        If tb_prod_num.Text = Nothing Then
            tb_prod_num.Text = prod_num
        End If
        With dgv_summary
            If .RowCount = 0 Then
                .Rows.Add(prod_num.ToString)
                .Rows(0).Cells("dgc_summary_desc").Value = desc
                .Rows(0).Cells("dgc_summary_mounter").Value = mounter
                .Rows(0).Cells("dgc_summary_stripped").Value = stripped
            Else
                For i As Integer = 0 To .RowCount - 1
                    If Strings.InStr(.Rows(i).Cells(0).Value, prod_num) > 0 Then
                        found = True
                        Exit For
                    End If
                Next
                'If found And Not s_department = "mounting" Then
                '    MsgBox("Job already added!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
                '    Exit Sub
                'Else
                .Rows.Add(prod_num.ToString)
                If Not s_department = "mounting" Then
                    .Rows(.RowCount - 1).Cells("dgc_summary_start").Value = .Rows(.RowCount - 2).Cells("dgc_summary_finish").Value
                    .Rows(.RowCount - 1).Cells("dgc_summary_start").Value = .Rows(.RowCount - 2).Cells("dgc_summary_finish").Value
                End If

                .Rows(.RowCount - 1).Cells("dgc_summary_mounter").Value = mounter
                .Rows(.RowCount - 1).Cells("dgc_summary_desc").Value = desc
                .Rows(.RowCount - 1).Cells("dgc_summary_stripped").Value = stripped

                ' End If
            End If
            export_summary()

        End With
        Select Case s_department
            Case "mounting"
                If Not stripped Then
                    acsis.enter_minute(prod_num)
                    sql.update("db_schedule", "status", 1, "prod_num", "=", prod_num)
                Else
                    sql.update("db_schedule", "status", 10, "prod_num", "=", prod_num)
                End If
            Case "printing"
                sql.update("db_schedule", "machine", "'" & machine_ & "'", "prod_num", "=", prod_num)
        End Select

    End Sub

    Sub add_rts(ByVal prod_num As Long, ByVal mat_num As Long, ByVal n_in As Integer, ByVal r_id As String, ByVal quantity As Decimal)
        With dgv_rts
            Dim r As Integer = .RowCount - 1
            .Rows(r).Cells("dgc_rts_num_in").Value = n_in
            .Rows(r).Cells("dgc_rts_item_id").Value = r_id
            .Rows(r).Cells("dgc_rts_quantity").Value = FormatNumber(quantity, 1)
            .Rows(r).Cells("dgc_rts_send").Value = True
            .Rows(r).Cells("dgc_rts_user_").Value = s_user
            .Rows(r).Cells("dgc_rts_date_").Value = tb_start_date.Text
            .Rows(r).Cells("dgc_rts_prod_num_orig").Value = prod_num
            .Rows(r).Cells("dgc_rts_mat_num").Value = mat_num

        End With
        calc_rts()
    End Sub

    Sub add_reject(ByVal mat_num As Integer, ByVal r_id As String, ByVal n_in As String, ByVal quantity As Decimal)
        With dgv_reject
            Dim r As Integer = .RowCount - 1
            .Rows(r).Cells("dgc_reject_item_no").Value = r_id
            .Rows(r).Cells("dgc_reject_quantity").Value = FormatNumber(quantity, 1)
            .Rows(r).Cells("dgc_reject_user_").Value = s_user
            .Rows(r).Cells("dgc_reject_num_in").Value = n_in
            .Rows(r).Cells("dgc_reject_mat_num").Value = mat_num
            .Rows(r).Cells("dgc_reject_date_").Value = tb_start_date.Text
        End With
    End Sub

    Sub calc_downtime(ByVal prod_num As Long)
        Dim bs01 As Decimal = sql.sum("db_downtime", "time_lost", "prod_num", "=", prod_num & " AND type='BS01' AND user_='" & s_user & "' AND date_='" & tb_start_date.Text & "'")
        Dim cs01 As Decimal = sql.sum("db_downtime", "time_lost", "prod_num", "=", prod_num & " AND type='CS01' AND user_='" & s_user & "' AND date_='" & tb_start_date.Text & "'")
        tb_downtime_bs01_total.Text = bs01
        tb_downtime_cs01_total.Text = cs01
    End Sub
    Sub calc_rts()
        dgv_rts_pending.Rows.Clear()
        With dgv_rts
            For r As Integer = 0 To .RowCount - 1
                If .Rows(r).Cells("dgc_rts_send").Value = True Then
                    Dim mat_num As Integer = .Rows(r).Cells("dgc_rts_mat_num").Value
                    Dim quant As Decimal = .Rows(r).Cells("dgc_rts_quantity").Value
                    If dgv_rts_pending.Rows.Count = 0 Then
                        dgv_rts_pending.Rows.Add(mat_num, quant, 1)
                    Else
                        Dim found As Boolean = False
                        For i As Integer = 0 To dgv_rts_pending.Rows.Count - 1
                            If dgv_rts_pending.Rows(i).Cells(0).Value = mat_num Then
                                dgv_rts_pending.Rows(i).Cells(1).Value = dgv_rts_pending.Rows(i).Cells(1).Value + quant
                                dgv_rts_pending.Rows(i).Cells(2).Value = dgv_rts_pending.Rows(i).Cells(2).Value + 1
                                found = True
                            End If
                        Next
                        If Not found Then
                            dgv_rts_pending.Rows.Add(mat_num, quant, 1)
                        End If
                    End If
                End If
            Next
        End With

        If dgv_rts_pending.Rows.Count = 0 Then
            b_rts_send.Enabled = False
        Else
            b_rts_send.Enabled = s_symple_access
        End If
    End Sub

    Sub calc_times(ByVal row As Integer)
        Dim str, prod_num, dt As String, d As Date, calc As Boolean = False
        Dim bs01, cs01, es01, sched, meters, runtime, setup As Double, mpm As Integer = 0, multiplier As Integer = 0
        d = tb_start_date.Text
        dt = "'" & tb_start_date.Text.ToString & "'"
        With dgv_summary
            If .RowCount = 0 Then Exit Sub
            prod_num = .Rows(row).Cells(0).Value
     
            'If found And Not s_department = "mounting" Then
            '    MsgBox("Job already added!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
            '    Exit Sub
            'Else
            Dim produced_kg As String = Nothing
            Dim scrap As String = Nothing
            Dim rts As String = Nothing
            Dim reject As String = Nothing
            Dim issued As String = Nothing
            Dim produced_km As String = Nothing
            calc_downtime(prod_num)
            .Rows(row).Cells("dgc_summary_bs01").Value = FormatNumber(tb_downtime_bs01_total.Text, 2)
            .Rows(row).Cells("dgc_summary_cs01").Value = FormatNumber(tb_downtime_cs01_total.Text, 2)
            Dim filter As String = "user_='" & s_user & "' AND date_=" & dt & " AND shift='" & s_shift & "'"
            If Not dt_produced Is Nothing Then

                reject = dt_produced.Compute("Sum(reject)", filter).ToString
                rts = dt_produced.Compute("Sum(rts)", filter).ToString
                scrap = dt_produced.Compute("Sum(scrap)", filter).ToString
                produced_km = dt_produced.Compute("Sum(km_out)", filter).ToString
                If s_department = "bagplant" And Not produced_km = Nothing Then
                    produced_kg = get_conversion_factor(False, tb_mat_num_out.Text, False) * produced_km / 1000
                Else
                    produced_kg = dt_produced.Compute("Sum(kg_out)", filter).ToString
                End If
                issued = dt_produced.Compute("Sum(kg_in)", filter).ToString
            End If

            If Not IsNumeric(issued) Then
                issued = 0
            End If
            If Not IsNumeric(produced_kg) Then
                produced_kg = 0
            End If
            If Not IsNumeric(produced_km) Then
                produced_km = 0
            End If
            If Not IsNumeric(scrap) Then
                scrap = 0
            End If
            If Not IsNumeric(rts) Then
                rts = 0
            End If
            If Not IsNumeric(reject) Then
                reject = 0
            End If
            If s_extruder Then
                issued = 0
                For c As Integer = 0 To .ColumnCount - 1
                    If .Columns(c).Name.Contains("res_") And IsNumeric(.Rows(row).Cells(c).Value) Then
                        issued = issued + CDec(.Rows(row).Cells(c).Value)
                    End If
                Next
                .Rows(row).Cells("dgc_summary_issued").Value = FormatNumber(issued, 1, , , TriState.False)
                .Rows(row).Cells("dgc_summary_scrap").Value = FormatNumber(scrap, 1, , , TriState.False)
            Else
                .Rows(row).Cells("dgc_summary_issued").Value = FormatNumber(issued, 1, , , TriState.False)
                .Rows(row).Cells("dgc_summary_scrap").Value = FormatNumber(scrap, 1, , , TriState.False)
                .Rows(row).Cells("dgc_summary_rts").Value = FormatNumber(rts, 1, , , TriState.False)
                .Rows(row).Cells("dgc_summary_reject").Value = FormatNumber(reject, 1, , , TriState.False)
            End If
            .Rows(row).Cells("dgc_summary_produced").Value = FormatNumber(produced_kg, 1, , , TriState.False)
            Dim dp As Integer = Strings.Right(.Columns("dgc_summary_km_out").DefaultCellStyle.Format.ToString, 1)
            Select Case s_department
                Case "printing"
                    .Rows(row).Cells("dgc_summary_km_out").Value = FormatNumber(produced_km, dp, , , TriState.False)
                Case "bagplant", "filmsline"
                    .Rows(row).Cells("dgc_summary_km_out").Value = FormatNumber(produced_km, dp, , , TriState.False)
                    Dim dr() As DataRow = dt_machines.Select("name='" & s_machine & "'")
                    If Not IsNull(dr(0)("multiplier")) Then multiplier = dr(0)("multiplier")
                    'multiplier = sql.read("db_machines", "name", "=", "'" & s_machine & "' AND plant=" & s_plant, "multiplier")
                Case Else
                    .Rows(row).Cells("dgc_summary_km_out").Value = FormatNumber(produced_km, dp, , , TriState.False)
            End Select
            'calc used
            .Rows(row).Cells("dgc_summary_used").Value = FormatNumber(issued - rts - reject, 1, , , TriState.False)
            If .Rows(row).Cells("dgc_summary_start").Value = Nothing Or .Rows(row).Cells("dgc_summary_finish").Value = Nothing Then
                .Rows(row).Cells("dgc_summary_run").Value = Nothing
                Exit Sub
            End If
            str = .Rows(row).Cells("dgc_summary_start").Value.ToString.PadLeft(4, "0"c)
            .Rows(row).Cells("dgc_summary_start").Value = str



            Dim t1 As New TimeSpan(Strings.Left(str, 2), Strings.Right(str, 2), 0)
            Dim start As New DateTime(d.Year, d.Month, d.Day, t1.Hours, t1.Minutes, 0)
            str = .Rows(row).Cells("dgc_summary_finish").Value.ToString.PadLeft(4, "0"c)
            .Rows(row).Cells("dgc_summary_finish").Value = str
            t1 = New TimeSpan(Strings.Left(str, 2), Strings.Right(str, 2), 0)
            Dim finish As New DateTime(d.Year, d.Month, d.Day, t1.Hours, t1.Minutes, 0)
            If finish < start Then
                dt = d.AddDays(1)
                finish = New DateTime(d.Year, d.Month, d.Day, t1.Hours, t1.Minutes, 0)
            End If
            Dim diff As TimeSpan = finish - start

            str = .Rows(row).Cells("dgc_summary_bs01").Value
            If Not str = Nothing Then bs01 = str

            str = .Rows(row).Cells("dgc_summary_cs01").Value
            If Not str = Nothing Then cs01 = str

            str = .Rows(row).Cells("dgc_summary_es01").Value
            If Not str = Nothing Then es01 = str

            str = .Rows(row).Cells("dgc_summary_schd_out").Value
            If Not str = Nothing Then sched = str

            str = .Rows(row).Cells("dgc_summary_setup").Value
            If Not str = Nothing Then setup = str

            If lb_multiple_job_list.Items.Count > 0 Then
                runtime = FormatNumber(((diff.Hours + (diff.Minutes / 60)) / lb_multiple_job_list.Items.Count) - setup - es01 - sched, 2)
            Else
                runtime = FormatNumber((diff.Hours + (diff.Minutes / 60)) - setup - es01 - sched, 2)
            End If
            If runtime < 0 Then
                'adjust for day change
                runtime = runtime + 24
            End If

            .Rows(row).Cells("dgc_summary_run").Value = runtime

            Select Case s_department

            End Select
            str = .Rows(row).Cells("dgc_summary_km_out").Value
            If Not str = Nothing Then
                meters = str
            End If
            If s_department = "printing" Then
                meters = meters * 1000
            End If
            If Not (runtime - bs01 - cs01) = 0 Then
                If meters / ((runtime - bs01 - cs01) * 60) < 0 Then
                    mpm = meters / ((runtime - cs01) * 60)
                Else
                    mpm = meters / ((runtime - bs01 - cs01) * 60)

                End If
            End If
            str = .Rows(row).Cells("dgc_summary_cyl_in").Value
            If Not str = Nothing Then
                If Not str.ToString = "0" Then
                    If bs01 + cs01 = runtime Then
                        .Rows(row).Cells("dgc_summary_min_deck").Value = (setup * 60) / str
                    Else
                        .Rows(row).Cells("dgc_summary_min_deck").Value = (setup * 60) / str
                    End If
                End If
            Else
                .Rows(row).Cells("dgc_summary_min_deck").Value = 0
            End If
            If s_department = "filmsline" Then
                If s_extruder And Not IsNull(s_selected_job(0)("ply_separated")) Then
                    If s_selected_job(0)("ply_separated") Then mpm = mpm / 2
                Else
                    Dim w_in, w_out As Integer
                    w_in = get_material_info(True).width
                    w_out = get_material_info().width
                    If Not w_in = w_out Then
                        mpm = mpm / CInt(w_in / w_out)
                    End If
                End If
            End If
            .Rows(row).Cells("dgc_summary_mpm").Value = mpm.ToString
            For i As Integer = 0 To row - 1
                If Strings.InStr(.Rows(i).Cells(0).Value, prod_num) > 0 Then
                    .Rows(i).Cells("dgc_summary_mpm").Value = 0
                    .Rows(i).Cells("dgc_summary_km_out").Value = 0
                    .Rows(i).Cells("dgc_summary_used").Value = 0
                    .Rows(i).Cells("dgc_summary_rts").Value = 0
                    .Rows(i).Cells("dgc_summary_reject").Value = 0
                    .Rows(i).Cells("dgc_summary_issued").Value = 0

                    Exit For
                End If
            Next
        End With
    End Sub

    Sub calc_times_multiple(ByVal summary As Boolean)
        With dgv_summary
            tb_multiple_setup.Text = 0
            tb_multiple_run.Text = 0
            Dim multiple(lb_multiple_job_list.Items.Count - 1) As String
            lb_multiple_job_list.Items.CopyTo(multiple, 0)

            For i As Integer = 0 To lb_multiple_job_list.Items.Count - 1
                For r As Integer = 0 To .RowCount - 1
                    If lb_multiple_job_list.Items(i).ToString = .Rows(r).Cells(0).Value.ToString Then
                        If summary Then calc_times(r)
                        .Rows(r).Cells("dgc_summary_multiple").Value = String.Join(",", multiple)
                        If IsNumeric(.Rows(r).Cells("dgc_summary_setup").Value) Then tb_multiple_setup.Text = FormatNumber(CDec(tb_multiple_setup.Text) + .Rows(r).Cells("dgc_summary_setup").Value, 2, , , TriState.False)
                        If IsNumeric(.Rows(r).Cells("dgc_summary_run").Value) Then tb_multiple_run.Text = FormatNumber(CDec(tb_multiple_run.Text) + .Rows(r).Cells("dgc_summary_run").Value, 2, , , TriState.False)
                    End If
                Next
            Next
        End With
    End Sub






    Sub display_stations(ByVal count As Integer)
        'textboxes19-28
        'labels16-25
        tb_station_4.Location = New Point(tb_station_3.Location.X + 58, tb_station_3.Location.Y)
        tb_station_5.Location = New Point(tb_station_4.Location.X + 58, tb_station_4.Location.Y)
        tb_station_6.Location = New Point(tb_station_1.Location.X, tb_station_1.Location.Y + 37)
        tb_station_7.Location = New Point(tb_station_6.Location.X + 58, tb_station_6.Location.Y)
        tb_station_8.Location = New Point(tb_station_7.Location.X + 58, tb_station_7.Location.Y)
        tb_station_9.Location = New Point(tb_station_8.Location.X + 58, tb_station_8.Location.Y)
        tb_station_10.Location = New Point(tb_station_9.Location.X + 58, tb_station_9.Location.Y)
        tb_station_7.Visible = True
        tb_station_8.Visible = True
        tb_station_9.Visible = True
        tb_station_10.Visible = True
        Label22.Visible = True
        Label23.Visible = True
        Label24.Visible = True
        Label25.Visible = True
        If count = 6 Then
            tb_station_4.Location = tb_station_6.Location
            tb_station_5.Location = tb_station_7.Location
            tb_station_6.Location = tb_station_8.Location
            tb_station_7.Visible = False
            tb_station_8.Visible = False
            tb_station_9.Visible = False
            tb_station_10.Visible = False
            Label22.Visible = False
            Label23.Visible = False
            Label24.Visible = False
            Label25.Visible = False
        ElseIf count = 8 Then
            tb_station_5.Location = tb_station_6.Location
            tb_station_6.Location = tb_station_7.Location
            tb_station_7.Location = tb_station_8.Location
            tb_station_8.Location = tb_station_9.Location
            tb_station_9.Visible = False
            tb_station_10.Visible = False
            Label24.Visible = False
            Label25.Visible = False
        End If
        Label19.Location = New Point(tb_station_4.Location.X + (tb_station_4.Width / 2) - (Label19.Width / 2), tb_station_4.Location.Y - 16)
        Label20.Location = New Point(tb_station_5.Location.X + (tb_station_5.Width / 2) - (Label20.Width / 2), tb_station_5.Location.Y - 16)
        Label21.Location = New Point(tb_station_6.Location.X + (tb_station_6.Width / 2) - (Label21.Width / 2), tb_station_6.Location.Y - 16)
        Label22.Location = New Point(tb_station_7.Location.X + (tb_station_7.Width / 2) - (Label22.Width / 2), tb_station_7.Location.Y - 16)
        Label23.Location = New Point(tb_station_8.Location.X + (tb_station_8.Width / 2) - (Label23.Width / 2), tb_station_8.Location.Y - 16)
        Label24.Location = New Point(tb_station_9.Location.X + (tb_station_9.Width / 2) - (Label24.Width / 2), tb_station_9.Location.Y - 16)
        Label25.Location = New Point(tb_station_10.Location.X + (tb_station_10.Width / 2) - (Label25.Width / 2), tb_station_10.Location.Y - 16)
    End Sub
    Sub clear_job()
        s_loading = True
        tb_prod_num.Clear()
        dgv_job_requirements.Rows.Clear()
        dgv_production.Rows.Clear()
        dgv_downtime_bs01.Rows.Clear()
        dgv_downtime_cs01.Rows.Clear()
        dgv_rts.Rows.Clear()
        dgv_rts_pending.Rows.Clear()
        dgv_reject.Rows.Clear()
        'nud_items_out.Value = 1
        b_main_mat_in_add.Enabled = False
        tb_roll_length.Clear()
        tb_roll_length.Clear()
        tb_mat_num_out.Clear()
        tb_mat_desc_out.Clear()
        tb_label_desc.Clear()
        cbo_mat_num_in.Items.Clear()
        cbo_mat_desc_in.Items.Clear()
        Me.Text = "Sym-PLE-Fied"

        tb_cyl_size.Clear()
        tb_no_up.Clear()
        b_main_show_zaw.Text = Nothing

        lb_inks_list.Items.Clear()
        lb_inks.Items.Clear()

        cb_main_symple_entered.Checked = False
        cb_main_final_done.Checked = False


        clear_inks()
        s_loading = False

    End Sub
    Sub display_job(Optional ByVal prep As Boolean = False)
        Dim l, w As Integer, d, g As Decimal
        Dim o As Integer = 0
        If tb_prod_num.Text = Nothing Then Exit Sub
        'If s_user = Nothing And Not s_department = "graphics" Then Exit Sub
        s_loading = True
        Dim ink As String = Nothing, prod_num As Long ', job_info As DataTable = Nothing
        If s_selected_job.Length = 0 Then
            clear_job()
            Exit Sub
        End If
        l_main_sched_text.Text = Nothing
        prod_num = s_selected_job(0)("prod_num")
        tb_prod_num.Text = prod_num
        cbo_mat_num_in.Items.Clear()
        cbo_mat_desc_in.Items.Clear()
        b_main_mat_in_add.Enabled = True
        cb_production_ply_separator.Checked = False
        dgv_job_requirements.Rows.Clear()
        'TextBox1.Text = sql.read("DP_LOIPRO01_E2AFVOL", "aufnr", prod_num, "docnum", "symple")
        'TextBox10.Text = sql.read("DP_LOIPRO01_E2AFKOL", "aufnr", prod_num, "docnum", "symple")
        'TextBox35.Text = sql.read("DP_LOIPRO01_E2AFKOL", "aufnr", prod_num, "gmein", "symple")
        'TextBox34.Text = sql.read("DP_LOIPRO01_E2AFKOL", "aufnr", prod_num, "gamng", "symple")
        'TextBox33.Text = sql.read("DP_LOIPRO01_E2AFKOL", "aufnr", prod_num, "matnr", "symple")
        'str = "'" & Trim(TextBox33.Text).ToString.PadLeft(18, "0") & "'"
        'TextBox36.Text = sql.read("DP_ZMATCHARH", "matnr", str, "maktx", "symple")
        'TextBox37.Text = sql.read("DP_ZMATCHARH", "matnr", str, "inspno", "symple")
        'TextBox38.Text = Trim(sql.read("DP_ZMATCHARH", "matnr", str, "prntlb1", "symple")) & " " & Trim(sql.read("DP_ZMATCHARH", "matnr", str, "prntlb2", "symple"))
        'DP_LOIPRO01_Z2KNA1M = s_label data?
        If Not tc_main.SelectedTab Is tp_scheduling Then

            If s_selected_job(0)("default_quant").ToString = Nothing Then
                tb_roll_length.Text = get_material_info(False).length
                sql.update("db_schedule", "default_quant", tb_roll_length.Text, "prod_num", "=", prod_num)
                s_selected_job(0)("default_quant") = tb_roll_length.Text
            Else
                tb_roll_length.Text = s_selected_job(0)("default_quant")
            End If
        End If
        If Not IsNull(s_selected_job(0)("mat_num_in")) Then
            cbo_mat_num_in.Items.Add(s_selected_job(0)("mat_num_in").ToString)
            cbo_mat_desc_in.Items.Add(s_selected_job(0)("mat_desc_in").ToString)
            cbo_mat_num_in.Text = s_selected_job(0)("mat_num_in").ToString
            cbo_mat_desc_in.Text = s_selected_job(0)("mat_desc_in").ToString
        End If
        If Not IsNull(s_selected_job(0)("mat_num_in_1")) Then
            cbo_mat_num_in.Items.Add(s_selected_job(0)("mat_num_in_1").ToString)
            cbo_mat_desc_in.Items.Add(s_selected_job(0)("mat_desc_in_1").ToString)
        End If
        If Not IsNull(s_selected_job(0)("mat_num_in_2")) Then
            cbo_mat_num_in.Items.Add(s_selected_job(0)("mat_num_in_2").ToString)
            cbo_mat_desc_in.Items.Add(s_selected_job(0)("mat_desc_in_2").ToString)
        End If
        If Not IsNull(s_selected_job(0)("sched_text")) Then
            l_main_sched_text.Text = "Sched Text: " & s_selected_job(0)("sched_text")
        End If
        Select Case s_department
            Case "printing"
                If IsNull(s_selected_job(0)("mat_num_semi")) Then
                    tb_mat_num_out.Text = s_selected_job(0)("mat_num_fin").ToString
                    tb_mat_desc_out.Text = s_selected_job(0)("mat_desc_fin").ToString
                Else
                    tb_mat_num_out.Text = s_selected_job(0)("mat_num_semi").ToString
                    tb_mat_desc_out.Text = s_selected_job(0)("mat_desc_semi").ToString
                End If
                If IsNull(s_selected_job(0)("label_desc")) Then
                    tb_label_desc.Text = s_selected_job(0)("mat_desc_fin").ToString
                Else
                    tb_label_desc.Text = s_selected_job(0)("label_desc").ToString
                End If

            Case "bagplant"
                If IsNull(s_selected_job(0)("mat_num_fin")) Then
                    tb_mat_num_out.Text = s_selected_job(0)("mat_num_semi")
                    tb_mat_desc_out.Text = s_selected_job(0)("mat_desc_semi")
                Else
                    tb_mat_num_out.Text = s_selected_job(0)("mat_num_fin")
                    tb_mat_desc_out.Text = s_selected_job(0)("mat_desc_fin")
                End If
                tb_label_desc.Text = s_selected_job(0)("customer").ToString
                If Not tc_main.SelectedTab Is tp_scheduling Then frm_bag_checks.display_data()
            Case "filmsline"
                If IsNull(s_selected_job(0)("mat_num_fin")) Then
                    tb_mat_num_out.Text = s_selected_job(0)("mat_num_semi")
                    tb_mat_desc_out.Text = s_selected_job(0)("mat_desc_semi")
                Else
                    tb_mat_num_out.Text = s_selected_job(0)("mat_num_fin")
                    tb_mat_desc_out.Text = s_selected_job(0)("mat_desc_fin")
                End If
                If s_extruder Then
                    If Not IsNull(s_selected_job(0)("ply_separated")) Then
                        cb_production_ply_separator.Checked = s_selected_job(0)("ply_separated")
                    Else
                        s_selected_job(0)("ply_separated") = 0
                        sql.update("db_schedule", "ply_separated", 0, "prod_num", "=", s_selected_job(0)("prod_num"))
                    End If
                    s_loading = False
                    If Not IsNull(s_selected_job(0)("extrusion_type")) Then
                        select_cbo_item(cbo_job_info_type, s_selected_job(0)("extrusion_type"))
                    Else
                        cbo_job_info_type.SelectedIndex = 0
                        s_selected_job(0)("extrusion_type") = cbo_job_info_type.Text
                        sql.update("db_schedule", "extrusion_type", "'" & cbo_job_info_type.Text & "'", "prod_num", "=", s_selected_job(0)("prod_num"))
                    End If
                    s_loading = True

                End If
                If IsNull(s_selected_job(0)("customer")) Then
                    If Not IsNull(s_selected_job(0)("sales_line")) Then
                        If IsNull(s_selected_job(0)("sales_doc")) Then
                            tb_label_desc.Text = s_selected_job(0)("sales_line")
                        Else
                            Dim str1() As String = Split(s_selected_job(0)("sales_doc"))
                            Dim str2() As String = Split(s_selected_job(0)("sales_line"))
                            Dim s As String = Nothing
                            For i As Integer = 0 To UBound(str1)
                                If s = Nothing Then
                                    s = str1(i) & "_" & str2(i)
                                Else
                                    s = s & " / " & str1(i) & "_" & str2(i)
                                End If
                            Next
                            tb_label_desc.Text = s
                        End If
                    End If
                Else
                    tb_label_desc.Text = s_selected_job(0)("customer")
                End If
            Case Else
                tb_label_desc.Text = s_selected_job(0)("customer").ToString
        End Select
        Me.Text = "Sym-PLE-Fied - " & tb_label_desc.Text

        dgv_job_requirements.Rows.Add()
        dgv_job_requirements.Rows(0).Cells(0).Value = s_selected_job(0)("req_uom_1")

        Select Case s_selected_job(0)("req_uom_1").ToString
            Case "ROL"
                dgv_job_requirements.Rows(0).Cells(1).Value = CInt(s_selected_job(0)("req_quant_1"))
            Case "M"
                dgv_job_requirements.Rows(0).Cells(1).Value = FormatNumber(s_selected_job(0)("req_quant_1"), 0, , , TriState.False)
            Case "KM"
                dgv_job_requirements.Rows(0).Cells(1).Value = FormatNumber(s_selected_job(0)("req_quant_1"), 2, , , TriState.False)
            Case Else
                dgv_job_requirements.Rows(0).Cells(1).Value = FormatNumber(s_selected_job(0)("req_quant_1"), 1, , , TriState.False)
        End Select



        'ToolTip1.SetToolTip(tb_label_desc, tb_label_desc.Text)
        dgv_job_requirements.Rows.Add()
        Dim m_info As materialInfo = get_material_info(False)
        If m_info.material_type = material_type.film Or m_info.material_type = material_type.laminate Then
            cb_report_use_estimate.Visible = True
            tb_roll_length.Visible = True
        Else
            cb_report_use_estimate.Visible = False
            tb_roll_length.Visible = False
        End If
        Select Case s_selected_job(0)("req_uom_2").ToString
            Case Nothing
                Select Case s_selected_job(0)("req_uom_1").ToString
                    Case "ROL"
                        dgv_job_requirements.Rows(1).Cells(0).Value = "KM"
                        dgv_job_requirements.Rows(1).Cells(1).Value = _
                            FormatNumber(s_selected_job(0)("req_quant_1") * get_material_info(False).length / 1000, 2, , , TriState.False)
                        dgv_job_requirements.Rows(1).Cells(5).Value = 1
                        dgv_job_requirements.Rows.Add()
                        dgv_job_requirements.Rows(2).Cells(0).Value = "KG"
                        Select Case get_material_info(False).formulation
                            Case "A203"
                                o = 5
                            Case "DL19"
                                o = 5
                            Case Else
                                g = g
                        End Select
                        d = get_item_desity(s_selected_job(0)("mat_desc_fin"))
                        l = m_info.length
                        w = m_info.width + o
                        g = m_info.gauge

                        dgv_job_requirements.Rows(2).Cells(1).Value = _
                            FormatNumber((d * w * l * g / 1000000) * dgv_job_requirements.Rows(0).Cells(1).Value, 1, , , TriState.False)

                        dgv_job_requirements.Rows(2).Cells(5).Value = 1
                    Case "KG"
                        dgv_job_requirements.Rows(1).Cells(0).Value = "M"
                        ' Dim str() As String = Split(dgv_production.Columns("dgc_production_km_out").HeaderText, " ")
                        'If Not str(0) = "M" Then
                        dgv_job_requirements.Rows(1).Cells(1).Value = _
                           FormatNumber(get_conversion_factor(False, tb_mat_num_out.Text, False) * dgv_job_requirements.Rows(0).Cells(1).Value * 1000, 0, , , TriState.False)
                        'Else
                        '    dgv_job_requirements.Rows(1).Cells(1).Value = _
                        '     FormatNumber(get_conversion_factor(True) * dgv_job_requirements.Rows(0).Cells(1).Value, 0, , , TriState.False)
                        'End If
                        dgv_job_requirements.Rows(1).Cells(5).Value = 1
                    Case "KM"
                        dgv_job_requirements.Rows(1).Cells(0).Value = "KG"
                        dgv_job_requirements.Rows(1).Cells(1).Value = _
                            FormatNumber(get_conversion_factor(False, tb_mat_num_out.Text, True) * dgv_job_requirements.Rows(0).Cells(1).Value, 1, , , TriState.False)
                    Case "MU"
                        dgv_job_requirements.Rows.RemoveAt(1)
                        'dgv_job_requirements.Rows(1).Cells(0).Value = "KG"
                        'dgv_job_requirements.Rows(1).Cells(1).Value = _
                        '    FormatNumber(get_conversion_factor(False, True) * dgv_job_requirements.Rows(0).Cells(1).Value, 1, , , TriState.False)
                    Case Else
                        dgv_job_requirements.Rows(1).Cells(1).Value = FormatNumber(s_selected_job(0)("req_quant_1"), 1, , , TriState.False)
                        ink = ink
                End Select
                ink = ink

            Case Else
                dgv_job_requirements.Rows(1).Cells(0).Value = s_selected_job(0)("req_uom_2")
                Select Case s_selected_job(0)("req_uom_2")
                    Case "KM"
                        dgv_job_requirements.Rows(1).Cells(1).Value = FormatNumber(s_selected_job(0)("req_quant_2"), 2, , , TriState.False)
                    Case Else
                        dgv_job_requirements.Rows(1).Cells(1).Value = FormatNumber(s_selected_job(0)("req_quant_2"), 1, , , TriState.False)
                End Select
        End Select

        If Not IsNull(s_selected_job(0)("cyl_size")) Then
            If s_extruder Then
                tb_cyl_size.Text = FormatNumber(s_selected_job(0)("cyl_size"), 1, , , TriState.False)
            Else
                tb_cyl_size.Text = CInt(s_selected_job(0)("cyl_size"))
            End If
        Else
            tb_cyl_size.Text = Nothing
        End If
        tb_no_up.Text = s_selected_job(0)("no_up").ToString
        b_main_show_zaw.Text = get_material_info.zaw
        If b_main_show_zaw.Text = Nothing Then
            b_main_show_zaw.Visible = False
        Else
            b_main_show_zaw.Visible = True
        End If
        If s_production Then

            s_label = get_label(tb_mat_num_out.Text, tb_mat_desc_out.Text)
            s_pallet = get_label(tb_mat_num_out.Text, tb_mat_desc_out.Text, pallet:=True)
            If Not cbo_mat_num_in.Text = Nothing Then
                If Not s_extruder Then s_label_rts = get_label(cbo_mat_num_in.Text, cbo_mat_desc_in.Text, rts:=True)

            End If
        End If
        Select Case s_department
            Case "printing", "filmsline"
                If s_selected_job(0)("items_out").ToString = Nothing Then
                    nud_items_out.Value = nud_items_out.Minimum
                    sql.update("db_schedule", "items_out", nud_items_out.Value, "prod_num", "=", prod_num)
                    s_selected_job(0)("items_out") = nud_items_out.Value
                    frm_status.Close()
                ElseIf Not s_selected_job(0)("items_out") = 0 Then
                    nud_items_out.Maximum = s_selected_job(0)("items_out") + 10
                    nud_items_out.Value = s_selected_job(0)("items_out")
                Else
                    nud_items_out.Value = 1
                End If
            Case "bagplant"
                Dim grid_item As String = "Carton Quantity"
                If Not tc_main.SelectedTab Is tp_scheduling Then
                    If cbo_machine.SelectedItem("rewinder") = True Then
                        l_items_out.Text = "Meters/Roll"
                        grid_item = "Length"
                    Else
                        l_items_out.Text = "Bags/Carton"
                    End If

                    If s_selected_job(0)("items_out").ToString = Nothing Then
                        acsis.open_job_screen(prod_num, False)
                        show_status("Getting label details")
                        acsis.click_button(acsis.get_button_access("Label Reprint", 0), False, 0)
                        acsis.select_cb_item(s_label, 0)
                        acsis.select_grid_item(grid_item, True)
                        If Not Clipboard.GetText = "FAILED" Then
                            nud_items_out.Value = Clipboard.GetText
                        Else
                            nud_items_out.Value = 700
                        End If
                        sql.update("db_schedule", "items_out", nud_items_out.Value, "prod_num", "=", prod_num)
                        s_selected_job(0)("items_out") = nud_items_out.Value
                        s_selected_job = dt_schedules.Select("prod_num=" & prod_num)
                        frm_status.Close()
                    Else
                        nud_items_out.Value = s_selected_job(0)("items_out")
                    End If
                End If
                frm_bag_checks.import_bursts()
        End Select

        If Not tc_main.SelectedTab Is tp_scheduling Then
            lb_inks_list.Items.Clear()
            lb_inks.Items.Clear()

            If s_department = "printing" Then
                If Not IsNull(s_selected_job(0)("inks")) Then
                    clear_inks()

                    Dim inks_current As String = s_selected_job(0)("inks")
                    Dim inks_previous As String = Nothing
                    Dim inks_set As String = Nothing
                    With lb_inks_list.Items
                        .Clear()
                        .AddRange(Split(inks_current, ","))
                    End With
                    If IsNull(s_selected_job(0)("set_inks")) Then
                        With dgv_summary
                            If .CurrentRow.Index > 0 Then
                                inks_previous = sql.read("db_schedule", "prod_num", "=", .Rows(.CurrentRow.Index - 1).Cells("dgc_summary_prod_num").Value, "set_inks")
                            End If
                        End With
                    Else
                        inks_set = s_selected_job(0)("set_inks")
                    End If
                    Dim inks() As String
                    Dim add_inks As Boolean = False
                    With lb_inks
                        If inks_set = Nothing Then
                            .Items.Clear()
                            .Items.AddRange(Split(inks_current, ","))
                            inks = Split(inks_previous, ",")
                            Dim found As Integer = 0

                            If Not inks_previous = Nothing Then
                                For i As Integer = 0 To .Items.Count - 1
                                    For Each ink In inks
                                        Dim str() As String = Split(ink, ":")
                                        If .Items(i) = str(1) Then
                                            found = found + 1
                                            Exit For
                                        End If
                                    Next
                                Next
                                If found = .Items.Count Then
                                    add_inks = True
                                    s_selected_job(0)("set_inks") = inks_previous
                                    sql.update("db_schedule", "set_inks", sql.convert_value(inks_previous), "prod_num", "=", prod_num & "AND plant=" & s_plant)
                                End If
                            End If
                        Else
                            inks = Split(inks_set, ",")
                            add_inks = True
                        End If
                        If add_inks Then

                            .Items.Clear()
                            For Each ink In inks
                                Dim str() As String = Split(ink, ":")
                                Select Case CInt(str(0))
                                    Case 1
                                        tb_station_1.Text = str(1)
                                    Case 2
                                        tb_station_2.Text = str(1)
                                    Case 3
                                        tb_station_3.Text = str(1)
                                    Case 4
                                        tb_station_4.Text = str(1)
                                    Case 5
                                        tb_station_5.Text = str(1)
                                    Case 6
                                        tb_station_6.Text = str(1)
                                    Case 7
                                        tb_station_7.Text = str(1)
                                    Case 8
                                        tb_station_8.Text = str(1)
                                    Case 9
                                        tb_station_9.Text = str(1)
                                    Case 10
                                        tb_station_10.Text = str(1)
                                End Select
                            Next

                        End If

                    End With

                End If
            End If

            If Not dgv_summary.CurrentRow Is Nothing Then
                If IsNull(dgv_summary.CurrentRow.Cells("dgc_summary_symple_times").Value) Then
                    cb_main_symple_entered.Checked = False
                Else
                    cb_main_symple_entered.Checked = dgv_summary.CurrentRow.Cells("dgc_summary_symple_times").Value
                End If
                If IsNull(dgv_summary.CurrentRow.Cells("dgc_summary_final_conf").Value) Then
                    cb_main_final_done.Checked = False
                Else
                    cb_main_final_done.Checked = dgv_summary.CurrentRow.Cells("dgc_summary_final_conf").Value
                End If
            End If

            If prep Then
                s_loading = False
                Exit Sub
            End If

            If lb_inks.Items.Count > 0 Then
                lb_inks.SelectedIndex = 0
            End If


            If s_department = "printing" And Not IsNull(s_selected_job(0)("inks")) And IsNull(s_selected_job(0)("set_inks")) Then
                tc_main.SelectTab(5)
            End If
            import_downtime("BS01")
            import_downtime("CS01")
            import_rts(prod_num)
            import_reject(prod_num)
        End If

        If Not s_department = "mounting" And s_production Then import_report()
        ' check_downtime()
        dgv_job_requirements.ClearSelection()
        If tc_main.SelectedTab Is tp_scheduling Then
            dt_produced = sql.get_table("db_produced", "prod_num", "=", "'" & tb_prod_num.Text & "'")
        Else
            dt_produced = dgv_to_dt(dgv_production)

        End If
        calc_totals()

        s_loading = False
    End Sub

    Sub export_downtime(ByVal type As String)
        Dim str, columns, values As String, prod_num As Long, dgv As DataGridView
        If s_loading Then Exit Sub
        prod_num = tb_prod_num.Text

        If type = "BS01" Then
            dgv = dgv_downtime_bs01
        ElseIf type = "CS01" Then
            dgv = dgv_downtime_cs01
        Else
            dgv = dgv_downtime_bs01
        End If

        sql.delete("db_downtime", "user_", "=", "'" & s_user & "' AND date_='" & tb_start_date.Text & _
                   "' AND prod_num=" & prod_num & " AND type='" & type & "'")

        With dgv
            For r As Integer = 0 To .RowCount - 1
                str = Nothing
                columns = Nothing
                values = Nothing
                For c As Integer = 0 To .ColumnCount - 1

                    If IsDBNull(.Rows(r).Cells(c).Value) Then
                        str = .Rows(r).Cells(c).Value.ToString
                    Else
                        str = .Rows(r).Cells(c).Value
                    End If

                    If IsDate(str) Then
                        str = "'" & str & "'"
                    Else
                        str = sql.convert_value(str)
                    End If

                    If columns = Nothing Then
                        columns = LCase(.Columns(c).HeaderText.Replace(" ", "_").Replace("downtime_", Nothing))
                        values = str
                    Else
                        columns = columns & "," & LCase(.Columns(c).HeaderText.Replace(" ", "_").Replace("downtime_", Nothing))
                        values = values & "," & str
                    End If

                Next
                columns = columns & "," & "row,date_,user_,machine,prod_num,shift,type,access"
                values = values & "," & r & ",'" & tb_start_date.Text & "','" & s_user & "','" & s_machine & "'," & _
                    prod_num & ",'" & s_shift & "','" & type & "'," & Convert.ToSByte(s_sql_access)

                sql.insert("db_downtime", columns, values)

            Next
        End With

    End Sub
    Sub export_plates()
        Dim bs As String = "'" & Strings.Left(cbo_plate_boxsize.Text, 1) & "'"
        sql.delete("db_plate_location", "plant", "=", s_plant & " AND box_size=" & bs)
        With dgv_plates
            For i As Integer = 0 To .RowCount - 1
                Dim pn, loc, desc, check, plates, com As String

                loc = .Rows(i).Cells(1).Value

                If IsNull(.Rows(i).Cells(0).Value) Then
                    pn = "NULL"
                Else
                    pn = "'" & .Rows(i).Cells(0).Value & "'"
                End If

                If IsNull(.Rows(i).Cells(2).Value) Then
                    desc = "NULL"
                Else
                    desc = "'" & Replace(.Rows(i).Cells(2).Value, "'", "''") & "'"
                End If
                If IsNull(.Rows(i).Cells(3).Value) Then
                    com = "NULL"
                Else
                    com = "'" & Replace(.Rows(i).Cells(3).Value, "'", "''") & "'"
                End If
                If IsNull(.Rows(i).Cells(4).Value) Then
                    check = "NULL"
                Else
                    check = "'" & .Rows(i).Cells(4).Value & "'"
                End If
                If IsNull(.Rows(i).Cells(5).Value) Then
                    plates = "NULL"
                Else
                    plates = "'" & .Rows(i).Cells(5).Value & "'"
                End If
                Dim col As Integer = CInt(Mid(loc, 2, 3))
                Dim row As String = "'" & Strings.Right(loc, 1) & "'"

                sql.insert("db_plate_location", "plate_num,box_size,col,row,description,comment,plates,checked,plant", pn & "," & bs & "," & col & "," & row & "," & desc & "," & com & "," & plates & "," & check & "," & s_plant)
            Next

        End With

    End Sub
    Sub export_reject()
        Dim str, columns, values, c_name As String, prod_num As Long
        prod_num = tb_prod_num.Text
        sql.delete("db_rejects", "prod_num", "=", prod_num)
        With dgv_reject
            For r As Integer = 0 To .RowCount - 1
                str = Nothing
                columns = Nothing
                values = Nothing
                For c As Integer = 0 To .ColumnCount - 1
                    c_name = Replace(.Columns(c).Name, "dgc_reject_", Nothing)
                    If IsDBNull(.Rows(r).Cells(c).Value) Then
                        str = .Rows(r).Cells(c).Value.ToString
                    Else
                        str = .Rows(r).Cells(c).Value
                    End If
                    If IsDate(str) Then
                        str = "'" & str & "'"
                    Else
                        str = sql.convert_value(str)
                    End If
                    If columns = Nothing Then
                        columns = c_name
                        values = str
                    Else
                        columns = columns & "," & c_name
                        values = values & "," & str
                    End If

                Next
                columns = columns & ",prod_num,row,access,location,machine"
                values = values & "," & prod_num & "," & r & "," & Convert.ToSByte(s_sql_access) & ",'" & cbo_reject_destination.Text & "','" & s_machine & "'"
                sql.insert("db_rejects", columns, values)
            Next
        End With

    End Sub

    Sub export_report(Optional ByVal row As Integer = -1)
        ' Exit Sub
        Dim str, columns, values, c_name As String, prod_num As Long, r_start As Integer, r_end As Integer, r_export As Integer
        sql.delete("db_produced", "prod_num", "IS", "NULL")


        prod_num = tb_prod_num.Text
        If row = -1 Then sql.delete("db_produced", "prod_num", "=", prod_num)

        With dgv_production
            If row = -1 Then
                r_start = 0
                r_end = .RowCount - 1
            Else
                r_start = row
                r_end = row
            End If
            r_export = r_start
            For r As Integer = r_start To r_end
                If .Rows(r).Visible Then

                    str = Nothing
                    columns = Nothing
                    values = Nothing
                    For c As Integer = 0 To .ColumnCount - 1
                        c_name = Replace(.Columns(c).Name, "dgc_production_", Nothing)
                        If IsDBNull(.Rows(r).Cells(c).Value) Then
                            str = .Rows(r).Cells(c).Value.ToString
                        Else
                            str = .Rows(r).Cells(c).Value
                        End If
                        If IsDate(str) Then
                            str = "'" & str & "'"
                        Else
                            str = sql.convert_value(str)
                        End If

                        If columns = Nothing Then
                            If row = -1 Then
                                columns = c_name
                                values = str
                            Else
                                columns = c_name & "=" & str
                            End If
                        Else
                            If row = -1 Then
                                columns = columns & "," & c_name
                                values = values & "," & str
                            Else
                                columns = columns & "," & c_name & "=" & str
                            End If
                        End If

                    Next
                    If row = -1 Then
                        columns = columns & ",row,access,dept"
                        values = values & "," & r_export & "," & Convert.ToSByte(s_sql_access) & "," & sql.convert_value(s_department)
                        sql.insert("db_produced", columns, values)
                    Else
                        columns = columns & ",access=" & Convert.ToSByte(s_sql_access)
                        sql.update("db_produced", columns & ",dept", sql.convert_value(s_department), "prod_num", "=", prod_num & " AND row=" & row)
                        'sql.insert("db_produced", columns, values)

                    End If
                    r_export = r_export + 1
                End If
            Next
        End With
        ' dt_produced = dgv_to_dt(dgv_production) 'sql.get_table("db_produced", "prod_num", "=", prod_num)

    End Sub

    Sub export_rts()
        Dim str, columns, values, c_name As String, prod_num As Long = tb_prod_num.Text
        sql.delete("db_rts", "prod_num", "=", prod_num)
        With dgv_rts
            If .RowCount > 0 Then
                For r As Integer = 0 To .RowCount - 1
                    str = Nothing
                    columns = Nothing
                    values = Nothing
                    For c As Integer = 0 To .Columns.Count - 2
                        c_name = Replace(.Columns(c).Name, "dgc_rts_", Nothing)
                        If Not c_name = "send" Then

                            If IsDBNull(.Rows(r).Cells(c).Value) Then
                                str = .Rows(r).Cells(c).Value.ToString
                            Else
                                str = .Rows(r).Cells(c).Value
                            End If

                            If IsDate(str) Then
                                str = "'" & str & "'"
                            Else
                                str = sql.convert_value(str)
                            End If

                            If columns = Nothing Then
                                columns = c_name
                                values = str
                            Else
                                columns = columns & "," & c_name
                                values = values & "," & str
                            End If

                        End If


                    Next
                    columns = columns & ",row,prod_num,access,machine"
                    values = values & "," & r & "," & prod_num & "," & Convert.ToSByte(s_sql_access) & ",'" & s_machine & "'"
                    sql.insert("db_rts", columns, values)

                Next
            End If
        End With
    End Sub

    Sub export_summary()
        Dim str, columns, values As String, prod_num As Long, tmp_user As String
        If s_user = Nothing Then Exit Sub
        tmp_user = s_user

        If s_department = "mounting" Then
            's_user = "PMNT" & s_shift_name
        ElseIf s_department = "graphics" Then
            Exit Sub
        End If
        sql.delete("db_summary", "user_", "=", "'" & s_user & "' AND date_='" & tb_start_date.Text & "' AND machine='" & s_machine & "' AND shift='" & s_shift & "'")
        With dgv_summary

            For r As Integer = 0 To .RowCount - 1
                str = Nothing
                columns = Nothing
                values = Nothing
                prod_num = .Rows(r).Cells(0).Value
                For c As Integer = 0 To .Columns.Count - 1

                    If Not .Columns(c).ReadOnly Or c = 0 Then
                        If IsDBNull(.Rows(r).Cells(c).Value) Then
                            str = .Rows(r).Cells(c).Value.ToString
                        Else
                            str = .Rows(r).Cells(c).Value
                        End If

                        If IsDate(str) Then
                            str = "'" & str & "'"
                        Else
                            str = sql.convert_value(str)
                        End If

                        If columns = Nothing Then
                            columns = "prod_num"
                            values = str
                        Else
                            columns = columns & "," & Replace(.Columns(c).Name, "dgc_summary_", Nothing)
                            values = values & "," & str
                        End If
                    End If

                Next
                columns = columns & ",user_,row,shift,machine,date_,access,assistant,dept,plant"
                values = values & ",'" & s_user & "'," & r & ",'" & s_shift & "','" & s_machine & "','" & tb_start_date.Text & "'," & Convert.ToSByte(s_sql_access) & ",'" & s_assistant & "','" & s_department & "'," & s_plant
                sql.insert("db_summary", columns, values)
            Next
        End With
        If s_department = "mounting" Then
            s_user = tmp_user
        End If
    End Sub

    Sub import_downtime(ByVal type As String)
        Dim dt As DataTable, prod_num As Long, dgv As DataGridView, tb As TextBox, t As Decimal = 0
        If type = "BS01" Then
            dgv = dgv_downtime_bs01
            tb = tb_downtime_bs01_total
        ElseIf type = "CS01" Then
            dgv = dgv_downtime_cs01
            tb = tb_downtime_cs01_total
        Else
            dgv = dgv_downtime_bs01
            tb = tb_downtime_bs01_total
        End If
        prod_num = tb_prod_num.Text
        dt = sql.get_table("db_downtime", "prod_num", "=", prod_num & " AND user_='" & s_user & "' AND date_='" & tb_start_date.Text & "' AND shift='" & s_shift & "' AND type='" & type & "'")
        With dgv
            .Rows.Clear()
            If Not IsNothing(dt) Then
                For i As Integer = 0 To dt.Compute("MAX(row)", "")
                    .Rows.Add()
                Next
                For r As Integer = 0 To dt.Rows.Count - 1
                    Dim dgvrow As Integer = dt(r)("row")
                    .Rows(dgvrow).Cells(0).Value = dt(r)("time_lost")
                    .Rows(dgvrow).Cells(1).Value = dt(r)("main")
                    .Rows(dgvrow).Cells(2).Value = dt(r)("reason")
                    .Rows(dgvrow).Cells(3).Value = dt(r)("comment")
                    .Rows(dgvrow).Cells(4).Value = dt(r)("setup")
                    .Rows(dgvrow).Cells(5).Value = dt(r)("five_why")
                    t = t + dt(r)("time_lost")
                Next
            End If
            If .RowCount > 0 Then
                .ClearSelection()
                .Rows(.RowCount - 1).Selected = True
                tb.Text = t
            End If
        End With
    End Sub

    Sub import_report()
        Dim tempdt As New DataTable, row As Integer
        Dim dr As DataRow()
        tempdt = sql.get_table("db_produced", "prod_num", "=", tb_prod_num.Text)
        If Not IsNothing(tempdt) Then
            dr = tempdt.Select("row=0")
            If dr.Length = 0 Then
                MsgBox("An error has happened trying to load the report, you will need to re-enter the items")
                Exit Sub
            End If
        End If

        ' TabControl1.SelectedIndex = 1
        s_loading = True

        With dgv_production
            .Rows.Clear()
            If Not IsNothing(tempdt) Then
                For Each r As DataRow In tempdt.Rows
                    .Rows.Add()
                Next
                For Each r As DataRow In tempdt.Rows
                    row = r("row")
                    For c As Integer = 0 To .ColumnCount - 1
                        Dim c_name As String = Replace(.Columns(c).Name, "dgc_production_", Nothing)
                        If Not IsNull(r(c_name)) Then

                            Select Case c_name
                                Case "num_in", "item_id", "kg_in", "num_out"
                                    .Rows(row).Cells(c).Value = "-"
                                    If IsNumeric(r(c_name)) Then
                                        If r(c_name) > 0 Then
                                            .Rows(row).Cells(c).Value = r(c_name)
                                        End If
                                    Else
                                        .Rows(row).Cells(c).Value = r(c_name)
                                    End If
                                Case "burst_info"
                                    .Rows(row).Cells(c).Value = CBool(r(c_name))
                                Case Else
                                    If Not IsNull(r(c_name)) Then
                                        If .Columns(c).DefaultCellStyle.Format.ToString.Contains("N") Then
                                            .Rows(row).Cells(c).Value = r(c_name)
                                            If IsNumeric(r(c_name)) Then
                                                Dim dp As Integer = Strings.Right(.Columns(c).DefaultCellStyle.Format, 1)
                                                .Rows(row).Cells(c).Value = FormatNumber(r(c_name), dp, , , False)
                                            End If
                                        Else
                                            .Rows(row).Cells(c).Value = r(c_name)
                                        End If
                                    End If
                            End Select
                        End If
                    Next

                Next

            End If
            .PerformLayout()
            Dim col As String
            Select Case s_department
                Case "printing", "filmsline"
                    col = "dgc_production_kg_out"
                Case "bagplant"
                    col = "dgc_production_km_out"
                Case Else
                    col = Nothing
            End Select
            If .RowCount > 0 Then
                .FirstDisplayedScrollingRowIndex = .RowCount - 1
                Dim current_done As Boolean = False
                For i As Integer = .RowCount - 1 To 0 Step -1
                    If Not IsNull(.Rows(i).Cells("dgc_production_kg_out").Value) Or Not IsNull(.Rows(i).Cells("dgc_production_km_out").Value) Then
                        If .RowCount - 1 > i Then
                            If i = 0 Then
                                If IsNull(.Rows(i).Cells(col).Value) Then
                                    .CurrentCell = .Rows(i).Cells(col)
                                Else
                                    .CurrentCell = .Rows(i + 1).Cells(col)
                                End If
                            Else
                                .CurrentCell = .Rows(i).Cells(col)
                                .CurrentCell = .Rows(i + 1).Cells(col)

                            End If
                            current_done = True
                        Else
                            .CurrentCell = .Rows(i).Cells(col)
                            current_done = True
                        End If
                        Exit For
                    End If
                Next
                If Not current_done Then
                    .CurrentCell = .Rows(0).Cells(col)
                End If
            End If
            If tc_main.SelectedTab Is tp_production Then
                If Not .CurrentCell Is Nothing Then get_production_info(.CurrentRow.Index)

            End If
        End With
        'If nud_items_out.Value > 1 Then
        '    b_report_merge.Enabled = False
        'Else
        '    b_report_merge.Enabled = True
        'End If
        s_loading = False

    End Sub
    Sub import_reject(ByVal prod_num As Long)
        Dim tempdt As New DataTable, row As Integer
        If cb_show_all.Checked Then
            tempdt = sql.get_table("db_rejects", "prod_num", "=", prod_num)
        Else
            tempdt = sql.get_table("db_rejects", "prod_num", "=", prod_num & " AND user_='" & s_user & "' AND date_='" & tb_start_date.Text & "'")
        End If
        With dgv_reject
            .Rows.Clear()
            If Not IsNothing(tempdt) Then
                For Each r As DataRow In tempdt.Rows
                    .Rows.Add()
                Next
                For Each r As DataRow In tempdt.Rows
                    row = r("row")
                    For c As Integer = 0 To .ColumnCount - 1
                        Dim c_name As String = Replace(.Columns(c).Name, "dgc_reject_", Nothing)
                        If Not IsNull(r(c_name)) Then
                            Select Case c_name
                                Case "quantity"
                                    .Rows(row).Cells(c).Value = FormatNumber(r(c_name), 1, , , TriState.False)
                                Case "code"
                                    .Rows(row).Cells(c).Value = r(c_name).ToString.PadLeft(3, "0")
                                Case Else
                                    .Rows(row).Cells(c).Value = r(c_name)
                            End Select
                        End If
                    Next

                Next
                If Not IsNull(tempdt(0)("location")) Then
                    select_cbo_item(cbo_reject_destination, tempdt(0)("location"))
                End If
            End If
        End With
    End Sub

    Sub import_rts(ByVal prod_num As Long)
        Dim tempdt As New DataTable, row As Integer

        tempdt = sql.get_table("db_rts", "prod_num", "=", prod_num & " AND user_='" & s_user & "' AND date_='" & tb_start_date.Text & "'")
        With dgv_rts
            .Rows.Clear()
            If Not IsNothing(tempdt) Then
                For Each r As DataRow In tempdt.Rows
                    .Rows.Add()
                Next
                For Each r As DataRow In tempdt.Rows
                    row = r("row")
                    For c As Integer = 0 To .ColumnCount - 1
                        Dim c_name As String = Replace(.Columns(c).Name, "dgc_rts_", Nothing)
                        If Not c_name = "send" Then

                            If Not IsNull(r(c_name)) Then
                                Select Case c_name
                                    Case "quantity"
                                        .Rows(row).Cells(c).Value = FormatNumber(r(c_name), 1, , , TriState.False)
                                    Case Else
                                        .Rows(row).Cells(c).Value = r(c_name)
                                End Select
                            End If
                        End If
                    Next
                    If Not IsNull(.Rows(row).Cells("dgc_rts_quantity").Value) And Not IsNull(.Rows(row).Cells("dgc_rts_quantity_symple").Value) Then

                        If Not .Rows(row).Cells("dgc_rts_quantity").Value.ToString = .Rows(row).Cells("dgc_rts_quantity_symple").Value.ToString Then
                            .Rows(row).Cells("dgc_rts_send").Value = True
                        Else
                            .Rows(row).Cells("dgc_rts_send").Value = False
                        End If
                    ElseIf Not IsNull(.Rows(row).Cells("dgc_rts_quantity").Value) And IsNull(.Rows(row).Cells("dgc_rts_quantity_symple").Value) Then
                        .Rows(row).Cells("dgc_rts_send").Value = True
                    End If
                Next
            End If
        End With
        calc_rts()
    End Sub
    Sub plate_search(Optional ByVal by_box As Boolean = False)
        If tb_plate_search.Text = Nothing Then
            If Not gb_box_control.Enabled Then
                Exit Sub
            End If
        End If
        If by_box Then
            dt_plates = sql.get_table("db_plate_location", "plant", "=", s_plant & " AND box_size='" & Strings.Left(cbo_plate_boxsize.Text, 1) & "'")
        Else
            dt_plates = sql.get_table("db_plate_location", "plant", "=", s_plant & " AND plate_num LIKE '%" & tb_plate_search.Text & "%' OR description LIKE '%" & tb_plate_search.Text & "%'")
        End If
        With dgv_plates
            .Rows.Clear()
            If Not dt_plates Is Nothing Then
                Dim dv As DataView = dt_plates.DefaultView
                dv.Sort = "box_size ASC, col ASC,row ASC,plate_num ASC"
                dt_plates = dv.ToTable
                'dt_plates.AcceptChanges()
                'dgv_plates.DataSource = dt_plates
                For i As Integer = 0 To dt_plates.Rows.Count - 1
                    Dim loc As String = dt_plates(i)("box_size").ToString & dt_plates(i)("col").ToString.PadLeft(3, "0") & dt_plates(i)("row").ToString
                    .Rows.Add(dt_plates(i)("plate_num").ToString, loc, dt_plates(i)("description"), dt_plates(i)("comment"), dt_plates(i)("checked"), dt_plates(i)("plates"), dt_plates(i)("archived"))
                Next
                .ClearSelection()
                .CurrentCell = Nothing
                '.Sort(.Columns(0), 0)
                '.Sort(.Columns(1), 0)
                .PerformLayout()
                If .RowCount = 1 Then
                    .CurrentCell = .Rows(0).Cells(0)
                    .Rows(0).Selected = True
                    tb_plate_desc.Text = .Rows(0).Cells(2).Value.ToString
                    tb_plate_comment.Text = .Rows(0).Cells(3).Value.ToString
                    b_plate_change_desc.Enabled = True
                End If
            End If

        End With
    End Sub
    Sub populate_plates()
        s_loading = True
        Dim dept As String = s_department
        Select Case dept
            Case "mounting", "graphics"
                dept = "printing"
        End Select

        Dim dt As DataTable = sql.get_table("plate_options", "box_size", "=", "'" & cbo_plate_boxsize.Text & "' AND plant=" & s_plant & " AND dept='" & dept & "'")
        tb_plate_total.Clear()
        tb_plate_used.Clear()
        tb_plate_remaining.Clear()
        If dt Is Nothing Then
            Exit Sub
        Else
            nud_plate_boxes.Value = dt(0)("quant")
            nud_plate_columns.Value = dt(0)("col")
            nud_plate_rows.Value = dt(0)("row")
            tb_plate_total.Text = dt(0)("quant") * dt(0)("col") * dt(0)("row")
        End If

        dt = Nothing
        dt = sql.get_table("db_plate_location", "box_size", "=", "'" & Strings.Left(cbo_plate_boxsize.Text, 1) & "' AND plate_num IS NOT NULL AND plant=" & s_plant)
        If dt Is Nothing Then
            tb_plate_used.Text = 0
        Else
            tb_plate_used.Text = dt.Rows.Count
        End If
        tb_plate_remaining.Text = tb_plate_total.Text - tb_plate_used.Text
        s_loading = False
    End Sub

    Sub print_box_label(ByVal location As String, ByVal plate_number As String)
        MsgBox("A box label will be printed, please make sure you have a label in the printer." & vbCr & _
               "Curent printer is: " & get_printer_name())
        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
        xlApp.ScreenUpdating = False
        xlWorkBook = xlApp.Workbooks.Open(s_resource_path & "box labels.xlsx", [ReadOnly]:=True)
        xlWorkSheets = xlWorkBook.Worksheets
        xlWorkSheet = xlWorkSheets("Label")
        xlWorkSheet.Range("B2").Value = plate_number
        xlWorkSheet.Range("B3").Value = location

        xlApp.ScreenUpdating = True
        xlWorkSheet.PrintOutEx(1, 1, 1, False)
        xlWorkBook.Close(False)
        close_excel()
    End Sub
    Sub print_rts()
        If dgv_rts.RowCount = 0 Then Exit Sub
        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
        xlApp.ScreenUpdating = False
        xlWorkBook = xlApp.Workbooks.Open(s_resource_path & "rts slip.xls", [ReadOnly]:=True)
        xlWorkSheets = xlWorkBook.Worksheets
        xlWorkSheet = xlWorkSheets(1)
        xlWorkSheet.Unprotect()

        xlWorkSheet.Range("A6").Value = s_shift
        xlWorkSheet.Range("F6").Value = Now
        With dgv_rts_pending
            For i As Integer = 0 To .RowCount - 1
                xlWorkSheet.Range("A" & i + 9).Value = tb_prod_num.Text
                xlWorkSheet.Range("B" & i + 9).Value = .Rows(i).Cells(0).Value
                Select Case True
                    Case .Rows(i).Cells(0).Value = s_selected_job(0)("mat_num_in")
                        xlWorkSheet.Range("C" & i + 9).Value = s_selected_job(0)("mat_desc_in")
                    Case .Rows(i).Cells(0).Value = s_selected_job(0)("mat_num_in_1")
                        xlWorkSheet.Range("C" & i + 9).Value = s_selected_job(0)("mat_desc_in_1")
                    Case .Rows(i).Cells(0).Value = s_selected_job(0)("mat_num_in_2")
                        xlWorkSheet.Range("C" & i + 9).Value = s_selected_job(0)("mat_desc_in_2")
                End Select
                xlWorkSheet.Range("D" & i + 9).Value = .Rows(i).Cells(1).Value
                xlWorkSheet.Range("E" & i + 9).Value = .Rows(i).Cells(2).Value
                xlWorkSheet.Range("F" & i + 9).Value = s_user

            Next
        End With
        xlWorkSheet.PrintOutEx(1, 1, 1, False)
        xlWorkBook.Close(False)
        close_excel()
        MsgBox("The rts slip has been printed to " & get_printer_name())
    End Sub
    Sub print_reject()
        If dgv_reject.RowCount = 0 Then Exit Sub
        If cbo_reject_destination.Text = Nothing Then
            MsgBox("Select the s_department you are rejecting to", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Select Location")
            Exit Sub
        End If
        Dim super As String = Nothing
        Do While super = Nothing
            super = InputBox("Enter supervisor", "Supervisor")
        Loop
        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
        xlApp.ScreenUpdating = False
        xlWorkBook = xlApp.Workbooks.Open(s_resource_path & "reject slip.xls", [ReadOnly]:=True)
        xlWorkSheets = xlWorkBook.Worksheets
        xlWorkSheet = xlWorkSheets(1)
        xlWorkSheet.Unprotect()
        Dim multiple(cbo_mat_num_in.Items.Count - 1) As String
        cbo_mat_num_in.Items.CopyTo(multiple, 0)

        xlWorkSheet.Range("B2").Value = s_department
        xlWorkSheet.Range("D2").Value = tb_prod_num.Text
        xlWorkSheet.Range("B3").Value = cbo_reject_destination.Text
        xlWorkSheet.Range("B5").Value = String.Join(",", multiple)

        ReDim multiple(cbo_mat_num_in.Items.Count - 1)
        cbo_mat_desc_in.Items.CopyTo(multiple, 0)

        xlWorkSheet.Range("D5").Value = String.Join(",", multiple)
        xlWorkSheet.Range("C47").Value = s_user
        xlWorkSheet.Range("K47").Value = super
        xlWorkSheet.Range("P3").Value = tb_start_date.Text

        With dgv_reject
            Dim row As Integer
            For r = 0 To .RowCount - 1
                row = 11 + r
                xlWorkSheet.Range("B" & row).Value = .Rows(r).Cells("dgc_reject_prev_prod_num").Value
                xlWorkSheet.Range("D" & row).Value = .Rows(r).Cells("dgc_reject_item_no").Value
                xlWorkSheet.Range("F" & row).Value = .Rows(r).Cells("dgc_reject_quantity").Value
                xlWorkSheet.Range("H" & row).Value = .Rows(r).Cells("dgc_reject_rack").Value
                xlWorkSheet.Range("I" & row).Value = .Rows(r).Cells("dgc_reject_shift").Value
                xlWorkSheet.Range("J" & row).Value = .Rows(r).Cells("dgc_reject_code").Value
                xlWorkSheet.Range("L" & row).Value = .Rows(r).Cells("dgc_reject_reason").Value
                If Not IsNull(.Rows(r).Cells("dgc_reject_reason").Value) Then
                    If r < 10 Then
                        xlWorkSheet.Range("B" & row + 25).Value = .Rows(r).Cells("dgc_reject_reason").Value
                    Else
                        xlWorkSheet.Range("L" & row + 25).Value = .Rows(r).Cells("dgc_reject_reason").Value
                    End If
                End If
            Next

        End With
        xlApp.ScreenUpdating = True
        xlWorkSheet.PrintOutEx(1, 1, 1, False)
        For Each s As Microsoft.Office.Interop.Excel.Shape In xlWorkSheet.Shapes
            If s.Name = "WordArt 51" Then
                s.Visible = True
                Exit For
            End If
        Next
        xlWorkSheet.PrintOutEx(1, 1, 1, False)
        xlWorkBook.Close(False)
        close_excel()
        MsgBox("The reject slip has been printed to " & get_printer_name())
    End Sub

    Sub print_report()

        If dgv_production.RowCount = 0 Or Not s_production Then Exit Sub
        'add operators/helpers
        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
        xlApp.ScreenUpdating = False
        xlWorkBook = xlApp.Workbooks.Open(s_resource_path & "production report.xls", [ReadOnly]:=True)
        xlWorkSheets = xlWorkBook.Worksheets
        If s_extruder Then
            xlWorkSheet = xlWorkSheets("extruder")
        Else
            xlWorkSheet = xlWorkSheets(s_department)
        End If
        Dim m_info As materialInfo = get_material_info(False)
        xlWorkSheet.Unprotect()
        Dim multiple(cbo_mat_num_in.Items.Count - 1) As String
        cbo_mat_num_in.Items.CopyTo(multiple, 0)
        Dim sd As Date = tb_start_date.Text
        With dgv_production
            Select Case s_department
                Case "printing"

                    xlWorkSheet.Range("E3").Value = sd
                    xlWorkSheet.Range("E4").Value = tb_prod_num.Text
                    xlWorkSheet.Range("M4").Value = cbo_machine.Text
                    xlWorkSheet.Range("E5").Value = tb_mat_num_out.Text
                    xlWorkSheet.Range("T7").Value = String.Join(",", multiple)

                    ReDim multiple(cbo_mat_num_in.Items.Count - 1)
                    cbo_mat_desc_in.Items.CopyTo(multiple, 0)

                    xlWorkSheet.Range("H7").Value = String.Join(",", multiple)
                    xlWorkSheet.Range("Y7").Value = "REQ " & dgv_job_requirements.Rows(0).Cells(0).Value
                    xlWorkSheet.Range("AB7").Value = dgv_job_requirements.Rows(0).Cells(1).Value
                    xlWorkSheet.Range("AE7").Value = dgv_job_requirements.Rows(0).Cells(0).Value & " LEFT"
                    xlWorkSheet.Range("AH7").Value = dgv_job_requirements.Rows(0).Cells(3).Value
                    xlWorkSheet.Range("H9").Value = tb_label_desc.Text
                    xlWorkSheet.Range("AH9").Value = b_main_show_zaw.Text
                    xlWorkSheet.Range("AA11").Value = dgv_job_requirements.Rows(1).Cells(1).Value
                    xlWorkSheet.Range("AD11").Value = dgv_job_requirements.Rows(1).Cells(0).Value
                    xlWorkSheet.Range("AH11").Value = dgv_job_requirements.Rows(1).Cells(3).Value
                    xlWorkSheet.Range("F13").Value = tb_cyl_size.Text
                    xlWorkSheet.Range("G17").Value = tb_station_1.Text
                    xlWorkSheet.Range("J17").Value = tb_station_2.Text
                    xlWorkSheet.Range("M17").Value = tb_station_3.Text
                    xlWorkSheet.Range("P17").Value = tb_station_4.Text
                    xlWorkSheet.Range("S17").Value = tb_station_5.Text
                    xlWorkSheet.Range("V17").Value = tb_station_6.Text
                    xlWorkSheet.Range("Y17").Value = tb_station_7.Text
                    xlWorkSheet.Range("AB17").Value = tb_station_8.Text
                    xlWorkSheet.Range("AE17").Value = tb_station_9.Text
                    xlWorkSheet.Range("AH17").Value = tb_station_10.Text
                    For i As Integer = 0 To dgv_production.RowCount - 1
                        Dim done As Boolean = False
                        For r As Integer = 22 To 109
                            If xlWorkSheet.Range("AK" & r).Value = 1 Then
                                xlWorkSheet.Range("A" & r).Value = .Rows(i).Cells("dgc_production_num_in").Value
                                xlWorkSheet.Range("B" & r).Value = .Rows(i).Cells("dgc_production_item_id").Value
                                xlWorkSheet.Range("E" & r).Value = .Rows(i).Cells("dgc_production_kg_in").Value
                                xlWorkSheet.Range("G" & r).Value = .Rows(i).Cells("dgc_production_num_out").Value
                                xlWorkSheet.Range("H" & r).Value = .Rows(i).Cells("dgc_production_kg_out").Value
                                xlWorkSheet.Range("J" & r).Value = .Rows(i).Cells("dgc_production_km_out").Value
                                xlWorkSheet.Range("L" & r).Value = .Rows(i).Cells("dgc_production_date_").Value & " " & .Rows(i).Cells("dgc_production_user_").Value
                                If Not IsNull(.Rows(i).Cells("dgc_production_comments").Value) Then
                                    xlWorkSheet.Range("L" & r).Value = xlWorkSheet.Range("L" & r).Value & " - " & .Rows(i).Cells("dgc_production_comments").Value
                                End If
                                xlWorkSheet.Range("AC" & r).Value = .Rows(i).Cells("dgc_production_rts").Value
                                xlWorkSheet.Range("AE" & r).Value = .Rows(i).Cells("dgc_production_reject").Value
                                xlWorkSheet.Range("AG" & r).Value = .Rows(i).Cells("dgc_production_barcode").Value
                                xlWorkSheet.Range("AI" & r).Value = .Rows(i).Cells("dgc_production_rho_left").Value
                                xlWorkSheet.Range("AJ" & r).Value = .Rows(i).Cells("dgc_production_rho_right").Value
                                done = True
                                Exit For
                            End If
                        Next
                        If Not done Then
                            i = i - 1
                            xlApp.ScreenUpdating = True
                            xlWorkSheet.PrintOutEx(1, 2, 1, False)
                            xlApp.ScreenUpdating = False
                            xlWorkSheet.Range("A22:AJ53").Value = Nothing
                            xlWorkSheet.Range("A63:AJ109").Value = Nothing
                            xlWorkSheet.Range("U4").Value = xlWorkSheet.Range("U4").Value + 1
                        End If
                    Next
                    xlWorkSheet.PrintOutEx(1, 2, 1, False)
                Case "filmsline"
                    If s_extruder Then
                        Dim row As DataGridViewRow = dgv_summary.CurrentRow
                        Dim d As Date = tb_start_date.Text
                        Dim str As String = row.Cells("dgc_summary_start").Value.ToString.PadLeft(4, "0"c)
                        Dim t1 As New TimeSpan(Strings.Left(str, 2), Strings.Right(str, 2), 0)
                        Dim start As New DateTime(d.Year, d.Month, d.Day, t1.Hours, t1.Minutes, 0)
                        str = row.Cells("dgc_summary_finish").Value.ToString.PadLeft(4, "0"c)
                        t1 = New TimeSpan(Strings.Left(str, 2), Strings.Right(str, 2), 0)
                        Dim finish As New DateTime(d.Year, d.Month, d.Day, t1.Hours, t1.Minutes, 0)
                        If finish < start Then
                            finish = New DateTime(d.Year, d.Month, d.Day, t1.Hours, t1.Minutes, 0)
                        End If
                        Dim diff As TimeSpan = finish - start
                        Dim t_kg(), t_w() As String
                        Dim s() As String = Split(l_production_kg_in.Text, vbCr)
                        t_kg = Split(s(1), " ") ' kg=2
                        t_w = Split(s(0), " ") ' w=3
                        xlWorkSheet.Range("D3").Value = s_user_name
                        xlWorkSheet.Range("I3").Value = sd
                        xlWorkSheet.Range("O3").Value = tb_prod_num.Text
                        xlWorkSheet.Range("D4").Value = s_assistant
                        xlWorkSheet.Range("H5").Value = s_shift_name
                        xlWorkSheet.Range("I4").Value = s_shift
                        xlWorkSheet.Range("O5").Value = tb_mat_num_out.Text
                        xlWorkSheet.Range("D6").Value = m_info.formulation
                        xlWorkSheet.Range("I6").Value = row.Cells("dgc_summary_cyl_in").Value
                        xlWorkSheet.Range("L6").Value = row.Cells("dgc_summary_ink_in").Value
                        xlWorkSheet.Range("I7").Value = tb_cyl_size.Text
                        xlWorkSheet.Range("O7").Value = row.Cells("dgc_summary_start").Value
                        xlWorkSheet.Range("D8").Value = m_info.width
                        xlWorkSheet.Range("G8").Value = t_w(3)
                        xlWorkSheet.Range("O8").Value = row.Cells("dgc_summary_finish").Value
                        xlWorkSheet.Range("D9").Value = tb_roll_length.Text
                        xlWorkSheet.Range("G9").Value = t_kg(2)
                        xlWorkSheet.Range("O9").Value = CInt(diff.TotalMinutes)
                        xlWorkSheet.Range("D10").Value = cbo_job_info_type.Text
                        xlWorkSheet.Range("G10").Value = get_item_desity(tb_mat_desc_out.Text)
                        xlWorkSheet.Range("D11").Value = m_info.gauge

                        For Each row In .Rows
                            If row.Cells("dgc_production_user_").Value.ToString = s_user _
                                And row.Cells("dgc_production_date_").Value.ToString = tb_start_date.Text _
                                And row.Cells("dgc_production_shift").Value.ToString = s_shift Then
                                For r As Integer = 14 To 55
                                    If xlWorkSheet.Range("B" & r).Value = Nothing Then
                                        xlWorkSheet.Range("B" & r).Value = row.Cells("dgc_production_num_out").Value
                                        xlWorkSheet.Range("C" & r).Value = row.Cells("dgc_production_kg_out").Value
                                        xlWorkSheet.Range("D" & r).Value = row.Cells("dgc_production_km_out").Value
                                        xlWorkSheet.Range("E" & r).Value = row.Cells("dgc_production_width_").Value
                                        xlWorkSheet.Range("F" & r).Value = row.Cells("dgc_production_gauge").Value
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                        xlWorkSheet.PrintOutEx(1, 1, 1, False)

                    Else
                        'slitter
                    End If
            End Select
        End With
        close_excel()
        MsgBox("The report has been printed to " & get_printer_name())
    End Sub

    Sub print_summary()
        If dgv_summary.RowCount = 0 Then Exit Sub

        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
        xlApp.ScreenUpdating = False
        xlWorkBook = xlApp.Workbooks.Open(s_resource_path & "summary.xls", [ReadOnly]:=True)
        xlWorkSheets = xlWorkBook.Worksheets
        If s_extruder Then
            xlWorkSheet = xlWorkSheets("extruder")
        Else
            xlWorkSheet = xlWorkSheets(s_department)
        End If
        xlWorkSheet.Unprotect()
        Select Case s_department
            Case "printing"
                xlWorkSheet.Range("B2").Value = s_machine
                xlWorkSheet.Range("B4").Value = tb_start_date.Text
                '  Dim str = Strings.Right(sql.read("db_machines", "machine", "'" & s_machine & "'", "name"), 2) & "||" & sql.read("db_shifts", "id", s_shift, "name")
                xlWorkSheet.Range("E4").Value = cbo_machine.SelectedItem("machine") 'Strings.Right(sql.read("db_machines", "name", "=", "'" & s_machine & "'", "machine"), 2)
                xlWorkSheet.Range("G4").Value = s_shift
                xlWorkSheet.Range("H4").Value = s_shift_name 'sql.read("db_shifts", "id", "=", s_shift, "name")
            Case "mounting"
                xlWorkSheet.Range("B2").Value = s_shift_name
                xlWorkSheet.Range("H2").Value = tb_start_date.Text
        End Select

        With dgv_summary
            Dim row As Integer
            For r = 0 To .RowCount - 1
                Select Case s_department
                    Case "printing"
                        row = 9 + (r * 3)
                        xlWorkSheet.Range("B" & row).Value = s_user
                        xlWorkSheet.Range("C" & row).Value = .Rows(r).Cells("dgc_summary_prod_num").Value
                        xlWorkSheet.Range("H" & row).Value = .Rows(r).Cells("dgc_summary_start").Value
                        xlWorkSheet.Range("I" & row).Value = .Rows(r).Cells("dgc_summary_finish").Value
                        xlWorkSheet.Range("K" & row).Value = .Rows(r).Cells("dgc_summary_cyl_in").Value
                        xlWorkSheet.Range("L" & row).Value = .Rows(r).Cells("dgc_summary_ink_in").Value
                        xlWorkSheet.Range("M" & row).Value = .Rows(r).Cells("dgc_summary_setup").Value
                        xlWorkSheet.Range("M" & row + 2).Value = .Rows(r).Cells("dgc_summary_setup").ToolTipText
                        xlWorkSheet.Range("N" & row + 2).Value = .Rows(r).Cells("dgc_summary_mpm").ToolTipText
                        xlWorkSheet.Range("O" & row).Value = .Rows(r).Cells("dgc_summary_bs01").Value
                        xlWorkSheet.Range("P" & row).Value = .Rows(r).Cells("dgc_summary_cs01").Value
                        xlWorkSheet.Range("Q" & row).Value = .Rows(r).Cells("dgc_summary_es01").Value
                        xlWorkSheet.Range("R" & row).Value = .Rows(r).Cells("dgc_summary_issued").Value
                        xlWorkSheet.Range("S" & row).Value = .Rows(r).Cells("dgc_summary_used").Value
                        xlWorkSheet.Range("T" & row).Value = .Rows(r).Cells("dgc_summary_produced").Value
                        xlWorkSheet.Range("S" & row + 2).Value = .Rows(r).Cells("dgc_summary_km_out").Value
                        xlWorkSheet.Range("U" & row).Value = .Rows(r).Cells("dgc_summary_rts").Value
                        xlWorkSheet.Range("V" & row).Value = .Rows(r).Cells("dgc_summary_reject").Value
                        xlWorkSheet.Range("W" & row).Value = .Rows(r).Cells("dgc_summary_schd_out").Value
                    Case "mounting"
                        row = 4 + r
                        xlWorkSheet.Range("A" & row).Value = .Rows(r).Cells("dgc_summary_prod_num").Value
                        xlWorkSheet.Range("B" & row).Value = .Rows(r).Cells("dgc_summary_start").Value
                        xlWorkSheet.Range("C" & row).Value = .Rows(r).Cells("dgc_summary_finish").Value
                        xlWorkSheet.Range("D" & row).Value = .Rows(r).Cells("dgc_summary_cyl_in").Value
                        xlWorkSheet.Range("E" & row).Value = .Rows(r).Cells("dgc_summary_ink_in").Value
                        xlWorkSheet.Range("F" & row).Value = .Rows(r).Cells("dgc_summary_run").Value
                        xlWorkSheet.Range("G" & row).Value = .Rows(r).Cells("dgc_summary_mounter").Value
                        xlWorkSheet.Range("H" & row).Value = .Rows(r).Cells("dgc_summary_desc").Value
                    Case "filmsline"
                        Dim d As Date = tb_start_date.Text
                        xlWorkSheet.Range("B4").Value = d
                        xlWorkSheet.Range("D4").Value = s_shift
                        xlWorkSheet.Range("E4").Value = s_shift_name
                        xlWorkSheet.Range("J4").Value = s_user_name
                        xlWorkSheet.Range("K4").Value = s_assistant
                        s_selected_job = dt_schedules.Select("prod_num=" & .Rows(r).Cells("dgc_summary_prod_num").Value)
                        Dim dt As DataTable = sql.get_table("db_produced", "prod_num", "=", s_selected_job(0)("prod_num"))
                        Dim dr As DataRow() = dt.Select("user_='" & s_user & "' AND date_='" & tb_start_date.Text & "' AND shift='" & s_shift & "' AND kg_out IS NOT NULL")
                        Dim m_info As materialInfo = get_material_info(False)
                        xlWorkSheet.Range("C8").Value = m_info.formulation
                        xlWorkSheet.Range("J8").Value = dr.Length
                        xlWorkSheet.Range("C9").Value = m_info.width
                        xlWorkSheet.Range("C10").Value = m_info.gauge
                        xlWorkSheet.Range("J10").Value = dr.Length
                        xlWorkSheet.Range("C11").Value = s_selected_job(0)("prod_num")
                        xlWorkSheet.Range("C12").Value = s_selected_job(0)("mat_num_semi")
                        xlWorkSheet.Range("J12").Value = .Rows(r).Cells("dgc_summary_produced").Value
                        xlWorkSheet.Range("C13").Value = .Rows(r).Cells("dgc_summary_start").Value.ToString
                        xlWorkSheet.Range("C14").Value = .Rows(r).Cells("dgc_summary_finish").Value.ToString
                        xlWorkSheet.Range("C15").Value = .Rows(r).Cells("dgc_summary_run").Value
                        xlWorkSheet.Range("D19").Value = .Rows(r).Cells("dgc_summary_res_1").Value
                        xlWorkSheet.Range("D20").Value = .Rows(r).Cells("dgc_summary_res_2").Value
                        xlWorkSheet.Range("D21").Value = .Rows(r).Cells("dgc_summary_res_3").Value
                        xlWorkSheet.Range("D22").Value = .Rows(r).Cells("dgc_summary_res_4").Value
                        xlWorkSheet.Range("D41").Value = get_extruder_kg_hr(m_info)
                        xlWorkSheet.Range("D43").Value = ConvertTime(.Rows(r).Cells("dgc_summary_bs01").Value)
                        xlWorkSheet.Range("E43").Value = ConvertTime(.Rows(r).Cells("dgc_summary_cs01").Value)
                        xlWorkSheet.Range("G43").Value = ConvertTime(.Rows(r).Cells("dgc_summary_es01").Value)
                        xlWorkSheet.Range("K43").Value = ConvertTime(.Rows(r).Cells("dgc_summary_setup").Value)
                        xlWorkSheet.Range("K44").Value = ConvertTime(.Rows(r).Cells("dgc_summary_run").Value)
                        xlWorkSheet.Range("K45").Value = ConvertTime(.Rows(r).Cells("dgc_summary_bs01").Value)
                        dt = sql.get_table("db_comments", "prod_num", "=", s_selected_job(0)("prod_num") & " AND user_='" & s_user & "'")
                        Dim s As String = Nothing
                        If Not dt Is Nothing Then
                            For Each i As DataRow In dt.Rows
                                s = add_to_string(s, i("text"), vbCrLf, True)
                            Next
                        End If
                        xlWorkSheet.Range("A51").Value = s

                        If row < .RowCount - 1 Then
                            xlWorkSheet.PrintOutEx(1, 1, 1, False)
                        End If
                End Select

            Next
            s_selected_job = dt_schedules.Select("prod_num=" & .CurrentRow.Cells("dgc_summary_prod_num").Value)
        End With
        xlApp.ScreenUpdating = True
        xlWorkSheet.PrintOutEx(1, 1, 1, False)
        xlWorkBook.Close(False)
        MsgBox("Sheet Printed to " & get_printer_name() & vbCr & "If there is nothing at the printer then there is something wrong with the printer.")
        close_excel()

    End Sub

    Sub set_colour(sender As Object, e As EventArgs) Handles tb_station_10.Click, tb_station_9.Click, tb_station_8.Click, tb_station_7.Click, tb_station_6.Click, tb_station_5.Click, tb_station_4.Click, tb_station_3.Click, tb_station_2.Click, tb_station_1.Click
        Dim tb As TextBox = CType(sender, TextBox)
        If tb.Text = Nothing Then
            If lb_inks.SelectedIndex = -1 Then Exit Sub
            tb.Text = lb_inks.Items(lb_inks.SelectedIndex)
            lb_inks.Items.RemoveAt(lb_inks.SelectedIndex)
            If lb_inks.Items.Count > 0 Then lb_inks.SelectedIndex = 0
        Else
            lb_inks.Items.Add(tb.Text)
            lb_inks.SelectedIndex = lb_inks.Items.Count - 1
            tb.Clear()
        End If
        Dim s As String = Nothing
        If Not tb_station_1.Text = Nothing Then s = "1:" & tb_station_1.Text
        If Not tb_station_2.Text = Nothing Then s = add_to_string(s, "2:" & tb_station_2.Text)
        If Not tb_station_3.Text = Nothing Then s = add_to_string(s, "3:" & tb_station_3.Text)
        If Not tb_station_4.Text = Nothing Then s = add_to_string(s, "4:" & tb_station_4.Text)
        If Not tb_station_5.Text = Nothing Then s = add_to_string(s, "5:" & tb_station_5.Text)
        If Not tb_station_6.Text = Nothing Then s = add_to_string(s, "6:" & tb_station_6.Text)
        If Not tb_station_7.Text = Nothing Then s = add_to_string(s, "7:" & tb_station_7.Text)
        If Not tb_station_8.Text = Nothing Then s = add_to_string(s, "8:" & tb_station_8.Text)
        If Not tb_station_9.Text = Nothing Then s = add_to_string(s, "9:" & tb_station_9.Text)
        If Not tb_station_10.Text = Nothing Then s = add_to_string(s, "10:" & tb_station_10.Text)
        sql.update("db_schedule", "set_inks", sql.convert_value(s), "prod_num", "=", s_selected_job(0)("prod_num") & "AND plant=" & s_plant)
        s_selected_job(0)("set_inks") = s
        lb_inks.Focus()
    End Sub
    Sub set_dgv(ByVal dgv_type As String)
        Dim dept As String = s_department
        If cbo_machine.SelectedItem Is Nothing Then Exit Sub
        If cbo_machine.SelectedItem("extruder") Then dept = "extruder"
        Dim dt As DataTable = sql.get_table("settings_departments", "plant", "=", s_plant & " AND item='dgv_col_" & dgv_type & "' AND dept='" & dept & "'")
        Dim dgv As DataGridView
        If dgv_type = "Summary" Then
            dgv = dgv_summary
        Else
            dgv = dgv_production
        End If
        With dgv
            Dim w As Integer = 0
            If Not dt Is Nothing Then
                Dim str() As String = Split(dt(0)("value"), ",")
                Dim s() As String = Nothing
                For i As Integer = 0 To str.Length - 1
                    s = Split(str(i), ":")
                    ' dgv_department_columns.Columns.Add(s(1), s(2))
                    If s(3) Then
                        .Columns(s(1)).DisplayIndex = s(0)
                    End If
                    .Columns(s(1)).HeaderText = s(2)
                    .Columns(s(1)).Visible = s(3)
                    .Columns(s(1)).Width = s(4)
                    .Columns(s(1)).DefaultCellStyle.Format = s(5)
                    If s(5).Contains("N") Then
                        .Columns(s(1)).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End If
                Next

            Else
                For c As Integer = 0 To .Columns.Count - 1
                    .Columns(c).Visible = True
                Next
                With frm_settings_admin
                    .StartPosition = FormStartPosition.CenterParent
                    .tb_password.Text = "lmuser1234"
                    .TabControl1.SelectedIndex = 4
                    .TabControl1.Enabled = True
                    .ShowDialog()
                End With

            End If
        End With
    End Sub
    Sub set_department()
        s_loading = True
        If s_old_department = s_department Then
            If s_production Then
                set_dgv("Summary")
                set_dgv("Report")
            ElseIf s_department = "mounting" Then
                set_dgv("Summary")
            Else
                If acsis.running(0) Then
                    s_user_name = Nothing
                    s_user = Nothing
                    s_pass = Nothing
                    w_sessions(0).process.Kill()
                End If
            End If

        Else
            If s_production Then
                set_dgv("Summary")
                set_dgv("Report")
            End If
            s_old_department = s_department
            tc_main.TabPages.Remove(tp_summary)
            tc_main.TabPages.Remove(tp_production)
            tc_main.TabPages.Remove(tp_downtime)
            tc_main.TabPages.Remove(tp_rts)
            tc_main.TabPages.Remove(tp_reject)
            tc_main.TabPages.Remove(tp_setup)
            tc_main.TabPages.Remove(tp_plates)
            tc_main.TabPages.Remove(tp_scheduling)

            Select Case s_department
                Case "printing"
                    tc_main.TabPages.Insert(0, tp_summary)
                    tc_main.TabPages.Insert(1, tp_production)
                    tc_main.TabPages.Insert(2, tp_downtime)
                    tc_main.TabPages.Insert(3, tp_rts)
                    tc_main.TabPages.Insert(4, tp_reject)
                    tc_main.TabPages.Insert(5, tp_setup)
                    tc_main.TabPages.Insert(6, tp_plates)
                Case "filmsline"
                    tc_main.TabPages.Insert(0, tp_summary)
                    tc_main.TabPages.Insert(1, tp_production)
                    tc_main.TabPages.Insert(2, tp_downtime)
                    tc_main.TabPages.Insert(3, tp_rts)
                    tc_main.TabPages.Insert(4, tp_reject)
                Case "graphics"
                    tc_main.TabPages.Insert(0, tp_plates)
                Case "planning"
                    tc_main.TabPages.Insert(0, tp_scheduling)
                Case "bagplant"
                    tc_main.TabPages.Insert(0, tp_summary)
                    tc_main.TabPages.Insert(1, tp_production)
                    tc_main.TabPages.Insert(2, tp_downtime)
                    tc_main.TabPages.Insert(3, tp_rts)
                    tc_main.TabPages.Insert(4, tp_reject)
                Case "mounting"
                    tc_main.TabPages.Insert(0, tp_summary)
                    tc_main.TabPages.Insert(1, tp_plates)
            End Select
        End If
        '3,42 location
        Dim p As New Point(3, 42)
        With tc_main
            If s_department = "mounting" Then
            End If
        End With
        If frm_bag_checks.Visible Then frm_bag_checks.Close()
        tb_roll_length.Visible = False
        cbo_mat_desc_in.Visible = True
        cbo_mat_num_in.Visible = True
        tb_cyl_size.Visible = False
        tb_no_up.Visible = False
        gb_printing.Visible = True
        cb_production_ply_separator.Visible = False
        cb_report_use_estimate.Visible = False

        l_job_info_cyl_size.Visible = False
        l_material_in.Visible = True
        l_items_out.Visible = False
        l_job_info_num_up.Visible = False

        b_summary_multiple_job_add.Visible = False
        b_report_merge.Visible = False
        b_report_enter_item.Visible = False
        b_report_part_roll.Visible = False
        b_main_mat_in_add.Visible = True
        b_main_show_zaw.Visible = True
        gb_multiple_job.Visible = False
        gb_item_info.Visible = False
        cbo_job_info_type.Visible = False
        tb_cyl_size.ReadOnly = True
        PrepareJobToolStripMenuItem.Text = "&Prepare or Add Job"
        With nud_items_out
            .Visible = False
            .Minimum = 1
            .Increment = 1
            .Value = 1
            If Not s_selected_job Is Nothing Then

                If Not IsNull(s_selected_job(0)("items_out")) Then
                    .Maximum = s_selected_job(0)("items_out") + 10
                    .Value = s_selected_job(0)("items_out")
                Else
                    .Maximum = 10
                End If
            End If
        End With

        ''678
        'Dim i1 As Integer = 0
        'For i As Integer = 0 To dgv_production.ColumnCount - 1
        '    If dgv_production.Columns(i).Visible And Not dgv_production.Columns(i).Name = "comments" Then
        '        i1 = i1 + dgv_production.Columns(i).Width
        '    End If
        'Next
        'If Not i1 = 401 Then
        '    dgv_production.Columns("comments").Width = 678 - i1
        '    dgv_production.PerformLayout()
        'End If
        Select Case s_department
            Case "printing"
                gb_item_info.Text = "Roll Info"
                gb_item_info.Visible = True
                b_report_enter_item.Text = "Enter Rolls"
                b_report_part_roll.Visible = True
                b_report_merge.Visible = True
                b_report_enter_item.Visible = True
                l_job_info_cyl_size.Visible = True
                l_job_info_num_up.Visible = True
                tb_cyl_size.Visible = True
                tb_no_up.Visible = True
                l_items_out.Text = "Roll(s) Out"
                l_items_out.Visible = True
                l_job_info_cyl_size.Text = "Cyl. Size"
                l_job_info_num_up.Text = "No. Up"
                With nud_items_out
                    .Visible = True
                    .Maximum = 10
                    .Minimum = 1
                    .Increment = 1
                    .Value = 1
                End With
            Case "bagplant"
                gb_item_info.Visible = True
                b_report_enter_item.Text = "Roll In"
                b_report_enter_item.Visible = True
                b_report_part_roll.Visible = True
                b_report_part_roll.Text = "Carton Out"
                l_items_out.Text = "Bags/Carton"
                l_items_out.Visible = True
                With nud_items_out
                    .Visible = True
                    .Maximum = 2000
                    .Minimum = 100
                    .Increment = 100
                    .Value = 1000
                End With
                dt_run_data = sql.get_run_data()
                If Not frm_bag_checks.Visible Then frm_bag_checks.Show()
            Case "mounting"
                PrepareJobToolStripMenuItem.Text = "&Print Placement Printing"
            Case "planning"
                scheduling_setting_load()
            Case "filmsline"
                b_report_enter_item.Text = "Item In"
                b_report_part_roll.Text = "Item(s) Out"
                b_report_part_roll.Visible = True
                tb_roll_length.Visible = True
                cb_report_use_estimate.Visible = True
                If s_extruder Then
                    tb_cyl_size.Visible = True
                    tb_cyl_size.ReadOnly = False
                    l_job_info_cyl_size.Text = "Shoe (inch)"
                    l_material_in.Visible = False
                    l_job_info_num_up.Text = "Type"
                    l_job_info_num_up.Visible = True
                    l_job_info_cyl_size.Visible = True
                    cbo_mat_desc_in.Visible = False
                    cbo_mat_num_in.Visible = False
                    b_main_mat_in_add.Visible = False
                    b_main_show_zaw.Visible = False
                    l_production_unused.Visible = False
                    l_production_kg_in.Visible = True
                    cb_production_ply_separator.Visible = True
                    cbo_job_info_type.Visible = True
                Else
                    l_items_out.Visible = True
                    l_items_out.Text = "Item(s) Out"
                    gb_item_info.Text = "Roll Info"
                    gb_item_info.Visible = True
                    b_summary_multiple_job_add.Visible = True
                    b_report_merge.Visible = True
                    b_report_enter_item.Visible = True
                    With nud_items_out
                        .Visible = True
                        .Minimum = 1
                        .Increment = 1
                        .Value = 1
                    End With
                End If
            Case "graphics"
                populate_plates()
                PrepareJobToolStripMenuItem.Text = "&Print Placement Printing"
                gb_printing.Visible = False

        End Select
        s_loading = False
    End Sub

    Sub start_up()
        s_loading = True
        Dim exeDir As New IO.FileInfo(Reflection.Assembly.GetExecutingAssembly.FullName)
        s_resource_path = exeDir.DirectoryName & "\Resources\"
        's_resource_path = s_resource_path.Substring(0, s_resource_path.Length - 39) & "\Resources\"
        show_status("Checking Database Connections (main)...")
        sql.connect("lean", True)
        dt_popups = sql.get_table("acsis_popups", "plant", "=", s_plant)
        dt_department = sql.get_table("db_departments", "plant", "=", s_plant) ' & " AND production=1")

        If s_version Is Nothing Then
            If IO.File.Exists(get_file_path("symple")) Then
                s_version = FileVersionInfo.GetVersionInfo(get_file_path("symple")).FileVersion
            Else
                s_version = "4.503"
            End If

        End If


        dt_acsis_options = sql.get_table("acsis_options", "plant", "=", s_plant & " AND version='" & s_version & "'")
        With lb_scheduling_print_departments
            .Items.Clear()
            For Each item In dt_department.Rows
                If item("production") Then
                    .Items.Add(LCase(item("dept")).ToString)
                End If
            Next
            .SelectedIndex = 0
        End With


        Dim dr() As DataRow = dt_acsis_options.Select("item='navigation'")
        Dim str() As String = Split(dr(0)("val"), ",")
        ' Dim str() As String = Split(sql.read("acsis_options", "item", "=", "'navigation' AND plant=" & s_plant & " AND version='" & s_version & "'", "val"), ",")
        Dim do_nav As Boolean = False
        For Each s In str

            dt_navigation = sql.get_table("acsis_navigation", "plant", "=", s_plant & " AND target='" & s.ToString & "' AND version='" & s_version & "'")
            If dt_navigation Is Nothing Then
                do_nav = True
                Exit For
            End If
        Next
        ' Dim login As String = sql.read("acsis_options", "plant", "=", s_plant & " AND item='login_id'" & " AND version='" & s_version & "'", "val")
        dr = dt_acsis_options.Select("item='login_id'")
        Dim login As String = dr(0)("val")
        If IsNull(login) Then do_nav = True

        If do_nav Then
            frm_status.Close()
            acsis.get_controls(0, True, False, True)
            's_version = FileVersionInfo.GetVersionInfo(w_sessions(0).process.MainModule.FileName).FileVersion
            dt_navigation = sql.get_table("acsis_navigation", "plant", "=", s_plant & " AND version='" & s_version & "'")
            With frm_settings_admin
                .StartPosition = FormStartPosition.CenterParent
                .ShowDialog()
                .tb_password.Text = "lmuser1234"
                .TabControl1.SelectedIndex = 3
                .TabControl1.Enabled = True
            End With

            Exit Sub
        Else
            dt_navigation = sql.get_table("acsis_navigation", "plant", "=", s_plant & " AND version='" & s_version & "'")
        End If
        acsis.get_controls(0, True, False, True)

        For i As Integer = 0 To 1
            With w_sessions(i)
                Try
                    If .process.Responding Then
                        b_main_logoff.Enabled = True
                        If .user = Nothing Then
                            acsis.get_logon_details(i)
                        End If
                        If s_version = Nothing Then
                            s_version = FileVersionInfo.GetVersionInfo(.process.MainModule.FileName).FileVersion
                        End If
                    End If
                Catch ex As Exception
                    'symple not open 
                    b_main_logoff.Enabled = False
                End Try

            End With
        Next
        show_status("Checking Database Connections (Sym-PLE)...")

        sql.connect("symple", True)
        tb_scale_tare.Text = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Scale", "Tare", Nothing)

        If s_sql_access Then
            Dim s As String = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Export", "Time", Nothing)
            Dim dt As Date = Now
            If s = Nothing Then
                WriteIni(Environ("USERPROFILE") & "\settings.ini", "Export", "Time", Now)
            Else
                dt = s
            End If
            Dim m As Integer = DateDiff(DateInterval.Minute, dt, Now)
            If m > 30 Then
                WriteIni(Environ("USERPROFILE") & "\settings.ini", "Export", "Time", Now)

                sql.export_tables()
            End If

        End If
        tb_start_date.Text = Format(Now, "dd/MM/yyyy").ToString

        dt_boxes = sql.get_table("plate_options", "plant", "=", s_plant)

        dt_schedules = sql.get_table("db_schedule", "plant", "=", s_plant)

        dt_shifts = sql.get_table("db_shifts", "plant", "=", s_plant)

        With cbo_plate_boxsize
            .Items.Clear()
            For i As Integer = 0 To dt_boxes.Rows.Count - 1
                .Items.Add(dt_boxes(i)("box_size"))
            Next
            .SelectedIndex = 0
        End With

        dt_machines = sql.get_table("db_machines", "plant", "=", s_plant) '& " AND dept='" & s_department & "' AND enabled=1")
        s_pass_plate = sql.read("settings_general", "item", "=", "'general_password_plate' AND plant=" & s_plant, "val")

        dt_downtime_reasons = sql.get_table("settings_downtime", "plant", "=", s_plant & " AND dept='" & s_department & "'")
        str = Split(sql.read("settings_general", "item", "=", "'reject_locations' AND plant=" & s_plant, "val"), ",")
        With cbo_reject_destination
            .Items.Clear()
            For i As Integer = 0 To str.Length - 1
                .Items.Add(str(i))
            Next
        End With

        'acsis.window_name = sql.read("acsis_session", "plant", "=", s_plant, "main")
        acsis.logon_name = sql.read("acsis_session", "plant", "=", s_plant, "logon")
        dgv_downtime_bs01.Columns(0).ValueType = GetType(Decimal)
        s_loading = True
        cbo_machine.DataSource = dr_to_dt(dt_machines.Select("dept='" & s_department & "' AND enabled=1"))
        cbo_machine.DisplayMember = "name"
        s_loading = False
        With dgv_rts
            .AllowUserToOrderColumns = False
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With

        With dgv_rts_pending
            .AllowUserToOrderColumns = False
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        With dgv_summary
            .AllowUserToOrderColumns = False

            'For i As Integer = 0 To .ColumnCount - 1
            '    .Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            '    .Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            '    If i > 6 Then .Columns(i).ValueType = GetType(Decimal)
            'Next
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With

        With dgv_plates
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        dr = dt_department.Select("dept='" & s_department & "'")
        s_production = dr(0)("production")
        'select_cbo_item(cbo_machine, s_machine, "name")
        's_extruder = cbo_machine.SelectedItem("extruder")
        set_department()

        With dgv_production
            .AllowUserToOrderColumns = False
            '.Columns("dgc_production_kg_in").ValueType = GetType(Decimal)
            '.Columns("dgc_production_kg_out").ValueType = GetType(Decimal)
            '.Columns("dgc_production_km_out").ValueType = GetType(Decimal)
            '.Columns("dgc_production_rts").ValueType = GetType(Decimal)
            '.Columns("dgc_production_reject").ValueType = GetType(Decimal)
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With

        dt_uom = sql.get_table_full("db_uom", True)
        dt_num_up = sql.get_table_full("db_num_up", True)
        If s_symple_access Then
            dt_labels = sql.get_table("DP_Print_Templates", "Label_Location", "=", "'Full' AND Label_Type='Item' ORDER BY List_Order", "symple")
            cbo_scheduling_labels.DataSource = dt_labels
            cbo_scheduling_labels.DisplayMember = "Template_Code"
        End If
        scheduling_setting_load()
        If Not s_restored Then
            If acsis.running(0) Then acsis.get_logon_details(0)
            Select Case s_department
                Case "graphics", "planning"
                    s_user = "non_prod"
            End Select

            If Not s_user = Nothing Then
                'Dim restore As Object
                Dim dt As DataTable = sql.get_table("db_restore", "user_", "=", "'" & s_user & "' AND restore_=1 AND NOT shift IS NULL")
                'restore = sql.read("db_restore", "user_", "=", "'" & s_user & "'", "restore_")
                If Not IsNull(dt) Then
                    If CBool(dt(0)("restore_")) = True Then
                        If s_user = "non_prod" Then
                            s_restored = True
                        Else
                            'Dim m As String = sql.read("db_restore", "user_", "=", "'" & s_user & "'", "machine")
                            'Dim d As String = sql.read("db_restore", "user_", "=", "'" & s_user & "'", "date_")
                            's_shift = sql.read("db_restore", "user_", "=", "'" & s_user & "'", "shift")
                            s_shift = dt(0)("shift")

                            dt_restore = sql.get_table("db_summary", "user_", "=", "'" & s_user & "' AND date_='" & dt(0)("date_") & _
                                                      "' AND machine='" & dt(0)("machine") & "' AND shift='" & s_shift & "'")

                            frm_restore.restore_it(dt(0)("tab_page")) 'dt(0)("date_"))
                            SetForegroundWindow(Me.Handle)
                            s_restored = True
                        End If
                        sql.update("db_restore", "restore_", 0, "user_", "=", "'" & s_user & "'")

                    End If
                End If
            End If
        End If
        If Not s_restored Then
            With frm_startup
                .StartPosition = FormStartPosition.CenterParent
                .ShowDialog()
            End With
        End If
        If s_user = "non_prod" Then s_user = Nothing
        s_loading = False
        Timer1.Enabled = True
        s_start_up = True
        SetForegroundWindow(Me.Handle)
    End Sub

    Sub update_files()
        Dim folder As New System.IO.DirectoryInfo(".\"), files As System.IO.FileInfo() = folder.GetFiles(), _
            dt As New DataTable, code(), table As String, del As Boolean
        code = Nothing
        For Each File In files
            If File.Extension = ".dtb" Then
                If Not IsNothing(dt) Then dt.Clear()
                table = Replace(Replace(File.Name, ".dtb", Nothing), Environ("COMPUTERNAME") & "-", Nothing)
                dt = sql.get_table_full(table, False)
                If Not IsNothing(dt) Then
                    Dim dr As DataRow = dt(0)
                    If dr.Table.Columns.Contains("access") Then
                        Dim result() As DataRow = dt.Select("access=0")
                        If result.Length > 0 Then
                            Dim value, column, str As String
                            For Each r In result
                                del = False

                                If dr.Table.Columns.Contains("date_") Then
                                    If code Is Nothing Then
                                        ReDim Preserve code(0)
                                        code(0) = r.Item("user_") & r.Item("date_")
                                        del = True
                                    Else
                                        If code.ToList.IndexOf(r.Item("user_").ToString & r.Item("date_").ToString) = -1 Then
                                            ReDim Preserve code(code.Length)
                                            code(0) = r.Item("user_").ToString & r.Item("date_").ToString
                                            del = True
                                        End If
                                    End If
                                Else
                                    'schedule

                                End If

                                value = Nothing
                                column = Nothing

                                For c = 0 To dt.Columns.Count - 1
                                    str = Nothing
                                    If Not dt.Columns(c).ColumnName = "access" Then
                                        If IsDBNull(r.Item(dt.Columns(c).ColumnName)) Then
                                            str = "NULL"
                                        Else
                                            str = r.Item(dt.Columns(c).ColumnName)
                                            If Not IsNumeric(str) Then str = "'" & str & "'"
                                            If str = "''" Then str = "NULL"
                                        End If

                                        If column = Nothing Then
                                            column = dt.Columns(c).ColumnName
                                            value = str
                                        Else
                                            column = column & "," & dt.Columns(c).ColumnName
                                            value = value & "," & str
                                        End If

                                    End If

                                Next
                                column = column & ",access"
                                value = value & ",1"
                                Select Case table
                                    Case "db_schedule"
                                        value = value
                                    Case Else
                                        If del Then sql.delete(table, "user_", "=", "'" & r.Item("user_") & _
                                            "' AND date_='" & r.Item("date_") & "' AND access=0")

                                        sql.insert(table, column, value)
                                End Select

                            Next
                        End If
                    End If
                End If

            End If
        Next
        ' sql.export_tables()
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles tsm_production_delete_row.Click
        With dgv_production
            If Not .Rows(.CurrentRow.Index).Cells("dgc_production_num_in").Value.ToString = "-" Then
                Dim i_info As item_info = get_item_info(.CurrentRow.Index, True)
                For r As Integer = .RowCount - 1 To i_info.row_start Step -1
                    .Rows.RemoveAt(r)
                Next
            Else
                .Rows.RemoveAt(.CurrentRow.Index)
            End If

        End With
        export_report()
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles tsm_summary_add_job.Click
        If s_department = "mounting" And lb_personel.Items.Count = 0 Then
            MsgBox("You need to add a mounter first", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "No mounter added")
        Else
            With frm_schedules
                If Not .Visible Then

                    .StartPosition = FormStartPosition.CenterParent
                    .ShowDialog()
                End If
            End With
        End If
    End Sub

    Private Sub AboveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsm_production_insert_row_above.Click
        insert_row(dgv_production.CurrentRow.Index - 1, True)
    End Sub

    Private Sub BelowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsm_production_insert_row_below.Click
        insert_row(dgv_production.CurrentRow.Index, True)
    End Sub

    Private Sub ToolStripMenuItem1_DropDownOpening(sender As Object, e As EventArgs) Handles tsm_production_insert_row.DropDownOpening
        DirectCast(DirectCast(sender, ToolStripDropDownItem).DropDown, ToolStripDropDownMenu).ShowImageMargin = False
    End Sub

    Private Sub DataGridView2_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv_summary.CellMouseClick
        tsm_summary_setup.Visible = False
        tsm_summary_run.Visible = False
        tsm_summary_edit_item.Visible = False
        With dgv_summary
            If .CurrentRow.Cells("dgc_summary_mounter").Selected Then .ClearSelection()
            If e.ColumnIndex > -1 And e.RowIndex > -1 Then
                's_selected_job = dt_schedule.Select("prod_num=" & .CurrentRow.Cells(0).Value)
                If e.Button = Windows.Forms.MouseButtons.Right Then
                    .ClearSelection()
                    Dim cell As DataGridViewCell = .Rows(e.RowIndex).Cells(e.ColumnIndex)
                    .Rows(e.RowIndex).Selected = True
                    .CurrentCell = cell
                    If .CurrentCell.ReadOnly And e.ColumnIndex > 0 Then tsm_summary_edit_item.Visible = True

                    tsm_summary_remove_job.Enabled = True

                    If Not IsNull(.Rows(e.RowIndex).Cells("dgc_summary_setup").Value) Then
                        If .Rows(e.RowIndex).Cells("dgc_summary_setup").Value > 0 Then
                            tsm_summary_setup.Visible = True
                            tsm_summary_setup.Text = "Setup: " & FormatPercent(.Rows(e.RowIndex).Cells("dgc_summary_st_setup").Value / .Rows(e.RowIndex).Cells("dgc_summary_setup").Value, 0)
                        End If
                    End If
                    If Not IsNull(.Rows(e.RowIndex).Cells("dgc_summary_st_speed").Value) Then
                        If .Rows(e.RowIndex).Cells("dgc_summary_st_speed").Value > 0 Then
                            tsm_summary_run.Visible = True
                            tsm_summary_run.Text = "Run: " & FormatPercent(.Rows(e.RowIndex).Cells("dgc_summary_mpm").Value / .Rows(e.RowIndex).Cells("dgc_summary_st_speed").Value, 0)
                        End If
                    End If

                    cms_summary.Show(Cursor.Position)
                Else

                    If Not s_selected_job(0)("prod_num") = .CurrentRow.Cells("dgc_summary_prod_num").Value Then
                        s_selected_job = dt_schedules.Select("prod_num=" & .CurrentRow.Cells("dgc_summary_prod_num").Value)

                        display_job()
                    End If
                End If

            End If
        End With
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles tsm_summary_remove_job.Click
        With dgv_summary
            If .RowCount = 0 Then Exit Sub
            s_loading = True
            sql.delete("db_summary", "prod_num", "=", .CurrentRow.Cells(0).Value & " AND date_='" & tb_start_date.Text & "' AND user_='" & s_user & "' AND shift='" & s_shift & "' AND machine='" & s_machine & "'")
            .Rows.RemoveAt(.CurrentCell.RowIndex)
            export_summary()
            If .RowCount > 0 Then
                s_selected_job = dt_schedules.Select("prod_num='" & .CurrentRow.Cells(0).Value & "'")
                display_job()
            End If

            s_loading = False
        End With

    End Sub

    Private Sub DataGridView2_MouseClick(sender As Object, e As MouseEventArgs) Handles dgv_summary.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            tsm_summary_remove_job.Enabled = False
            cms_summary.Show(Cursor.Position)
        End If
    End Sub


    Private Sub DataGridView2_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgv_summary.DataError
        e.ThrowException = False
        e.Cancel = True
    End Sub

    Private Sub AreYouSureToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsm_production_clear_form.Click
        dgv_production.Rows.Clear()
        export_report()
    End Sub

    Private Sub DataGridView4_Scroll(sender As Object, e As ScrollEventArgs) Handles dgv_production.Scroll
        l_production_top_row.Text = "Top Row: " & dgv_production.FirstDisplayedCell.RowIndex + 1
    End Sub


    Private Sub DataGridView4_MouseLeave(sender As Object, e As EventArgs) Handles dgv_production.MouseLeave
        l_production_blank_row.Visible = False
    End Sub

    Private Sub dgv_production_MouseMove(sender As Object, e As MouseEventArgs) Handles dgv_production.MouseMove

        Dim bottom As Boolean = False, middle As Boolean = False, top As Boolean = False
        Dim v As Decimal = 0, conversion As Decimal = 0, c1 As Integer = -1, show_conversion As Boolean = False, km As Boolean = False, f As Decimal = 0
        Dim dp As Integer = 1

        If e.Button = Windows.Forms.MouseButtons.Left Or ModifierKeys = Keys.Control Then
            With dgv_production
                If .CurrentCell Is Nothing Then Exit Sub
                Dim str() As String = Split(.Columns(.CurrentCell.ColumnIndex).HeaderText, " ")
                If .SelectedCells.Count > 1 Then
                    For r As Integer = 0 To .RowCount - 1
                        For c As Integer = 0 To .Columns.Count - 1
                            If .Rows(r).Cells(c).Selected Then
                                If c1 = -1 And .Columns(c).Name.Contains("kg") Or c1 = -1 And .Columns(c).Name.Contains("km") Then
                                    c1 = c
                                    show_conversion = True
                                ElseIf Not c1 = c Then
                                    show_conversion = False
                                    conversion = 0
                                End If
                                If IsNumeric(.Rows(r).Cells(c).Value) And Not IsNull(.Rows(r).Cells(c).Value) Then
                                    v = v + .Rows(r).Cells(c).Value
                                    If show_conversion Then
                                        If s_department = "bagplant" Then
                                            conversion = v
                                        Else
                                            Select Case .Columns(c).Name
                                                Case "dgc_production_kg_out", "dgc_production_kg_in"
                                                    dp = Strings.Right(.Columns("dgc_production_km_out").DefaultCellStyle.Format, 1)
                                                    f = get_conversion_factor(True, tb_mat_num_out.Text, False)
                                                    If f = 0 Then
                                                        conversion = FormatNumber(get_roll_length(v, s_selected_job(0)("mat_desc_semi")) / 1000, dp, , , TriState.False)
                                                    Else
                                                        conversion = FormatNumber(v * f, dp, , , TriState.False)
                                                    End If
                                                    km = True
                                                Case "dgc_production_km_out"
                                                    dp = Strings.Right(.Columns("dgc_production_kg_out").DefaultCellStyle.Format, 1)
                                                    f = get_conversion_factor(True, tb_mat_num_out.Text, False)
                                                    If f = 0 Then
                                                        conversion = FormatNumber(get_roll_length(v, s_selected_job(0)("mat_desc_semi")) / 1000, dp, , , TriState.False)
                                                    Else
                                                        conversion = FormatNumber(v * f, dp, , , TriState.False)
                                                    End If
                                                    km = False
                                            End Select
                                        End If

                                    End If

                                End If
                            End If
                        Next
                    Next
                    l_production_selected_totals.Visible = True
                    If show_conversion Then
                        dp = Strings.Right(.Columns(.CurrentCell.ColumnIndex).DefaultCellStyle.Format, 1)
                        If km Then
                            str = Split(.Columns("dgc_production_km_out").HeaderText, " ")
                            l_production_selected_totals.Text = FormatNumber(v, dp, , , False) & " KG (estimated " & conversion & " " & str(0) & ")"
                        Else
                            If s_department = "bagplant" Then
                                l_production_selected_totals.Text = FormatNumber(v, dp, , , False) & " " & str(0)
                            Else
                                l_production_selected_totals.Text = FormatNumber(v, dp, , , False) & " " & str(0) & " (estimated " & conversion & " KG)"
                            End If
                        End If
                    Else
                        l_production_selected_totals.Text = v

                    End If
                    l_production_selected_totals.Location = New Point(e.X + 10, e.Y + 10)
                    '  End If
                Else
                    l_production_selected_totals.Visible = False
                End If
            End With
        Else
            If Not ModifierKeys = Keys.Control Then l_production_selected_totals.Visible = False
        End If

        With dgv_production
            If .RowCount = 0 Then Exit Sub
            'If .FirstDisplayedCell.RowIndex = 0 Then Exit Sub

            For r As Integer = 0 To .FirstDisplayedCell.RowIndex
                If can_delete(r, True) Then
                    If r < .FirstDisplayedCell.RowIndex Then
                        l_production_blank_row.Visible = True
                        Exit Sub
                    End If
                End If
            Next
        End With
        l_production_blank_row.Visible = False
        'If Not bottom And Not middle And Not top Then
        '    l_production_blank_row.Visible = True
        'ElseIf Not bottom And Not middle And top Then
        'Else
        '    l_production_blank_row.Visible = False
        'End If
    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        Select Case True
            Case tc_main.SelectedTab Is tp_summary
                print_summary()
            Case tc_main.SelectedTab Is tp_production
                print_report()
            Case tc_main.SelectedTab Is tp_reject
                print_reject()
        End Select
    End Sub

    Private Sub ImportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportToolStripMenuItem.Click

        Dim path As String = "Z:\Production\Printing Dept\Printing (Prod logon)\History\Printing\Production Sheets\"
        Dim pn As Long = InputBox("Enter the Production Number for the job you want to import data from", "Import Data")
        Dim fn As String = path & pn & ".xls"
        If System.IO.File.Exists(fn) Then
            xlApp = CreateObject("Excel.Application")
            xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
            'xlApp.ScreenUpdating = False
            xlWorkBook = xlApp.Workbooks.Open(fn, [ReadOnly]:=True)
            xlWorkSheets = xlWorkBook.Worksheets
            xlWorkSheet = xlWorkSheets("1")
            xlWorkSheet.Unprotect()
            s_selected_job = dt_schedules.Select("prod_num = " & pn)
            If Not tb_prod_num.Text = Nothing Then s_current_job = tb_prod_num.Text
            display_job(True)
            dgv_production.Rows.Clear()
            tc_main.SelectedIndex = 1
            Dim plain_num As Long = xlWorkSheet.Range("T7").Value
            Dim plain_desc As String = xlWorkSheet.Range("H7").Value
            cbo_mat_num_in.Items.Add(plain_num)
            cbo_mat_num_in.Text = plain_num
            cbo_mat_desc_in.Items.Add(plain_desc)
            cbo_mat_desc_in.Text = plain_desc

            If Not sql.exists("db_production", "prod_num", "=", pn) Then
                sql.insert("db_production", "prod_num", pn)
            End If
            s_loading = True
            For sh = 1 To xlWorkSheets.Count - 1
                If sh = 1 Then
                    sql.update("db_production", "ink_1='" & xlWorkSheet.Range("G17").Value & _
                               "',ink_2='" & xlWorkSheet.Range("J17").Value & _
                               "',ink_3='" & xlWorkSheet.Range("M17").Value & _
                               "',ink_4='" & xlWorkSheet.Range("P17").Value & _
                               "',ink_5='" & xlWorkSheet.Range("S17").Value & _
                               "',ink_6='" & xlWorkSheet.Range("V17").Value & _
                               "',ink_7='" & xlWorkSheet.Range("Y17").Value & _
                               "',ink_8='" & xlWorkSheet.Range("AB17").Value & _
                               "',ink_9='" & xlWorkSheet.Range("AE17").Value & _
                               "',ink_10='" & xlWorkSheet.Range("AH17").Value & _
                               "',mat_num_plain='" & plain_num & "',mat_desc_plain='" & _
                               plain_desc & "',access", Convert.ToSByte(s_sql_access), "prod_num", "=", pn)

                    nud_items_out.Value = xlWorkSheet.Range("B20").Value
                End If
                xlWorkSheet = xlWorkSheets(CStr(sh))
                If xlWorkSheet.Range("AO109").Value = 0 Then xlWorkSheet.Range("AO109").Clear()
                If xlWorkSheet.Range("AP109").Value = 0 Then xlWorkSheet.Range("AP109").Clear()
                For r As Integer = 22 To 109
                    If xlWorkSheet.Range("AV" & r).Value > 1 Then
                        With dgv_production
                            .Rows.Add()
                            Dim row As Integer = .RowCount - 1
                            .Rows(row).Cells("num_in").Value = xlWorkSheet.Range("A" & r).Value
                            .Rows(row).Cells("item_id").Value = xlWorkSheet.Range("B" & r).Value
                            .Rows(row).Cells("kg_in").Value = xlWorkSheet.Range("E" & r).Value
                            .Rows(row).Cells("num_out").Value = xlWorkSheet.Range("G" & r).Value
                            .Rows(row).Cells("kg_out").Value = xlWorkSheet.Range("H" & r).Value
                            .Rows(row).Cells("km_out").Value = xlWorkSheet.Range("J" & r).Value
                            .Rows(row).Cells("comments").Value = xlWorkSheet.Range("L" & r).Value
                            .Rows(row).Cells("rts").Value = xlWorkSheet.Range("AC" & r).Value
                            .Rows(row).Cells("reject").Value = xlWorkSheet.Range("AE" & r).Value
                            .Rows(row).Cells("barcode").Value = xlWorkSheet.Range("AG" & r).Value
                            .Rows(row).Cells("rho_left").Value = xlWorkSheet.Range("AI" & r).Value
                            .Rows(row).Cells("rho_right").Value = xlWorkSheet.Range("AJ" & r).Value
                            .Rows(row).Cells("pallet_id").Value = xlWorkSheet.Range("AL" & r).Value
                            .Rows(row).Cells("prod_num").Value = pn
                            .Rows(row).Cells("shift").Value = s_shift
                            If Not IsNull(xlWorkSheet.Range("AK" & r).Value) Then
                                Dim str() As String = Split(xlWorkSheet.Range("AK" & r).Value, "/")

                                Dim u As String = get_user(str(0))
                                .Rows(row).Cells("user_").Value = u
                                .Rows(row).Cells("date_").Value = Replace(str(0), u, Nothing).ToString.PadLeft(2, "0") & "/" & str(1).ToString.PadLeft(2, "0") & "/" & Strings.Left(str(2), 4)
                            End If
                            .Rows(row).Cells("assistant").Value = xlWorkSheet.Range("AE5").Value

                        End With
                    End If
                Next
            Next
            xlWorkBook.Close(False)

            close_excel()
            If nud_items_out.Value > 1 Then
                With dgv_production
                    For r As Integer = 0 To .RowCount - 2
                        If IsNull(.Rows(r + 1).Cells("kg_in").Value) Then
                            .Rows(r + 1).Cells("kg_in").Value = "-"
                        End If
                        If r = 18 Then
                            r = r
                        End If
                        If IsNumeric(.Rows(r).Cells("kg_in").Value) And .Rows(r + 1).Cells("kg_in").Value.ToString = "-" And Not IsNull(.Rows(r).Cells("kg_out").Value) Then
                            .Rows(r + 1).Cells("kg_in").Value = FormatNumber(.Rows(r).Cells("kg_in").Value - .Rows(r).Cells("kg_out").Value, 1)
                            .Rows(r).Cells("kg_in").Value = FormatNumber(.Rows(r).Cells("kg_out").Value, 1)
                        End If
                    Next
                End With
            End If
            s_loading = False

            export_report()
            If Not s_current_job = Nothing Then
                s_selected_job = dt_schedules.Select("prod_num='" & s_current_job & "'")
                display_job()
            End If
            import_report()
        End If
    End Sub

    Sub close_excel()
        Dim xlHWND As Integer = xlApp.Hwnd
        'this will have the process ID after call to GetWindowThreadProcessId
        Dim ProcIdXL As Integer = 0
        'get the process ID
        GetWindowThreadProcessId(xlHWND, ProcIdXL)
        'get the process
        Dim xproc As Process = Process.GetProcessById(ProcIdXL)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
        'set to nothing
        xlApp = Nothing
        'kill it with glee
        If Not xproc.HasExited Then
            xproc.Kill()
        End If
    End Sub

    Function get_user(ByVal str As String) As String
        Dim s() As String = Split(UCase(str), "/")
        If s.Length > 0 Then
            If sql.exists("DP_Logins", "Login_Id", "=", "'" & Strings.Left(s(0), s(0).Length - 1) & "'", "symple") Then
                Return Strings.Left(s(0), s(0).Length - 1)
            ElseIf sql.exists("DP_Logins", "Login_Id", "=", "'" & Strings.Left(s(0), s(0).Length - 2) & "'", "symple") Then
                Return Strings.Left(s(0), s(0).Length - 2)
            Else
                Return "N0TF0UND"
            End If
        Else
            Return "N0TF0UND"
        End If
    End Function


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_machine.SelectedIndexChanged
        If s_loading Then Exit Sub
        s_machine = cbo_machine.SelectedItem("name")
        s_extruder = cbo_machine.SelectedItem("extruder")

        If s_department = "printing" Then display_stations(cbo_machine.SelectedItem("stations"))
        set_department()
    End Sub

    Private Sub RestoreToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestoreToolStripMenuItem.Click
        With frm_restore
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub

    Private Sub EditToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditToolStripMenuItem.Click
        If Not dgv_job_requirements.Rows(0).Cells(2).Value = 0 Then
            ReportToolStripMenuItem.Enabled = False
        Else
            ReportToolStripMenuItem.Enabled = True
        End If
    End Sub

    Private Sub ReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportToolStripMenuItem.Click
        dgv_production.Rows.Clear()
        export_report()
    End Sub

    Private Sub SummaryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SummaryToolStripMenuItem.Click
        s_loading = True
        export_summary()
        dgv_summary.Rows.Clear()
        tb_start_date.Text = Format(Now, "dd/MM/yyyy")
        s_shift = Nothing
        s_loading = False
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tc_main.SelectedIndexChanged
        If dgv_production.RowCount = 0 Or s_loading Then Exit Sub
        Select Case tc_main.SelectedTab.Text
            Case "Summary" 'summary
                ' If nud_items_out.Value > 1 And s_department = "printing" Then
                If nud_items_out.Value > 1 Then
                    calc_times(dgv_summary.CurrentRow.Index)
                    If s_department = "printing" Or s_department = "filmsline" Then
                        With dgv_production
                            For i As Integer = 0 To .RowCount - 2
                                If Not .Rows(i).Cells("dgc_production_kg_in").Value.ToString = "-" And .Rows(i + 1).Cells("dgc_production_kg_in").Value.ToString = _
                                    "-" And Not IsNull(.Rows(i).Cells("dgc_production_kg_out").Value) And Not IsNull(.Rows(i).Cells("dgc_production_user_").Value) Then

                                    If .Rows(i).Cells("dgc_production_kg_in").Value - .Rows(i).Cells("dgc_production_kg_out").Value > 9 Then

                                        If .Rows(i).Cells("dgc_production_user_").Value.ToString = s_user And Not i = .RowCount - 1 Then
                                            Dim resp As MsgBoxResult = MsgBox("Part of a roll has been run, is there anything of this roll left for the next shift?" & vbCr & _
                                                                              "This is needed for your totals to be correct", MsgBoxStyle.YesNo, "Part roll detected")
                                            If resp = MsgBoxResult.Yes Then
                                                .Rows(i + 1).Cells("dgc_production_kg_in").Value = _
                                                    FormatNumber(get_item_info(i, True).remaining, 1, , , False)
                                                .Rows(i).Cells("dgc_production_kg_in").Value = _
                                                    FormatNumber(.Rows(i).Cells("dgc_production_kg_out").Value, 1)
                                                export_report(i)
                                                export_report(i + 1)
                                                dt_produced = dgv_to_dt(dgv_production)
                                                calc_times(dgv_summary.CurrentRow.Index)
                                            End If

                                            Exit For
                                        End If
                                    End If
                                End If
                            Next
                        End With
                    End If
                End If

                dgv_summary.Focus()
            Case "Production Report" ' production
                dgv_production.PerformLayout()
                dgv_production.Focus()
            Case 2 ' downtime
            Case 3 ' rts
            Case 4 ' reject
            Case 5 ' job info
            Case "Plate Directory" ' plate directory
                dt_boxes = sql.get_table("plate_options", "plant", "=", s_plant)
                With cbo_plate_boxsize
                    .Items.Clear()
                    For i As Integer = 0 To dt_boxes.Rows.Count - 1
                        .Items.Add(dt_boxes(i)("box_size"))
                    Next
                    .SelectedIndex = 0
                End With


                populate_plates()
            Case "Planning"

        End Select
        If Not tb_prod_num.Text = Nothing Then calc_times(get_prod_num_row(tb_prod_num.Text))
        'dt_produced = sql.get_table("db_produced", "prod_num", "=", tb_prod_num.Text)

    End Sub

    Private Sub b_personel_Click(sender As Object, e As EventArgs) Handles b_main_personel.Click
        With frm_helper
            .StartPosition = FormStartPosition.CenterParent
            .Button2.Text = "Cancel"
            .Button3.Enabled = False
            .ShowDialog()
        End With
    End Sub

    Private Sub lb_personel_DoubleClick(sender As Object, e As EventArgs) Handles lb_personel.DoubleClick
        If lb_personel.SelectedIndex > -1 Then
            lb_personel.Items.RemoveAt(lb_personel.SelectedIndex)
            s_assistant = Nothing
            For i As Integer = 0 To lb_personel.Items.Count - 1
                If s_assistant = Nothing Then
                    s_assistant = lb_personel.Items(i)
                Else
                    s_assistant = s_assistant & "/" & lb_personel.Items(i)
                End If
            Next
        End If
        export_summary()
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles cb_main_loading.CheckedChanged
        s_loading = cb_main_loading.Checked
    End Sub

    Private Sub tb_shift_DoubleClick(sender As Object, e As EventArgs) Handles tb_shift.DoubleClick
        With frm_startup
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub


    Private Sub EditItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsm_production_edit_item.Click
        Dim str As String
        With dgv_production
            'modify for dashes
            Dim i_info As item_info = get_item_info(.CurrentRow.Index, True)
            If .Rows(i_info.row_start).Cells("dgc_production_item_id").Value.ToString.Contains("&") Then
                MsgBox("You can't edit a merged item")
                Exit Sub
            ElseIf .Rows(i_info.row_start).Cells("dgc_production_km_in").Value > 0 Then
                'bring up roll info box
                MsgBox("You can't edit this item")
                Exit Sub
            End If


            str = InputBox("Enter the original production number" & vbCrLf & "Roll: " & i_info.num_in, "Edit item IN", i_info.prod_num)
            If str = Nothing Then
                Exit Sub
            ElseIf IsNumeric(str) Then
                i_info.prod_num = str
            End If

            str = InputBox("Enter the material number" & vbCrLf & "Roll: " & i_info.num_in, "Edit item IN", i_info.mat_num)
            If str = Nothing Then
                Exit Sub
            ElseIf IsNumeric(str) Then
                i_info.mat_num = str
            End If
            If cbo_mat_num_in.FindStringExact(i_info.mat_num) > -1 Then
                i_info.item_id = InputBox("Enter the Item ID", "Change Item ID", i_info.item_id)
                If i_info.item_id = Nothing Then
                    Exit Sub
                End If
            Else
                MsgBox("That material number does not match!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error!")
                Exit Sub
            End If

            str = InputBox("Enter the IN weight", "Change Item Weight", i_info.kg_in)

            If IsNumeric(str) Then
                If str > 0 Then
                    i_info.kg_in = str
                    i_info.remaining = i_info.kg_in
                    .Rows(i_info.row_start).Cells("dgc_production_kg_in").Value = FormatNumber(i_info.kg_in, 1)
                    .Rows(i_info.row_start).Cells("dgc_production_item_id").Value = i_info.item_id
                    .Rows(i_info.row_start).Cells("dgc_production_kg_in_orig").Value = FormatNumber(i_info.kg_in, 1)
                    .Rows(i_info.row_start).Cells("dgc_production_mat_num").Value = i_info.mat_num
                    .Rows(i_info.row_start).Cells("dgc_production_prod_num_orig").Value = i_info.prod_num
                    If Not i_info.row_start = i_info.row_end Then
                        For r As Integer = i_info.row_start To i_info.row_end
                            If r < i_info.row_last_item_row Then
                                Dim kg_out As Decimal = 0
                                If IsNumeric(.Rows(r).Cells("dgc_production_kg_out").Value) Then
                                    kg_out = .Rows(r).Cells("dgc_production_kg_out").Value
                                End If
                                If s_department = "bagplant" And IsNumeric(.Rows(r).Cells("dgc_production_km_out").Value) Then
                                    kg_out = get_conversion_factor(True, tb_mat_num_out.Text, False) * .Rows(r).Cells("dgc_production_km_out").Value / 1000
                                End If

                                If i_info.remaining - kg_out < 0 Then
                                    .Rows(r).Cells("dgc_production_kg_in").Value = i_info.remaining
                                    i_info.remaining = 0
                                Else
                                    .Rows(r).Cells("dgc_production_kg_in").Value = kg_out
                                    i_info.remaining = i_info.remaining - kg_out
                                End If

                            ElseIf r = i_info.row_last_item_row Then
                                .Rows(r).Cells("dgc_production_kg_in").Value = i_info.remaining
                            Else
                                .Rows(r).Cells("dgc_production_kg_in").Value = "-"
                            End If
                        Next

                    End If

                    'report_kg_out(dgv_production)
                    export_report()
                    dt_produced = dgv_to_dt(dgv_production)
                    calc_totals()
                End If
            End If
        End With
    End Sub

    Private Sub RollTicketsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RollTicketsToolStripMenuItem.Click

        Dim mn As String = InputBox("Please Enter the material Number", "Material Number")

        If IsNumeric(mn) Then
            Dim str As String = InputBox("How many Roll Tickets?", "Roll Ticket", 1)
            If IsNumeric(str) And Not str = "0" Then

                xlApp = New Microsoft.Office.Interop.Excel.Application
                xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
                xlApp.ScreenUpdating = False
                xlWorkBook = xlApp.Workbooks.Open(s_resource_path & "roll tickets.xls", [ReadOnly]:=True)
                xlWorkSheets = xlWorkBook.Worksheets
                xlWorkSheet = xlWorkSheets(1)
                xlWorkSheet.Unprotect()
                xlWorkSheet.Range("A5").Value = acsis.get_material_desc(mn)
                xlWorkSheet.Range("H1").Value = mn
                xlApp.ScreenUpdating = True
                xlWorkSheet.PrintOutEx(1, 1, str, False)
                xlWorkBook.Close(False)
                close_excel()
                MsgBox("Tickets printed to " & get_printer_name())
            End If
        End If
    End Sub

    Private Sub PalletsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PalletsToolStripMenuItem.Click
        Dim str As String = InputBox("How many Pallet Forms?", "Pallet Form", 1)
        If IsNumeric(str) And Not str = "0" Then
            xlApp = New Microsoft.Office.Interop.Excel.Application
            xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
            xlApp.ScreenUpdating = False
            xlWorkBook = xlApp.Workbooks.Open(s_resource_path & "pallet label.xls", [ReadOnly]:=True)
            xlWorkSheets = xlWorkBook.Worksheets
            xlWorkSheet = xlWorkSheets(1)
            xlApp.ScreenUpdating = True
            If s_department = "filmsline" Then
                xlWorkSheet.PrintOutEx(1, 2, str, False)
            Else
                xlWorkSheet.PrintOutEx(1, 1, str, False)
            End If
            xlWorkBook.Close(False)
            close_excel()
            MsgBox("Pallet Forms printed to " & get_printer_name())
        End If
    End Sub

    Private Sub b_downtime_add_Click(sender As Object, e As EventArgs) Handles b_downtime_bs01_add.Click, b_downtime_cs01_add.Click
        Dim btn As Button = CType(sender, Button), dgv As DataGridView
        s_dt_edit = False
        Select Case True
            Case btn.Name.Contains("bs01")
                dgv = dgv_downtime_bs01
                s_downtime_type = "BS01"
            Case btn.Name.Contains("cs01")
                dgv = dgv_downtime_cs01
                s_downtime_type = "CS01"
            Case Else
                dgv = dgv_downtime_bs01
        End Select
        dgv.ClearSelection()
        With frm_downtime
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub

    Private Sub b_downtime_remove_Click(sender As Object, e As EventArgs) Handles b_downtime_bs01_remove.Click, b_downtime_cs01_remove.Click
        Dim btn As Button = CType(sender, Button), dgv As DataGridView
        Select Case True
            Case btn.Name.Contains("bs01")
                dgv = dgv_downtime_bs01
                s_downtime_type = "BS01"
            Case btn.Name.Contains("cs01")
                dgv = dgv_downtime_cs01
                s_downtime_type = "CS01"
            Case Else
                dgv = dgv_downtime_bs01
                s_downtime_type = "BS01"
        End Select
        With dgv
            If .RowCount > 0 Then
                If .CurrentRow.Index > -1 Then
                    .Rows.RemoveAt(.CurrentRow.Index)
                    export_downtime(s_downtime_type)
                End If
            End If
        End With

    End Sub

    Private Sub dgv_downtime_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_downtime_bs01.CellClick, dgv_downtime_bs01.CellDoubleClick, dgv_downtime_cs01.CellDoubleClick, dgv_downtime_cs01.CellClick
        Dim dgv As DataGridView = CType(sender, DataGridView)
        'Select Case True
        '    Case dgv.Name.Contains("bs01")
        '        dgv = dgv_downtime_bs01
        '    Case dgv.Name.Contains("cs01")
        '        dgv = dgv_downtime_cs01
        '    Case Else
        '        export_downtime("BS01")
        'End Select
        If e.RowIndex > -1 Then
            dgv.Rows(e.RowIndex).Selected = True
        End If
    End Sub

    Private Sub dgv_downtime_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_downtime_bs01.CellEndEdit, dgv_downtime_cs01.CellEndEdit
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Select Case True
            Case dgv.Name.Contains("bs01")
                export_downtime("BS01")
            Case dgv.Name.Contains("cs01")
                export_downtime("CS01")
            Case Else
                export_downtime("BS01")
        End Select
    End Sub

    Private Sub dgv_downtime_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv_downtime_bs01.CellMouseClick, dgv_downtime_cs01.CellMouseClick
        Dim dgv As DataGridView = CType(sender, DataGridView)
        With dgv
            If e.Button = Windows.Forms.MouseButtons.Right And e.ColumnIndex > -1 And e.RowIndex > -1 Then

                .ClearSelection()
                Dim cell As DataGridViewCell = .Rows(e.RowIndex).Cells(e.ColumnIndex)
                .Rows(e.RowIndex).Selected = True
                .CurrentCell = cell
                EditToolStripMenuItem1.Enabled = True
                DeleteToolStripMenuItem.Enabled = True

                cms_downtime.Show(Cursor.Position)


            End If
            Dim enable As Boolean = True
            If .RowCount = 0 Then enable = False

            Select Case True
                Case dgv.Name.Contains("bs01")
                    b_downtime_bs01_edit.Enabled = enable
                    b_downtime_bs01_remove.Enabled = enable
                Case dgv.Name.Contains("cs01")
                    b_downtime_cs01_edit.Enabled = enable
                    b_downtime_cs01_remove.Enabled = enable
            End Select
        End With

    End Sub

    Private Sub dgv_downtime_MouseClick(sender As Object, e As MouseEventArgs) Handles dgv_downtime_bs01.MouseClick, dgv_downtime_cs01.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Dim dgv As DataGridView = CType(sender, DataGridView)
            Select Case True
                Case dgv.Name.Contains("bs01")
                    s_downtime_type = "BS01"
                Case dgv.Name.Contains("cs01")
                    s_downtime_type = "CS01"
                Case Else
                    s_downtime_type = "BS01"
            End Select
            EditToolStripMenuItem1.Enabled = False
            DeleteToolStripMenuItem.Enabled = False
            cms_downtime.Show(Cursor.Position)
        End If
    End Sub

    Private Sub EditToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles EditToolStripMenuItem1.Click
        s_dt_edit = True

        With frm_downtime
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub

    Private Sub AddDowntimeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddDowntimeToolStripMenuItem.Click
        s_dt_edit = False
        Select Case s_downtime_type
            Case "BS01"
                dgv_downtime_bs01.ClearSelection()
            Case "CS01"
                dgv_downtime_cs01.ClearSelection()
            Case Else
                dgv_downtime_bs01.ClearSelection()
        End Select
        With frm_downtime
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub

    Private Sub b_downtime_edit_Click(sender As Object, e As EventArgs) Handles b_downtime_bs01_edit.Click, b_downtime_cs01_edit.Click
        s_dt_edit = True
        Dim btn As Button = CType(sender, Button), dgv As DataGridView
        Select Case True
            Case btn.Name.Contains("bs01")
                dgv = dgv_downtime_bs01
                s_downtime_type = "BS01"
            Case btn.Name.Contains("cs01")
                dgv = dgv_downtime_cs01
                s_downtime_type = "CS01"
            Case Else
                dgv = dgv_downtime_bs01
                s_downtime_type = "BS01"
        End Select
        With dgv
            If .RowCount > 0 Then
                If .CurrentRow.Index > -1 Then
                    With frm_downtime
                        .StartPosition = FormStartPosition.CenterParent
                        .ShowDialog()
                    End With
                End If
            End If
        End With

    End Sub

    Private Sub b_main_change_mat_Click(sender As Object, e As EventArgs) Handles b_main_mat_in_add.Click
        With frm_material_info
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub

    Private Sub dgv_production_MouseUp(sender As Object, e As MouseEventArgs) Handles dgv_production.MouseUp
        If dgv_production.SelectedCells.Count > 1 And Not ModifierKeys = Keys.Control Then
            dgv_production.ClearSelection()
        End If
    End Sub

    Private Sub dgv_production_KeyUp(sender As Object, e As KeyEventArgs) Handles dgv_production.KeyUp
        If e.KeyCode = Keys.ControlKey And dgv_production.SelectedCells.Count > 1 Then
            dgv_production.ClearSelection()
            l_production_selected_totals.Visible = False
        End If
    End Sub

    Private Sub tb_start_date_Click(sender As Object, e As EventArgs) Handles tb_start_date.Click
        'Dim dt As Date = tb_start_date.Text
        'If ModifierKeys = Keys.Shift Then
        '    dt = dt.AddDays(-1)
        'Else
        '    dt = dt.AddDays(1)
        'End If
        'tb_start_date.Text = Format(dt, "dd/MM/yyyy").ToString

    End Sub

    Private Sub dgv_summary_MouseMove(sender As Object, e As MouseEventArgs) Handles dgv_summary.MouseMove
        Dim v As Decimal = 0
        If e.Button = Windows.Forms.MouseButtons.Left Or ModifierKeys = Keys.Control Then
            With dgv_summary
                'If .SelectedCells.Count > 1 Then
                For r As Integer = 0 To .RowCount - 1
                    For c As Integer = 0 To .Columns.Count - 1
                        If .Rows(r).Cells(c).Selected Then
                            If IsNumeric(.Rows(r).Cells(c).Value) And Not IsNull(.Rows(r).Cells(c).Value) Then
                                v = v + .Rows(r).Cells(c).Value
                            End If
                        End If
                    Next
                Next
                l_summary_selected_totals.Visible = True
                l_summary_selected_totals.Text = v
                l_summary_selected_totals.Location = New Point(e.X + 10, e.Y + 10)
                'End If

            End With
        Else
            If Not ModifierKeys = Keys.Control Then l_summary_selected_totals.Visible = False
        End If
    End Sub

    Private Sub dgv_summary_KeyUp(sender As Object, e As KeyEventArgs) Handles dgv_summary.KeyUp
        If e.KeyCode = Keys.ControlKey And dgv_summary.SelectedCells.Count > 1 Then
            dgv_summary.ClearSelection()
            l_summary_selected_totals.Visible = False
        End If
    End Sub

    Private Sub dgv_summary_MouseUp(sender As Object, e As MouseEventArgs) Handles dgv_summary.MouseUp
        If dgv_summary.SelectedCells.Count > 1 And Not ModifierKeys = Keys.Control Then
            dgv_summary.ClearSelection()
        End If
    End Sub

    Private Sub EditValueToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsm_summary_edit_item.Click
        With dgv_summary
            Dim i As Decimal = .CurrentCell.Value
            Dim s = InputBox("Enter the new value" & vbCr & "This change will not be saved, it is to fix the printout." & vbCr & "Please be sure to enter into the debug comment why you needed to do this", "Edit Item", .CurrentCell.Value)
            If IsNumeric(s) Then .CurrentCell.Value = CDec(s)
        End With
    End Sub

    Private Sub Button18_Click_1(sender As Object, e As EventArgs) Handles Button18.Click
        With AboutBox1
            .StartPosition = FormStartPosition.CenterScreen
            .ShowDialog()

        End With
    End Sub


    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles b_main_show_running.Click
        With frm_job_status
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub

    Private Sub PrepareJobToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrepareJobToolStripMenuItem.Click
        If s_department = "mounting" And lb_personel.Items.Count = 0 Then
            MsgBox("You need to add a mounter first", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "No mounter added")
        Else
            With frm_schedules
                .StartPosition = FormStartPosition.CenterParent
                .ShowDialog()
            End With
        End If
    End Sub

    Private Sub tsm_reprint_Click(sender As Object, e As EventArgs) Handles tsm_production_reprint.Click
        With dgv_production
            Dim quant As Decimal = 0
            If IsNumeric(.CurrentRow.Cells("dgc_production_km_out").Value) Then
                If s_selected_job(0)("req_uom_1") = "IMP" Then
                    quant = FormatNumber((.CurrentRow.Cells("dgc_production_km_out").Value * 1000) / (tb_cyl_size.Text / 1000), 2)
                Else
                    quant = .CurrentRow.Cells("dgc_production_km_out").Value
                End If
            End If
            Dim meters As String = .CurrentRow.Cells("dgc_production_km_out").Value
            Dim str() As String = Split(.Columns("dgc_production_km_out").HeaderText, " ")
            If str(0) = "KM" Then
                meters = meters * 1000
            End If

            ticket_print(meters)

        End With
    End Sub


    Private Sub LaminateFormToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LaminateFormToolStripMenuItem.Click
        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
        xlApp.ScreenUpdating = False
        xlWorkBook = xlApp.Workbooks.Open(s_resource_path & "laminates.xls", [ReadOnly]:=True)
        xlWorkSheets = xlWorkBook.Worksheets
        xlWorkSheet = xlWorkSheets(1)
        xlWorkSheet.Range("C4").Value = tb_label_desc.Text
        xlWorkSheet.Range("C6").Value = tb_prod_num.Text
        xlWorkSheet.Range("C7").Value = tb_mat_num_out.Text
        xlWorkSheet.Range("C8").Value = cbo_machine.Text
        xlWorkSheet.Range("G6").Value = tb_start_date.Text
        xlWorkSheet.Range("G7").Value = s_user_name
        xlApp.ScreenUpdating = True
        xlWorkSheet.PrintOutEx(1, 1, 1, False)
        xlWorkBook.Close(False)
        close_excel()
        MsgBox("The laminate form has been sent to the printer " & get_printer_name())
    End Sub
    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.F4 Then
            ' Call your sub method here  .....

            Application.Exit()
            ' then avoid the key to reach the current control
            e.Handled = False
        ElseIf Not e.Modifiers Then
            'Select Case tc_main.SelectedIndex
            '    Case 0 'summary
            '        dgv_summary.Focus()
            '    Case 1 'production
            '        dgv_production.Focus()
            '    Case 2 'downtime
            '    Case 3 'rts
            '    Case 4 ' reject
            '        dgv_reject.Focus()
            'End Select

        End If
    End Sub

    Private Sub ReprintTicketsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReprintTicketsToolStripMenuItem.Click
        With frm_reprint
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub

    Private Sub tb_plate_search_TextChanged(sender As Object, e As EventArgs) Handles tb_plate_search.TextChanged
        If s_loading Then Exit Sub
        b_plate_change_desc.Enabled = False
        plate_search()
    End Sub


    Private Sub tb_plate_password_TextChanged(sender As Object, e As EventArgs) Handles tb_plate_password.TextChanged
        If tb_plate_password.Text = s_pass_plate Then
            gb_box_control.Enabled = True
            If tb_plate_search.Text = Nothing Then
                plate_search(True)
            Else
                plate_search()
            End If
        Else
            gb_box_control.Enabled = False
        End If
    End Sub


    Private Sub cbo_plate_boxsize_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_plate_boxsize.SelectedIndexChanged
        s_loading = True
        tb_plate_search.Clear()
        populate_plates()
        plate_search(True)
        s_loading = False
    End Sub



    Private Sub dgv_plates_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_plates.CellClick
        With dgv_plates
            If Not .CurrentCell Is Nothing Then
                If IsNull(.CurrentRow.Cells(2).Value) Then
                    tb_plate_desc.Clear()
                Else
                    tb_plate_desc.Text = .CurrentRow.Cells(2).Value.ToString
                End If
                If IsNull(.CurrentRow.Cells(3).Value) Then
                    tb_plate_comment.Clear()
                Else
                    tb_plate_comment.Text = .CurrentRow.Cells(3).Value.ToString
                End If
                .CurrentRow.Selected = True
                b_plate_check.Enabled = True
                b_plate_remove_box.Enabled = True
                If Not IsNull(.CurrentRow.Cells(0).Value) Then
                    b_plate_contents.Enabled = True
                    b_plate_change_desc.Enabled = True
                    b_plate_label_reprint.Enabled = True
                Else
                    b_plate_contents.Enabled = False
                    b_plate_change_desc.Enabled = False
                    b_plate_label_reprint.Enabled = False
                End If
                s_loading = True
                tb_plate_search.Text = .CurrentRow.Cells(0).Value
                s_loading = False
            Else
                b_plate_check.Enabled = False
                b_plate_remove_box.Enabled = False
                b_plate_contents.Enabled = False
                b_plate_change_desc.Enabled = False
                b_plate_label_reprint.Enabled = False
            End If
        End With
    End Sub

    Private Sub b_plate_change_desc_Click(sender As Object, e As EventArgs) Handles b_plate_change_desc.Click
        If cb_plate_num_lock.Checked Then
            sql.update("db_plate_location", "description", "'" & Replace(tb_plate_desc.Text, "'", "''") & _
                       "'", "plate_num", "=", "'" & dgv_plates.CurrentRow.Cells(0).Value & "'")
            sql.update("db_plate_location", "comment", "'" & Replace(tb_plate_comment.Text, "'", "''") & _
                       "'", "plate_num", "=", "'" & dgv_plates.CurrentRow.Cells(0).Value & "'")
        Else
            export_plates()
        End If
        plate_search()
    End Sub


    Private Sub b_plate_check_Click(sender As Object, e As EventArgs) Handles b_plate_check.Click
        With dgv_plates
            If Not .CurrentRow Is Nothing Then
                sql.update("db_plate_location", "checked", "'" & tb_start_date.Text & "'", "plate_num", "=", "'" & _
                           dgv_plates.CurrentRow.Cells(0).Value & "'")
                plate_search()
            Else
                MsgBox("Select a row first!")
                b_plate_check.Enabled = False
            End If
        End With
    End Sub

    Private Sub box_count_change(sender As Object, e As EventArgs) Handles nud_plate_columns.ValueChanged, nud_plate_rows.ValueChanged, _
        nud_plate_boxes.ValueChanged

        If s_loading Then Exit Sub
        Dim response As MsgBoxResult
        Dim n_old, n_new As Integer
        n_old = CInt(tb_plate_total.Text)
        n_new = nud_plate_boxes.Value * nud_plate_columns.Value * nud_plate_rows.Value
        If n_new < n_old Then
            response = MessageBox.Show("You have chosen to reduce the box count, boxes will now have to be removed from storage" & _
                                       vbCr & "Do you want to proceed?", "Warning", MessageBoxButtons.YesNo)

            If response = MsgBoxResult.Yes Then
                With frm_plate_removal
                    .StartPosition = FormStartPosition.CenterParent
                    .ShowDialog()
                End With
            End If
        ElseIf n_new > n_old Then
            response = MessageBox.Show("You have chosen to increase the box count, you will be shown the new storage locations" & _
                                       vbCr & "Do you want to proceed?", "Warning", MessageBoxButtons.YesNo)

            If response = MsgBoxResult.Yes Then
                With frm_plate_add
                    .StartPosition = FormStartPosition.CenterParent
                    .ShowDialog()
                End With
            End If

        End If
        populate_plates()
    End Sub

    Private Sub b_plate_contents_Click(sender As Object, e As EventArgs) Handles b_plate_contents.Click
        With frm_plate_contents
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub


    Private Sub b_plate_search_Click(sender As Object, e As EventArgs) Handles b_plate_search.Click
        With frm_plate_search
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub


    Private Sub b_plate_add_box_Click(sender As Object, e As EventArgs) Handles b_plate_add_box.Click
        With frm_plate_add_box
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With

    End Sub

    Private Sub b_plate_label_reprint_Click(sender As Object, e As EventArgs) Handles b_plate_label_reprint.Click
        With dgv_plates
            print_box_label(.CurrentRow.Cells(1).Value, .CurrentRow.Cells(0).Value)
        End With
    End Sub

    Private Sub b_plate_remove_box_Click(sender As Object, e As EventArgs) Handles b_plate_remove_box.Click
        Dim response As MsgBoxResult
        response = MsgBox("Are you sure you want to remove this box from the inventory?" & _
                          vbCr & "This will delete ALL boxes if there are more than one.", _
                          MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Delete Box")

        If response = MsgBoxResult.Yes Then
            With dgv_plates
                sql.update("db_plate_location", "checked", "NULL", "plate_num", "=", "'" & .CurrentRow.Cells(0).Value & "'")
                sql.update("db_plate_location", "description", "NULL", "plate_num", "=", "'" & .CurrentRow.Cells(0).Value & "'")
                sql.update("db_plate_location", "comment", "NULL", "plate_num", "=", "'" & .CurrentRow.Cells(0).Value & "'")
                sql.update("db_plate_location", "plates", "NULL", "plate_num", "=", "'" & .CurrentRow.Cells(0).Value & "'")
                sql.update("db_plate_location", "archived", "NULL", "plate_num", "=", "'" & .CurrentRow.Cells(0).Value & "'")
                sql.update("db_plate_location", "plate_num", "NULL", "plate_num", "=", "'" & .CurrentRow.Cells(0).Value & "'")
            End With
            populate_plates()
            plate_search(True)
            b_plate_remove_box.Enabled = False
        End If

    End Sub

    Private Sub cbx_reject_show_all_CheckedChanged(sender As Object, e As EventArgs) Handles cb_show_all.CheckedChanged
        import_reject(tb_prod_num.Text)
    End Sub

    Private Sub dgv_job_requirements_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_job_requirements.CellEndEdit
        With dgv_job_requirements
            s_selected_job = dt_schedules.Select("prod_num='" & tb_prod_num.Text & "'")
            s_selected_job(0)("req_quant_1") = .Rows(0).Cells(1).Value

            sql.update("db_schedule", "req_quant_1", .Rows(0).Cells(1).Value, "prod_num", "=", tb_prod_num.Text & " AND plant=" & s_plant)
            display_job()
            If .RowCount > 1 Then
                .Rows(1).Cells(1).Value = .Rows(0).Cells(1).Value * get_conversion_factor(False, tb_mat_num_out.Text, False)
                If .Rows(1).Cells(0).Value = "M" Then
                    .Rows(1).Cells(1).Value = .Rows(1).Cells(1).Value * 1000
                End If
                .Rows(1).Cells(1).Value = FormatNumber(.Rows(1).Cells(1).Value, 1, , , TriState.False)
                sql.update("db_schedule", "req_quant_2", .Rows(1).Cells(1).Value, "prod_num", "=", tb_prod_num.Text & " AND plant=" & s_plant)
                s_selected_job(0)("req_quant_2") = .Rows(1).Cells(1).Value
            End If
            dgv_focus(sender, e)
            calc_totals()
            .CurrentCell = Nothing
        End With
    End Sub


    Private Sub dgv_job_requirements_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgv_job_requirements.EditingControlShowing
        allow_decimal = True
        RemoveHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf CellKeyPress
        AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf CellKeyPress
    End Sub

    Private Sub dgv_production_SelectionChanged(sender As Object, e As EventArgs) Handles dgv_production.SelectionChanged
        If s_loading Then Exit Sub
        With dgv_production
            If .RowCount > 0 Then
                get_production_info(.CurrentRow.Index)
                'Dim i_info As item_info = get_item_info(.CurrentRow.Index, True)
                'If  i_info.row_last_item_row < .CurrentRow.Index Then
                '    s_loading = True
                '    .CurrentCell = .Rows(i_info.row_last_item_row).Cells(.CurrentCell.ColumnIndex)
                '    s_loading = False
                'End If
            End If

        End With
    End Sub

    Private Sub nud_min_per_roll_ValueChanged(sender As Object, e As EventArgs) Handles nud_min_per_roll.ValueChanged
        With dgv_production
            If Not .CurrentRow Is Nothing Then
                get_production_info(.CurrentRow.Index)
            End If
        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GetSerialPortNames()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Button2.Text = ReceiveSerialData(44)
    End Sub

    Private Sub b_main_tare_Click(sender As Object, e As EventArgs) Handles b_main_tare.Click
        tb_scale_tare.Text = tb_scale_weight.Text
    End Sub

    Private Sub b_main_use_weight_Click(sender As Object, e As EventArgs) Handles b_main_use_weight.Click
        Dim kg As String = tb_scale_tared.Text
        Dim est As Decimal = 0
        If IsNumeric(kg) And Not kg = "0" Then
            Dim m As String = tb_roll_length.Text
            If s_department = "printing" Then
                est = FormatNumber(kg * get_conversion_factor(True, tb_mat_num_out.Text, False) * 1000, 0, , , TriState.False)
            Else
                est = FormatNumber(kg * get_conversion_factor(True, tb_mat_num_out.Text, False), 0, , , TriState.False)
            End If
            If cb_report_use_estimate.Checked Then
                m = est
            End If
            m = InputBox("How many Meters?" & vbCr & "Estimated Meters: " & est, "KG Out", m)


            'If s_department = "filmsline" Then
            'Else
            ' m = 0
            'End If
            If IsNumeric(m) Then mod_frm_main.items_out(kg, m, True)

        End If
        s_loading = False
    End Sub

    Private Sub cb_plate_num_lock_CheckedChanged(sender As Object, e As EventArgs) Handles cb_plate_num_lock.CheckedChanged
        If s_loading Then Exit Sub
        dgv_plates.Columns(0).ReadOnly = cb_plate_num_lock.Checked
    End Sub

    Private Sub b_summary_multiple_job_add_Click(sender As Object, e As EventArgs) Handles b_summary_multiple_job_add.Click
        With lb_multiple_job_list
            .Items.Add(dgv_summary.CurrentRow.Cells(0).Value)
            .SelectedIndex = .Items.Count - 1
            calc_times_multiple(False)
            export_summary()
            gb_multiple_job.Visible = True
        End With
        b_summary_multiple_job_add.Enabled = False
    End Sub

    Private Sub lb_multiple_job_list_DoubleClick(sender As Object, e As EventArgs) Handles lb_multiple_job_list.DoubleClick
        With lb_multiple_job_list
            For r As Integer = 0 To dgv_summary.RowCount - 1
                If dgv_summary.Rows(r).Cells(0).Value = .SelectedItem Then
                    dgv_summary.Rows(r).Cells("dgc_summary_multiple").Value = Nothing
                    Exit For
                End If
            Next
            .Items.RemoveAt(.SelectedIndex)
            If .Items.Count > 0 Then
                .SelectedIndex = 0
                calc_times_multiple(False)
            Else
                gb_multiple_job.Visible = False
            End If
        End With
        export_summary()
    End Sub

    Private Sub lb_multiple_job_list_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_multiple_job_list.SelectedIndexChanged
        With lb_multiple_job_list
            If .SelectedIndex > -1 Then
                s_selected_job = dt_schedules.Select("prod_num='" & .SelectedItem & "'")
                With dgv_summary
                    If Not .CurrentRow.Cells(0).Value = lb_multiple_job_list.SelectedItem Then

                        .CurrentCell = Nothing
                        For i As Integer = 0 To .RowCount - 1
                            If .Rows(i).Cells(0).Value = lb_multiple_job_list.SelectedItem Then
                                .CurrentCell = .Rows(i).Cells(0)
                            End If
                        Next
                    End If
                End With
            End If
            display_job()
        End With
    End Sub

    Private Sub tb_scale_tare_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb_scale_tare.KeyPress
        number_only(sender, e, True)
    End Sub

    Private Sub tb_roll_length_Click(sender As Object, e As EventArgs) Handles tb_roll_length.Click
        tb_roll_length.SelectAll()
    End Sub

    Private Sub tb_roll_length_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_roll_length.KeyDown
        If e.KeyValue = Keys.Enter Or e.KeyValue = Keys.Tab Then
            sql.update("db_schedule", "default_quant", tb_roll_length.Text, "prod_num", "=", tb_prod_num.Text)
            s_selected_job(0)("default_quant") = tb_roll_length.Text
            dgv_production.Focus()
        End If
    End Sub

    Private Sub tb_roll_length_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb_roll_length.KeyPress
        number_only(sender, e, False)
    End Sub

    Private Sub tb_roll_length_KeyUp(sender As Object, e As KeyEventArgs) Handles tb_roll_length.KeyUp
        If e.KeyCode = Keys.Enter Then
            dgv_focus(sender, e)
        End If
    End Sub

    Private Sub tb_roll_length_TextChanged(sender As Object, e As EventArgs) Handles tb_roll_length.TextChanged
        If s_loading Then Exit Sub
        If s_extruder Then
            Dim l, o, w, g As Integer, d As Decimal
            Select Case get_material_info(False).formulation
                Case "A203"
                    o = 5
                Case "DL19"
                    o = 5
                Case Else
                    g = g
            End Select
            d = get_item_desity(s_selected_job(0)("mat_desc_semi"))
            l = tb_roll_length.Text
            w = get_material_info(False).width + o
            g = get_material_info(False).gauge
            l_production_kg_in.Text = _
                "Aim Width (mm): " & w & vbCr & _
                "Target Weight: " & FormatNumber(d * w * l * g / 1000000, 0, , , TriState.False) & " kg"
        End If
    End Sub

    Private Sub cbo_mat_num_in_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_mat_num_in.SelectedIndexChanged, cbo_mat_desc_in.SelectedIndexChanged
        If s_loading Then Exit Sub
        s_loading = True
        Dim cbo As ComboBox = CType(sender, ComboBox)
        If cbo_mat_desc_in.Items.Count = cbo_mat_num_in.Items.Count Then
            If cbo Is cbo_mat_desc_in Then
                cbo_mat_num_in.SelectedIndex = cbo.SelectedIndex
            Else
                cbo_mat_desc_in.SelectedIndex = cbo.SelectedIndex
            End If
            s_label_rts = get_label(cbo_mat_num_in.Text, cbo_mat_desc_in.Text, rts:=True)
        End If


        s_loading = False
    End Sub

    Private Sub b_plate_print_list_Click(sender As Object, e As EventArgs) Handles Button3.Click
        xlApp = New Microsoft.Office.Interop.Excel.Application
        With xlApp
            .UserControl = True
            .Visible = True
        End With
        xlWorkBook = xlApp.Workbooks.Open(s_resource_path & "box labels.xlsx", [ReadOnly]:=True)
        xlWorkSheets = xlWorkBook.Worksheets
        xlWorkSheet = xlWorkSheets("Printout")
        Dim dt As DataTable = sql.get_table("db_plate_location", "plate_num", "IS NOT", "NULL")
        For Each item In cbo_plate_boxsize.Items
            xlWorkSheet.Range("A1").Value = item & " Box Plate Locations"
            xlWorkSheet.Range("A4:H41").Value = Nothing
            Dim dr() As DataRow = dt.Select("box_size='" & Strings.Left(item, 1) & "'", "plate_num ASC")
            Dim row As Integer = 4
            Dim col As Integer = 1
            For Each r In dr
                If row = 42 Then
                    row = 4
                    If col = 7 Then
                        col = 1
                        xlWorkSheet.PrintOutEx(Copies:=1)
                        xlWorkSheet.Range("A4:H41").Value = Nothing
                    Else
                        col = col + 3
                    End If
                End If
                xlWorkSheet.Cells(row, col).value = r("plate_num")
                xlWorkSheet.Cells(row, col + 1).value = r("box_size") & r("col").ToString.PadLeft(3, "0") & r("row")
                row = row + 1
            Next
            xlWorkSheet.PrintOutEx(Copies:=1)
        Next
        close_excel()
    End Sub

    Private Sub b_plate_add_size_Click(sender As Object, e As EventArgs) Handles b_plate_add_size.Click
        MsgBox("this does nothing at the moment!")
    End Sub

    Private Sub frm_main_Load(sender As Object, e As EventArgs) Handles Me.Load
        UpdateControls(Me)
    End Sub

    Private Sub b_scheduling_build_Click(sender As Object, e As EventArgs) Handles b_scheduling_build.Click
        With lb_scheduling_departments
            display_scheduled_jobs()
            If .SelectedIndex > -1 Then
                build_schedules(.SelectedItem)
                dt_schedules = sql.get_table("db_schedule", "plant", "=", s_plant & " AND NOT status=10")
            Else
                MsgBox("Please select a department")
            End If
        End With

    End Sub

    Private Sub lb_scheduling_machines_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_scheduling_machines.SelectedIndexChanged
        If lb_scheduling_machines.SelectedIndex = -1 Then Exit Sub
        display_scheduled_jobs()

    End Sub

    Private Sub lb_scheduling_departments_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_scheduling_departments.SelectedIndexChanged
        Dim str() As String
        b_scheduling_upload.Enabled = False
        str = Split(s_machines, ",")
        lb_scheduling_machines.Items.Clear()
        dgv_scheduling.Rows.Clear()
        gb_scheduling_totals.Visible = False
        Dim quant() As quantInfo = Nothing
        Dim total As Integer = 0
        Dim count As Integer = 0
        If jobs Is Nothing Then Exit Sub
        For i As Integer = 0 To UBound(str)
            count = 0
            For c As Integer = 0 To UBound(jobs)
                If Strings.Left(str(i), 6) = jobs(c).machine Then
                    count = count + 1
                End If

            Next

            If count > 0 Then

                If lb_scheduling_departments.SelectedItem = "All Departments" Then
                    lb_scheduling_machines.Items.Add(str(i) & " (" & count & ")")
                    b_scheduling_upload.Enabled = True
                    total = total + count
                Else
                    'b_scheduling_labels.Enabled = False

                    If lb_scheduling_departments.SelectedItem = "bagplant" Then
                        b_scheduling_upload.Enabled = True
                    End If
                    gb_scheduling_totals.Visible = True
                    If get_department(str(i)) = lb_scheduling_departments.SelectedItem Then

                        lb_scheduling_machines.Items.Add(str(i) & " (" & count & ")")
                        total = total + count
                        l_scheduling_totals.Text = "Total Jobs: " & total

                    End If
                End If
            End If
        Next
        For c As Integer = 0 To UBound(jobs)
            If jobs(c).dept = lb_scheduling_departments.SelectedItem Then

                If quant Is Nothing Then
                    ReDim quant(0)
                    quant(0).uom = jobs(c).req_uom_1
                End If

                Dim found1 As Boolean = False
                Dim found2 As Boolean = False

                For q As Integer = 0 To UBound(quant)
                    If quant(q).uom = jobs(c).req_uom_1 Then
                        quant(q).quant = CDec(quant(q).quant) + CDec(jobs(c).req_quant_1)
                        found1 = True
                    End If
                    If quant(q).uom = jobs(c).req_uom_2 Then
                        quant(q).quant = CDec(quant(q).quant) + CDec(jobs(c).req_quant_2)
                        found2 = True
                    End If
                Next

                If Not found1 Or Not found2 Then
                    If Not found1 Then
                        ReDim Preserve quant(quant.Length)
                        quant(quant.Length - 1).uom = jobs(c).req_uom_1
                        quant(quant.Length - 1).quant = jobs(c).req_quant_1
                    ElseIf Not found2 And Not IsNull(jobs(c).req_uom_2) Then
                        ReDim Preserve quant(quant.Length)
                        quant(quant.Length - 1).uom = jobs(c).req_uom_2
                        quant(quant.Length - 1).quant = jobs(c).req_quant_2
                    End If
                End If
            End If
        Next
        If Not quant Is Nothing Then
            If Not lb_scheduling_departments.SelectedItem = "All Departments" Then
                For q As Integer = 0 To UBound(quant)
                    l_scheduling_totals.Text = l_scheduling_totals.Text & vbCr & quant(q).uom.PadRight(3, " ") & FormatNumber(quant(q).quant, 3).ToString.PadLeft(13, " ")
                Next
            End If
        End If

        If lb_scheduling_machines.Items.Count > 1 Then lb_scheduling_machines.Items.Add("All Machines (" & total & ")")
        lb_scheduling_machines.SelectedIndex = lb_scheduling_machines.Items.Count - 1
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        If Not xlApp Is Nothing Then

            If xlApp.Hwnd > 0 Then
                If xlApp.Workbooks.Count = 0 Then
                    close_excel()
                End If
            End If
        End If
        scheduling_setting_save()
    End Sub


    Private Sub b_scheduling_move_up_Click(sender As Object, e As EventArgs) Handles b_scheduling_move_up.Click

        move_lb_item(lb_scheduling_print_order)
        scheduling_setting_save()
    End Sub

    Private Sub b_scheduling_move_down_Click(sender As Object, e As EventArgs) Handles b_scheduling_move_down.Click
        move_lb_item(lb_scheduling_print_order, True)
        scheduling_setting_save()
    End Sub

    Private Sub lb_scheduling_print_order_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_scheduling_print_order.SelectedIndexChanged
        With lb_scheduling_print_order

            If .SelectedIndex = -1 Or .Items.Count = 1 Then
                b_scheduling_move_up.Enabled = False
                b_scheduling_move_down.Enabled = False
            Else
                If .SelectedIndex = 0 Then
                    b_scheduling_move_up.Enabled = False
                Else
                    b_scheduling_move_up.Enabled = True
                End If
                If .SelectedIndex = .Items.Count - 1 Then
                    b_scheduling_move_down.Enabled = False
                Else
                    b_scheduling_move_down.Enabled = True
                End If

            End If
        End With
    End Sub

    Private Sub lb_scheduling_print_departments_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_scheduling_print_departments.SelectedIndexChanged
        Dim s As String = Nothing
        With lb_scheduling_print_departments
            If .SelectedIndex = -1 Or s_loading Then Exit Sub
            nud_scheduling_print_copies.Value = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Print Copies", .SelectedItem, 1)
            s = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Print Order", .SelectedItem, Nothing)
        End With
        With lb_scheduling_print_order

            .Items.Clear()
            If Not s = Nothing Then
                .Items.AddRange(Split(s, ","))
            Else
                For Each item In dt_machines.Rows
                    If item("dept") = lb_scheduling_print_departments.SelectedItem Then
                        .Items.Add(item("machine"))
                    End If
                Next
                scheduling_setting_save()
            End If
        End With
    End Sub

    Private Sub scheduling_format_CheckedChanged(sender As Object, e As EventArgs) Handles cb_scheduling_format_bold.CheckedChanged, _
        cb_scheduling_format_strikeout.CheckedChanged, cb_scheduling_format_underline.CheckedChanged
        If s_loading Then Exit Sub
        format_text(0)
        scheduling_setting_save()

    End Sub

    Private Sub dgv_scheduling_format_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_scheduling_format.CellClick
        s_loading = True
        cb_scheduling_format_bold.Enabled = False
        cb_scheduling_format_strikeout.Enabled = False
        cb_scheduling_format_underline.Enabled = False
        l_scheduling_colour_interior.BackColor = Color.White
        l_scheduling_colour_font.BackColor = Color.Black
        cb_scheduling_format_bold.Checked = False
        cb_scheduling_format_strikeout.Checked = False
        cb_scheduling_format_underline.Checked = False

        If IsNull(dgv_scheduling_format.CurrentRow.Cells(0).Value) Then Exit Sub

        If e.RowIndex > -1 Then
            b_scheduling_format_edit.Enabled = True
            b_scheduling_format_delete.Enabled = True
            cb_scheduling_format_strikeout.Enabled = True
            cb_scheduling_format_bold.Enabled = True
            cb_scheduling_format_underline.Enabled = True
            With dgv_scheduling_format.CurrentRow
                If Not IsNull(.Cells(2).Value) Then
                    Dim str() As String = Split(.Cells(2).Value, ",")
                    For i As Integer = 0 To UBound(str)
                        Select Case True
                            Case str(i).Contains("Bold")
                                cb_scheduling_format_bold.Checked = True
                            Case str(i).Contains("Strikeout")
                                cb_scheduling_format_strikeout.Checked = True
                            Case str(i).Contains("Underline")
                                cb_scheduling_format_underline.Checked = True
                            Case Else
                                str(i) = str(i)
                        End Select
                    Next
                End If
                If Not IsNull(.Cells(3).Value) Then
                    Dim str() As String = Split(.Cells(3).Value, ",")
                    l_scheduling_colour_interior.BackColor = Color.FromName(str(0))
                    l_scheduling_colour_font.BackColor = Color.FromName(str(1))
                End If
                tb_scheduling_format.Text = .Cells(0).Value
                cb_scheduling_format_logo.Checked = .Cells(1).Value
            End With
            format_text(dgv_scheduling_format.CurrentRow.Index)
        Else
            b_scheduling_format_edit.Enabled = False
            b_scheduling_format_delete.Enabled = False
        End If
        s_loading = False
    End Sub


    Private Sub dgv_scheduling_format_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_scheduling_format.CellEndEdit
        format_text(dgv_scheduling_format.CurrentRow.Index)
        scheduling_setting_save()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles b_scheduling_print.Click
        Dim order() As String = Split(ReadIni(Environ("USERPROFILE") & "\settings.ini", "Print Order", lb_scheduling_departments.SelectedItem, Nothing), ",")
        Dim copies As Integer = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Print Copies", lb_scheduling_departments.SelectedItem, 1)
        Dim resp As MsgBoxResult = MsgBox("Print all " & copies & " copies?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, "Print Schedules")
        Select Case resp
            Case MsgBoxResult.Cancel
                Exit Sub
            Case MsgBoxResult.No
                copies = 1
        End Select
        Dim print_missing As Boolean = True
        xlApp.ScreenUpdating = False
        For i As Integer = 1 To copies
            If copies > 1 Then
                If lb_scheduling_departments.SelectedItem = "printing" And dgv_scheduling_missing_pdf.Rows.Count > 0 And print_missing Then
                    xlWorkSheet = xlWorkSheets("Missing PDFs")
                    xlWorkSheet.PrintOutEx(Copies:=1)
                    print_missing = False
                End If
                For Each ws As Microsoft.Office.Interop.Excel.Worksheet In xlWorkBook.Worksheets
                    If ws.Visible Then
                        If ws.Name.Contains("Priority") And LCase(ws.Name).Contains(lb_scheduling_departments.SelectedItem) Then
                            ws.PrintOutEx(Copies:=1)
                            Exit For
                        End If
                    End If
                Next
            End If
            For Each machine In order
                For Each ws As Microsoft.Office.Interop.Excel.Worksheet In xlWorkBook.Worksheets
                    If ws.Visible Then
                        If Not ws.Name.Contains("Priority") And ws.Name.Contains(machine) Then
                            ws.PrintOutEx(Copies:=1)
                            Exit For
                        End If
                    End If
                Next
            Next

        Next
        xlApp.ScreenUpdating = True
        If resp = MsgBoxResult.Yes Then
            export_jobs()
        End If
    End Sub

    Private Sub b_scheduling_font_Click(sender As Object, e As EventArgs) Handles b_scheduling_font.Click
        With ColorDialog1
            .Color = l_scheduling_colour_font.BackColor
            If Not .ShowDialog = Windows.Forms.DialogResult.Cancel Then
                l_scheduling_colour_font.BackColor = .Color
                format_text(dgv_scheduling_format.CurrentRow.Index)
            End If
        End With

    End Sub

    Private Sub b_scheduling_interior_Click(sender As Object, e As EventArgs) Handles b_scheduling_interior.Click
        With ColorDialog1
            .Color = l_scheduling_colour_interior.BackColor
            If Not .ShowDialog = Windows.Forms.DialogResult.Cancel Then
                l_scheduling_colour_interior.BackColor = .Color
                format_text(dgv_scheduling_format.CurrentRow.Index)
            End If

        End With

    End Sub

    Private Sub nud_scheduling_print_copies_ValueChanged(sender As Object, e As EventArgs) Handles nud_scheduling_print_copies.ValueChanged
        If lb_scheduling_print_departments.SelectedIndex > -1 Then
            WriteIni(Environ("USERPROFILE") & "\settings.ini", "Print Copies", lb_scheduling_print_departments.SelectedItem, nud_scheduling_print_copies.Value)
        End If
    End Sub

    Private Sub dgv_scheduling_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_scheduling.CellEndEdit
        With dgv_scheduling
            Dim c_name As String = Replace(.Columns(e.ColumnIndex).Name, "dgc_scheduling_", Nothing)
            Dim dr() As DataRow
            For i As Integer = 0 To UBound(jobs)
                If s_selected_job(0)("prod_num") = jobs(i).prod_num Then
                    Select Case c_name
                        Case "prod_num"
                            jobs(i).prod_num = .CurrentCell.Value
                        Case "req_quant_1"
                            jobs(i).req_quant_1 = .CurrentCell.Value
                            Select Case jobs(i).req_uom_2
                                Case "M"
                                    If jobs(i).mat_num_fin > 0 Then
                                        dr = dt_uom.Select("mat_num=" & jobs(i).mat_num_fin)
                                    Else
                                        dr = dt_uom.Select("mat_num=" & jobs(i).mat_num_semi)
                                    End If
                                    If dr.Length > 0 Then
                                        jobs(i).req_quant_2 = FormatNumber((dr(0)("denom") / dr(0)("numer")) * jobs(i).req_quant_1 * 1000, 2, , , TriState.False)
                                    Else
                                        If jobs(i).mat_num_semi = 0 Then
                                            jobs(i).req_quant_2 = get_roll_length(jobs(i).req_quant_1, jobs(i).mat_desc_fin)
                                        Else
                                            jobs(i).req_quant_2 = get_roll_length(jobs(i).req_quant_1, jobs(i).mat_desc_semi)
                                        End If
                                    End If
                                Case Nothing
                                Case Else
                                    i = i
                            End Select
                            .CurrentRow.Cells("dgc_scheduling_req_quant_2").Value = jobs(i).req_quant_2
                            'MsgBox("add calc here")
                        Case "req_quant_2"
                            jobs(i).req_quant_2 = .CurrentCell.Value
                            MsgBox("add calc here")
                        Case "status"
                            jobs(i).status = .CurrentCell.Value
                        Case Else
                            .CurrentCell.Value = s_selected_job(0)(c_name)
                    End Select
                    Exit For
                End If
            Next
        End With
    End Sub

    Private Sub dgv_scheduling_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_scheduling.CellEnter
        If s_loading Then Exit Sub
        'Dim dt As DataTable = dgv_to_dt(dgv_scheduling)
        If Not dgv_scheduling.Rows(e.RowIndex).Cells("dgc_scheduling_prod_num").Value = Nothing Then
            s_selected_job = dt_schedules.Select("prod_num=" & dgv_scheduling.Rows(e.RowIndex).Cells("dgc_scheduling_prod_num").Value)
        End If
        If Not s_selected_job Is Nothing Then

            If Not s_selected_job.Length = 0 Then

                'If Not s_selected_job(0)("prod_num") = dgv_scheduling.Rows(e.RowIndex).Cells("dgc_scheduling_prod_num").Value Then
                '    s_selected_job = dt.Select("prod_num='" & dgv_scheduling.CurrentRow.Cells("dgc_scheduling_prod_num").Value & "'")

                s_department = s_selected_job(0)("dept")
                tb_prod_num.Text = s_selected_job(0)("prod_num")
                display_job()
                s_department = "planning"
                '  End If
            End If
        End If
    End Sub

    Private Sub cb_production_ply_separator_CheckedChanged(sender As Object, e As EventArgs) Handles cb_production_ply_separator.CheckedChanged
        If s_loading Then Exit Sub
        Dim b As Boolean = cb_production_ply_separator.CheckState
        s_selected_job(0)("ply_separated") = b
        sql.update("db_schedule", "ply_separated", sql.convert_value(b), "prod_num", "=", s_selected_job(0)("prod_num"))
    End Sub

    Private Sub tb_cyl_size_Click(sender As Object, e As EventArgs) Handles tb_cyl_size.Click
        If s_extruder Then tb_cyl_size.SelectAll()

    End Sub

    Private Sub tb_cyl_size_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb_cyl_size.KeyPress
        number_only(sender, e, True)
    End Sub

    Private Sub tb_cyl_size_KeyUp(sender As Object, e As KeyEventArgs) Handles tb_cyl_size.KeyUp
        If e.KeyCode = Keys.Enter Then
            dgv_focus(sender, e)
        End If
    End Sub

    Private Sub tb_cyl_size_LostFocus(sender As Object, e As EventArgs) Handles tb_cyl_size.LostFocus
        If s_extruder Then tb_cyl_size.Text = FormatNumber(tb_cyl_size.Text, 1, , , TriState.False)
    End Sub

    Private Sub tb_cyl_size_TextChanged(sender As Object, e As EventArgs) Handles tb_cyl_size.TextChanged
        If s_loading Or tb_prod_num.Text = Nothing Or Not s_extruder Then Exit Sub
        s_selected_job(0)("cyl_size") = tb_cyl_size.Text
        sql.update("db_schedule", "cyl_size", tb_cyl_size.Text, "prod_num", "=", tb_prod_num.Text)
    End Sub

    Private Sub cbo_job_info_type_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_job_info_type.SelectedIndexChanged
        If s_loading Or tb_prod_num.Text = Nothing Or Not s_extruder Then Exit Sub
        Dim l, w As Integer, d, g, m As Decimal
        Dim o As Integer = 0

        Select Case get_material_info(False).formulation
            Case "A203"
                o = 5
            Case "DL19", "D321"
                o = 5
            Case Else
                g = g
        End Select
        If cbo_job_info_type.Text = "NCF" Then
            m = 2
        Else
            m = 1
        End If
        If IsNull(s_selected_job(0)("mat_desc_semi")) Then
            d = get_item_desity(s_selected_job(0)("mat_desc_fin"))
        Else
            d = get_item_desity(s_selected_job(0)("mat_desc_semi"))
        End If
        l = tb_roll_length.Text
        w = get_material_info(False).width + o
        g = get_material_info(False).gauge
        l_production_kg_in.Text = _
            "Aim Width (mm): " & w & vbCr & _
            "Target Weight: " & FormatNumber(d * w * l * g * m / 1000000, 0, , , TriState.False) & " kg"

        s_selected_job(0)("extrusion_type") = cbo_job_info_type.Text
        sql.update("db_schedule", "extrusion_type", "'" & cbo_job_info_type.Text & "'", "prod_num", "=", tb_prod_num.Text)
    End Sub


    Private Sub dgv_scheduling_label_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_scheduling_label.CellClick

        With dgv_scheduling_label.CurrentRow
            If .Index = -1 Then
                b_scheduling_label_edit.Enabled = False
                b_scheduling_label_delete.Enabled = False
            Else
                b_scheduling_label_delete.Enabled = True
                b_scheduling_label_edit.Enabled = True
                tb_scheduling_label.Text = .Cells(0).Value
                cbo_scheduling_labels.Text = .Cells(1).Value
            End If
        End With
    End Sub

    Private Sub b_scheduling_format_add_Click(sender As Object, e As EventArgs) Handles b_scheduling_format_add.Click
        With dgv_scheduling_format
            Dim r As Integer = .RowCount
            .Rows.Add()
            .Rows(r).Cells(0).Value = tb_scheduling_format.Text
            .Rows(r).Cells(1).Value = cb_scheduling_format_logo.Checked
            .Rows(r).Cells(2).Value = tb_scheduling_format.Font.Style.ToString
            .Rows(r).Cells(3).Value = l_scheduling_colour_interior.BackColor.Name & "," & l_scheduling_colour_font.BackColor.Name
            scheduling_setting_save()

        End With
    End Sub

    Private Sub b_scheduling_format_edit_Click(sender As Object, e As EventArgs) Handles b_scheduling_format_edit.Click
        With dgv_scheduling_format.CurrentRow
            .Cells(0).Value = tb_scheduling_format.Text
            .Cells(1).Value = cb_scheduling_format_logo.Checked
            .Cells(2).Value = tb_scheduling_format.Font.Style.ToString
            .Cells(3).Value = l_scheduling_colour_interior.BackColor.Name & "," & l_scheduling_colour_font.BackColor.Name
            b_scheduling_format_add.Enabled = False
            scheduling_setting_save()
        End With
    End Sub

    Private Sub b_scheduling_label_add_Click(sender As Object, e As EventArgs) Handles b_scheduling_label_add.Click
        With dgv_scheduling_label
            Dim r As Integer = .RowCount
            .Rows.Add()
            .Rows(r).Cells(0).Value = tb_scheduling_label.Text
            .Rows(r).Cells(1).Value = cbo_scheduling_labels.Text
            scheduling_setting_save()
        End With
    End Sub

    Private Sub b_scheduling_label_edit_Click(sender As Object, e As EventArgs) Handles b_scheduling_label_edit.Click
        With dgv_scheduling_label.CurrentRow
            .Cells(0).Value = tb_scheduling_label.Text
            .Cells(1).Value = cbo_scheduling_labels.Text
            b_scheduling_label_add.Enabled = False
            scheduling_setting_save()
        End With
    End Sub

    Private Sub tb_scheduling_format_TextChanged(sender As Object, e As EventArgs) Handles tb_scheduling_format.TextChanged
        Dim add As Boolean = True
        Dim t() As String = Split(UCase(tb_scheduling_format.Text), ",")
        With dgv_scheduling_format
            For r As Integer = 0 To .RowCount - 1
                Dim s() As String = Split(UCase(.Rows(r).Cells(0).Value), ",")
                For txt As Integer = 0 To UBound(t)
                    For i As Integer = 0 To UBound(s)
                        If s(i) = t(txt) Then
                            add = False
                        End If
                    Next
                Next
            Next
        End With
        If tb_scheduling_format.Text = Nothing Then
            b_scheduling_format_add.Enabled = add
            b_scheduling_format_edit.Enabled = False
        Else
            b_scheduling_format_add.Enabled = add
            b_scheduling_format_edit.Enabled = False
            If dgv_scheduling_format.RowCount > 0 Then
                If dgv_scheduling_format.CurrentRow.Index > -1 Then
                    b_scheduling_format_edit.Enabled = True
                End If
            End If
        End If
    End Sub

    Private Sub tb_scheduling_label_TextChanged(sender As Object, e As EventArgs) Handles tb_scheduling_label.TextChanged
        Dim add As Boolean = True
        Dim t() As String = Split(UCase(tb_scheduling_label.Text), ",")
        With dgv_scheduling_label
            For r As Integer = 0 To .RowCount - 1
                Dim s() As String = Split(UCase(.Rows(r).Cells(0).Value), ",")
                For txt As Integer = 0 To UBound(t)
                    For i As Integer = 0 To UBound(s)
                        If s(i) = t(txt) Then
                            add = False
                        End If
                    Next
                Next
            Next
        End With
        If tb_scheduling_label.Text = Nothing Then
            b_scheduling_label_add.Enabled = add
            b_scheduling_label_edit.Enabled = False
        Else
            b_scheduling_label_add.Enabled = add
            b_scheduling_label_edit.Enabled = False
            If dgv_scheduling_label.RowCount > 0 Then
                If dgv_scheduling_label.CurrentRow.Index > -1 Then
                    b_scheduling_label_edit.Enabled = True
                End If
            End If

        End If
    End Sub

    Private Sub b_scheduling_remove_job_Click(sender As Object, e As EventArgs) Handles b_scheduling_remove_job.Click
        Dim prod_num As Long = dgv_scheduling.CurrentRow.Cells(0).Value
        Dim i_d As Integer = lb_scheduling_departments.SelectedIndex
        Dim i_m As Integer = lb_scheduling_machines.SelectedIndex
        Dim j(jobs.Length - 2) As JobInfo
        Dim j_i As Integer = 0
        For i As Integer = 0 To jobs.Length - 1
            If Not jobs(i).prod_num = prod_num Then
                j(j_i) = jobs(i)
                j_i = j_i + 1
            End If
        Next
        jobs = j
        lb_scheduling_departments.SelectedIndex = -1
        lb_scheduling_departments.SelectedIndex = i_d
        lb_scheduling_machines.SelectedIndex = i_m
        ' display_scheduled_jobs()
    End Sub

    Private Sub ProductionReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductionReportToolStripMenuItem.Click
        ReprintTicketsToolStripMenuItem.Enabled = s_symple_access
    End Sub

    Private Sub b_scheduling_labels_Click(sender As Object, e As EventArgs) Handles b_scheduling_labels.Click
        '  If lb_scheduling_departments.SelectedItem = "bagplant" Then
        Dim dtable As DataTable = sql.get_table("db_schedule", "dept", "=", "'bagplant' AND status=-2")
        If dtable Is Nothing Then
            MsgBox("No jobs found to print!")
            Exit Sub
        End If
        Dim dr() As DataRow = dtable.Select("status=-2", "prod_num ASC")
        Dim resp As MsgBoxResult = MsgBox(dr.Length & " jobs have been found, do you want to proceed with printing?", MsgBoxStyle.YesNo)
        If resp = MsgBoxResult.No Then Exit Sub
        If Not acsis.running(0) Then
            With frm_logon
                .StartPosition = FormStartPosition.CenterParent
                .ShowDialog()
            End With
        End If
        s_shift = "1"
        Dim prod_num As Long
        If acsis.running(0) Then
            For r As Integer = 0 To dr.Length - 1
                prod_num = dr(r)("prod_num")
                s_selected_job = dtable.Select("prod_num=" & prod_num)
                tssl_status.Text = "Status: Printing " & r + 1 & " of " & dr.Length & " (" & prod_num & ")"
                Me.Refresh()
                '  Next
                'With dgv_scheduling
                '    For r As Integer = 0 To .RowCount - 1
                '        prod_num = .Rows(r).Cells("dgc_scheduling_prod_num").Value
                '        If .Rows(r).Cells("dgc_scheduling_dept").Value = lb_scheduling_departments.SelectedItem And _
                '            Strings.Left(prod_num, 1) = "5" And _
                '            .Rows(r).Cells("dgc_scheduling_status").Value = -1 Then
                acsis.open_job_screen(prod_num, False)
                If Not s_popup = 12 Then
                    acsis.click_button(acsis.get_button_access("Label Reprint", 0), False, 0)
                    Dim special As Boolean = False
                    For i As Integer = 0 To dgv_scheduling_label.RowCount - 1
                        Dim s() As String = Split(dgv_scheduling_label.Rows(i).Cells(0).Value, ",")
                        For Each customer In s
                            'If .Rows(r).Cells("dgc_scheduling_customer").Value.ToString.Contains(customer) Then
                            If dr(r)("customer").ToString.Contains(customer) Then
                                acsis.select_cb_item(dgv_scheduling_label.Rows(i).Cells(1).Value, 0)
                                special = True
                                Exit For
                            End If
                        Next
                        If special Then Exit For
                    Next
                    Dim machine As String = dr(r)("machine")
                    Dim text_to_find As String = "Carton Quantity"
                    If Strings.Left(machine, 4) = "RWGE" Then
                        ' If Not special Then
                        s_label = acsis.get_acsis_option("label_rollbag")
                        'End If
                        text_to_find = "Length"
                    ElseIf Not special Then
                        s_label = get_label(dr(r)("mat_num_fin"), dr(r)("mat_desc_fin"))
                    End If
                    tssl_symple_info_label.Text = "Printer: " & s_printer & " Label: " & s_label
                    Me.Refresh()
                    acsis.select_cb_item(s_label, 0)

                    acsis.select_cb_item(s_printer, 0)
                    Dim quant As Integer = 0 ' 
                    If IsNull(dr(r)("items_out")) Then
                        acsis.select_grid_item(text_to_find, True)
                        quant = Clipboard.GetText
                        sql.update("db_schedule", "items_out", quant & ",status=-3", "prod_num", "=", prod_num & " AND plant=" & s_plant)
                    Else
                        quant = dr(r)("items_out")
                    End If
                    Dim cq As Integer = CInt(dr(r)("req_quant_1") * 1000) / quant
                    acsis.insert_string(quant, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("production_label_quant")).handle)
                    acsis.insert_string(1, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("production_label_from")).handle)
                    acsis.insert_string(cq + 3, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("production_label_to")).handle)
                    acsis.click_button(acsis.get_acsis_option("production_label_confirm"), False, 0)
                    'System.Threading.Thread.Sleep(2000)
                    acsis.click_button(acsis.get_button_back(0), False, 0)
                End If
                s_popup = 0
            Next
            'sql.update("db_schedule", "status", -3, "status", "=", -2 & " AND dept='bagplant' AND plant=" & s_plant)

            'End With
        Else
            MsgBox("You need to logon to print labels")
        End If
        'End If
        tssl_status.Text = "Status: Printing Complete!"

    End Sub

    Private Sub TabControl1_SelectedIndexChanged1(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        With TabControl1
            If .SelectedTab.Text = "Setup" Then
                If tb_scheduling_format.Text = Nothing Then dgv_scheduling_format.ClearSelection()

                If tb_scheduling_label.Text = Nothing Then dgv_scheduling_label.ClearSelection()
            End If
        End With
    End Sub

    Private Sub cb_main_ofline_mode_CheckedChanged(sender As Object, e As EventArgs) Handles cb_main_ofline_mode.CheckedChanged
        If s_loading Then Exit Sub
        s_loading = True
        frm_startup.cb_ofline_mode.Checked = cb_main_ofline_mode.Checked
        s_loading = False

    End Sub

    Private Sub b_scheduling_upload_Click(sender As Object, e As EventArgs) Handles b_scheduling_upload.Click
        export_jobs()
        MsgBox("Upload Complete!")
    End Sub

    Private Sub b_scheduling_format_delete_Click(sender As Object, e As EventArgs) Handles b_scheduling_format_delete.Click
        Dim resp As MsgBoxResult = MsgBox("Are you sure you want to delete this item?", MsgBoxStyle.YesNo)
        If resp = MsgBoxResult.Yes Then
            dgv_scheduling_format.Rows.RemoveAt(dgv_scheduling_format.CurrentRow.Index)
            b_scheduling_format_edit.Enabled = False
            b_scheduling_format_delete.Enabled = False
            b_scheduling_format_add.Enabled = True
            scheduling_setting_save()
        End If
    End Sub

    Private Sub b_scheduling_label_delete_Click(sender As Object, e As EventArgs) Handles b_scheduling_label_delete.Click
        Dim resp As MsgBoxResult = MsgBox("Are you sure you want to delete this item?", MsgBoxStyle.YesNo)
        If resp = MsgBoxResult.Yes Then
            dgv_scheduling_label.Rows.RemoveAt(dgv_scheduling_label.CurrentRow.Index)
            b_scheduling_label_edit.Enabled = False
            b_scheduling_label_delete.Enabled = False
            b_scheduling_label_add.Enabled = True
            scheduling_setting_save()
        End If
    End Sub


    Private Sub tb_scale_tare_TextChanged(sender As Object, e As EventArgs) Handles tb_scale_tare.TextChanged
        If IsNumeric(tb_scale_tare.Text) And Not s_loading Then
            WriteIni(Environ("USERPROFILE") & "\settings.ini", "Scale", "Tare", tb_scale_tare.Text)
        End If
    End Sub
End Class

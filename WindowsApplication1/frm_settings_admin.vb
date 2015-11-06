Imports System.Text
Public Class frm_settings_admin

    Dim password As String

    Private Sub tb_password_TextChanged(sender As Object, e As EventArgs) Handles tb_password.TextChanged
        If tb_password.Text = password Then
            TabControl1.Enabled = True
            tb_password_change.Enabled = True
        Else
            TabControl1.Enabled = False
            tb_password_change.Enabled = False
        End If
    End Sub


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub b_save_Click(sender As Object, e As EventArgs) Handles b_save.Click
        If tb_password_plate.Text = Nothing Then
            MsgBox("There is no password set for the plate inventory." & vbCr & "Please set it now.")
            TabControl1.SelectedIndex = 2
            tb_password_plate.Focus()
            Exit Sub
        Else
            sql.update("settings_file_locations", "location", "'" & tb_symple_location.Text & "'", "name='symple' AND plant", "=", s_plant)
            sql.delete("settings_general", "plant", "=", s_plant & " AND item LIKE 'general_%'")
            sql.insert("settings_general", "item,val,plant", "'general_ink_systems','" & tb_ink_systems.Text & "'," & s_plant)
            sql.insert("settings_general", "item,val,plant", "'general_uom1','" & tb_uom_1.Text & "'," & s_plant)
            sql.insert("settings_general", "item,val,plant", "'general_uom2','" & tb_uom_1.Text & "'," & s_plant)
            sql.insert("settings_general", "item,val,plant", "'general_password','" & tb_password.Text & "'," & s_plant)
            sql.insert("settings_general", "item,val,plant", "'general_password_plate','" & tb_password_plate.Text & "'," & s_plant)
            s_pass_plate = tb_password_plate.Text

        End If

        Me.Close()
    End Sub

    Private Sub dgv_downtime_KeyDown(sender As Object, e As KeyEventArgs) Handles dgv_downtime.KeyDown
        With dgv_downtime
            If e.Control Then
                If e.KeyCode = Keys.V Then
                    .CurrentRow.Cells(.SelectedCells(0).ColumnIndex).Value = Clipboard.GetText()
                End If
            End If
        End With

    End Sub

    Private Sub settings_admin_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        s_loading = True
        s_plant = ReadIni(Environ("USERPROFILE") & "\settings.ini", "Location", "Plant", 3203)
        cbo_plant.DataSource = sql.get_table("db_plants", "plant", "IS NOT", "NULL")
        cbo_plant.DisplayMember = "name"
        select_cbo_item(cbo_plant, s_plant, "plant")
        hideunhide()
        load_plant()
        load_deparment()
        s_loading = False

    End Sub
    Sub load_plant()
        Dim result() As DataRow
        Dim str() As String
        Dim dt_settings_general, dt_settings_acsis As DataTable

        dt_settings_general = sql.get_table("settings_general", "plant", "=", s_plant)
        dt_settings_acsis = sql.get_table("acsis_options", "plant", "=", s_plant & " AND version='" & s_version & "'")
        result = dt_settings_general.Select("item='general_password'")

        If result.Length > 0 Then
            password = result(0)("val")
        Else
            MsgBox("There is no password set, please set it now.")
            tb_password_change.Enabled = True
            tb_password_change.Focus()
            tb_password.Enabled = False
            Exit Sub
        End If

        result = dt_settings_general.Select("item='downtime_codes'")
        If result.Length > 0 Then
            dgv_downtime_codes.Rows.Clear()
            str = Split(result(0)("val"), ",")
            For i As Integer = 0 To str.Length - 1
                dgv_downtime_codes.Rows.Add(str(i))
                cbo_downtime_code.Items.Add(str(i))
            Next
            cbo_downtime_code.SelectedIndex = 0
        End If

        cbo_department_select.DataSource = sql.get_table("db_departments", "plant", "=", s_plant)
        cbo_department_select.DisplayMember = "dept"
        select_cbo_item(cbo_department_select, s_department, "dept")

        tb_symple_location.Text = sql.read("settings_file_locations", "plant=" & s_plant & " AND name", "=", "'symple'", "location").ToString
        tb_symple_version.Text = FileVersionInfo.GetVersionInfo(tb_symple_location.Text).FileVersion

        result = dt_settings_general.Select("item='general_ink_systems'")
        If result.Length = 0 Then
            sql.insert("settings_general", "item,plant", "'general_ink_systems'," & s_plant)
        Else
            tb_ink_systems.Text = result(0)("val")
        End If

        result = dt_settings_general.Select("item ='general_uom1'")
        If result.Length = 0 Then
            sql.insert("settings_general", "item,plant", "'general_uom1'," & s_plant)
        Else
            tb_uom_1.Text = result(0)("val")
        End If

        result = dt_settings_general.Select("item ='general_uom2'")
        If result.Length = 0 Then
            sql.insert("settings_general", "item,plant", "'general_uom2'," & s_plant)
        Else
            tb_uom_2.Text = result(0)("val")
        End If

        result = dt_settings_general.Select("item ='general_password_plate'")
        If result.Length = 0 Then
            sql.insert("settings_general", "item,plant", "'general_password_plate'," & s_plant)
        Else
            tb_password_plate.Text = result(0)("val")
        End If

        Dim l1, l2, l3, l4, l5, l6, p1, p2, p3, p4, p5, p6 As DataTable
        l1 = sql.get_table("DP_Print_Templates", "Label_Location", "=", "'Full' AND Label_Type='Item' ORDER BY List_Order", "symple")
        l2 = l1.Copy
        l3 = l1.Copy
        l4 = l1.Copy
        l5 = l1.Copy
        l6 = l1.Copy
        cbo_label_barrier.DataSource = l1
        cbo_label_barrier.DisplayMember = "Template_Code"
        cbo_label_film.DataSource = l2
        cbo_label_film.DisplayMember = "Template_Code"
        cbo_label_laminate.DataSource = l3
        cbo_label_laminate.DisplayMember = "Template_Code"
        cbo_label_rollbag.DataSource = l4
        cbo_label_rollbag.DisplayMember = "Template_Code"
        cbo_label_semi_finished.DataSource = l5
        cbo_label_semi_finished.DisplayMember = "Template_Code"
        cbo_label_carton.DataSource = l6
        cbo_label_carton.DisplayMember = "Template_Code"

        p1 = sql.get_table("DP_Print_Templates", "Label_Location", "=", "'Full' AND Label_Type='Pallet' ORDER BY List_Order", "symple")
        p2 = p1.Copy
        p3 = p1.Copy
        p4 = p1.Copy
        p5 = p1.Copy
        p6 = p1.Copy

        cbo_pallet_barrier.DataSource = p1
        cbo_pallet_barrier.DisplayMember = "Template_Code"
        cbo_pallet_film.DataSource = p2
        cbo_pallet_film.DisplayMember = "Template_Code"
        cbo_pallet_laminate.DataSource = p3
        cbo_pallet_laminate.DisplayMember = "Template_Code"
        cbo_pallet_rollbag.DataSource = p4
        cbo_pallet_rollbag.DisplayMember = "Template_Code"
        cbo_pallet_carton.DataSource = p5
        cbo_pallet_carton.DisplayMember = "Template_Code"
        cbo_pallet_semi_finished.DataSource = p6
        cbo_pallet_semi_finished.DisplayMember = "Template_Code"

        result = dt_settings_acsis.Select("item='label_barrier'")
        If result.Length = 0 Then
            sql.insert("acsis_options", "item,plant,version", "'label_label_barrier'," & s_plant & ",'" & s_version & "'")
        Else
            select_cbo_item(cbo_label_barrier, result(0)("val").ToString, "Template_Code")
        End If

        result = dt_settings_acsis.Select("item='label_film'")
        If result.Length = 0 Then
            sql.insert("acsis_options", "item,plant,version", "'label_film'," & s_plant & ",'" & s_version & "'")
        Else
            select_cbo_item(cbo_label_film, result(0)("val").ToString, "Template_Code")
        End If

        result = dt_settings_acsis.Select("item='label_laminate'")
        If result.Length = 0 Then
            sql.insert("acsis_options", "item,plant,version", "'label_laminate'," & s_plant & ",'" & s_version & "'")
        Else
            select_cbo_item(cbo_label_laminate, result(0)("val").ToString, "Template_Code")
        End If

        result = dt_settings_acsis.Select("item='label_rollbag'")
        If result.Length = 0 Then
            sql.insert("acsis_options", "item,plant,version", "'label_rollbag'," & s_plant & ",'" & s_version & "'")
        Else
            select_cbo_item(cbo_label_rollbag, result(0)("val").ToString, "Template_Code")
        End If

        result = dt_settings_acsis.Select("item='label_carton'")
        If result.Length = 0 Then
            sql.insert("acsis_options", "item,plant,version", "'label_carton'," & s_plant & ",'" & s_version & "'")
        Else
            select_cbo_item(cbo_label_carton, result(0)("val").ToString, "Template_Code")
        End If

        result = dt_settings_acsis.Select("item='label_semi_finished'")
        If result.Length = 0 Then
            sql.insert("acsis_options", "item,plant,version", "'label_semi_finished'," & s_plant & ",'" & s_version & "'")
        Else
            select_cbo_item(cbo_label_semi_finished, result(0)("val").ToString, "Template_Code")
        End If

        result = dt_settings_acsis.Select("item='pallet_barrier'")
        If result.Length = 0 Then
            sql.insert("acsis_options", "item,plant,version", "'pallet_barrier'," & s_plant & ",'" & s_version & "'")
        Else
            select_cbo_item(cbo_pallet_barrier, result(0)("val").ToString, "Template_Code")
        End If

        result = dt_settings_acsis.Select("item='pallet_film'")
        If result.Length = 0 Then
            sql.insert("acsis_options", "item,plant,version", "'pallet_film'," & s_plant & ",'" & s_version & "'")
        Else
            select_cbo_item(cbo_pallet_film, result(0)("val").ToString, "Template_Code")
        End If

        result = dt_settings_acsis.Select("item='pallet_laminate'")
        If result.Length = 0 Then
            sql.insert("acsis_options", "item,plant,version", "'pallet_laminate'," & s_plant & ",'" & s_version & "'")
        Else
            select_cbo_item(cbo_pallet_laminate, result(0)("val").ToString, "Template_Code")
        End If

        result = dt_settings_acsis.Select("item='pallet_rollbag'")
        If result.Length = 0 Then
            sql.insert("acsis_options", "item,plant,version", "'pallet_rollbag'," & s_plant & ",'" & s_version & "'")
        Else
            select_cbo_item(cbo_pallet_rollbag, result(0)("val").ToString, "Template_Code")
        End If

        result = dt_settings_acsis.Select("item='pallet_carton'")
        If result.Length = 0 Then
            sql.insert("acsis_options", "item,plant,version", "'pallet_carton'," & s_plant & ",'" & s_version & "'")
        Else
            select_cbo_item(cbo_pallet_carton, result(0)("val").ToString, "Template_Code")
        End If

        result = dt_settings_acsis.Select("item='pallet_semi_finished'")
        If result.Length = 0 Then
            sql.insert("acsis_options", "item,plant,version", "'pallet_semi_finished'," & s_plant & ",'" & s_version & "'")
        Else
            select_cbo_item(cbo_pallet_semi_finished, result(0)("val").ToString, "Template_Code")
        End If

        result = dt_settings_acsis.Select("item='navigation'")
        If result.Length > 0 Then
            str = Split(result(0)("val"), ",")
            cbo_navigate_target.Items.Clear()
            For i As Integer = 0 To str.Length - 1
                cbo_navigate_target.Items.Add(str(i))
            Next
        End If
    End Sub

    Sub load_deparment()
        Dim dt As DataTable = sql.get_table("settings_email", "dept", "=", "'" & s_department & "'")
        Select Case s_department
            Case "filmsline"
                cb_department_extruder.Visible = True
            Case Else
                cb_department_extruder.Visible = False
        End Select

        With dgv_email
            .Rows.Clear()
            If Not dt Is Nothing Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    .Rows.Add(dt(i)("name"), dt(i)("email"), dt(i)("manager"), dt(i)("graphics"))
                Next
            End If

        End With
        get_downtime()
    End Sub

    Private Sub cb_downtime_new_section_CheckedChanged(sender As Object, e As EventArgs) Handles cb_downtime_new_section.CheckedChanged
        If cb_downtime_new_section.Checked Then
            tb_section_main.Visible = True
            cbo_section_main.Visible = False
        Else
            tb_section_main.Visible = False
            cbo_section_main.Visible = True
        End If
    End Sub

    Private Sub b_downtime_edit_Click(sender As Object, e As EventArgs) Handles b_downtime_edit.Click
        'Button2.Text = "Save Item"
        hideunhide()
    End Sub

    Sub hideunhide()
        Dim h As Integer = 40
        If Me.Height > b_save.Location.Y + b_save.Height + h Then
            Me.Height = b_save.Location.Y + b_save.Height + h
            b_downtime_edit.Text = "Edit Items"
        Else
            Me.Height = b_downtime_add.Location.Y + b_downtime_add.Height + h
            b_downtime_edit.Text = "Finish Edit"
        End If
    End Sub

    Private Sub cbo_label_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_label_barrier.SelectedIndexChanged, _
        cbo_label_film.SelectedIndexChanged, cbo_label_laminate.SelectedIndexChanged, cbo_label_rollbag.SelectedIndexChanged, _
        cbo_label_semi_finished.SelectedIndexChanged, cbo_pallet_barrier.SelectedIndexChanged, cbo_pallet_film.SelectedIndexChanged, _
        cbo_pallet_laminate.SelectedIndexChanged, cbo_pallet_rollbag.SelectedIndexChanged, cbo_label_carton.SelectedIndexChanged, _
        cbo_pallet_carton.SelectedIndexChanged, cbo_pallet_semi_finished.SelectedIndexChanged

        If s_loading Then Exit Sub
        Select Case Me.ActiveControl.Name
            Case "cbo_label_barrier"
                sql.update("acsis_options", "val", "'" & cbo_label_barrier.Text & "'", "item='label_barrier' AND plant", "=", s_plant & " AND version='" & s_version & "'")
            Case "cbo_label_film"
                sql.update("acsis_options", "val", "'" & cbo_label_film.Text & "'", "item='label_film' AND plant", "=", s_plant & " AND version='" & s_version & "'")
            Case "cbo_label_laminate"
                sql.update("acsis_options", "val", "'" & cbo_label_laminate.Text & "'", "item='label_laminate' AND plant", "=", s_plant & " AND version='" & s_version & "'")
            Case "cbo_label_rollbag"
                sql.update("acsis_options", "val", "'" & cbo_label_rollbag.Text & "'", "item='label_rollbag' AND plant", "=", s_plant & " AND version='" & s_version & "'")
            Case "cbo_label_carton"
                sql.update("acsis_options", "val", "'" & cbo_label_carton.Text & "'", "item='label_carton' AND plant", "=", s_plant & " AND version='" & s_version & "'")
            Case "cbo_label_semi_finished"
                sql.update("acsis_options", "val", "'" & cbo_label_semi_finished.Text & "'", "item='label_semi_finished' AND plant", "=", s_plant & " AND version='" & s_version & "'")
            Case "cbo_pallet_barrier"
                sql.update("acsis_options", "val", "'" & cbo_pallet_barrier.Text & "'", "item='pallet_barrier' AND plant", "=", s_plant & " AND version='" & s_version & "'")
            Case "cbo_pallet_film"
                sql.update("acsis_options", "val", "'" & cbo_pallet_film.Text & "'", "item='pallet_film' AND plant", "=", s_plant & " AND version='" & s_version & "'")
            Case "cbo_pallet_laminate"
                sql.update("acsis_options", "val", "'" & cbo_pallet_laminate.Text & "'", "item='pallet_laminate' AND plant", "=", s_plant & " AND version='" & s_version & "'")
            Case "cbo_pallet_tubing"
                sql.update("acsis_options", "val", "'" & cbo_pallet_rollbag.Text & "'", "item='pallet_tubing' AND plant", "=", s_plant & " AND version='" & s_version & "'")
            Case "cbo_pallet_carton"
                sql.update("acsis_options", "val", "'" & cbo_pallet_carton.Text & "'", "item='pallet_carton' AND plant", "=", s_plant & " AND version='" & s_version & "'")
            Case "cbo_pallet_semi_finished"
                sql.update("acsis_options", "val", "'" & cbo_pallet_semi_finished.Text & "'", "item='pallet_semi_finished' AND plant", "=", s_plant & " AND version='" & s_version & "'")
        End Select
        dt_acsis_options = sql.get_table("acsis_options", "plant", "=", s_plant & " AND version='" & s_version & "'")
    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_downtime.CellEnter
        If dgv_downtime.CurrentRow Is Nothing Then Exit Sub
        Dim pl, co, ma, re As String

        With dgv_downtime
            pl = .CurrentRow.Cells("plant").Value
            co = .CurrentRow.Cells("code").Value.ToString
            ma = .CurrentRow.Cells("main").Value
            re = .CurrentRow.Cells("reason").Value
            .Rows(.CurrentRow.Index).Selected = True
        End With

        Dim dt As DataTable = sql.get_table("settings_downtime", "plant", "=", s_plant & " AND dept='" & s_department & "'")

        cbo_section_main.Items.Clear()
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim str As String = dt(i)("main")
            If cbo_section_main Is Nothing Then
                cbo_section_main.Items.Add(str)

            ElseIf cbo_section_main.Items.IndexOf(str) = -1 Then
                cbo_section_main.Items.Add(str)
            End If
        Next
        For i As Integer = 0 To cbo_section_main.Items.Count - 1
            If cbo_section_main.Items(i) = ma Then
                cbo_section_main.SelectedIndex = i
            End If
        Next
        For i As Integer = 0 To cbo_downtime_code.Items.Count - 1
            If cbo_downtime_code.Items(i) = co Then
                cbo_downtime_code.SelectedIndex = i
            End If
        Next
        tb_section_reason.Text = re
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles b_downtime_add.Click
        ' Button2.Text = "Add Item"
        Dim ma, re, co
        If cb_downtime_new_section.Checked Then
            ma = tb_section_main.Text
        Else
            ma = cbo_section_main.Text
        End If
        re = tb_section_reason.Text
        co = cbo_downtime_code.Text
        sql.insert("settings_downtime", "main,reason,dept,plant,code,row", "'" & ma & "','" & re & "','" & s_department & "'," & _
                   s_plant & ",'" & co & "'," & dgv_downtime.RowCount)

        get_downtime()
        dgv_downtime.CurrentCell = dgv_downtime.Rows(dgv_downtime.RowCount - 1).Cells(0)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles b_downtime_save.Click
        Dim ma, re, co As String, row As Integer
        row = dgv_downtime.CurrentRow.Index
        If cb_downtime_new_section.Checked Then
            ma = tb_section_main.Text
        Else
            ma = cbo_section_main.Text
        End If
        re = tb_section_reason.Text
        co = cbo_downtime_code.Text
        sql.update("settings_downtime", "main='" & ma & "', reason='" & re & "', dept='" & s_department & "', plant=" & s_plant & _
                   ", code", "'" & co & "'", "row", "=", dgv_downtime.CurrentRow.Cells("row").Value)

        get_downtime()
        dgv_downtime.CurrentCell = dgv_downtime.Rows(row).Cells(0)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles b_downtime_remove.Click
        With dgv_downtime
            If .RowCount = 0 Then Exit Sub
            Dim row, r As Integer
            sql.delete("settings_downtime", "plant", "=", s_plant & " AND dept='" & cbo_department_select.Text & "'")
            r = 0
            row = .CurrentRow.Index
            For i As Integer = 0 To .RowCount - 1
                If Not i = row Then
                    sql.insert("settings_downtime", "main,reason,dept,plant,code,row", "'" & .Rows(i).Cells("main").Value & _
                               "','" & .Rows(i).Cells("reason").Value & "','" & .Rows(i).Cells("dept").Value & "'," & s_plant & _
                               ",'" & .Rows(i).Cells("code").Value & "'," & r)

                    r = r + 1
                Else
                    r = r
                End If
            Next
        End With

        get_downtime()

    End Sub

    Private Sub DataGridView2_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_email.CellEndEdit
        export_email()
    End Sub

    Sub export_email()
        With dgv_email
            If Not IsNull(.CurrentRow.Cells(0).Value) And Not IsNull(.CurrentRow.Cells(1).Value) Then
                sql.delete("settings_email", "plant", "=", s_plant)
                For r As Integer = 0 To .RowCount - 2
                    Dim m As Integer = 0
                    Dim g As Integer = 0
                    If Not IsNull(.Rows(r).Cells(2).Value) Then
                        If .Rows(r).Cells(2).Value = "True" Then
                            m = 1
                        End If
                    End If
                    If Not IsNull(.Rows(r).Cells(3).Value) Then
                        If .Rows(r).Cells(3).Value.ToString = "True" Then
                            g = 1
                        End If
                    End If
                    sql.insert("settings_email", "name,email,plant,dept,manager,graphics", "'" & .Rows(r).Cells(0).Value & "','" & _
                               .Rows(r).Cells(1).Value & "'," & s_plant & ",'" & cbo_department_select.Text & "'," & m & "," & g)

                Next
            End If
        End With
    End Sub

    Private Sub dgv_email_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv_email.CellMouseUp
        If s_loading Then Exit Sub
        With dgv_email

            If .CurrentCell.ColumnIndex = 2 Or .CurrentCell.ColumnIndex = 3 Then
                Dim s As String = .CurrentCell.Value.ToString
                If s = "True" Then
                    .CurrentCell.Value = False
                Else
                    .CurrentCell.Value = True
                End If
                export_email()
            End If
        End With
    End Sub

    Private Sub b_password_change_Click(sender As Object, e As EventArgs) Handles b_password_change.Click
        Dim s As Object = sql.read("settings_general", "item", "=", "'password' AND plant=" & s_plant, "item")
        password = tb_password_change.Text
        tb_password_change.Clear()
        If IsNull(s) Then
            sql.insert("settings_general", "item,val,plant", "'password','" & password & "'," & s_plant)
            tb_password.Enabled = True
        Else
            sql.update("settings_general", "val", "'" & tb_password.Text & "'", "item='password' AND plant", "=", s_plant)
        End If
        tb_password.Text = password
    End Sub

    Private Sub b_navigate_refresh_Click(sender As Object, e As EventArgs) Handles b_navigate_refresh.Click
        RefreshControls(True)
        enable_button()

    End Sub

    Sub RefreshControls(ByVal all As Boolean)
        s_loading = True
        acsis.get_controls(0, True, False, True)
        Dim access As String = cbo_navigate_access.Text
        Dim back As String = cbo_navigate_exit.Text
        For Each ctrl In gb_navigate_main.Controls
            If TypeOf ctrl Is ComboBox Then
                Dim cbx As ComboBox = ctrl
                cbx.Enabled = True
                If Not cbx Is cbo_navigate_target Then cbx.Items.Clear()
            End If
        Next
        For i As Integer = 0 To w_sessions(0).winWnd(0).btnWnd.Length - 1
            cbo_navigate_access.Items.Add(w_sessions(0).winWnd(0).btnWnd(i).text(0))
            cbo_navigate_exit.Items.Add(w_sessions(0).winWnd(0).btnWnd(i).text(0))
        Next
        If Not all And w_sessions(0).winWnd.Length = 1 Then
            cbo_navigate_access.Text = access
            cbo_navigate_exit.Text = back
        End If
        If w_sessions(0).winWnd.Length = 1 Then

            If Not w_sessions(0).winWnd(0).optWnd Is Nothing Then
                For i As Integer = 0 To w_sessions(0).winWnd(0).optWnd.Length - 1
                    If w_sessions(0).winWnd(0).optWnd(i).enabled Then
                        cbo_navigate_selection.Items.Add(w_sessions(0).winWnd(0).optWnd(i).text(0))
                    End If
                    s_loading = True

                    Select Case i
                        Case 0
                            cbo_navigate_opt_1.Items.Add(w_sessions(0).winWnd(0).optWnd(i).text(0))
                            cbo_navigate_opt_1.SelectedIndex = 0
                        Case 1
                            cbo_navigate_opt_2.Items.Add(w_sessions(0).winWnd(0).optWnd(i).text(0))
                            cbo_navigate_opt_2.SelectedIndex = 0
                        Case 2
                            cbo_navigate_opt_3.Items.Add(w_sessions(0).winWnd(0).optWnd(i).text(0))
                            cbo_navigate_opt_3.SelectedIndex = 0
                        Case Else
                            cbo_navigate_opt_1.Items.Add(w_sessions(0).winWnd(0).optWnd(i).text(0))
                            cbo_navigate_opt_2.Items.Add(w_sessions(0).winWnd(0).optWnd(i).text(0))
                            cbo_navigate_opt_3.Items.Add(w_sessions(0).winWnd(0).optWnd(i).text(0))
                    End Select
                    s_loading = False
                    'ComboBox17.Items.Add(get_window_controls.optWnd(i).text(0))
                    'ComboBox18.Items.Add(get_window_controls.optWnd(i).text(0))
                Next
            End If
        Else
            cbo_navigate_opt_1.Enabled = False
            cbo_navigate_opt_2.Enabled = False
            cbo_navigate_opt_3.Enabled = False
            cbo_navigate_selection.Enabled = False
            If Not cbo_navigate_target.Text = "Production Confirmation" Then cbo_navigate_access.Enabled = False
        End If

        tb_navigate_current_window.Text = w_sessions(0).winWnd(0).name
        tb_navigate_window_count.Text = w_sessions(0).winWnd.Length
        enable_button()
        If Not cbo_navigate_target.Text = Nothing And Not dt_navigation Is Nothing Then
            dr_navigation = dt_navigation.Select("sub_target='" & cbo_navigate_target.Text & "' AND target_reached=1")
            If dr_navigation.Length = 1 Then
                cbo_navigate_access.Enabled = False
                cbo_navigate_exit.Enabled = False
                cbo_navigate_selection.Enabled = False
                b_navigate_set.Enabled = True
            End If
        End If


        s_loading = False

    End Sub

    Sub SelectOption(ByVal name As String, ByVal checked As Boolean)
        If s_loading Or Not checked Or name = Nothing Then Exit Sub
        For i As Integer = 0 To w_sessions(0).winWnd(0).optWnd.Length - 1
            If w_sessions(0).winWnd(0).optWnd(i).text(0) = name Then
                acsis.set_check(w_sessions(0).winWnd(0).optWnd(i).handle, checked)
                Exit For
            End If
        Next
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedTab.Text = "Sym-PLE Menu Navigation" Then
            Dim s() As String, file As String = get_file_path("symple")
            If Not w_sessions Is Nothing Then
                s_version = FileVersionInfo.GetVersionInfo(w_sessions(0).process.MainModule.FileName).FileVersion
            Else
                s_version = FileVersionInfo.GetVersionInfo(file).FileVersion
            End If
            Dim login As String = sql.read("acsis_options", "plant", "=", s_plant & " AND item='login_id' AND version='" & s_version & "'", "val")
            If login Is Nothing Then
                If Not acsis.running(0) Then
                    w_sessions(0).process = New Process
                    With w_sessions(0).process
                        s = Split(file, "\")
                        Dim n As String = s(s.Length - 1)
                        .StartInfo.FileName = file
                        .StartInfo.WorkingDirectory = Replace(file, n, Nothing)

                        .StartInfo.CreateNoWindow = True
                        .Start()
                        .WaitForInputIdle()

                    End With
                    ' get_session(True)
                    setup_login()
                Else
                    RefreshControls(True)
                    reset_navigation(gb_navigate_main)
                End If
            Else
                If acsis.running(0) Then acsis.get_logon_details(0)
                If s_pass = Nothing And s_user = Nothing Then
                    With frm_logon
                        .StartPosition = FormStartPosition.CenterScreen
                        .ShowDialog()
                    End With
                ElseIf Not acsis.running(0) Then
                    acsis.logon()
                End If
                RefreshControls(True)
                reset_navigation(gb_navigate_main)
            End If
        End If
        tb_navigation_version.Text = s_version
    End Sub

    Private Sub b_navigate_go_Click(sender As Object, e As EventArgs) Handles b_navigate_go.Click
        acsis.get_controls(0, True, False, True)

        show_status("Navigating to next window")

        For i As Integer = 0 To w_sessions(0).winWnd(0).optWnd.Length - 1
            If cbo_navigate_selection.Text = w_sessions(0).winWnd(0).optWnd(i).text(0) Then
                acsis.set_check(w_sessions(0).winWnd(0).optWnd(i).handle, True)
                Exit For
            End If
        Next
        For i As Integer = 0 To w_sessions(0).winWnd(0).btnWnd.Length - 1
            If cbo_navigate_access.Text = w_sessions(0).winWnd(0).btnWnd(i).text(0) Then
                acsis.click_button(w_sessions(0).winWnd(0).btnWnd(i).text(0), False, 0)
                Exit For
            End If
        Next
        With dgv_navigate
            .Rows.Add(cbo_navigate_access.Text, cbo_navigate_exit.Text, cbo_navigate_opt_1.Text, cbo_navigate_opt_2.Text, _
                      cbo_navigate_opt_3.Text, nud_navigate_step.Value, cbo_navigate_selection.Text, _
                      tb_navigate_current_window.Text, cbo_navigate_target.Text, _
                      tb_navigate_window_count.Text, cbo_navigate_target.Text)

        End With
        RefreshControls(False)
        export_navigation()
        nud_navigate_step.Value = define_step(True)
        frm_status.Close()
    End Sub

    Sub copy_combobox(ByVal cbx_from As ComboBox, ByVal cbx_to As ComboBox, Optional ByVal enable As Boolean = True)
        cbx_to.Items.Clear()
        For Each s As String In cbx_from.Items
            If Not s = cbx_from.Text Then
                cbx_to.Items.Add(s)
                If Not IsNumeric(s) Then cbx_to.DropDownWidth = 150
            End If
        Next
        cbx_to.Enabled = enable
    End Sub
    Enum efill_type
        button
        radio
        textbox
        combobox
        checkbox
        _nothing
    End Enum
    Sub fill_combobox(ByVal cbx1 As ComboBox, Optional ByVal cbx2 As ComboBox = Nothing, _
                        Optional ByVal type As efill_type = efill_type._nothing, Optional ByVal n As Integer = 0, _
                        Optional ByVal enable1 As Boolean = True, Optional ByVal enable2 As Boolean = True, _
                        Optional ByVal keep_values As Boolean = True)

        s_loading = True
        Dim s1 As String = cbx1.Text
        Dim s2 As String = Nothing
        cbx1.Items.Clear()
        If Not cbx2 Is Nothing Then
            s2 = cbx2.Text
            cbx2.Items.Clear()
        End If
        Select Case type
            Case efill_type.button
                For i As Integer = 0 To w_sessions(0).winWnd(0).btnWnd.Length - 1
                    If Not w_sessions(0).winWnd(0).btnWnd(i).text(0) = Me.ActiveControl.Text And Not _
                        w_sessions(0).winWnd(0).btnWnd(i).text(0) = s2 Then
                        cbx1.Items.Add(w_sessions(0).winWnd(0).btnWnd(i).text(0))
                        If Not IsNumeric(w_sessions(0).winWnd(0).btnWnd(i).text(0)) Then cbx1.DropDownWidth = 150
                    End If
                    If Not cbx2 Is Nothing Then
                        If Not w_sessions(0).winWnd(0).btnWnd(i).text(0) = Me.ActiveControl.Text And Not _
                            w_sessions(0).winWnd(0).btnWnd(i).text(0) = s1 Then
                            cbx2.Items.Add(w_sessions(0).winWnd(0).btnWnd(i).text(0))
                            If Not IsNumeric(w_sessions(0).winWnd(0).btnWnd(i).text(0)) Then cbx2.DropDownWidth = 150
                        End If
                    End If
                Next
            Case efill_type.radio
                For i As Integer = 0 To w_sessions(0).winWnd(0).optWnd.Length - 1
                    If Not w_sessions(0).winWnd(0).optWnd(i).text(0) = Me.ActiveControl.Text And Not _
                        w_sessions(0).winWnd(0).optWnd(i).text(0) = s2 Then
                        cbx1.Items.Add(w_sessions(0).winWnd(0).optWnd(i).text(0))
                        If Not IsNumeric(w_sessions(0).winWnd(0).optWnd(i).text(0)) Then cbx1.DropDownWidth = 150
                    End If
                    If Not cbx2 Is Nothing Then
                        If Not w_sessions(0).winWnd(0).optWnd(i).text(0) = Me.ActiveControl.Text And Not w_sessions(0).winWnd(0).optWnd(i).text(0) = s1 Then
                            cbx2.Items.Add(w_sessions(0).winWnd(0).optWnd(i).text(0))
                            If Not IsNumeric(w_sessions(0).winWnd(0).optWnd(i).text(0)) Then cbx2.DropDownWidth = 150
                        End If
                    End If

                Next
            Case efill_type.textbox
                For i As Integer = 0 To n
                    If Not i.ToString = Me.ActiveControl.Text And Not i.ToString = s2 Then
                        cbx1.Items.Add(i)
                        If Not IsNumeric(i) Then cbx1.DropDownWidth = 150
                    End If
                    If Not cbx2 Is Nothing Then
                        If Not i.ToString = Me.ActiveControl.Text And Not i.ToString = s1 Then
                            cbx2.Items.Add(i)
                            If Not IsNumeric(i) Then cbx2.DropDownWidth = 150
                        End If
                    End If

                Next
            Case efill_type.combobox
                For i As Integer = 0 To w_sessions(0).winWnd(0).cbWnd(n).text.Length - 1
                    If Not w_sessions(0).winWnd(0).cbWnd(n).text(i) = Me.ActiveControl.Text And Not w_sessions(0).winWnd(0).cbWnd(n).text(i) = s2 Then
                        cbx1.Items.Add(w_sessions(0).winWnd(0).cbWnd(n).text(i))
                        If Not IsNumeric(w_sessions(0).winWnd(0).cbWnd(n).text(i)) Then cbx1.DropDownWidth = 150
                    End If
                    If Not cbx2 Is Nothing Then
                        If Not w_sessions(0).winWnd(0).cbWnd(n).text(i) = Me.ActiveControl.Text And Not w_sessions(0).winWnd(0).cbWnd(n).text(i) = s1 Then
                            cbx2.Items.Add(w_sessions(0).winWnd(0).cbWnd(n).text(i))
                            If Not IsNumeric(w_sessions(0).winWnd(0).cbWnd(i).text(i)) Then cbx2.DropDownWidth = 150
                        End If
                    End If

                Next
            Case efill_type.checkbox
                For i As Integer = 0 To w_sessions(0).winWnd(0).cbxWnd.Length - 1
                    If Not w_sessions(0).winWnd(0).cbxWnd(i).text(0) = Me.ActiveControl.Text And Not w_sessions(0).winWnd(0).cbxWnd(i).text(0) = s2 Then
                        cbx1.Items.Add(w_sessions(0).winWnd(0).cbxWnd(i).text(0))
                        If Not IsNumeric(w_sessions(0).winWnd(0).cbxWnd(i).text(0)) Then cbx1.DropDownWidth = 150
                    End If
                    If Not cbx2 Is Nothing Then
                        If Not w_sessions(0).winWnd(0).cbxWnd(i).text(0) = Me.ActiveControl.Text And Not w_sessions(0).winWnd(0).cbxWnd(i).text(0) = s1 Then
                            cbx2.Items.Add(w_sessions(0).winWnd(0).cbxWnd(i).text(0))
                            If Not IsNumeric(w_sessions(0).winWnd(0).cbxWnd(i).text(0)) Then cbx2.DropDownWidth = 150
                        End If
                    End If

                Next

            Case Else
                cbx1 = cbx1
        End Select
        If keep_values Then cbx1.Text = s1
        cbx1.Enabled = enable1
        If Not cbx2 Is Nothing Then
            If keep_values Then cbx2.Text = s2
            cbx2.Enabled = enable2
        End If
        s_loading = False
    End Sub

    Sub enable_button()
        b_navigate_go.Enabled = False
        b_navigate_target_reached.Enabled = False
        If cbo_navigate_access.Enabled = False And Not cbo_navigate_exit.Text = Nothing _
            Or Not cbo_navigate_access.Text = Nothing And Not cbo_navigate_exit.Text = Nothing Then
            b_navigate_target_reached.Enabled = True
        End If
        If Not cbo_navigate_access.Text = Nothing And Not cbo_navigate_exit.Text = Nothing Then
            If w_sessions(0).winWnd.Length = 1 Then
                If Not cbo_navigate_opt_1.Text = Nothing And Not cbo_navigate_opt_2.Text = Nothing _
                    And Not cbo_navigate_opt_3.Text = Nothing And Not cbo_navigate_selection.Text = Nothing _
                    And Not cbo_navigate_target.Text = Nothing Then
                    b_navigate_go.Enabled = True
                End If
            End If
        End If
        If s_loading Then Exit Sub
        Select Case Me.ActiveControl.Name
            Case cbo_navigate_access.Name
                fill_combobox(cbo_navigate_access, cbo_navigate_exit, efill_type.button)
            Case cbo_navigate_exit.Name
                If cbo_navigate_access.Enabled Then fill_combobox(cbo_navigate_exit, cbo_navigate_access, efill_type.button)
            Case cbo_navigate_opt_1.Name
                fill_combobox(cbo_navigate_opt_2, cbo_navigate_opt_3, efill_type.radio)
                fill_combobox(cbo_navigate_opt_3, cbo_navigate_opt_2, efill_type.radio)
            Case cbo_navigate_opt_2.Name
                fill_combobox(cbo_navigate_opt_1, cbo_navigate_opt_3, efill_type.radio)
                fill_combobox(cbo_navigate_opt_3, cbo_navigate_opt_1, efill_type.radio)
            Case cbo_navigate_opt_3.Name
                fill_combobox(cbo_navigate_opt_1, cbo_navigate_opt_2, efill_type.radio)
                fill_combobox(cbo_navigate_opt_2, cbo_navigate_opt_1, efill_type.radio)
        End Select
    End Sub

    Private Sub opt_navigate_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_navigate_opt_1.SelectedIndexChanged, _
        cbo_navigate_opt_2.SelectedIndexChanged, cbo_navigate_opt_3.SelectedIndexChanged, cbo_navigate_exit.SelectedIndexChanged, _
        cbo_navigate_access.SelectedIndexChanged, cbo_navigate_selection.SelectedIndexChanged
        enable_button()
    End Sub

    Sub goto_start(ByVal refresh As Boolean)
        With dgv_navigate
            Dim s As String = Nothing, target As String = Nothing, found As Boolean = False

            show_status("Returning to start...")

            'exit current window
            acsis.get_controls(0, True, False, True)
            Do While w_sessions(0).winWnd.Length > 1
                acsis.get_controls(w_sessions(0).winWnd(w_sessions(0).winWnd.Length - 1).handle, True, False, True)
                If Not dt_navigation Is Nothing Then
                    dr_navigation = dt_navigation.Select("window_count=" & w_sessions(0).winWnd.Length & " AND opt_target='" & acsis.get_selected_option(0) & "'")
                    If dr_navigation.Length > 0 Then s = dr_navigation(0)("button_exit")
                End If

                ' s = sql.read("acsis_navigation", "plant", "=", s_plant & " AND window_name='" & get_window_controls.winWnd(0).text(0) & "'", "button_exit")
                If s = Nothing Then
                    frm_status.Close()
                    MsgBox("You are in a screen that has not been recognised." & vbCr & _
                           "Please exit the current screen in Sym-PLE screen.")
                    SetForegroundWindow(w_sessions(0).process.MainWindowHandle)
                    'add auto menu detection with timer here
                    MsgBox("Click OK to proceed")
                    acsis.get_controls(0, True, False, True)
                    acsis.get_controls(w_sessions(0).winWnd(w_sessions(0).winWnd.Length - 1).handle, True, False, True)
                    If Not dt_navigation Is Nothing Then
                        dr_navigation = dt_navigation.Select("window_count=" & w_sessions(0).winWnd.Length & " AND opt_target='" & acsis.get_selected_option(0) & "'")
                        If dr_navigation.Length > 0 Then s = dr_navigation(0)("button_exit")
                        ' s = sql.read("acsis_navigation", "plant", "=", s_plant & " AND window_name='" & get_window_controls.winWnd(0).text(0) & "'", "button_exit")
                    End If
                End If

                show_status("Returning to start...")


                If Not dt_navigation Is Nothing Then
                    If dr_navigation.Length > 0 Then
                        target = dr_navigation(0)("target")
                        acsis.click_button(dr_navigation(0)("button_exit"), False, 0)
                    End If
                End If
            Loop
            'find current position in Sym-PLE
            If Not target = Nothing Then
                dr_navigation = dt_navigation.Select("target='" & target & _
                                              "' AND window_name='" & w_sessions(0).winWnd(0).name & "'")

            Else
                If dt_navigation Is Nothing Then
                    'put something here
                Else
                    For i As Integer = 0 To w_sessions(0).winWnd(0).optWnd.Length - 1
                        dr_navigation = dt_navigation.Select("opt_1='" & w_sessions(0).winWnd(0).optWnd(i).text(0) & _
                                                      "' AND window_name='" & w_sessions(0).winWnd(0).name & "'")

                        If dr_navigation.Length > 0 Then
                            If dr_navigation.Length = 1 Then
                                found = True
                                Exit For
                            Else
                                For o As Integer = 0 To w_sessions(0).winWnd(0).optWnd.Length - 1
                                    If s = Nothing Then
                                        s = w_sessions(0).winWnd(0).optWnd(o).text(0)
                                    Else
                                        s = s & "," & w_sessions(0).winWnd(0).optWnd(o).text(0)
                                    End If
                                Next
                                For r As Integer = 0 To dr_navigation.Length - 1
                                    If s.Contains(dr_navigation(r)("opt_1")) And _
                                       s.Contains(dr_navigation(r)("opt_2")) And _
                                       s.Contains(dr_navigation(r)("opt_3")) Then

                                        dr_navigation = dt_navigation.Select("target='" & dt_navigation(r)("target") & _
                                                                      "' AND step<=" & dt_navigation(r)("step"))

                                        found = True
                                        Exit For
                                    End If
                                Next
                                If found Then Exit For
                            End If

                        End If
                    Next
                End If

                If Not found Then
                    frm_status.Close()
                    MsgBox("A menu has been detected that is not recorded" & vbCr & "Please navigate to the very start of Sym-PLE.")
                    SetForegroundWindow(w_sessions(0).process.MainWindowHandle)
                    MsgBox("Click OK if you are at the start of the Sym-PLE menus")
                    RefreshControls(True)
                    Exit Sub
                End If

            End If
            nud_navigate_step.Value = define_step(False)

            If dr_navigation.Length > 0 Then
                nud_navigate_step.Value = define_step(True)
                For i As Integer = dr_navigation.Length - 1 To 0 Step -1
                    Dim dr As DataRow() = dt_navigation.Select("step='" & i & "'")
                    acsis.click_button(dr(0)("button_exit"), False, 0)
                Next
            End If

            frm_status.Close()
            If refresh Then RefreshControls(True)
        End With
    End Sub

    Private Sub b_navigate_target_reached_Click(sender As Object, e As EventArgs) Handles b_navigate_target_reached.Click
        With dgv_navigate
            If .RowCount = 0 Then
                MsgBox("You have not chosen anything!")
                Exit Sub
            End If
            .Rows.Add(cbo_navigate_access.Text, cbo_navigate_exit.Text, cbo_navigate_opt_1.Text, cbo_navigate_opt_1.Text, _
                      cbo_navigate_opt_2.Text, nud_navigate_step.Value, acsis.get_selected_option(0), _
                      tb_navigate_current_window.Text, cbo_navigate_target.Text, _
                       w_sessions(0).winWnd.Length, cbo_navigate_target.Text, 1)

            export_navigation()
        End With
        nud_navigate_step.Value = define_step(True)

        Select Case cbo_navigate_target.Text
            Case "Production Confirmation"
                setup_production()
            Case "Print Pallet Labels"
                setup_pallet()
            Case "Stock Overview"
                setup_overview()
            Case "Warehouse Reciepts - Other"
                setup_rts()
            Case "Re-Labeling"
                setup_relabeling()
            Case "Modify Pallet Label"
                setup_modify_pallet()
            Case Else
                Dim s As String = "d"
        End Select
    End Sub

    Private Sub cbo_navigate_target_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles cbo_navigate_target.SelectedIndexChanged

        goto_start(True)
        With dgv_navigate
            .Rows.Clear()

            dt_navigation = sql.get_table("acsis_navigation", "plant", "=", s_plant & _
                                          " AND sub_target='" & cbo_navigate_target.Text & "' AND version='" & s_version & "'")
            If Not dt_navigation Is Nothing Then
                For i As Integer = 0 To dt_navigation.Rows.Count - 1
                    Dim ok, back, opt1, opt2, opt3, stp, opt_target, target, window, _
                        window_count, sub_target, target_reached

                    ok = dt_navigation(i)("button_access").ToString
                    back = dt_navigation(i)("button_exit").ToString
                    opt1 = dt_navigation(i)("opt_1").ToString
                    opt2 = dt_navigation(i)("opt_2").ToString
                    opt3 = dt_navigation(i)("opt_3").ToString
                    stp = dt_navigation(i)("step").ToString
                    opt_target = dt_navigation(i)("opt_target").ToString
                    window = dt_navigation(i)("window_name").ToString
                    target = dt_navigation(i)("target").ToString
                    window_count = dt_navigation(i)("window_count").ToString
                    sub_target = dt_navigation(i)("sub_target").ToString
                    target_reached = dt_navigation(i)("target_reached").ToString

                    .Rows.Add(ok, back, opt1, opt2, opt3, stp, opt_target, window, target, _
                              window_count, sub_target, target_reached)

                Next

                dr_navigation = dt_navigation.Select("sub_target='" & cbo_navigate_target.Text & "' AND target_reached=1")
                If dr_navigation.Length > 0 Then

                    Select Case cbo_navigate_target.Text
                        Case "Production Confirmation"
                            Dim dt As DataTable
                            dt = sql.get_table("acsis_options", "item", "LIKE", "'production_%' AND plant=" & s_plant)
                            If dt Is Nothing Then setup_production()
                        Case "Print Pallet Labels"
                            Dim dt As DataTable
                            dt = sql.get_table("acsis_options", "item", "LIKE", "'pallet_screen_%' AND plant=" & s_plant)
                            If dt Is Nothing Then setup_pallet()
                        Case "Stock Overview"
                            If IsNull(.Rows(.RowCount - 1).Cells(9).Value) Then
                                setup_overview()
                            End If
                        Case "Warehouse Reciepts - Other"
                            Dim dt As DataTable
                            dt = sql.get_table("acsis_options", "item", "LIKE", "'rts_screen_%' AND plant=" & s_plant)
                            If dt Is Nothing Then setup_rts()
                        Case "Re-Labeling"
                            Dim dt As DataTable
                            dt = sql.get_table("acsis_options", "item", "LIKE", "'relabel_screen_%' AND plant=" & s_plant)
                            If dt Is Nothing Then setup_relabeling()
                        Case "Modify Pallet Label"
                            Dim dt As DataTable
                            dt = sql.get_table("acsis_options", "item", "LIKE", "'modify_pallet_%' AND plant=" & s_plant)
                            If dt Is Nothing Then setup_modify_pallet()
                        Case Else
                            Dim s As String = "s"
                    End Select
                Else
                    MsgBox("The steps for this navigation are not complete")
                    goto_window()
                    RefreshControls(True)
                End If
            End If
            nud_navigate_step.Value = define_step(True)
            If .RowCount > 0 Then
                b_navigate_set.Enabled = True
            Else
                b_navigate_set.Enabled = False
            End If
        End With
    End Sub

    Function define_step(ByVal next_step As Boolean) As Integer
        define_step = 0
        With dgv_navigate
            If Not .RowCount = 0 Then
                For i As Integer = 0 To .RowCount - 1
                    If .Rows(i).Cells(5).Value > define_step Then
                        define_step = .Rows(i).Cells(5).Value
                    End If
                Next
                If next_step Then define_step = define_step + 1
            End If

        End With
    End Function

    Private Sub b_navigate_test_Click(sender As Object, e As EventArgs) Handles b_navigate_test.Click
        SetForegroundWindow(w_sessions(0).process.MainWindowHandle)
        goto_window()
        RefreshControls(True)
    End Sub

    Private Sub b_production_next_Click(sender As Object, e As EventArgs) Handles b_production_next.Click
        With dgv_navigate
            If Not acsis.running(0) Then
                acsis.logon()
            End If
            acsis.get_controls(0, True, False, True)
            acsis.get_controls(w_sessions(0).winWnd(0).handle, True, False, True)

            SetForegroundWindow(w_sessions(0).process.MainWindowHandle)
            SendMessage(w_sessions(0).winWnd(0).tbWnd(acsis.get_textbox(cls_ACSIS.eProdbox.Production_Number, 0, 0)).handle, WM_SETFOCUS, 0, 0)
            w_sessions(0).process.WaitForInputIdle()
            acsis.insert_string(tb_production_prod_num.Text, w_sessions(0).winWnd(0).tbWnd(acsis.get_textbox(cls_ACSIS.eProdbox.Production_Number, 0, 0)).handle)
            SendMessage(w_sessions(0).winWnd(0).tbWnd(acsis.get_textbox(cls_ACSIS.eProdbox.User, 0, 0)).handle, WM_SETFOCUS, 0, 0)

            w_sessions(0).process.WaitForInputIdle()
            acsis.popup_check(True, 0)
            If s_popup = 4 Then
                While Not s_popup = 0
                    s_popup = 0
                    acsis.popup_check(True, 0)
                End While
                MsgBox(tb_production_prod_num.Text & " is an invalid Production Order Number!")
                Exit Sub
            End If
            acsis.click_button(.Rows(.RowCount - 1).Cells(0).Value, False, 0)
            acsis.get_controls(0, True, False, True)
            acsis.get_controls(w_sessions(0).winWnd(0).handle, True, False, True)
            tb_navigate_current_window.Text = w_sessions(0).winWnd(0).name
            .Rows.Add(.Rows(.RowCount - 2).Cells(0).Value, .Rows(.RowCount - 2).Cells(1).Value, Nothing, Nothing, _
                      Nothing, nud_navigate_step.Value, acsis.get_selected_option(0), tb_navigate_current_window.Text, "Production Entry", _
                      w_sessions(0).winWnd.Length, cbo_navigate_target.Text)

            nud_navigate_step.Value = define_step(True)
            export_navigation()
            SetForegroundWindow(Me.Handle)
            fill_combobox(cbo_production_label, type:=efill_type.button)
            b_production_next.Enabled = False
            tb_production_prod_num.Enabled = False
        End With

    End Sub

    Private Sub tb_navigate_prod_num_TextChanged(sender As Object, e As EventArgs) Handles tb_production_prod_num.TextChanged
        If tb_production_prod_num.Text.Length = 10 Then
            b_production_next.Enabled = True
        Else
            b_production_next.Enabled = False
        End If
    End Sub

    Private Sub menu_button_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_production_label.SelectedIndexChanged, _
        cbo_production_view_prod.SelectedIndexChanged, cbo_production_setup.SelectedIndexChanged, cbo_production_run.SelectedIndexChanged, _
        cbo_production_downtime.SelectedIndexChanged, cbo_production_final.SelectedIndexChanged
        acsis.get_controls(0, True, False, True)
        Dim cbx As ComboBox = CType(sender, ComboBox)
        If Not cbx.Text = Nothing Then

            show_status("Opening a window by clicking " & cbx.Text)

            cbx.Enabled = False
            acsis.click_button(cbx.Text, False, 0)
            acsis.get_controls(w_sessions(0).winWnd(0).handle, False, True, True)
            Select Case cbx.Name
                Case cbo_production_label.Name
                    fill_combobox(cbo_production_label_confirm, type:=efill_type.button, keep_values:=False)
                Case cbo_production_view_prod.Name
                    fill_combobox(cbo_production_view_prod_back, type:=efill_type.button, keep_values:=False)
                Case cbo_production_setup.Name
                    fill_combobox(cbo_production_setup_confirm, type:=efill_type.button, keep_values:=False)
                Case cbo_production_run.Name
                    fill_combobox(cbo_production_run_confirm, type:=efill_type.button, keep_values:=False)
                Case cbo_production_downtime.Name
                    fill_combobox(cbo_production_downtime_confirm, type:=efill_type.button, keep_values:=False)
                Case cbo_production_final.Name
                    fill_combobox(cbo_production_final_confirm, type:=efill_type.button, keep_values:=False)
            End Select
            tb_navigate_current_window.Text = w_sessions(0).winWnd(0).name
            tb_navigate_window_count.Text = w_sessions(0).winWnd.Length
            frm_status.Close()
            SetForegroundWindow(Me.Handle)
        End If
    End Sub

    Private Sub production_combobox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_production_label_back.SelectedIndexChanged, _
        cbo_production_view_prod_back.SelectedIndexChanged, cbo_production_setup_back.SelectedIndexChanged, cbo_production_run_confirm.SelectedIndexChanged, _
        cbo_production_run_back.SelectedIndexChanged, cbo_production_downtime_start.SelectedIndexChanged, cbo_production_downtime_end.SelectedIndexChanged, _
        cbo_production_downtime_back.SelectedIndexChanged, cbo_production_downtime_start.SelectedIndexChanged, cbo_production_final_back.SelectedIndexChanged, _
        cbo_production_setup_start.SelectedIndexChanged, cbo_production_setup_end.SelectedIndexChanged, cbo_production_run_start.SelectedIndexChanged, _
        cbo_production_run_end.SelectedIndexChanged, cbo_production_label_quant.SelectedIndexChanged, cbo_production_run_quant.SelectedIndexChanged, _
        cbo_production_label_to.SelectedIndexChanged, cbo_production_label_from.SelectedIndexChanged, cbo_production_label_confirm.SelectedIndexChanged, _
        cbo_production_downtime_confirm.SelectedIndexChanged, cbo_production_setup_confirm.SelectedIndexChanged, cbo_production_final_confirm.SelectedIndexChanged, cbo_production_final_check.SelectedIndexChanged

        Dim cbx As ComboBox = CType(sender, ComboBox), export As Boolean = False
        If Not cbx.Text = Nothing Then
            With cbx
                Dim name As String = cbx.Name
                Dim str() As String = Split(name, "_")
                cbx.Enabled = False
                Select Case True
                    Case name.Contains("_label_")
                        Select Case str(str.Length - 1)
                            Case "confirm"
                                copy_combobox(cbx, cbo_production_label_back)
                            Case "back"
                                cbo_production_label_quant.Items.Clear()
                                For i As Integer = 0 To w_sessions(0).winWnd(0).tbWnd.Length - 1
                                    cbo_production_label_quant.Items.Add(i)
                                Next
                                cbo_production_label_quant.Enabled = True
                            Case "quant"
                                copy_combobox(cbx, cbo_production_label_from)
                            Case "from"
                                copy_combobox(cbx, cbo_production_label_to)
                            Case "to"
                                With dgv_navigate
                                    .Rows.Add(cbo_production_label.Text, cbo_production_label_back.Text, Nothing, Nothing, _
                                              Nothing, nud_navigate_step.Value, acsis.get_selected_option(0), _
                                              tb_navigate_current_window.Text, "Label Reprint", w_sessions(0).winWnd.Length, _
                                              cbo_navigate_target.Text)

                                End With
                                acsis.click_button(cbo_production_label_back.Text, False, 0)
                                copy_combobox(cbo_production_label, cbo_production_view_prod)
                        End Select
                    Case name.Contains("_view_prod_")
                        With dgv_navigate
                            .Rows.Add(cbo_production_view_prod.Text, cbo_production_view_prod_back.Text, Nothing, Nothing, _
                                      Nothing, nud_navigate_step.Value, acsis.get_selected_option(0), tb_navigate_current_window.Text, _
                                      "View Production Order", w_sessions(0).winWnd.Length, cbo_navigate_target.Text)

                        End With
                        acsis.click_button(cbo_production_view_prod_back.Text, False, 0)
                        copy_combobox(cbo_production_view_prod, cbo_production_setup)
                    Case name.Contains("_setup_")
                        Select Case str(str.Length - 1)
                            Case "confirm"
                                copy_combobox(cbx, cbo_production_setup_back)
                            Case "back"
                                cbo_production_setup_start.Items.Clear()
                                For i As Integer = 0 To w_sessions(0).winWnd(0).tbWnd.Length - 1
                                    cbo_production_setup_start.Items.Add(i)
                                Next
                                cbo_production_setup_start.Enabled = True
                            Case "start"
                                copy_combobox(cbx, cbo_production_setup_end)
                            Case "end"
                                With dgv_navigate
                                    .Rows.Add(cbo_production_setup.Text, cbo_production_setup_back.Text, Nothing, Nothing, _
                                              Nothing, nud_navigate_step.Value, acsis.get_selected_option(0), tb_navigate_current_window.Text, _
                                              "Setup Confirmation", w_sessions(0).winWnd.Length, cbo_navigate_target.Text)

                                End With
                                acsis.click_button(cbo_production_setup_back.Text, False, 0)
                                copy_combobox(cbo_production_setup, cbo_production_run)
                        End Select
                    Case name.Contains("_run_")
                        Select Case str(str.Length - 1)
                            Case "confirm"
                                copy_combobox(cbx, cbo_production_run_back)
                            Case "back"
                                cbo_production_run_start.Items.Clear()
                                For i As Integer = 0 To w_sessions(0).winWnd(0).tbWnd.Length - 1
                                    cbo_production_run_start.Items.Add(i)
                                Next
                                cbo_production_run_start.Enabled = True
                            Case "start"
                                copy_combobox(cbx, cbo_production_run_end)
                            Case "end"
                                copy_combobox(cbx, cbo_production_run_quant)
                            Case "quant"
                                With dgv_navigate
                                    .Rows.Add(cbo_production_run.Text, cbo_production_run_back.Text, Nothing, Nothing, _
                                              Nothing, nud_navigate_step.Value, acsis.get_selected_option(0), _
                                              tb_navigate_current_window.Text, "Run Confirmation", _
                                              w_sessions(s_current_session).winWnd.Length, cbo_navigate_target.Text)

                                End With
                                acsis.click_button(cbo_production_run_back.Text, False, 0)
                                copy_combobox(cbo_production_run, cbo_production_downtime)
                        End Select
                    Case name.Contains("_downtime_")
                        Select Case str(str.Length - 1)
                            Case "confirm"
                                copy_combobox(cbx, cbo_production_downtime_back)
                            Case "back"
                                cbo_production_downtime_start.Items.Clear()
                                For i As Integer = 0 To w_sessions(0).winWnd(0).tbWnd.Length - 1
                                    cbo_production_downtime_start.Items.Add(i)
                                Next
                                cbo_production_downtime_start.Enabled = True
                            Case "start"
                                copy_combobox(cbx, cbo_production_downtime_end)
                            Case "end"
                                With dgv_navigate
                                    .Rows.Add(cbo_production_downtime.Text, cbo_production_downtime_back.Text, Nothing, Nothing, Nothing, _
                                              nud_navigate_step.Value, acsis.get_selected_option(0), tb_navigate_current_window.Text, _
                                              "DT & Scrap Recording", w_sessions(0).winWnd.Length, cbo_navigate_target.Text)
                                End With
                                acsis.click_button(cbo_production_downtime_back.Text, False, 0)
                                copy_combobox(cbo_production_downtime, cbo_production_final)
                        End Select
                    Case name.Contains("_final_")
                        Select Case str(str.Length - 1)
                            Case "confirm"
                                copy_combobox(cbx, cbo_production_final_back)
                            Case "back"
                                fill_combobox(cbo_production_final_check, type:=efill_type.checkbox)
                            Case "check"
                                With dgv_navigate
                                    .Rows.Add(cbo_production_final.Text, cbo_production_final_back.Text, Nothing, Nothing, _
                                              Nothing, nud_navigate_step.Value, acsis.get_selected_option(0), _
                                              tb_navigate_current_window.Text, "Final Confirmation", _
                                              w_sessions(0).winWnd.Length, cbo_navigate_target.Text)

                                End With
                                export = True
                        End Select
                    Case Else
                        cbx = cbx
                End Select
                tb_navigate_current_window.Text = w_sessions(0).winWnd(0).name
                If export Then
                    export_navigation()
                    sql.delete("acsis_options", "item", "LIKE", "'production_%' AND plant=" & s_plant & " AND version='" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_label_confirm','" & cbo_production_label_confirm.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_label_quant','" & cbo_production_label_quant.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_label_from','" & cbo_production_label_from.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_label_to','" & cbo_production_label_to.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_setup_confirm','" & cbo_production_setup_confirm.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_setup_start','" & cbo_production_setup_start.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_setup_end','" & cbo_production_setup_end.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_run_confirm','" & cbo_production_run_confirm.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_run_start','" & cbo_production_run_start.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_run_end','" & cbo_production_run_end.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_run_quant','" & cbo_production_run_quant.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_downtime_confirm','" & cbo_production_downtime_confirm.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_downtime_start','" & cbo_production_downtime_start.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_downtime_end','" & cbo_production_downtime_end.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_final_confirm','" & cbo_production_final_confirm.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'production_final_check','" & cbo_production_final_check.Text & "'," & s_plant & ",'" & s_version & "'")
                    dt_acsis_options = sql.get_table("acsis_options", "plant", "=", s_plant & " AND version='" & s_version & "'")
                    acsis.click_button(cbo_production_final_back.Text, False, 0)
                    reset_navigation(gb_navigate_main)
                End If
            End With
        End If

    End Sub

    Private Sub b_navigate_menu_reset_Click(sender As Object, e As EventArgs) Handles b_navigate_menu_reset.Click
        sql.delete("acsis_navigation", "sub_target", "=", "'" & cbo_navigate_target.Text & "' AND plant=" & s_plant & " AND version='" & s_version & "'")
        goto_start(True)
        dt_navigation = sql.get_table("acsis_navigation", "plant", "=", s_plant & " AND version='" & s_version & "'")
        Select Case True
            Case gb_navigate_main.Visible
                reset_navigation(gb_navigate_main)
                With dgv_navigate
                    .Rows.Clear()
                    nud_navigate_step.Value = define_step(False)
                    ' sql.delete("acsis_navigation", "sub_target", "is", "null AND plant=" & s_plant)
                    If w_sessions(0).winWnd.Length > 1 Then
                        frm_status.Close()
                        MsgBox("A menu has been detected that is not recorded" & vbCr & "Please navigate to the very start of Sym-PLE.")
                        SetForegroundWindow(w_sessions(0).process.MainWindowHandle)
                        MsgBox("Click OK if you are at the start of the Sym-PLE menus")
                    End If
                    RefreshControls(True)
                End With
            Case gb_navigate_pallet.Visible
                setup_pallet()
            Case gb_navigate_overview.Visible
                setup_overview()
            Case gb_navigate_prod.Visible
                With dgv_navigate
                    For i As Integer = 0 To .RowCount - 1
                        If Not IsNull(.Rows(i).Cells(.ColumnCount - 1).Value) Then
                            If .RowCount - 2 - i = 6 Then ' 6 is the number of controls accessed in Production Confirmation
                                Exit For
                            ElseIf Not .RowCount - 1 = i Then
                                For r As Integer = .RowCount - 1 To i + 1 Step -1
                                    .Rows.RemoveAt(r)
                                Next
                                export_navigation()
                                Exit For
                            End If
                        End If
                    Next
                End With
                setup_production()
            Case gb_navigate_rts.Visible
                setup_rts()
            Case gb_navigate_relabel.Visible
                setup_relabeling()
        End Select
        Select Case cbo_navigate_target.Text
            Case "Production Confirmation"
                sql.delete("acsis_options", "item", "LIKE", "'production_%' AND plant=" & s_plant & " AND version='" & s_version & "'")
            Case "Print Pallet Labels"
                sql.delete("acsis_options", "item", "LIKE", "'pallet_screen_%' AND plant=" & s_plant & " AND version='" & s_version & "'")
            Case "Modify Pallet Label"
                sql.delete("acsis_options", "item", "LIKE", "'modify_pallet_%' AND plant=" & s_plant & " AND version='" & s_version & "'")
            Case "Stock Overview"
            Case "Warehouse Reciepts - Other"
                sql.delete("acsis_options", "item", "LIKE", "'rts_screen_%' AND plant=" & s_plant & " AND version='" & s_version & "'")
            Case "Re-Labeling"
                sql.delete("acsis_options", "item", "LIKE", "'relabel_%' AND plant=" & s_plant & " AND version='" & s_version & "'")
        End Select

    End Sub

    Private Sub cbo_pallet_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_pallet_movement.SelectedIndexChanged, _
        cbo_pallet_seq.SelectedIndexChanged, cbo_pallet_quant.SelectedIndexChanged, cbo_pallet_prod_num.SelectedIndexChanged, _
        cbo_pallet_mat_num.SelectedIndexChanged, cbo_pallet_label_data.SelectedIndexChanged, cbo_pallet_id.SelectedIndexChanged, _
        cbo_pallet_last_pallet.SelectedIndexChanged, cbo_pallet_print.SelectedIndexChanged

        Dim cbx As ComboBox = CType(sender, ComboBox)
        If Not cbx.Text = Nothing Then
            Dim name As String = cbx.Name
            cbx.Enabled = False
            Select Case True
                Case name.Contains("movement")
                    copy_combobox(cbx, cbo_pallet_prod_num)
                Case name.Contains("prod_num")
                    copy_combobox(cbx, cbo_pallet_mat_num)
                Case name.Contains("mat_num")
                    copy_combobox(cbx, cbo_pallet_seq)
                Case name.Contains("seq")
                    copy_combobox(cbx, cbo_pallet_quant)
                Case name.Contains("quant")
                    copy_combobox(cbx, cbo_pallet_id)
                Case name.Contains("_id")
                    With cbo_pallet_label_data
                        .Items.Clear()
                        For i As Integer = 0 To w_sessions(0).winWnd(0).cbWnd.Length - 1
                            .Items.Add(w_sessions(0).winWnd(0).cbxWnd(i).text(0))
                        Next
                        .Enabled = True
                    End With
                Case name.Contains("label_data")
                    copy_combobox(cbx, cbo_pallet_last_pallet)
                Case name.Contains("last_pallet")
                    fill_combobox(cbo_pallet_print, type:=efill_type.button)
                Case name.Contains("print")
                    sql.delete("acsis_options", "item", "LIKE", "'pallet_screen_%' AND plant=" & s_plant & " AND version='" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'pallet_screen_movement','" & cbo_pallet_movement.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'pallet_screen_prod_num','" & cbo_pallet_prod_num.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'pallet_screen_mat_num','" & cbo_pallet_mat_num.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'pallet_screen_seq','" & cbo_pallet_seq.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'pallet_screen_quant','" & cbo_pallet_quant.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'pallet_screen_id','" & cbo_pallet_id.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'pallet_screen_label_data','" & cbo_pallet_label_data.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'pallet_screen_last_pallet','" & cbo_pallet_last_pallet.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'pallet_screen_print','" & cbo_pallet_print.Text & "'," & s_plant & ",'" & s_version & "'")
                    dt_acsis_options = sql.get_table("acsis_options", "plant", "=", s_plant & " AND version='" & s_version & "'")
                    reset_navigation(gb_navigate_main)
            End Select

        End If

    End Sub

    Private Sub overview_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_overview_mat_num.SelectedIndexChanged, cbo_overview_show_stock.SelectedIndexChanged
        Dim cbx As ComboBox = CType(sender, ComboBox)
        If Not cbx.Text = Nothing Then
            Dim name As String = cbx.Name
            cbx.Enabled = False
            Select Case True
                Case name.Contains("mat_num")
                    fill_combobox(cbo_overview_show_stock, type:=efill_type.button)
                Case name.Contains("show_stock")
                    sql.delete("acsis_options", "item", "LIKE", "'overview_%' AND plant=" & s_plant & " AND version='" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'overview_mat_num','" & cbo_overview_mat_num.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'overview_show_stock','" & cbo_overview_show_stock.Text & "'," & s_plant & ",'" & s_version & "'")
                    dt_acsis_options = sql.get_table("acsis_options", "plant", "=", s_plant & " AND version='" & s_version & "'")
                    reset_navigation(gb_navigate_main)
            End Select

        End If

    End Sub

    Private Sub Number_only_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb_rts_mat_num.KeyPress, tb_production_prod_num.KeyPress
        number_only(sender, e, False)
    End Sub

    Private Sub b_rts_next_Click(sender As Object, e As EventArgs) Handles b_rts_next.Click
        b_rts_next.Enabled = False
        fill_combobox(cbo_rts_movement, type:=efill_type.combobox)

    End Sub

    Private Sub tb_rts_mat_num_TextChanged(sender As Object, e As EventArgs) Handles tb_rts_mat_num.TextChanged
        If tb_rts_mat_num.Text.Length = 9 Then
            b_rts_next.Enabled = True
        Else
            b_rts_next.Enabled = False
        End If
    End Sub

    Sub export_navigation()
        Dim str, values, columns As String
        With dgv_navigate

            'sql.delete("acsis_navigation", "sub_target", "IS", "NULL")
            sql.delete("acsis_navigation", "sub_target", "=", "'" & cbo_navigate_target.Text & "' AND plant=" & s_plant & " AND version='" & tb_navigation_version.Text & "'")
            For r As Integer = 0 To .RowCount - 1
                str = Nothing
                values = Nothing
                For c As Integer = 0 To .Columns.Count - 1
                    str = sql.convert_value(.Rows(r).Cells(c).Value)

                    If values = Nothing Then
                        values = str
                    Else
                        values = values & "," & str
                    End If


                Next
                columns = "button_access,button_exit,opt_1,opt_2,opt_3,step,opt_target,window_name,target" & _
                    ",window_count,sub_target,target_reached,plant,version"

                values = values & "," & s_plant & ",'" & tb_navigation_version.Text & "'"
                sql.insert("acsis_navigation", columns, values)
            Next
        End With
        dt_navigation = sql.get_table("acsis_navigation", "plant", "=", s_plant & " AND version='" & s_version & "'")

    End Sub

    Sub get_downtime()
        With dgv_downtime
            .DataSource = sql.get_table("settings_downtime", "plant", "=", s_plant & " AND dept='" & s_department & "'")
            If .RowCount > 0 Then
                .Columns("plant").Visible = False
                .Columns("row").Visible = False
            End If
        End With

    End Sub

    Sub goto_window()

        Dim index As cls_ACSIS.eWindowTarget
        Select Case cbo_navigate_target.Text
            Case "Production Confirmation"
                index = cls_ACSIS.eWindowTarget.ProductionConfirmation
            Case "Print Pallet Labels"
                index = cls_ACSIS.eWindowTarget.PrintPalletLabels
            Case "Modify Pallet Label"
                index = cls_ACSIS.eWindowTarget.ModifyPalletLabel
            Case "Stock Overview"
                index = cls_ACSIS.eWindowTarget.StockOverview
            Case "Re-Labeling"
                index = cls_ACSIS.eWindowTarget.ReLabeling
            Case "Warehouse Reciepts - Other"
                index = cls_ACSIS.eWindowTarget.WarehouseReceiptsOther
            Case Else
                index = cls_ACSIS.eWindowTarget.ProductionConfirmation
        End Select
        With dgv_navigate
            For i As Integer = 0 To .RowCount - 1
                If Not IsNull(.Rows(i).Cells(.ColumnCount - 1).Value) Then
                    If .Rows(i).Cells(.ColumnCount - 1).Value = True Then
                        acsis.open_window(index, 0)
                        Exit Sub
                    End If
                End If
            Next
            goto_start(True)

            show_status("Navigating to the last recorded position...")

            For r As Integer = 0 To .RowCount - 1
                acsis.get_controls(0, True, False, True)
                If w_sessions(0).winWnd.Length = 1 Then
                    For o As Integer = 0 To w_sessions(0).winWnd(0).optWnd.Length - 1
                        If w_sessions(0).winWnd(0).optWnd(o).text(0) = .Rows(r).Cells(6).Value Then
                            acsis.set_check(w_sessions(0).winWnd(0).optWnd(o).handle, True)
                            acsis.click_button(.Rows(r).Cells(0).Value, False, 0)
                            Exit For
                        End If
                    Next
                Else
                    Exit For
                End If
            Next
            frm_status.Close()
        End With

    End Sub
    Sub reset_groupbox(ByVal gb As GroupBox)
        For Each ctrl In gb.Controls
            If TypeOf ctrl Is ComboBox Then
                Dim cbx As ComboBox = ctrl
                cbx.Enabled = False
                If Not cbx Is cbo_navigate_target Then cbx.Items.Clear()
            ElseIf TypeOf ctrl Is Button Then
                Dim b As Button = ctrl
                If Not gb Is gb_navigate_main Then b.Enabled = False
            ElseIf TypeOf ctrl Is TextBox Then
                Dim tb As TextBox = ctrl
                tb.Enabled = True
                tb.Text = Nothing
            ElseIf TypeOf ctrl Is GroupBox Then
                reset_groupbox(ctrl)
            End If
        Next
    End Sub
    Sub reset_navigation(ByVal gb As GroupBox)
        Dim enter_window As Boolean = True, show_tb As Boolean = True, login As Boolean = False
        For Each gbc In TabPage4.Controls
            If TypeOf gbc Is GroupBox Then

                Dim n As String = gbc.Name
                If n.Contains("gb_navigate_") Then gbc.Visible = False
                reset_groupbox(gbc)
                Label25.Visible = True
                If gb Is gbc Then
                    gbc.Visible = True
                    gbc.Location = New Point(0, 1)
                    cbo_navigate_target.Enabled = True
                    Select Case vbTrue
                        Case gb Is gb_navigate_main
                            Label25.Visible = False
                            show_tb = False
                            enter_window = False
                        Case gb Is gb_navigate_overview
                        Case gb Is gb_navigate_pallet
                        Case gb Is gb_navigate_prod
                            nud_navigate_step.Value = define_step(True)
                            show_tb = False
                        Case gb Is gb_navigate_rts
                        Case gb Is gb_navigate_relabel
                        Case gb Is gb_navigate_modify_pallet
                        Case gb Is gb_navigate_login
                            enter_window = False
                            login = True
                    End Select
                End If

            End If
        Next
        If enter_window Then goto_window()
        If login Then
            acsis.get_controls(0, True, False, True)
            acsis.get_controls(w_sessions(0).winWnd(0).handle, True, False, True)
            For i As Integer = 0 To w_sessions(0).winWnd(0).tbWnd.Length - 1
                acsis.insert_string(i.ToString.PadLeft(i, i.ToString), w_sessions(0).winWnd(0).tbWnd(i).handle)
            Next
        Else
            If show_tb Then acsis.get_controls(w_sessions(0).winWnd(0).handle, True, True, True)
        End If

    End Sub

    Sub setup_login()
        SetForegroundWindow(Me.Handle)
        MsgBox("You now need to set up how a user logs in")
        reset_navigation(gb_navigate_login)
        fill_combobox(cbo_login_id, type:=efill_type.textbox, n:=w_sessions(0).winWnd(0).tbWnd.Length - 1)
    End Sub

    Sub setup_modify_pallet()
        MsgBox("You now need to set up how pallet labels are deleted")
        reset_navigation(gb_navigate_modify_pallet)
        fill_combobox(cbo_modify_pallet_id, type:=efill_type.textbox, n:=w_sessions(0).winWnd(0).tbWnd.Length - 1)
    End Sub

    Sub setup_overview()
        MsgBox("You now need to set up how current stock is accessed")
        reset_navigation(gb_navigate_overview)
        fill_combobox(cbo_overview_mat_num, type:=efill_type.textbox, n:=w_sessions(0).winWnd(0).tbWnd.Length - 1)
    End Sub

    Sub setup_pallet()
        MsgBox("You now need to set up how a pallet is entered")
        reset_navigation(gb_navigate_pallet)
        fill_combobox(cbo_pallet_movement, type:=efill_type.textbox, n:=w_sessions(0).winWnd(0).tbWnd.Length - 1)
    End Sub

    Sub setup_production()
        MsgBox("You now need to set up how a job is accessed")
        reset_navigation(gb_navigate_prod)

    End Sub

    Sub setup_relabeling()
        MsgBox("You now need to set up how Re-labeling is done")
        reset_navigation(gb_navigate_relabel)
        fill_combobox(cbo_relabel_material, type:=efill_type.radio)
    End Sub

    Sub setup_rts()
        MsgBox("You now need to set up how a RTS is done")
        reset_navigation(gb_navigate_rts)
    End Sub


    Private Sub cbo_rts_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_rts_prod_num.SelectedIndexChanged, _
        cbo_rts_items.SelectedIndexChanged, cbo_rts_mat_num.SelectedIndexChanged, cbo_rts_dest_type.SelectedIndexChanged, _
        cbo_rts_quant.SelectedIndexChanged, cbo_rts_dest_bin.SelectedIndexChanged, cbo_rts_description.SelectedIndexChanged, _
        cbo_rts_put_away.SelectedIndexChanged, cbo_rts_movement.SelectedIndexChanged, cbo_rts_source_type.SelectedIndexChanged, cbo_rts_source_bin.SelectedIndexChanged

        Dim cbx As ComboBox = CType(sender, ComboBox)
        If Not cbx.Text = Nothing Then
            Dim name As String = cbx.Name
            cbx.Enabled = False
            Select Case True
                Case name.Contains("movement")
                    fill_combobox(cbo_rts_prod_num, type:=efill_type.textbox, n:=w_sessions(0).winWnd(0).tbWnd.Length - 1)
                Case name.Contains("prod_num")
                    copy_combobox(cbx, cbo_rts_mat_num)
                Case name.Contains("mat_num")
                    copy_combobox(cbx, cbo_rts_quant)
                Case name.Contains("quant")
                    copy_combobox(cbx, cbo_rts_description)
                    cb_rts_description.Enabled = True
                Case name.Contains("description")
                    copy_combobox(cbx, cbo_rts_items)
                    cb_rts_description.Enabled = False
                Case name.Contains("items")
                    copy_combobox(cbx, cbo_rts_source_bin)
                Case name.Contains("source_bin")
                    copy_combobox(cbx, cbo_rts_source_type)
                Case name.Contains("source_type")
                    copy_combobox(cbx, cbo_rts_dest_bin)
                Case name.Contains("dest_bin")
                    copy_combobox(cbx, cbo_rts_dest_type)
                    SendMessage(w_sessions(0).winWnd(0).tbWnd(cbo_rts_dest_bin.Text).handle, WM_SETFOCUS, 0, 0)
                    mod_common.KeyPress(Keys.Tab, w_sessions(0).winWnd(0).tbWnd(cbo_rts_dest_bin.Text).handle)
                Case name.Contains("dest_type")
                    fill_combobox(cbo_rts_put_away, type:=efill_type.button)
                Case name.Contains("put_away")

                    sql.delete("acsis_options", "item", "LIKE", "'rts_screen_%' AND plant=" & s_plant & " AND version='" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'rts_screen_movement','" & cbo_rts_movement.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'rts_screen_prod_num','" & cbo_rts_prod_num.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'rts_screen_mat_num','" & cbo_rts_mat_num.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'rts_screen_quant','" & cbo_rts_quant.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'rts_screen_desc','" & cbo_rts_description.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'rts_screen_items','" & cbo_rts_items.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'rts_screen_source_bin','" & cbo_rts_source_bin.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'rts_screen_source_type','" & cbo_rts_source_type.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'rts_screen_dest_bin','" & cbo_rts_dest_bin.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'rts_screen_dest_type','" & cbo_rts_dest_type.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'rts_screen_put_away','" & cbo_rts_put_away.Text & "'," & s_plant & ",'" & s_version & "'")
                    dt_acsis_options = sql.get_table("acsis_options", "plant", "=", s_plant & " AND version='" & s_version & "'")
                    reset_navigation(gb_navigate_main)
            End Select

        End If

    End Sub


    Private Sub cb_rts_description_CheckedChanged(sender As Object, e As EventArgs) Handles cb_rts_description.CheckedChanged
        If cb_rts_description.Checked Then
            For i As Integer = 0 To w_sessions(0).winWnd(0).tbWnd.Length - 1
                acsis.insert_string(w_sessions(0).winWnd(0).tbWnd(i).text(0), w_sessions(0).winWnd(0).tbWnd(i).handle)
            Next
            SendMessage(w_sessions(0).winWnd(0).tbWnd(cbo_rts_mat_num.Text).handle, WM_SETFOCUS, 0, 0)
            acsis.insert_string(tb_rts_mat_num.Text, w_sessions(0).winWnd(0).tbWnd(cbo_rts_mat_num.Text).handle)
            SendMessage(w_sessions(0).winWnd(0).tbWnd(cbo_rts_quant.Text).handle, WM_SETFOCUS, 0, 0)
            w_sessions(0).process.WaitForInputIdle()
            cbo_rts_description.Enabled = False
        Else
            For i As Integer = 0 To w_sessions(0).winWnd(0).tbWnd.Length - 1
                acsis.insert_string(i, w_sessions(0).winWnd(0).tbWnd(i).handle)
            Next
            cbo_rts_description.Enabled = True
        End If
    End Sub

    Private Sub relabel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_relabel_material.SelectedIndexChanged, _
        cbo_relabel_of.SelectedIndexChanged, cbo_relabel_to.SelectedIndexChanged, cbo_relabel_from.SelectedIndexChanged, cbo_relabel_quant.SelectedIndexChanged, _
        cbo_relabel_mat_num.SelectedIndexChanged, cbo_relabel_date.SelectedIndexChanged, cbo_relabel_print.SelectedIndexChanged

        Dim cbx As ComboBox = CType(sender, ComboBox)
        If Not cbx.Text = Nothing Then
            Dim name As String = cbx.Name
            cbx.Enabled = False
            Select Case True
                Case name.Contains("material")
                    fill_combobox(cbo_relabel_mat_num, type:=efill_type.textbox, n:=w_sessions(0).winWnd(0).tbWnd.Length - 1)
                Case name.Contains("mat_num")
                    copy_combobox(cbx, cbo_relabel_date)
                Case name.Contains("date")
                    copy_combobox(cbx, cbo_relabel_quant)
                Case name.Contains("quant")
                    copy_combobox(cbx, cbo_relabel_from)
                Case name.Contains("from")
                    copy_combobox(cbx, cbo_relabel_to)
                Case name.Contains("abel_to")
                    copy_combobox(cbx, cbo_relabel_of)
                Case name.Contains("abel_of")
                    fill_combobox(cbo_relabel_print, type:=efill_type.button)
                Case name.Contains("print")
                    sql.delete("acsis_options", "item", "LIKE", "'relabel_screen_%' AND plant=" & s_plant & " AND version='" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'relabel_screen_material','" & cbo_relabel_material.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'relabel_screen_mat_num','" & cbo_relabel_mat_num.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'relabel_screen_quant','" & cbo_relabel_quant.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'relabel_screen_date','" & cbo_relabel_date.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'relabel_screen_from','" & cbo_relabel_from.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'relabel_screen_to','" & cbo_relabel_to.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'relabel_screen_of','" & cbo_relabel_of.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'relabel_screen_print','" & cbo_relabel_print.Text & "'," & s_plant & ",'" & s_version & "'")
                    dt_acsis_options = sql.get_table("acsis_options", "plant", "=", s_plant & " AND version='" & s_version & "'")
                    reset_navigation(gb_navigate_main)
            End Select

        End If
    End Sub


    Private Sub b_navigate_set_Click(sender As Object, e As EventArgs) Handles b_navigate_set.Click
        Select Case cbo_navigate_target.Text
            Case "Production Confirmation"
                setup_production()
            Case "Re-Labeling"
                setup_relabeling()
            Case "Warehouse Reciepts - Other"
                setup_rts()
            Case "Print Pallet Labels"
                setup_pallet()
            Case "Modify Pallet Label"
                setup_modify_pallet()
            Case "Stock Overview"
                setup_overview()
            Case "Re-Labeling"
                setup_relabeling()
        End Select
    End Sub

    Private Sub cb_modify_pallet_id_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_modify_pallet_id.SelectedIndexChanged, _
        cbo_modify_pallet_get_info.SelectedIndexChanged, cbo_modify_pallet_delete.SelectedIndexChanged

        Dim cbx As ComboBox = CType(sender, ComboBox)
        If Not cbx.Text = Nothing Then
            Dim name As String = cbx.Name
            cbx.Enabled = False
            Select Case True
                Case name.Contains("pallet_id")
                    fill_combobox(cbo_modify_pallet_get_info, type:=efill_type.button)
                Case name.Contains("get_info")
                    copy_combobox(cbx, cbo_modify_pallet_delete)
                Case name.Contains("delete")
                    sql.delete("acsis_options", "item", "LIKE", "'modify_pallet_%' AND plant=" & s_plant & " AND version='" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'modify_pallet_id','" & cbo_modify_pallet_id.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'modify_pallet_get_info','" & cbo_modify_pallet_get_info.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'modify_pallet_delete','" & cbo_modify_pallet_delete.Text & "'," & s_plant & ",'" & s_version & "'")
                    dt_acsis_options = sql.get_table("acsis_options", "plant", "=", s_plant & " AND version='" & s_version & "'")
                    reset_navigation(gb_navigate_main)
            End Select

        End If
    End Sub

    Private Sub cbo_login_id_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_login_id.SelectedIndexChanged, _
        cbo_login_pass.SelectedIndexChanged, cbo_login_ok.SelectedIndexChanged

        Dim cbx As ComboBox = CType(sender, ComboBox)
        If Not cbx.Text = Nothing Then
            Dim name As String = cbx.Name
            cbx.Enabled = False
            Select Case True
                Case name.Contains("_id")
                    copy_combobox(cbx, cbo_login_pass)
                Case name.Contains("_pass")
                    fill_combobox(cbo_login_ok, type:=efill_type.button)
                Case name.Contains("_ok")
                    sql.delete("acsis_options", "item", "LIKE", "'login_%' AND plant=" & s_plant & " AND version='" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'login_id','" & cbo_login_id.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'login_pass','" & cbo_login_pass.Text & "'," & s_plant & ",'" & s_version & "'")
                    sql.insert("acsis_options", "item,val,plant,version", "'login_button','" & cbo_login_ok.Text & "'," & s_plant & ",'" & s_version & "'")
                    dt_acsis_options = sql.get_table("acsis_options", "plant", "=", s_plant & " AND version='" & s_version & "'")
                    w_sessions(0).process.Kill()
                    acsis.get_controls(0, True, False, True)
                    With frm_logon
                        .StartPosition = FormStartPosition.CenterParent
                        .ShowDialog()
                    End With

                    reset_navigation(gb_navigate_main)
            End Select

        End If
    End Sub

    Sub get_dgv_info(Optional ByVal reset As Boolean = False)
        dgv_department_columns.Rows.Clear()
        dgv_department_preview.Columns.Clear()
        Dim dgv As DataGridView = Nothing
        If cbo_department_dgv.Text = Nothing Then Exit Sub
        Select Case cbo_department_dgv.Text
            Case "Summary"
                dgv = frm_main.dgv_summary
            Case "Report"
                dgv = frm_main.dgv_production
        End Select
        Dim dept As String = s_department
        If cb_department_extruder.Checked Then dept = "extruder"
        dgv_department_preview.Width = dgv.Width
        Dim dt As DataTable = sql.get_table("settings_departments", "plant", "=", s_plant & " AND item='dgv_col_" & cbo_department_dgv.Text & "' AND dept='" & dept & "'")
        Dim ordr As Integer = 0
        Dim str() As String = Nothing
        Dim str_add As Object = Nothing
        With dgv
            If Not dt Is Nothing Then str = Split(dt(0)("value"), ",")
            For c As Integer = 0 To .ColumnCount - 1
                If .Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable Then
                    If dt Is Nothing Or reset Then
                        dgv_department_columns.Rows.Add(str_add, .Columns(c).Name, .Columns(c).HeaderText, _
                                                        .Columns(c).Visible, .Columns(c).Width, _
                                                        .Columns(c).DefaultCellStyle.Format, .Columns(c).ReadOnly)
                    Else
                        Dim s() As String = Nothing
                        Dim add As Boolean = True
                        For i As Integer = 0 To str.Length - 1
                            s = Split(str(i), ":")

                            If s(1) = .Columns(c).Name Then
                                dgv_department_columns.Rows.Add(CInt(s(0)), s(1), s(2), s(3), s(4), s(5), dgv.Columns(s(1)).ReadOnly)
                                add = False
                                Exit For
                            End If
                            ' dgv_department_columns.Columns.Add(s(1), s(2))
                        Next
                        If add Then
                            dgv_department_columns.Rows.Add(str_add, .Columns(c).Name, .Columns(c).HeaderText, _
                                                            .Columns(c).Visible, .Columns(c).Width, _
                                                            .Columns(c).DefaultCellStyle.Format, .Columns(c).ReadOnly)
                        End If
                    End If
                    str_add = Nothing
                    If .Columns(c).Visible Then
                        str_add = CInt(ordr)
                        ordr = ordr + 1
                    End If
                End If
            Next
            dgv_department_columns.Sort(dgv_department_columns.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            dgv_department_columns.PerformLayout()

            If Not reset Then
                b_department_save.PerformClick()

            End If
            ' renumber_dgv()
            preview_dgv()
        End With
    End Sub

    Sub preview_dgv()
        Dim dgv As DataGridView = Nothing
        Select Case cbo_department_dgv.Text
            Case "Summary"
                dgv = frm_main.dgv_summary
            Case "Report"
                dgv = frm_main.dgv_production
        End Select
        s_loading = True
        With dgv_department_preview
            .Columns.Clear()
            For r As Integer = 0 To dgv_department_columns.RowCount - 1
                If dgv_department_columns.Rows(r).Cells(3).Value = True Then
                    .Columns.Add(dgv_department_columns.Rows(r).Cells(1).Value, dgv_department_columns.Rows(r).Cells(2).Value)
                    .Columns(.ColumnCount - 1).Width = dgv_department_columns.Rows(r).Cells(4).Value
                End If
            Next
            For c As Integer = 0 To .ColumnCount - 1
                For r As Integer = 0 To dgv_department_columns.Rows.Count - 1
                    If Not dgv_department_columns.Rows(r).Cells(0).Value = Nothing Then

                        If .Columns(c).Name = dgv_department_columns.Rows(r).Cells(1).Value And dgv_department_columns.Rows(r).Cells(3).Value = True Then
                            If Not dgv_department_columns.CurrentRow.Index = r Then
                                If Not dgv_department_columns.Rows(r).Cells(0).Value = 9999 Then .Columns(c).DisplayIndex = dgv_department_columns.Rows(r).Cells(0).Value

                            End If
                            .Columns(c).Width = dgv_department_columns.Rows(r).Cells(4).Value
                            Exit For
                        End If
                    End If
                Next
            Next

        End With
        s_loading = False


    End Sub
    Private Sub cbo_department_dgv_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_department_dgv.SelectedIndexChanged
        If Not s_loading Then
            get_dgv_info()
        End If
    End Sub

    Private Sub b_department_reload_Click(sender As Object, e As EventArgs) Handles b_department_reload.Click
        If Not s_loading Then
            get_dgv_info()
        End If
    End Sub

    Private Sub cbo_department_select_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_department_select.SelectedIndexChanged
        If Not s_loading Then
            s_department = LCase(cbo_department_select.SelectedItem("dept"))
            get_dgv_info()
            load_deparment()
            Select Case s_department
                Case "filmsline"
                    cb_department_extruder.Visible = True
                Case Else
                    cb_department_extruder.Visible = False
            End Select
        End If
    End Sub

    Private Sub b_department_save_Click(sender As Object, e As EventArgs) Handles b_department_save.Click
        Dim s, str As String
        str = Nothing
        With dgv_department_columns
            For r As Integer = 0 To .RowCount - 1
                s = .Rows(r).Cells(0).Value & ":" & .Rows(r).Cells(1).Value & ":" & .Rows(r).Cells(2).Value & _
                    ":" & .Rows(r).Cells(3).Value & ":" & .Rows(r).Cells(4).Value & ":" & .Rows(r).Cells(5).Value

                If str = Nothing Then
                    str = s
                Else
                    str = str & "," & s
                End If
            Next
        End With
        Dim dept As String = s_department
        If cb_department_extruder.Checked Then dept = "extruder"

        sql.delete("settings_departments", "plant", "=", s_plant & " AND item='dgv_col_" & cbo_department_dgv.Text & "' AND dept='" & dept & "'")
        sql.insert("settings_departments", "dept,plant,item,value", "'" & dept & "'," & s_plant & ",'dgv_col_" & cbo_department_dgv.Text & "','" & str & "'")
    End Sub

    Private Sub dgv_department_columns_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_department_columns.CellContentClick
        If e.ColumnIndex = 3 And e.RowIndex > -1 Then
            With dgv_department_columns

                If .CurrentCell.Value = True Then
                    .CurrentCell.Value = False
                    dgv_department_preview.Columns.Remove(.CurrentRow.Cells(1).Value)
                Else
                    .CurrentCell.Value = True
                    '.CurrentRow.Cells(0).Value = 0
                    dgv_department_preview.Columns.Add(.CurrentRow.Cells(1).Value, .CurrentRow.Cells(2).Value)
                    dgv_department_preview.Columns(dgv_department_preview.ColumnCount - 1).Width = .CurrentRow.Cells(4).Value
                    'For r As Integer = 0 To .Rows.Count - 1
                    '    If Not IsNull(.Rows(r).Cells(0).Value) Then
                    '        If .Rows(r).Cells(0).Value > .CurrentRow.Cells(0).Value Then
                    '            .CurrentRow.Cells(0).Value = .Rows(r).Cells(0).Value + 1
                    '        End If
                    '    End If
                    'Next
                End If
                renumber_dgv()
            End With
        End If
    End Sub

    Private Sub dgv_department_preview_ColumnDisplayIndexChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dgv_department_preview.ColumnDisplayIndexChanged
        If Not s_loading Then
            renumber_dgv()
        End If
    End Sub

    Sub renumber_dgv(Optional ByVal n As Integer = -1)
        If n > -1 Then
            For r As Integer = 0 To dgv_department_columns.Rows.Count - 1
                If dgv_department_columns.Rows(r).Cells(0).Value >= n Then
                    If n = 0 Then
                        dgv_department_columns.Rows(r).Cells(0).Value = 9999
                    Else
                        dgv_department_columns.Rows(r).Cells(0).Value = dgv_department_columns.Rows(r).Cells(0).Value - 1
                    End If

                End If
            Next
        Else
            For r As Integer = 0 To dgv_department_columns.Rows.Count - 1
                dgv_department_columns.Rows(r).Cells(0).Value = 9999
            Next
            For c As Integer = 0 To dgv_department_preview.ColumnCount - 1
                For r As Integer = 0 To dgv_department_columns.Rows.Count - 1
                    If dgv_department_preview.Columns(c).Name = dgv_department_columns.Rows(r).Cells(1).Value Then
                        dgv_department_columns.Rows(r).Cells(0).Value = dgv_department_preview.Columns(c).DisplayIndex
                        Exit For
                    End If
                Next
            Next
        End If
        dgv_department_columns.Sort(dgv_department_columns.Columns(0), 0)
        dgv_department_columns.PerformLayout()

    End Sub


    Private Sub dgv_downtime_codes_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_downtime_codes.CellEndEdit
        Dim s As String = Nothing
        With dgv_downtime_codes
            For i As Integer = 0 To .RowCount - 2
                If s = Nothing Then
                    s = .Rows(i).Cells(0).Value
                Else
                    s = s & "," & .Rows(i).Cells(0).Value
                End If
            Next

            sql.delete("settings_general", "item", "=", "'downtime_codes'")
            sql.insert("settings_general", "item,val,plant", "'downtime_codes','" & s & "'," & s_plant)
            s = cbo_downtime_code.Text

            cbo_downtime_code.Items.Clear()

            For i As Integer = 0 To .RowCount - 2
                cbo_downtime_code.Items.Add(.Rows(i).Cells(0).Value)
            Next
            select_cbo_item(cbo_downtime_code, s)
        End With

    End Sub


    Private Sub cbo_plant_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_plant.SelectedIndexChanged
        If Not s_loading Then
            s_plant = cbo_plant.SelectedItem("plant")
            load_plant()
            load_deparment()
        End If
    End Sub

    Private Sub dgv_department_columns_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_department_columns.CellEndEdit
        preview_dgv()
    End Sub

    Private Sub dgv_department_preview_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dgv_department_preview.ColumnWidthChanged
        If Not s_loading Then
            With dgv_department_columns
                For r As Integer = 0 To .RowCount - 1
                    If .Rows(r).Cells(1).Value = dgv_department_preview.Columns(e.Column.Index).Name Then
                        .Rows(r).Cells(4).Value = e.Column.Width
                        Exit For
                    End If
                Next
            End With
        End If
    End Sub


    Private Sub b_navigation_Click(sender As Object, e As EventArgs) Handles b_navigation.Click
        With frm_navigation
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
    End Sub

    Private Sub b_department_reset_Click(sender As Object, e As EventArgs) Handles b_department_reset.Click
        If Not s_loading Then
            get_dgv_info(True)
        End If
    End Sub

    Private Sub cb_department_extruder_CheckedChanged(sender As Object, e As EventArgs) Handles cb_department_extruder.CheckedChanged
        If Not s_loading Then
            get_dgv_info()
        End If
    End Sub

    Private Sub cbo_navigate_goal_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_navigate_goal.SelectedIndexChanged

    End Sub

    Private Sub dgv_navigate_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_navigate.CellEndEdit

    End Sub
End Class
Public Class frm_schedules

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        fill_list()
    End Sub
    Sub fill_list()
        If s_loading Then Exit Sub

        fill_schedule("status < 10 AND machine='" & ComboBox1.SelectedItem("machine") & "' AND dept='" & LCase(ComboBox3.Text) & "'")
        With DataGridView1
            If Not .CurrentRow Is Nothing Then .Rows(.CurrentRow.Index).Selected = True
        End With
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.CurrentCell Is Nothing Then Exit Sub
        s_loading = True
        TextBox1.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value
        Button4.Visible = False
        Button1.Enabled = False
        Button2.Enabled = False
        With DataGridView1

            s_selected_job = dt_schedules.Select("prod_num = " & TextBox1.Text)

            If Not .CurrentRow Is Nothing Then
                .Rows(.CurrentRow.Index).Selected = True
                If Not get_material_info.zaw Is Nothing Then
                    Button4.Visible = True
                    Button4.Text = "View ZAW: " & get_material_info.zaw
                End If
                Button1.Enabled = True
                If s_department = "printing" Then Button2.Enabled = True
            End If

        End With
        s_loading = False
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress

        number_only(sender, e, False)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If s_loading Then Exit Sub
        If TextBox1.Text = Nothing Then
            fill_schedule("machine='" & ComboBox1.SelectedItem("machine") & "' AND dept='" & ComboBox3.Text & "'")
        Else
            fill_schedule("prod_num LIKE '%" & TextBox1.Text & "%' AND dept='" & ComboBox3.Text & "'")

        End If
        ' hide_columns()
        With DataGridView1
            .CurrentCell = Nothing
            If .RowCount = 1 Then
                .CurrentCell = .Rows(0).Cells(0)
                .Rows(0).Selected = True
                Button1.Enabled = True
                Button2.Enabled = True

                s_selected_job = dt_schedules.Select("prod_num = " & .Rows(0).Cells(0).Value)

                If Not .CurrentRow Is Nothing Then
                    .Rows(.CurrentRow.Index).Selected = True
                    If Not get_material_info.zaw Is Nothing Then
                        Button4.Visible = True
                        Button4.Text = "View ZAW: " & get_material_info.zaw
                    End If
                    Button1.Enabled = True
                    If s_department = "printing" Then Button2.Enabled = True
                End If

            Else
                Button1.Enabled = False
                Button2.Enabled = False
                Button4.Visible = False
            End If

        End With

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If s_department = LCase(ComboBox3.Text) Or s_department = "mounting" And LCase(ComboBox3.Text) = "printing" Then

            If Not IsNull(DataGridView1.CurrentCell) Then
                s_selected_job = dt_schedules.Select("prod_num = " & DataGridView1.CurrentRow.Cells(0).Value)
                Dim md As String = s_selected_job(0)("mat_desc_fin").ToString
                If md = Nothing Then md = s_selected_job(0)("mat_desc_semi").ToString
                If s_department = "mounting" Then md = s_selected_job(0)("label_desc")
                frm_main.add_job(DataGridView1.CurrentRow.Cells(0).Value, md, ComboBox2.Text, CheckBox2.Checked, ComboBox1.SelectedItem("machine"))
                TextBox1.Clear()
                If frm_main.dgv_summary.RowCount = 1 Then
                    s_selected_job = dt_schedules.Select("prod_num = " & frm_main.dgv_summary.Rows(0).Cells(0).Value)
                    frm_main.tb_prod_num.Text = s_selected_job(0)("prod_num")
                    frm_main.display_job()
                End If
                s_selected_job = dt_schedules.Select("prod_num = " & frm_main.dgv_summary.Rows(0).Cells(0).Value)
                Me.Close()
            Else
                MsgBox("Select a Job!")
            End If
        Else
            MsgBox("You can't add a job from another department")
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        s_selected_job = dt_schedules.Select("prod_num = " & DataGridView1.CurrentRow.Cells(0).Value)
        If Not frm_main.tb_prod_num.Text = Nothing Then
            s_current_job = frm_main.tb_prod_num.Text
        Else
            frm_main.tb_prod_num.Text = DataGridView1.CurrentRow.Cells(0).Value
        End If
        frm_main.display_job(True)
        frm_main.import_report()
        frm_main.l_prep_warning.Visible = True
        frm_main.tc_main.SelectedIndex = 1
        frm_main.b_report_part_roll.Text = "End Prep"
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If s_loading Then Exit Sub
        If CheckBox1.Checked Then
            If TextBox1.Text = Nothing Then
                fill_schedule("dept='" & ComboBox3.Text & "'")
            Else
                fill_schedule("prod_num LIKE '%" & TextBox1.Text & "%' AND dept='" & ComboBox3.Text & "'")

            End If
        Else
            If TextBox1.Text = Nothing Then
                fill_schedule("status<10 AND machine='" & ComboBox1.SelectedItem("machine") & "' AND dept='" & ComboBox3.Text & "'")
            Else
                fill_schedule("prod_num LIKE '%" & TextBox1.Text & "%' AND dept='" & ComboBox3.Text & "'")


            End If
        End If
        With DataGridView1
            .CurrentCell = Nothing
            If .RowCount = 1 Then
                .CurrentCell = .Rows(0).Cells(0)
                .Rows(0).Selected = True
                Button1.Enabled = True
                Button2.Enabled = True
            Else
                Button1.Enabled = False
                Button2.Enabled = False

            End If

        End With
    End Sub

    Private Sub schedules_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'If dt_schedules Is Nothing Then
        s_loading = True
        dt_schedules = sql.get_table("db_schedule", "plant", "=", s_plant)
        s_loading = False
        ' End If
        TextBox1.Clear()
        ComboBox3.DataSource = dr_to_dt(dt_department.Select("production=1"))
        ComboBox3.DisplayMember = "dept"
        If s_department = Nothing Then
            ComboBox3.SelectedIndex = 0
        Else
            ComboBox3.Text = s_department
        End If
        ComboBox2.Visible = False
        Button4.Visible = False
        Button2.Visible = False
        Label1.Visible = False
        CheckBox2.Visible = False
        Button1.Enabled = False
        Button2.Enabled = False
        Select Case LCase(ComboBox3.Text)
            Case "printing"
                Button2.Visible = True
            Case "mounting"
                CheckBox2.Visible = True
                With ComboBox2
                    .Items.Clear()
                    .Visible = True
                    For i As Integer = 0 To frm_main.lb_personel.Items.Count - 1
                        .Items.Add(frm_main.lb_personel.Items(i))
                    Next
                    ComboBox2.SelectedIndex = 0
                End With
                Label1.Visible = True
            Case "bagplant"
            Case "graphics"
        End Select

        s_loading = True
        ComboBox1.DataSource = dr_to_dt(dt_machines.Select("dept='" & ComboBox3.Text & "' AND enabled=1"))
        ComboBox1.DisplayMember = "name"
        select_cbo_item(ComboBox1, s_machine, "name")

        fill_schedule("status<10 AND machine='" & ComboBox1.SelectedItem("machine") & "' AND dept='" & ComboBox3.Text & "' AND plant=" & s_plant)

        select_cbo_item(ComboBox1, s_machine, "name")

        'hide_columns()
        With DataGridView1
            .CurrentCell = Nothing
            .ClearSelection()
        End With
        TextBox1.Focus()
        s_loading = False
    End Sub
    Sub fill_schedule(ByVal filter As String)
        Dim dr() As DataRow '= dt_schedules.Select(filter)
        If CheckBox1.Checked Then
            dr = dt_schedules.Select(filter)
        Else
            dr = dt_schedules.Select(filter & " AND status<10")
        End If

        With DataGridView1
            .Rows.Clear()
            For i As Integer = 0 To dr.Length - 1
                .Rows.Add()
                .Rows(i).Cells(0).Value = dr(i)("prod_num")
                .Rows(i).Cells(2).Value = get_status(dr(i)("status"))
                .Rows(i).Cells(3).Value = dr(i)("index_").ToString
                Select Case LCase(ComboBox3.Text)
                    Case "printing"
                        .Rows(i).Cells(1).Value = dr(i)("label_desc")
                    Case "bagplant", "filmsline"
                        If IsNull(dr(i)("mat_num_fin")) Then
                            .Rows(i).Cells(1).Value = dr(i)("mat_desc_semi")
                        Else
                            .Rows(i).Cells(1).Value = dr(i)("mat_desc_fin")
                        End If
                    Case Else
                        .Rows(i).Cells(1).Value = dr(i)("label_desc")

                End Select
            Next
            If dr.Length > 0 Then
                .Sort(.Columns(3), System.ComponentModel.ListSortDirection.Ascending)
                .PerformLayout()
            End If
        End With

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        With frm_custom_job

            .StartPosition = FormStartPosition.CenterParent
            ' .Close()
            .ShowDialog()

        End With
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        open_zaw(Strings.Right(Button4.Text, 5))
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ComboBox1.DataSource = sql.get_table("db_machines", "plant", "=", s_plant & " AND dept='" & ComboBox3.Text & "' AND enabled=1")
        If s_machine = Nothing Then
            ComboBox1.SelectedIndex = 0
        Else
            ComboBox1.Text = s_machine
        End If

        fill_list()
    End Sub


End Class
Public Class frm_custom_job


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles b_add_job.Click
        Dim prod_num, mat_num_semi, mat_num_fin, mat_desc_semi, mat_desc_fin, cyl_size, _
            inksys, inks, label_desc, _
            req_quant_1, req_uom_1, req_quant_2, req_uom_2, machine, no_up
        prod_num = sql.convert_value(tb_prod_num.Text)
        Dim dt As DataTable = sql.get_table("db_schedule", "prod_num", "=", prod_num)
        If Not dt Is Nothing Then
            MsgBox("That Job number already exists!")
            Exit Sub
        End If
        mat_num_semi = sql.convert_value(tb_mat_num_semi.Text)
        mat_num_fin = sql.convert_value(tb_mat_num_finish.Text)
        mat_desc_semi = sql.convert_value(Label11.Text)
        mat_desc_fin = sql.convert_value(Label3.Text)
        label_desc = sql.convert_value(tb_label_desc.Text)
        cyl_size = sql.convert_value(tb_cyl_size.Text)
        inksys = sql.convert_value(cbo_ink_system.Text)
        inks = Nothing
        inks = add_to_string(inks, tb_station_1.Text, ",", True)
        inks = add_to_string(inks, tb_station_2.Text, ",", True)
        inks = add_to_string(inks, tb_station_3.Text, ",", True)
        inks = add_to_string(inks, tb_station_4.Text, ",", True)
        inks = add_to_string(inks, tb_station_5.Text, ",", True)
        inks = add_to_string(inks, tb_station_6.Text, ",", True)
        inks = add_to_string(inks, tb_station_7.Text, ",", True)
        inks = add_to_string(inks, tb_station_8.Text, ",", True)
        inks = add_to_string(inks, tb_station_9.Text, ",", True)
        inks = add_to_string(inks, tb_station_10.Text, ",", True)
        inks = sql.convert_value(inks)
        req_quant_1 = sql.convert_value(tb_req_quant_1.Text)
        req_uom_1 = sql.convert_value(cbo_req_uom_1.Text)
        req_quant_2 = sql.convert_value(tb_req_quant_2.Text)
        req_uom_2 = sql.convert_value(cbo_req_uom_2.Text)
        machine = sql.convert_value(frm_main.cbo_machine.Text)
        no_up = sql.convert_value(nud_num_up.Value)
        sql.insert("db_schedule", "prod_num,mat_num_semi,mat_num_fin,mat_desc_semi,mat_desc_fin,label_desc,cyl_size,inksys,inks,req_quant_1,req_uom_1,req_quant_2,req_uom_2,machine,no_up,dept,plant,status", _
                    prod_num & "," & mat_num_semi & "," & mat_num_fin & "," & mat_desc_semi & "," & mat_desc_fin & "," & _
                    label_desc & "," & cyl_size & "," & inksys & "," & inks & "," & req_quant_1 & "," & req_uom_1 & "," & _
                    req_quant_2 & "," & req_uom_2 & "," & machine & "," & no_up & ",'" & s_department & "'," & s_plant & ",0")

        dt_schedules = sql.get_table("db_schedule", "plant", "=", s_plant)

        MsgBox("Job Added!")
        s_loading = True
        tb_prod_num.Clear()
        tb_label_desc.Clear()
        tb_mat_num_semi.Clear()
        tb_mat_num_finish.Clear()
        Label11.Text = "Enter material number starting with 200"
        Label3.Text = "Enter material number starting with 100"
        tb_cyl_size.Clear()
        cbo_ink_system.SelectedIndex = 0
        tb_station_1.Clear()
        tb_station_2.Clear()
        tb_station_3.Clear()
        tb_station_4.Clear()
        tb_station_5.Clear()
        tb_station_6.Clear()
        tb_station_7.Clear()
        tb_station_8.Clear()
        tb_station_9.Clear()
        tb_station_10.Clear()
        tb_req_quant_1.Clear()
        cbo_req_uom_1.SelectedIndex = 0
        tb_req_quant_2.Clear()
        cbo_req_uom_2.SelectedIndex = 0
        nud_num_up.Value = 1
        s_loading = False
    End Sub

    Private Sub tb_mat_num_finish_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb_mat_num_finish.KeyPress
        number_only(sender, e, False)
    End Sub
    Sub populate(ByVal dt As DataTable)
        Dim r As Integer = dt.Rows.Count - 1
        tb_mat_num_finish.Text = dt(r)("mat_num_fin").ToString
        tb_mat_num_semi.Text = dt(r)("mat_num_semi").ToString
        Label3.Text = dt(r)("mat_desc_fin").ToString
        Label11.Text = dt(r)("mat_desc_semi").ToString
        If IsNull(dt(r)("label_desc")) Then
            tb_label_desc.Text = dt(r)("mat_desc_fin").ToString
        Else
            tb_label_desc.Text = dt(r)("label_desc").ToString
        End If
        tb_cyl_size.Text = dt(r)("cyl_size").ToString
        Dim nu As String = dt(r)("no_up").ToString
        ' If IsNumeric(nu) Then NumericUpDown1.Value = nu

        Dim inks() As String = Split(dt(0)("inks").ToString, ",")
        For i As Integer = 0 To UBound(inks)
            If Not inks(i) = Nothing Then

                For Each tb As Control In GroupBox2.Controls
                    If TypeOf tb Is TextBox Then
                        If tb.Name.StartsWith("tb_station_") Then
                            If Replace(tb.Name, "tb_station_", Nothing) = i + 1 Then
                                tb.Text = inks(i)
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If
        Next

        cbo_ink_system.Text = dt(r)("inksys").ToString
        cbo_req_uom_1.Text = dt(r)("req_uom_1").ToString
        cbo_req_uom_2.Text = dt(r)("req_uom_2").ToString
    End Sub
    Private Sub tb_mat_num_finish_TextChanged(sender As Object, e As EventArgs) Handles tb_mat_num_finish.TextChanged
        If s_loading Then Exit Sub
        If tb_mat_num_finish.Text.Length = 9 Then
            s_loading = True
            Dim dt As DataTable = sql.get_table("db_schedule", "mat_num_fin", "=", tb_mat_num_finish.Text & " AND plant=" & s_plant)
            If Not dt Is Nothing Then
                populate(dt)
            Else
                Label3.Text = acsis.get_material_desc(tb_mat_num_finish.Text)
            End If
            s_loading = False
        End If
        display_message()
    End Sub

    Private Sub tb_prod_num_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb_prod_num.KeyPress
        number_only(sender, e, False)
    End Sub

    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs) Handles tb_prod_num.TextChanged, tb_req_quant_1.TextChanged, tb_req_quant_2.TextChanged
        display_message()
    End Sub
    Sub display_message()
        Dim enabled As Boolean = False

        If tb_prod_num.Text = Nothing Then
            Label13.Text = "Enter the Production Number"
        ElseIf Not tb_prod_num.Text.Length = 10 Or CInt(Strings.Left(tb_prod_num.Text, 1)) < 5 Then
            Label13.Text = "The Production number you have entered is invalid"
        ElseIf tb_req_quant_1.Text = Nothing Then
            Label13.Text = "Enter the required quantity (" & cbo_req_uom_1.Text & ")"
        ElseIf tb_req_quant_2.Text = Nothing Then
            Label13.Text = "Enter the required quantity (" & cbo_req_uom_2.Text & ")"
        ElseIf tb_mat_num_finish.Text = Nothing And tb_mat_num_semi.Text = Nothing Then
            Label13.Text = "Enter the material numbers on the datacard"
        Else
            enabled = True
        End If

        If enabled Then
            If Not tb_mat_num_finish.Text = Nothing Then
                If Not CInt(Strings.Left(tb_mat_num_finish.Text, 1)) = 1 Or Not tb_mat_num_finish.TextLength = 9 Then
                    Label13.Text = "The ""100"" Material number you have entered is invalid"
                    enabled = False
                End If
            ElseIf Not tb_mat_num_semi.Text = Nothing Then
                If Not CInt(Strings.Left(tb_mat_num_semi.Text, 1)) = 2 Or Not tb_mat_num_semi.TextLength = 9 Then
                    Label13.Text = "The ""200"" Material number you have entered is invalid"
                    enabled = False
                End If
            End If
        End If

        If enabled Then
            Label13.Visible = False
            b_add_job.Enabled = True
        Else
            b_add_job.Enabled = False
            Label13.Visible = True
        End If
    End Sub
    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb_cyl_size.KeyPress
        number_only(sender, e, False)
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb_req_quant_1.KeyPress
        number_only(sender, e, False)
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb_req_quant_2.KeyPress
        number_only(sender, e, False)
    End Sub

    Private Sub frm_custom_job_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        frm_schedules.fill_list()
        'frm_schedules.ShowDialog()
    End Sub

    Private Sub frm_custom_job_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim s() As String, c As Object
        c = sql.read("settings_general", "item", "=", "'ink_systems' AND plant=" & s_plant, "val")
        If Not IsNull(c) Then
            s = Split(c, ",")
            cbo_ink_system.Items.Clear()
            For i As Integer = 0 To s.Count - 1
                cbo_ink_system.Items.Add(s(i))
            Next
            cbo_ink_system.SelectedIndex = 0
        End If

        c = sql.read("settings_general", "item", "=", "'uom1' AND plant=" & s_plant, "val")
        If Not IsNull(c) Then
            s = Split(c, ",")
            cbo_req_uom_1.Items.Clear()
            For i As Integer = 0 To s.Count - 1
                cbo_req_uom_1.Items.Add(s(i))
            Next
            cbo_req_uom_1.SelectedIndex = 0
        End If

        c = sql.read("settings_general", "item", "=", "'uom2' AND plant=" & s_plant, "val")
        If Not IsNull(c) Then
            s = Split(c, ",")
            cbo_req_uom_2.Items.Clear()
            For i As Integer = 0 To s.Count - 1
                cbo_req_uom_2.Items.Add(s(i))
            Next
            cbo_req_uom_2.SelectedIndex = 0
        End If
        display_message()
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb_mat_num_semi.KeyPress
        number_only(sender, e, False)
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles tb_mat_num_semi.TextChanged
        If s_loading Then Exit Sub
        If tb_mat_num_semi.Text.Length = 9 Then
            s_loading = True
            Dim dt As DataTable = sql.get_table("db_schedule", "mat_num_semi", "=", tb_mat_num_semi.Text & " AND plant=" & s_plant)
            If Not dt Is Nothing Then
                populate(dt)
            Else
                Label11.Text = acsis.get_material_desc(tb_mat_num_semi.Text)

            End If
            s_loading = False
        End If
        display_message()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles b_cancel.Click
        Me.Hide()
    End Sub

    Private Sub ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_req_uom_1.SelectedIndexChanged, cbo_req_uom_2.SelectedIndexChanged
        display_message()
    End Sub
End Class
Public Class frm_bag_checks


    Private Sub burst_Click(sender As Object, e As EventArgs) Handles l_burst_1.Click, l_burst_2.Click, l_burst_5.Click, _
        l_burst_4.Click, l_burst_3.Click, l_burst_12.Click, l_burst_10.Click, l_burst_11.Click, l_burst_9.Click, l_burst_8.Click, _
        l_burst_6.Click, l_burst_7.Click, l_burst_14.Click, l_burst_13.Click
        Dim lbl As Label = CType(sender, Label)
        For Each c As Control In GroupBox1.Controls
            If TypeOf c Is Label Then
                If IsNumeric(Replace(c.Name, "l_burst_", Nothing)) Then
                    With c
                        .BackColor = Color.LightGray
                    End With
                End If
            End If
        Next
        lbl.BackColor = Color.Green

    End Sub

    Private Sub frm_bag_checks_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If s_department = "bagplant" Then e.Cancel = True
    End Sub

    Private Sub frm_bag_checks_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        display_data()
    End Sub
    Sub display_data()
        If frm_main.tb_prod_num.Text = Nothing Then Exit Sub
        Dim f, s, g
        Dim w, l, b_aim, b_lap, b_uap, b_fail, w_aim, w_lap, w_uap, l_aim, l_lap, l_uap As Integer
        'Dim m_info As materialInfo = get_material_info(False, s_selected_job(0)("mat_desc_semi"))
        f = get_material_info(True).formulation
        l = get_material_info.length
        w = get_material_info(True).width
        s = get_material_info.seal_type
        g = get_material_info(True).gauge
        dr_burst = dt_run_data.Select("Formulation='" & f & "' AND [Width ]=" & w & " AND [Seal Type] LIKE '%" & s & "%'")
        If dr_burst.Length > 0 Then
            If IsNull(dr_burst(0)("Aussie Burst Fail")) Then
                CheckBox5.Checked = False
                CheckBox5.Enabled = False
            Else
                CheckBox5.Enabled = True
                If IsNull(s_selected_job(0)("au_burst_data")) Then
                    CheckBox5.Checked = False
                Else
                    CheckBox5.Checked = s_selected_job(0)("au_burst_data")
                End If
                'set check based on aussie
            End If
            If CheckBox5.Checked Then
                b_aim = dr_burst(0)("Aussie Burst Fail")
            Else
                b_aim = dr_burst(0)("Burst Aim")
            End If
            b_lap = dr_burst(0)("Burst LAP")
            b_uap = dr_burst(0)("Burst UAP")
            b_fail = dr_burst(0)("Burst Fail")
            l_aim = dr_burst(0)("Length Aim")
            l_lap = dr_burst(0)("Length LAP")
            l_uap = dr_burst(0)("Length UAP")
            w_aim = dr_burst(0)("Width ")
            w_lap = dr_burst(0)("Width Low")
            w_uap = dr_burst(0)("Width High")
        End If

        nud_width.Minimum = w_lap - 20
        nud_width.Maximum = w_uap + 20
        nud_width.Value = w_aim

        nud_length.Minimum = l + l_lap - 20
        nud_length.Maximum = l + l_uap + 20
        nud_length.Value = l_aim + l

        For Each ctrl As Control In GroupBox1.Controls
            If TypeOf ctrl Is Label Then
                If ctrl.Name.Contains("l_burst_") And IsNumeric(Replace(ctrl.Name, "l_burst_", Nothing)) Then

                    Dim num As Integer = b_uap - Replace(ctrl.Name, "l_burst_", Nothing) + 2
                    ctrl.Text = num
                    Select Case True
                        Case num = b_fail
                            l_burst_fail.Top = ctrl.Top
                        Case num = b_aim
                            l_burst_aim.Top = ctrl.Top
                        Case num = b_lap
                            l_burst_lap.Top = ctrl.Top
                        Case num = b_uap
                            l_burst_uap.Top = ctrl.Top
                    End Select
                End If
            End If
        Next
        Label2.Text = "Bag Dimensions" & vbCr & "Target:" & vbCr & w & "mm x " & l & "mm" & vbCr & "(Width x Length)"

    End Sub


    Private Sub sb_Scroll(sender As Object, e As EventArgs)
        Label2.Text = "Bag Dimensions" & vbCr & _
            "Target:" & vbCr & _
        get_material_info.width & "mm x " & get_material_info.length & "mm" & vbCr & _
        "(Width x Length)" & vbCr & _
        "Actual:" & vbCr & _
        nud_width.Value & "mm x " & nud_length.Value & "mm"

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim burst As Integer = 0
        Dim tc As String = Nothing
        For Each c As Control In GroupBox1.Controls
            If TypeOf c Is Label Then
                If IsNumeric(Replace(c.Name, "l_burst_", Nothing)) Then
                    With c
                        If .BackColor = Color.Green Then burst = .Text
                    End With
                End If
            ElseIf TypeOf c Is CheckBox Then
                Dim cb As CheckBox = c
                If cb.Checked Then
                    tc = add_to_string(tc, Strings.Left(cb.Text, 1))
                End If
            End If
        Next

        If burst = 0 Then Exit Sub

        With dgv_burst
            Dim r As Integer = .RowCount
            .Rows.Add()
            .Rows(r).Cells("dgc_burst_prod_num").Value = s_selected_job(0)("prod_num")
            .Rows(r).Cells("dgc_burst_mat_num").Value = s_selected_job(0)("mat_num_in")
            .Rows(r).Cells("dgc_burst_test_num").Value = r + 1
            .Rows(r).Cells("dgc_burst_carton").Value = nud_carton_num.Value
            .Rows(r).Cells("dgc_burst_test_code").Value = tc
            .Rows(r).Cells("dgc_burst_burst").Value = burst
            .Rows(r).Cells("dgc_burst_length").Value = nud_length.Value
            .Rows(r).Cells("dgc_burst_width").Value = nud_width.Value
            .Rows(r).Cells("dgc_burst_date_").Value = Now()
            .Rows(r).Cells("dgc_burst_machine").Value = frm_main.cbo_machine.Text
        End With
        export_bursts()
        display_test_data()
    End Sub
    Sub display_test_data()
        If Not Me.Visible Then
            Exit Sub
        End If
        With Chart1.Series
            .Clear()
            .Add("Burst Data")
            .Add("AIM")
            With Chart1.Series("Burst Data")
                .ChartType = DataVisualization.Charting.SeriesChartType.Line
                .Color = Color.Blue
                .IsValueShownAsLabel = True
                .LabelBackColor = Color.White
                .BorderWidth = 2
            End With
            With Chart1.Series("AIM")
                .ChartType = DataVisualization.Charting.SeriesChartType.Line
                .Color = Color.Green
                .BorderWidth = 2
            End With
            If Not CheckBox5.Checked Then
                .Add("UAP")
                .Add("LAP")
                .Add("FAIL")
                With Chart1.Series("FAIL")
                    .ChartType = DataVisualization.Charting.SeriesChartType.Line
                    .Color = Color.Red
                    .BorderWidth = 2
                End With
                With Chart1.Series("UAP")
                    .ChartType = DataVisualization.Charting.SeriesChartType.Line
                    .Color = Color.DarkViolet
                    .BorderWidth = 2
                End With
                With Chart1.Series("LAP")
                    .ChartType = DataVisualization.Charting.SeriesChartType.Line
                    .Color = Color.DarkViolet
                    .BorderWidth = 2
                End With
            End If

        End With
        For count As Integer = 0 To dgv_burst.Rows.Count - 1
            Chart1.Series("Burst Data").Points.AddXY(count + 1, dgv_burst.Rows(count).Cells(3).Value)
            Chart1.Series("AIM").Points.AddXY(count + 1, dr_burst(0)("Burst Aim")) ', dr_burst(0)("Length Aim"), dr_burst(0)("Width"))
            If Not CheckBox5.Checked Then
                Chart1.Series("UAP").Points.AddXY(count + 1, dr_burst(0)("Burst UAP")) ', dr_burst(0)("Length Aim"), dr_burst(0)("Width"))
                Chart1.Series("LAP").Points.AddXY(count + 1, dr_burst(0)("Burst LAP")) ', dr_burst(0)("Length Aim"), dr_burst(0)("Width"))
                Chart1.Series("FAIL").Points.AddXY(count + 1, dr_burst(0)("Burst FAIL")) ', dr_burst(0)("Length Aim"), dr_burst(0)("Width"))
            End If
        Next
        If CheckBox5.Checked Then
            Chart1.ChartAreas(0).AxisY.Minimum = l_burst_14.Text
            Chart1.ChartAreas(0).AxisY.Maximum = l_burst_1.Text
        Else
            Chart1.ChartAreas(0).AxisY.Minimum = dr_burst(0)("Burst FAIL") - 5
            Chart1.ChartAreas(0).AxisY.Maximum = dr_burst(0)("Burst UAP") + 5
        End If

        Chart1.ChartAreas(0).AxisY.Interval = 1
        Chart1.ChartAreas(0).AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
        If dgv_burst.RowCount - 15 > 0 Then
            Chart1.ChartAreas(0).AxisX.Minimum = dgv_burst.RowCount - 15
            Chart1.ChartAreas(0).AxisX.Maximum = dgv_burst.RowCount
        Else
            Chart1.ChartAreas(0).AxisX.Minimum = 1
            Chart1.ChartAreas(0).AxisX.Maximum = 15
        End If

    End Sub
    Sub import_bursts()
        Dim tempdt As New DataTable, row As Integer
        tempdt = sql.get_table("db_burst_tests", "prod_num", "=", s_selected_job(0)("prod_num"))
        With dgv_burst
            .Rows.Clear()
            If Not IsNothing(tempdt) Then
                For Each r As DataRow In tempdt.Rows
                    .Rows.Add()
                Next
                For Each r As DataRow In tempdt.Rows
                    row = r("test_num") - 1
                    For c As Integer = 0 To .ColumnCount - 1
                        Dim c_name As String = Replace(.Columns(c).Name, "dgc_burst_", Nothing)
                        If Not IsNull(r(c_name)) Then
                            .Rows(row).Cells(c).Value = r(c_name)
                        End If
                    Next

                Next

            End If
        End With
        display_test_data()
    End Sub

    Sub export_bursts()
        Dim str, columns, values, c_name As String, prod_num As Long = frm_main.tb_prod_num.Text
        sql.delete("db_burst_tests", "prod_num", "=", prod_num)
        With dgv_burst
            If .RowCount > 0 Then
                For r As Integer = 0 To .RowCount - 1
                    str = Nothing
                    columns = Nothing
                    values = Nothing
                    For c As Integer = 0 To .Columns.Count - 2
                        c_name = Replace(.Columns(c).Name, "dgc_burst_", Nothing)

                        If IsDBNull(.Rows(r).Cells(c).Value) Then
                            str = .Rows(r).Cells(c).Value.ToString
                        Else
                            str = .Rows(r).Cells(c).Value
                        End If

                        str = sql.convert_value(str)

                        If columns = Nothing Then
                            columns = c_name
                            values = str
                        Else
                            columns = columns & "," & c_name
                            values = values & "," & str
                        End If

                    Next
                    sql.insert("db_burst_tests", columns, values)

                Next
            End If
        End With
    End Sub
    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If s_loading Then Exit Sub
        s_loading = True
        If CheckBox5.Checked Then
            l_burst_uap.Visible = False
            l_burst_lap.Visible = False
            l_burst_fail.Visible = False
            l_burst_aim.Text = "Minimum"
        Else
            l_burst_uap.Visible = True
            l_burst_lap.Visible = True
            l_burst_fail.Visible = True
            l_burst_aim.Text = "Aim"
        End If
        display_data()
        s_loading = False
    End Sub
End Class
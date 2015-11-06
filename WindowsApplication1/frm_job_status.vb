Public Class frm_job_status
    Dim sw_status As New Stopwatch
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim ts As TimeSpan = sw_status.Elapsed

        If ts.Minutes = 10 Then
            sw_status.Restart()
            add_info()
        End If

        Label1.Text = "Auto-Refresh in " & 9 - ts.Minutes & ":" & (60 - ts.Seconds).ToString.PadLeft(2, "0")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        add_info()
        sw_status.Restart()
    End Sub

    Private Sub frm_job_status_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Timer1.Enabled = False
        sw_status.Stop()
    End Sub

    Private Sub frm_job_status_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        ComboBox1.DataSource = sql.get_table("db_departments", "plant", "=", s_plant & " AND production=1")
        ComboBox1.DisplayMember = "dept"
        If s_department = Nothing Then
            ComboBox1.SelectedIndex = 0
        Else
            ComboBox1.Text = s_department
        End If
        add_info()
        Timer1.Enabled = True
        sw_status.Start()
    End Sub
    Sub add_info()
        Dim dt As DataTable = sql.get_table("db_schedule", "plant", "=", s_plant & " AND dept='" & ComboBox1.Text & "'")
        populate(dt.Select("status=2"), True, Color.White)
        populate(dt.Select("status=9"), False, Color.LightPink)
        If Format(Now, "ddd") = "Mon" Then
            ' populate(dt.Select("fin_date>'" & Now.AddDays(-1) & "'"), False, Color.LightGreen)
            populate(dt.Select("fin_date>'" & Now.AddDays(-3) & "'"), False, Color.LightGreen)
        Else
            populate(dt.Select("fin_date>'" & Now.AddDays(-1) & "'"), False, Color.LightGreen)
        End If
    End Sub
    Sub populate(ByVal dr() As DataRow, ByVal clear As Boolean, ByVal color As Color)
        'Dim dt As DataTable = sql.get_table("db_schedule", "plant", "=", s_plant & " AND status=2  AND dept='" & s_department & "'")

        With DataGridView1
            sw_status.Restart()
            If clear Then .Rows.Clear()
            If dr Is Nothing Then Exit Sub
            For r As Integer = 0 To dr.Length - 1
                Dim prod_num As Long = dr(r)("prod_num"), uom As String = dr(r)("req_uom_1"), req As Decimal = dr(r)("req_quant_1"), total As String
                .Rows.Add()
                .Rows(.RowCount - 1).DefaultCellStyle.BackColor = color
                .Rows(.RowCount - 1).Cells(0).Value = prod_num
                If IsNull(dr(r)("mat_num_semi")) Then
                    .Rows(.RowCount - 1).Cells(1).Value = dr(r)("mat_num_fin")
                Else
                    .Rows(.RowCount - 1).Cells(1).Value = dr(r)("mat_num_semi")
                End If
                .Rows(.RowCount - 1).Cells(2).Value = dr(r)("label_desc")
                .Rows(.RowCount - 1).Cells(3).Value = uom
                .Rows(.RowCount - 1).Cells(4).Value = FormatNumber(req, 2)
                If uom = "M" Or uom = "KM" Then
                    total = sql.sum("db_produced", "km_out", "prod_num", "=", prod_num)
                ElseIf uom = "ROL" Then
                    total = sql.count("db_produced", "prod_num", "=", prod_num & " AND kg_out<>'0'")
                ElseIf uom = "IMP" Then
                    total = (sql.sum("db_produced", "km_out", "prod_num", "=", prod_num) * 1000) / (dr(r)("cyl_size") / 1000)
                Else
                    total = sql.sum("db_produced", "kg_out", "prod_num", "=", prod_num)
                End If
                If Not IsNumeric(total) Then
                    total = 0
                End If
                If uom = "ROL" Then
                    .Rows(.RowCount - 1).Cells(5).Value = CInt(total)
                    .Rows(.RowCount - 1).Cells(6).Value = CInt(req - total)
                Else
                    .Rows(.RowCount - 1).Cells(5).Value = FormatNumber(total, 2, , , TriState.False)
                    .Rows(.RowCount - 1).Cells(6).Value = FormatNumber(req - total, 2, , , TriState.False)
                End If
                .Rows(.RowCount - 1).Cells(7).Value = FormatPercent(total / req)
                .Rows(.RowCount - 1).Cells(8).Value = dr(r)("machine")
            Next
        End With

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        add_info()
        sw_status.Restart()
    End Sub
End Class
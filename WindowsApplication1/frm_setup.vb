Public Class frm_setup
    Dim result() As DataRow
    Dim cyl, ink, web, sign, speed As Decimal
    Private Sub frm_setup_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        s_loading = True
        ComboBox1.Items.Clear()
        If s_department = "mounting" Then
            Me.Width = 149
            Me.Height = 144
            Button1.Width = 121
            Button1.Location = New Point(7, 79)
            NumericUpDown3.Visible = False
            Label2.Text = "Plates"
            Label3.Visible = False
            Label4.Visible = False
            Label5.Visible = False
            NumericUpDown1.Maximum = 8
            NumericUpDown2.Maximum = 100
        Else
            Me.Width = 270
            Me.Height = 184
            Button1.Width = 241
            Button1.Location = New Point(7, 117)
            NumericUpDown3.Visible = True
            Label2.Text = "Inks Out"
            Label3.Visible = True
            Label4.Visible = True
            Label5.Visible = True
        End If
        With frm_main.dgv_summary
            For r As Integer = 0 To .RowCount - 1
                ComboBox1.Items.Add(.Rows(r).Cells("dgc_summary_prod_num").Value)
            Next
            ComboBox1.SelectedIndex = .CurrentRow.Index
        End With
        NumericUpDown1.Value = 0
        NumericUpDown2.Value = 0
        NumericUpDown3.Value = 0

        If s_department = "printing" Then get_values()
        calc()
        s_loading = False
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        calc()
    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        calc()
    End Sub

    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        calc()
    End Sub
    Sub get_values()
        s_loading = True
        result = dt_machines.Select("name='" & s_machine & "'")
        With frm_main.dgv_summary
            If Not IsNull(.Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_cyl_in").Value) Then
                If .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_cyl_in").Value > (result(0)("stations") * 3) Then
                    NumericUpDown1.Maximum = result(0)("stations") * 3
                    NumericUpDown1.Value = result(0)("stations") * 3
                    .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_cyl_in").Value = NumericUpDown1.Value
                End If
            End If

            If Not IsNull(.Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_cyl_in").Value) Then NumericUpDown1.Value = .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_cyl_in").Value
            If Not IsNull(.Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_ink_in").Value) Then NumericUpDown2.Value = .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_ink_in").Value
            If IsNull(.Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_mpm").ToolTipText) Then
                NumericUpDown3.Value = result(0)("speed")
            Else
                NumericUpDown3.Value = .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_mpm").ToolTipText
            End If
            NumericUpDown1.Maximum = result(0)("stations") * 3
            NumericUpDown2.Maximum = result(0)("stations")
        End With
        s_loading = False
    End Sub
    Sub calc()
        If s_loading Then Exit Sub

        If Not IsNull(result) And s_department = "printing" Then
            cyl = (NumericUpDown1.Value * result(0)("cylinder"))
            ink = (NumericUpDown2.Value * result(0)("colour"))
            sign = result(0)("sign_off")
            speed = NumericUpDown3.Value
            If RadioButton1.Checked Then
                web = 0
            ElseIf RadioButton2.Checked Then
                web = result(0)("web_ds")
            Else
                web = result(0)("web_ss")
            End If
            Label5.Text = FormatNumber(((cyl + ink + web) / 60) + sign, 2)
        End If
        send_data()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub RadioButton1_Click(sender As Object, e As EventArgs) Handles RadioButton1.Click
        calc()
    End Sub

    Private Sub RadioButton2_Click(sender As Object, e As EventArgs) Handles RadioButton2.Click
        calc()
    End Sub

    Private Sub RadioButton3_Click(sender As Object, e As EventArgs) Handles RadioButton3.Click
        calc()
    End Sub

    Sub send_data()
        With frm_main.dgv_summary
            .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_cyl_in").Value = CInt(NumericUpDown1.Value).ToString
            .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_ink_in").Value = CInt(NumericUpDown2.Value).ToString
            If s_department = "printing" Then
                .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_setup").ToolTipText = Label5.Text
                .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_min_deck").ToolTipText = result(0)("colour") + result(0)("cylinder")
                .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_mpm").ToolTipText = CInt(speed)
                .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_st_setup").Value = FormatNumber(Label5.Text, 2)
                .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_st_speed").Value = CInt(speed)
                .Rows(ComboBox1.SelectedIndex).Cells("dgc_summary_st_mindeck").Value = result(0)("colour") + result(0)("cylinder")
            End If
            frm_main.export_summary()
        End With
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If s_loading Then Exit Sub
        get_values()
    End Sub
End Class
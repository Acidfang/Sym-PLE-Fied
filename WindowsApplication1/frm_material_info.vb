Public Class frm_material_info

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Strings.Left(TextBox1.Text, 1) = "5" Then
            MsgBox(TextBox1.Text & " is not a material number!")
        Else
            add_mat()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(Keys.Enter) Or e.KeyChar = Chr(Keys.Tab) Then
            If Strings.Left(TextBox1.Text, 1) = "5" Then
                MsgBox(TextBox1.Text & " is not a material number!")
            Else
                add_mat()
            End If
        Else
            number_only(sender, e, False)
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Sub add_mat()
        If Not TextBox1.Text = Nothing And IsNumeric(TextBox1.Text) Then

            Dim str As String = Nothing
            Dim dr() As DataRow = dt_schedules.Select("mat_num_in=" & TextBox1.Text & " AND mat_desc_in IS NOT NULL")
            If dr.Length > 0 Then
                str = dr(0)("mat_desc_in")
            Else

                If Not str = Nothing Then
                    If Mid(str, 35, 1) = "N" Then
                        Dim resp As MsgBoxResult = MsgBox("That is a printed material number." & _
                                                          vbCr & "Are you sure you want to use this as the material going in?", _
                                                          MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Printed Material Number!")

                        If resp = MsgBoxResult.No Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                End If
                str = acsis.get_material_desc(TextBox1.Text)
            End If

            frm_main.cbo_mat_num_in.Items.Add(TextBox1.Text)
            frm_main.cbo_mat_num_in.Text = TextBox1.Text
            frm_main.cbo_mat_desc_in.Items.Add(str)
            frm_main.cbo_mat_desc_in.Text = str
            dr = dt_schedules.Select("prod_num=" & frm_main.tb_prod_num.Text)
            Dim s As String = Nothing

            If IsNull(dr(0)("mat_num_in")) Then
                s = Nothing
            ElseIf IsNull(dr(0)("mat_num_in_1")) Then
                s = "_1"
            Else
                s = "_2"
            End If
            sql.update("db_schedule", "mat_num_in" & s & "='" & TextBox1.Text & "',mat_desc_in" & s & "='" & _
                       str & "',access", Convert.ToSByte(s_sql_access), "prod_num", "=", frm_main.tb_prod_num.Text & " AND plant=" & s_plant)

            dr(0)("mat_num_in" & s) = TextBox1.Text
            dr(0)("mat_desc_in" & s) = str

            Me.Close()
            TextBox1.Clear()
            Me.Close()
        Else
            MsgBox("Enter a valid material number.")
        End If

    End Sub

    Private Sub material_info_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        TextBox1.Clear()
        TextBox1.Focus()
    End Sub


End Class
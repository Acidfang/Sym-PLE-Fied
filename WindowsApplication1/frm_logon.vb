Public Class frm_logon

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim logins As New DataTable, old_user As String = s_user, str As String = sql.read("DP_Logins", "Login_Id", "=", "'" & TextBox1.Text & "'", "Password", "symple")
        reset = False
        If IsNull(str) Then
            MsgBox("User does not exist, please check your entry!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "invalid login!")
            Exit Sub
        ElseIf Not UCase(str) = UCase(TextBox2.Text) Then
            MsgBox("Wrong password, please check your entry!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "invalid password!")
            Exit Sub
        End If
        s_user_name = sql.read("DP_Logins", "Login_Id", "=", "'" & TextBox1.Text & "'", "User_Name", "symple")
        s_user = UCase(TextBox1.Text)
        s_pass = TextBox2.Text
        If w_sessions(0).winWnd Is Nothing Then
            acsis.logon()
        ElseIf Not old_user = Nothing Then
            If Not old_user = s_user Then
                While Not w_sessions(0).winWnd Is Nothing
                    w_sessions(0).process.Kill()
                    acsis.get_controls(0, True, False, True)
                End While
                acsis.logon()
            End If
        End If
        frm_main.tb_start_date.Text = Format(Now, "dd/MM/yyyy").ToString
        Me.Close()
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            Button1_Click(Me, EventArgs.Empty)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim logins As New DataTable, old_user As String = s_user, str As String = sql.read("DP_Logins", "Login_Id", "=", "'" & TextBox1.Text & "'", "Password", "symple")
        reset = True
        If IsNull(str) Then
            MsgBox("User does not exist, please check your entry!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "invalid login!")
            Exit Sub
        ElseIf Not UCase(str) = UCase(TextBox2.Text) Then
            MsgBox("Wrong password, please check your entry!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "invalid password!")
            Exit Sub
        End If
        s_user_name = sql.read("DP_Logins", "Login_Id", "=", "'" & TextBox1.Text & "'", "User_Name", "symple")
        s_user = UCase(TextBox1.Text)
        s_pass = TextBox2.Text
        s_preparing = True
        acsis.logon()
        With frm_schedules
            .StartPosition = FormStartPosition.CenterParent
            .ShowDialog()
        End With
        ' symple.Kill()
        'frm_main.cb_sql_symple.Checked = False
        's_symple_access = False
        'If sw_symple.IsRunning Then
        '    sw_symple.Reset()
        'Else
        '    sw_symple.Start()
        'End If
        'frm_main.tb_start_date.Text = Format(Now, "dd/MM/yyyy").ToString
        Me.Close()
    End Sub

    Private Sub frm_logon_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If reset Then
            s_shift = Nothing
        End If
    End Sub
    Dim reset As Boolean = True
    Private Sub logon_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        reset = True
        TextBox1.Text = s_user
        TextBox1.Focus()
        If TextBox1.Text = Nothing Or TextBox2.Text = Nothing Then
            Button1.Enabled = False
            Button2.Enabled = False
        End If
        If Not s_user = Nothing Then
            TextBox2.Focus()
        End If
        TextBox2.Text = s_pass
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = Nothing Or TextBox2.Text = Nothing Then
            Button1.Enabled = False
            Button2.Enabled = False
        Else
            Button1.Enabled = True
            Button2.Enabled = True
        End If

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox1.Text = Nothing Or TextBox2.Text = Nothing Then
            Button1.Enabled = False
            Button2.Enabled = False
        Else
            Button1.Enabled = True
            Button2.Enabled = True
        End If
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.F4 Then
            ' Call your sub method here  .....

            Application.Exit()
            ' then avoid the key to reach the current control
            e.Handled = False
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        s_department = "graphics"
        s_shift = Nothing

        With frm_startup

            Dim rv As DataRowView, fieldinfo As System.Reflection.FieldInfo

            For Each rv In .cbo_department.Items
                If LCase(rv(0)) = s_department Then
                    fieldinfo = rv.Row.GetType.GetField("_rowID", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
                    .cbo_department.SelectedIndex = fieldinfo.GetValue(rv.Row) - 1
                End If
            Next
            If .Visible Then
                .Close()
            End If
        End With
        Me.Close()
    End Sub


End Class
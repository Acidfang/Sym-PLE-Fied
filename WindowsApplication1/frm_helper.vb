Public Class frm_helper

    Private Sub frm_helper_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        TextBox1.Clear()
        TextBox2.Clear()
        If s_department = "printing" Then
            Button1.Text = "Add Helper"
        Else
            Button1.Text = "Add Personel"
        End If
        Button1.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox1.Text = Nothing Or TextBox2.Text = Nothing Then
            Button1.Enabled = False
        ElseIf Not TextBox1.Text = s_user Then
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim logins As New DataTable, str As String = sql.read("DP_Logins", "Login_Id", "=", "'" & TextBox1.Text & "'", "Password", "symple")
        If IsNull(str) Then
            MsgBox("User does not exist, please check your entry!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "invalid login!")
            Exit Sub
        ElseIf Not UCase(str) = UCase(TextBox2.Text) Then
            MsgBox("Wrong password, please check your entry!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "invalid password!")
            Exit Sub
        End If
        str = sql.read("DP_Logins", "Login_Id", "=", "'" & TextBox1.Text & "'", "User_Name", "symple")
        frm_main.lb_personel.Items.Add(str)
        For i As Integer = 0 To frm_main.lb_personel.Items.Count - 1
            If s_assistant = Nothing Then
                s_assistant = frm_main.lb_personel.Items(i)
            Else
                s_assistant = s_assistant & "/" & frm_main.lb_personel.Items(i)
            End If
        Next
        frm_main.export_summary()
        Me.Close()
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            Button1_Click(Me, EventArgs.Empty)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = Nothing Then
            TextBox1.Text = s_user
            MsgBox("Enter your password to confirm")
        ElseIf Not TextBox2.Text = Nothing Then
            Dim logins As New DataTable, str As String = sql.read("DP_Logins", "Login_Id", "=", "'" & TextBox1.Text & "'", "Password", "symple")
            If IsNull(str) Then
                MsgBox("User does not exist, please check your entry!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "invalid login!")
                Exit Sub
            ElseIf Not UCase(str) = UCase(TextBox2.Text) Then
                MsgBox("Wrong password, please check your entry!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "invalid password!")
                Exit Sub
            End If
            s_assistant = s_user_name
            Me.Close()
        Else
            MsgBox("Enter your password to confirm")
        End If
    End Sub


End Class
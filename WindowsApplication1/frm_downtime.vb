Imports System.Net.Mail
Public Class frm_downtime

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim str, item As String
        ListBox2.Items.Clear()
        With frm_main.dgv_downtime_bs01
            If .RowCount > 0 Then
                If .CurrentRow.Index > -1 Then
                    item = .CurrentRow.Cells(2).Value
                Else
                    item = Nothing
                End If
            Else
                item = Nothing
            End If
        End With

        For r As Integer = 0 To dt_downtime_reasons.Rows.Count - 1
            str = dt_downtime_reasons(r)("main")
            If str = ListBox1.SelectedItem Then
                str = dt_downtime_reasons(r)("reason")
                If ListBox2.Items.IndexOf(str) = -1 Then
                    ListBox2.Items.Add(str)
                    If item = str And s_dt_edit Then
                        ListBox2.SelectedIndex = ListBox2.Items.Count - 1
                    End If
                End If
            End If


        Next
        ListBox2.Items.Add("Other")
        If item = "Other" And s_dt_edit Then ListBox2.SelectedIndex = ListBox2.Items.Count - 1
        okbutton()
        'If ListBox2.SelectedIndex = -1 Then ListBox2.SelectedIndex = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If NumericUpDown1.Value > 0 Then
            Dim dgv As DataGridView
            If s_downtime_type = "BS01" Then
                dgv = frm_main.dgv_downtime_bs01
            ElseIf s_downtime_type = "CS01" Then
                dgv = frm_main.dgv_downtime_cs01
            Else
                dgv = frm_main.dgv_downtime_bs01
            End If
            With dgv
                Dim r As Integer

                If s_dt_edit Then
                    r = .CurrentRow.Index
                Else
                    .Rows.Add()
                    r = .RowCount - 1
                End If

                .Rows(r).Cells(0).Value = NumericUpDown1.Value
                .Rows(r).Cells(1).Value = ListBox1.Items(ListBox1.SelectedIndex)
                .Rows(r).Cells(2).Value = ListBox2.Items(ListBox2.SelectedIndex)
                .Rows(r).Cells(3).Value = TextBox1.Text
                .Rows(r).Cells(4).Value = CheckBox1.Checked
                .Rows(r).Cells(5).Value = CheckBox2.Checked

            End With
            frm_main.export_downtime(s_downtime_type)
            frm_main.calc_downtime(frm_main.tb_prod_num.Text)
            frm_main.calc_times(frm_main.get_prod_num_row(frm_main.tb_prod_num.Text))
            Me.Close()
        Else
            MsgBox("You need to enter the time lost")
        End If
        If CheckBox2.Checked Then

            show_status("Sending 5 why email to " & ComboBox1.Text)

            Try
                Dim Smtp_Server As New SmtpClient("smtp.gmail.com")
                Dim e_mail As New MailMessage()
                Smtp_Server.UseDefaultCredentials = False
                Smtp_Server.Credentials = New Net.NetworkCredential("5whyactionneeded", "5whyaction")
                Smtp_Server.Port = 587
                Smtp_Server.EnableSsl = True
                Smtp_Server.Host = "smtp.gmail.com"
                e_mail = New MailMessage()
                e_mail.From = New MailAddress("5whyactionneeded@gmail.com")
                Dim email As String = ComboBox1.SelectedItem("email")
                Dim manager As String = sql.read("settings_email", "manager", "=", "1 AND dept='" & s_department & "'", "email")
                Dim graphics As String = sql.read("settings_email", "graphics", "=", "1 AND dept='" & s_department & "'", "email")
                If ListBox1.Items(ListBox1.SelectedIndex) = "Graphics" Then
                    email = email & "," & graphics
                End If
                e_mail.To.Add(email & "," & manager)
                e_mail.Subject = "5 Why's needed!"
                e_mail.IsBodyHtml = False
                Dim multiple(frm_main.cbo_mat_num_in.Items.Count - 1) As String
                frm_main.cbo_mat_num_in.Items.CopyTo(multiple, 0)
                e_mail.Body = "A 5 Whys has been requested." & vbCr & vbCr & _
                    "Requested by: " & s_user_name & " (" & s_user & ")" & vbCr & _
                    "Job Number: " & frm_main.tb_prod_num.Text & vbCr & _
                    "Material # In: " & String.Join(",", multiple) & vbCr & _
                    "Material # Out: " & frm_main.tb_mat_num_out.Text & vbCr & _
                    "Details: " & ListBox1.Items(ListBox1.SelectedIndex) & " - " & ListBox2.Items(ListBox2.SelectedIndex) & vbCr & _
                    "Time lost: " & NumericUpDown1.Value & " hrs" & vbCr & _
                    "Happened in setup: " & CheckBox1.Checked & vbCr & _
                    "Comments: " & TextBox1.Text & vbCr & vbCr & _
                    "This is an automated message, do not reply to this as it will not be read."
                Smtp_Server.Send(e_mail)
                frm_status.Close()

                'MsgBox("5 Whys email sent sucesfully to " & ComboBox1.Text)

            Catch error_t As Exception
                MsgBox(error_t.ToString)
            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub downtime_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        CheckBox2.Checked = False
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        dt_downtime_reasons = sql.get_table("settings_downtime", "plant", "=", s_plant & " AND dept='" & s_department & "' AND code='" & s_downtime_type & "'")
        ComboBox1.DataSource = sql.get_table("settings_email", "plant", "=", s_plant & " AND dept='" & s_department & "' AND manager=0 AND graphics=0")
        ComboBox1.DisplayMember = "name"
        NumericUpDown1.Value = 0
        TextBox1.Clear()
        Dim str, item As String, dgv As DataGridView
        If s_downtime_type = "BS01" Then
            dgv = frm_main.dgv_downtime_bs01
        ElseIf s_downtime_type = "CS01" Then
            dgv = frm_main.dgv_downtime_cs01
        Else
            dgv = frm_main.dgv_downtime_bs01
        End If
        With dgv
            If .RowCount > 0 Then
                If .CurrentRow.Index > -1 Then
                    item = .CurrentRow.Cells("downtime_main").Value
                Else
                    item = Nothing
                End If
            Else
                item = Nothing
            End If
        End With
        If Not dt_downtime_reasons Is Nothing Then

            For r As Integer = 0 To dt_downtime_reasons.Rows.Count - 1
                str = dt_downtime_reasons(r)("main")
                If ListBox1.Items.IndexOf(str) = -1 Then
                    ListBox1.Items.Add(str)
                    If item = str And s_dt_edit Then
                        With dgv

                            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                            NumericUpDown1.Value = .CurrentRow.Cells("downtime_time_lost").Value
                            TextBox1.Text = .CurrentRow.Cells("downtime_comment").Value
                            CheckBox1.Checked = .CurrentRow.Cells("downtime_setup").Value
                            CheckBox2.Checked = .CurrentRow.Cells("downtime_five_why").Value
                        End With
                    End If
                End If

            Next
        End If

        '    If ListBox1.SelectedIndex = -1 Then ListBox1.SelectedIndex = 0
    End Sub


    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        ComboBox1.Visible = CheckBox2.Checked
        Label3.Visible = CheckBox2.Checked
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        okbutton()
    End Sub
    Sub okbutton()
        If ListBox1.SelectedIndex > -1 And ListBox2.SelectedIndex > -1 Then
            If TextBox1.Text = Nothing Or NumericUpDown1.Value = 0 Then
                Label4.Visible = True
                Button1.Enabled = False
            Else
                Label4.Visible = False
                Button1.Enabled = True
            End If
        Else
            Label4.Text = "Please select the downtime reason"
            Label4.Visible = True
            Button1.Enabled = False
        End If
    End Sub
    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        If NumericUpDown1.Value = 0 Then
            Label4.Text = "You need to add time"
        ElseIf ListBox1.SelectedIndex > -1 Or ListBox2.SelectedIndex > -1 Then
            Label4.Text = "Please select the downtime reason"
        Else
            Label4.Text = "You need to add information"
        End If
        okbutton()
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        okbutton()
    End Sub
End Class
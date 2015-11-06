Public Class frm_plate_add_box

    Private Sub frm_plate_add_box_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        s_loading = True
        frm_main.plate_search(True)
        ComboBox1.Items.Clear()
        For i As Integer = 0 To dt_boxes.Rows.Count - 1
            ComboBox1.Items.Add(dt_boxes(i)("box_size"))
        Next
        ComboBox1.SelectedIndex = 0
        reset_boxes()
        s_loading = False
    End Sub
    Sub populate_boxes()
        Dim dt As DataTable = sql.get_table("db_plate_location", "box_size", "=", "'" & Strings.Left(ComboBox1.Text, 1) & "' AND plate_num IS NULL AND plant=" & s_plant)
        TextBox1.Text = frm_main.nud_plate_boxes.Value * frm_main.nud_plate_columns.Value * frm_main.nud_plate_rows.Value
        If dt Is Nothing Then
            TextBox3.Text = 0
        Else
            TextBox3.Text = dt.Rows.Count
        End If
        TextBox2.Text = TextBox1.Text - TextBox3.Text
        Dim rand As New Random
        Dim n = rand.Next(dt.Rows.Count - 1)
        Dim loc As String = dt(n)("box_size") & dt(n)("col").ToString.PadLeft(3, "0") & dt(n)("row")
        TextBox6.Text = loc
        dt = sql.get_table("db_plate_location", "box_size", "=", "'" & Strings.Left(ComboBox1.Text, 1) & "' AND col=" & dt(n)("col") & " AND row='" & dt(n)("row") & "' AND plate_num IS NULL AND plant=" & s_plant)
        Label5.Text = dt.Rows.Count
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If s_loading Then Exit Sub
        frm_main.cbo_plate_boxsize.SelectedIndex = ComboBox1.SelectedIndex
        populate_boxes()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = Nothing Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dt As DataTable = sql.get_table("db_plate_location", "plate_num", "=", "'" & TextBox4.Text & "' AND plant=" & s_plant)
        With frm_main.dgv_plates
            If Not dt Is Nothing Then
                Dim loc As String = dt(0)("box_size") & dt(0)("col").ToString.PadLeft(3, "0") & dt(0)("row")
                Dim response As MsgBoxResult
                response = MsgBox("That box number already exists in " & loc & vbCr & "Do you want to try to put these boxes together?" & vbCr & "Total: " & dt.Rows.Count & vbCr & "Description: " & dt(0)("description"), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Box already exists!")
                If response = MsgBoxResult.Yes Then

                    Dim bs, row, desc, comment, checked, archive, plates As String
                    Dim col As Integer = dt(0)("col")
                    bs = "'" & dt(0)("box_size") & "'"
                    row = "'" & dt(0)("row") & "'"
                    desc = dt(0)("description").ToString
                    comment = dt(0)("comment").ToString
                    checked = dt(0)("checked").ToString
                    plates = dt(0)("plates").ToString
                    archive = dt(0)("archived").ToString
                    With frm_main.cbo_plate_boxsize
                        For i As Integer = 0 To .Items.Count - 1
                            If Strings.Left(.Items(i), 1) = Replace(bs, "'", Nothing) Then
                                .SelectedIndex = i
                                Exit For
                            End If
                        Next
                    End With
                    dt = sql.get_table("db_plate_location", "box_size", "=", bs & " AND col=" & col & " AND row=" & row & " AND plate_num IS NULL AND plant=" & s_plant)
                    If Not dt Is Nothing Then
                        For r As Integer = 0 To .RowCount - 1
                            If IsNull(.Rows(r).Cells(0).Value) And .Rows(r).Cells(1).Value = loc Then
                                .Rows(r).Cells(0).Value = TextBox4.Text
                                .Rows(r).Cells(2).Value = desc
                                .Rows(r).Cells(3).Value = comment
                                .Rows(r).Cells(4).Value = checked
                                .Rows(r).Cells(5).Value = plates
                                .Rows(r).Cells(6).Value = archive
                                frm_main.print_box_label(loc, TextBox4.Text)
                                frm_main.export_plates()
                                frm_main.plate_search(True)
                                Exit Sub
                            End If

                        Next
                    Else
                        response = MsgBox("There is no space in " & loc & vbCr & "The box will be moved, do you want to continue?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Boxes need to move!")
                        If response = MsgBoxResult.Yes Then
                            Dim old_loc As String = loc
                            Dim first_row As Integer = 0
                            dt = sql.get_table("db_plate_location", "plate_num", "=", "'" & TextBox4.Text & "' AND plant=" & s_plant)
                            Dim count As Integer = dt.Rows.Count + 1
                            Dim old_count As Integer = count
                            Dim box_to_move As String = Nothing
                            'remove old plates from inventory
                            For r As Integer = 0 To .RowCount - 1
                                If .Rows(r).Cells(0).Value = TextBox4.Text Then
                                    .Rows(r).Cells(0).Value = Nothing
                                    .Rows(r).Cells(2).Value = Nothing
                                    .Rows(r).Cells(3).Value = Nothing
                                    .Rows(r).Cells(4).Value = Nothing
                                    .Rows(r).Cells(5).Value = Nothing
                                    .Rows(r).Cells(6).Value = Nothing
                                End If
                            Next
                            'find enough space
                            With frm_main.dgv_plates
                                Dim slots As Integer = 0
                                loc = Nothing
                                Do While loc = Nothing
                                    loc = find_empty_slot(frm_main.dgv_plates, count, False)
                                    'not found slots - find slots -1
                                    If loc = Nothing Then count = count - 1

                                Loop

                                Dim stored As Integer = 0
                                If old_count = count Then
                                    'put boxes away, print labels - nothing to remove
                                    For r As Integer = 0 To .RowCount - 1
                                        If .Rows(r).Cells(1).Value = loc And IsNull(.Rows(r).Cells(0).Value) And Not stored = count Then
                                            .Rows(r).Cells(0).Value = TextBox4.Text
                                            .Rows(r).Cells(2).Value = desc
                                            .Rows(r).Cells(3).Value = comment
                                            .Rows(r).Cells(4).Value = checked
                                            .Rows(r).Cells(5).Value = plates
                                            .Rows(r).Cells(6).Value = archive
                                            frm_main.print_box_label(loc, TextBox4.Text)
                                            stored = stored + 1
                                            'put box away, count them and stop
                                        End If

                                    Next
                                    frm_main.export_plates()
                                Else
                                    'find a box number to remove from loc
                                    'get value of box and store, show user what it is
                                    Dim removed As Integer = 0
                                    For r As Integer = 0 To .RowCount - 1
                                        If .Rows(r).Cells(1).Value = loc And Not IsNull(.Rows(r).Cells(0).Value) And Not removed = old_count - count Then
                                            MsgBox("You need to remove this box" & vbCr & "Plate Number: " & .Rows(r).Cells(0).Value & "Location: " & .Rows(r).Cells(1).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "A box needs to be removed")
                                            .Rows(r).Cells(0).Value = TextBox4.Text
                                            .Rows(r).Cells(2).Value = desc
                                            .Rows(r).Cells(3).Value = comment
                                            .Rows(r).Cells(4).Value = checked
                                            .Rows(r).Cells(5).Value = plates
                                            .Rows(r).Cells(6).Value = archive
                                            frm_main.print_box_label(loc, TextBox4.Text)
                                            removed = removed + 1
                                            'put box away, count them and stop
                                        End If

                                    Next
                                    frm_main.export_plates()

                                End If

                            End With

                        Else
                            Exit Sub
                        End If
                        bs = bs

                    End If
                Else
                    Exit Sub
                End If
            End If
            'Exit For
            '        End If
            '    Next
            For i As Integer = 0 To .Rows.Count - 1
                If IsNull(.Rows(i).Cells(0).Value) And .Rows(i).Cells(1).Value = TextBox6.Text Then
                    .Rows(i).Cells(0).Value = TextBox4.Text
                    .Rows(i).Cells(2).Value = TextBox5.Text
                    .Rows(i).Cells(4).Value = frm_main.tb_start_date.Text
                    frm_main.print_box_label(.Rows(i).Cells(1).Value, .Rows(i).Cells(0).Value)
                    Exit For
                End If
            Next
        End With
        frm_main.export_plates()
        frm_main.plate_search(True)
        reset_boxes()
    End Sub

    Sub reset_boxes()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        populate_boxes()

    End Sub

    Private Sub TextBox6_DoubleClick(sender As Object, e As EventArgs) Handles TextBox6.DoubleClick
        populate_boxes()
    End Sub

End Class
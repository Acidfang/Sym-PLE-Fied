Public Class Comments_
    Dim lastsel As Integer = -1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        insert_comment()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With DataGridView1
            .Rows.RemoveAt(.CurrentRow.Index)
            export_comments()
            .ClearSelection()
            Button2.Enabled = False
            If .RowCount = 0 Then
                Button1.Text = "Add Comment"
            End If
        End With
        lastsel = -1
    End Sub

    Sub export_comments()
        Dim str, values, name As String
        If CheckBox1.Checked Then
            name = "'" & 1234567890 & "'"
        Else
            name = "'" & frm_main.tb_prod_num.Text & "'"
        End If
        sql.delete("db_comments", "prod_num", "=", name)
        With DataGridView1
            For r As Integer = 0 To .RowCount - 1
                values = name
                For c = 0 To .Columns.Count - 1
                    str = Replace(.Rows(r).Cells(c).Value, "'", "''")
                    If str = Nothing Then str = "NULL"
                    If Not IsNumeric(str) Then str = "'" & str & "'"
                    values = values & "," & str
                Next
                values = values & "," & Convert.ToSByte(s_sql_access)
                sql.insert("db_comments", "prod_num,date_,user_,text,access", values)
            Next
        End With

    End Sub

    Sub get_comments()
        Dim str As String
        If CheckBox1.Checked Then
            str = "'" & 1234567890 & "'"
        Else
            str = "'" & frm_main.tb_prod_num.Text & "'"
        End If
        Dim cmnt As DataTable = sql.get_table("db_comments", "prod_num", "=", str)
        With DataGridView1
            .Rows.Clear()
            If Not IsNull(cmnt) Then
                For i As Integer = 0 To cmnt.Rows.Count - 1
                    .Rows.Add()
                    .Rows(i).Cells(0).Value = cmnt(i)("date_")
                    .Rows(i).Cells(1).Value = cmnt(i)("user_")
                    .Rows(i).Cells(2).Value = cmnt(i)("text")
                Next


            End If
            .ClearSelection()
        End With

    End Sub

    Sub insert_comment()

        With DataGridView1
            If .SelectedCells.Count = 0 Then
                get_comments()
                .Rows.Add()
                Dim row As Integer = .RowCount - 1
                .Rows(row).Cells(0).Value = Format(Now, "dd/MM/yyyy")
                .Rows(row).Cells(1).Value = s_user
                .Rows(row).Cells(2).Value = TextBox1.Text
            Else
                .Rows(.CurrentRow.Index).Cells(2).Value = TextBox1.Text
            End If
            .ClearSelection()
        End With
        TextBox1.Clear()
        export_comments()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            insert_comment()
        End If
    End Sub

    Private Sub CheckBox1_CheckStateChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckStateChanged
        get_comments()

    End Sub

    Private Sub Comments__Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Button2.Enabled = False
        Button1.Text = "Add Comment"
        If frm_main.tb_prod_num.Text = Nothing Then
            CheckBox1.Checked = True
            CheckBox1.Enabled = False
        Else
            CheckBox1.Checked = False
            CheckBox1.Enabled = True
        End If
        get_comments()
        DataGridView1.ClearSelection()
        TextBox1.Clear()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        With DataGridView1

            If .Rows(.CurrentRow.Index).Cells(1).Value = s_user Or s_user = "RATTANJ" Then
                If Button1.Text = "Edit Comment" And lastsel = .CurrentRow.Index Then
                    TextBox1.Clear()
                    Button1.Text = "Add Comment"
                    .ClearSelection()
                    Button2.Enabled = False
                Else
                    TextBox1.Text = .Rows(.CurrentRow.Index).Cells(2).Value
                    Button1.Text = "Edit Comment"
                    .Rows(.CurrentRow.Index).Selected = True
                    Button2.Enabled = True
                End If
            Else
                TextBox1.Clear()
                Button1.Text = "Add Comment"
                .ClearSelection()
                Button2.Enabled = False
            End If
            lastsel = .CurrentRow.Index
        End With

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = Nothing Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub

End Class
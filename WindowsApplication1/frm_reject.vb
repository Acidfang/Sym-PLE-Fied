Public Class frm_reject

    Dim rejectdb As DataTable
    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        If e.RowIndex > -1 Then
            DataGridView1.Rows(e.RowIndex).Selected = True
        End If
    End Sub

    Private Sub frm_reject_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        rejectdb = sql.get_table_full("db_reject_codes", True)
        rejectdb.DefaultView.Sort = "code asc"
        TextBox1.Clear()
        Dim rejectcode As String = Nothing
        If Not IsNull(frm_main.dgv_reject.CurrentRow.Cells(5).Value) Then rejectcode = frm_main.dgv_reject.CurrentRow.Cells(5).Value
        With DataGridView1
            .DataSource = rejectdb
            .Columns(0).Width = 35
            .Columns(1).Width = 150
            .Columns(2).Width = 350
            If Not IsNull(rejectcode) Then
                For r = 0 To .RowCount - 1
                    If .Rows(r).Cells(0).Value = rejectcode Then
                        .CurrentCell = .Rows(r).Cells(0)
                        .Rows(r).Selected = True
                        Exit For
                    End If
                Next
            End If
        End With


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim dt As DataTable = rejectdb
        dt.DefaultView.RowFilter = "reason LIKE '%" & TextBox1.Text & "%'"
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With frm_main.dgv_reject
            .CurrentRow.Cells("dgc_reject_code").Value = DataGridView1.CurrentRow.Cells(0).Value.ToString.PadLeft(3, "0")
            .CurrentRow.Cells("dgc_reject_reason").Value = DataGridView1.CurrentRow.Cells(1).Value
        End With
        frm_main.export_reject()
        Me.Close()
    End Sub
End Class
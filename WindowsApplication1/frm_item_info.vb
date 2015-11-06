Public Class frm_item_info

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With s_roll_info
            .item_id = DataGridView1.CurrentRow.Cells(0).Value
            .kg_in = DataGridView1.CurrentRow.Cells(1).Value
            .km_in = DataGridView1.CurrentRow.Cells(2).Value
            .mat_num = DataGridView1.CurrentRow.Cells(3).Value
        End With
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex > -1 Then
            DataGridView1.Rows(e.RowIndex).Selected = True
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If
    End Sub

    Private Sub frm_item_info_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        DataGridView1.ClearSelection()
        Button1.Enabled = False
    End Sub
End Class
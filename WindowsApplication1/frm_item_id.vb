Public Class frm_item_id

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles b_ok.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_itemlist.CellClick
        With dgv_itemlist
            If .RowCount > 0 And e.RowIndex > -1 Then
                .Rows(.CurrentRow.Index).Selected = True

                s_roll_info.num_in = .CurrentRow.Cells("dgc_itemlist_num_in").Value
                s_roll_info.item_id = .CurrentRow.Cells("dgc_itemlist_item_id").Value
                s_roll_info.kg_in = .CurrentRow.Cells("dgc_itemlist_kg_in").Value
                s_roll_info.prod_num = .CurrentRow.Cells("dgc_itemlist_prod_num").Value
                s_roll_info.mat_num = .CurrentRow.Cells("dgc_itemlist_mat_num").Value

                b_ok.Enabled = True
            Else
                b_ok.Enabled = False
            End If
        End With

    End Sub

    Private Sub frm_item_id_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        dgv_itemlist.ClearSelection()
    End Sub

End Class
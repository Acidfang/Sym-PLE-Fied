Public Class frm_merge_roll

    Private Sub frm_merge_roll_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        populate()
    End Sub

    Sub populate()
        With frm_main.dgv_production
            DataGridView1.Rows.Clear()
            DataGridView2.Rows.Clear()
            For r As Integer = 0 To .RowCount - 1

                If Not has_info(r) And Not .Rows(r).Cells("dgc_production_num_in").Value.ToString = "-" And .Rows(r).Visible Then
                    DataGridView1.Rows.Add(.Rows(r).Cells("dgc_production_num_in").Value, .Rows(r).Cells("dgc_production_item_id").Value, _
                                           .Rows(r).Cells("dgc_production_kg_in").Value, r, .Rows(r).Cells("dgc_production_kg_in_orig").Value, _
                                           .Rows(r).Cells("dgc_production_prod_num_orig").Value, .Rows(r).Cells("dgc_production_mat_num").Value)
                End If

            Next
        End With
        DataGridView1.ClearSelection()
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        With DataGridView1
            DataGridView2.Rows.Clear()
            .Rows(.CurrentRow.Index).Selected = True
            For r As Integer = .CurrentRow.Index + 1 To .RowCount - 1
                DataGridView2.Rows.Add(.Rows(r).Cells(0).Value, .Rows(r).Cells(1).Value, .Rows(r).Cells(2).Value, _
                                       .Rows(r).Cells(3).Value, .Rows(r).Cells(4).Value, .Rows(r).Cells(5).Value, _
                                       .Rows(r).Cells(6).Value)
            Next
            DataGridView2.ClearSelection()
        End With

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        DataGridView2.Rows(DataGridView2.CurrentRow.Index).Selected = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With frm_main.dgv_production
            Dim r As Integer = DataGridView1.CurrentRow.Cells(3).Value
            .Rows(r).Cells("dgc_production_num_in").Value = _
                (DataGridView1.CurrentRow.Cells(0).Value & " & " & DataGridView2.CurrentRow.Cells(0).Value).ToString

            .Rows(r).Cells("dgc_production_item_id").Value = _
                (DataGridView1.CurrentRow.Cells(1).Value & " & " & DataGridView2.CurrentRow.Cells(1).Value).ToString

            .Rows(r).Cells("dgc_production_kg_in").Value = _
                FormatNumber(CDec(DataGridView1.CurrentRow.Cells(2).Value) + CDec(DataGridView2.CurrentRow.Cells(2).Value), 1, , , TriState.False)

            .Rows(r).Cells("dgc_production_kg_in_orig").Value = _
                DataGridView1.CurrentRow.Cells(4).Value & "," & DataGridView2.CurrentRow.Cells(4).Value

            .Rows(r).Cells("dgc_production_prod_num_orig").Value = _
                DataGridView1.CurrentRow.Cells(5).Value & "," & DataGridView2.CurrentRow.Cells(5).Value

            .Rows(r).Cells("dgc_production_mat_num").Value = _
                DataGridView1.CurrentRow.Cells(6).Value & "," & DataGridView2.CurrentRow.Cells(6).Value

            For r = DataGridView2.CurrentRow.Cells(3).Value To .RowCount - 1
                If r = DataGridView2.CurrentRow.Cells(3).Value Or .Rows(r).Cells("dgc_production_num_in").Value.ToString = "-" Then
                    .Rows(r).Visible = False
                Else
                    Exit For

                End If

            Next
            frm_main.export_report()
            'frm_main.import_report()
        End With
        populate()
    End Sub


End Class
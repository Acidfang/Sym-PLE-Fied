Public Class frm_plate_add

    Private Sub frm_plate_add_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        frm_main.plate_search(True)
        dgv_plates.Rows.Clear()
        For col As Integer = 1 To frm_main.nud_plate_columns.Value
            For row As Integer = 1 To frm_main.nud_plate_rows.Value
                For boxes As Integer = 1 To frm_main.nud_plate_boxes.Value
                    dgv_plates.Rows.Add(Nothing, Strings.Left(frm_main.cbo_plate_boxsize.Text, 1) & col.ToString.PadLeft(3, "0") & Convert.ToChar(row + 64))
                Next
            Next
        Next
        For r As Integer = 0 To frm_main.dgv_plates.RowCount - 1
            Dim plate_num As String = frm_main.dgv_plates.Rows(r).Cells(0).Value.ToString

            If Not IsNull(plate_num) Then
                For new_r As Integer = 0 To dgv_plates.RowCount - 1
                    If IsNull(dgv_plates.Rows(new_r).Cells(0).Value) And frm_main.dgv_plates.Rows(r).Cells(1).Value = dgv_plates.Rows(new_r).Cells(1).Value Then

                        For c As Integer = 0 To frm_main.dgv_plates.Columns.Count - 1
                            dgv_plates.Rows(new_r).Cells(c).Value = frm_main.dgv_plates.Rows(r).Cells(c).Value
                        Next
                        Exit For
                    End If
                Next

            End If

        Next
        dgv_plates.ClearSelection()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frm_main.dgv_plates.Rows.Clear()
        For r As Integer = 0 To dgv_plates.RowCount - 1
            Dim pn, loc, desc, check, plates, com As String
            If IsNull(dgv_plates.Rows(r).Cells(0).Value) Then
                pn = Nothing
            Else
                pn = dgv_plates.Rows(r).Cells(0).Value
            End If
            If IsNull(dgv_plates.Rows(r).Cells(1).Value) Then
                loc = Nothing
            Else
                loc = dgv_plates.Rows(r).Cells(1).Value
            End If

            If IsNull(dgv_plates.Rows(r).Cells(2).Value) Then
                desc = Nothing
            Else
                desc = dgv_plates.Rows(r).Cells(2).Value
            End If

            If IsNull(dgv_plates.Rows(r).Cells(3).Value) Then
                com = Nothing
            Else
                com = dgv_plates.Rows(r).Cells(3).Value
            End If

            If IsNull(dgv_plates.Rows(r).Cells(4).Value) Then
                check = Nothing
            Else
                check = dgv_plates.Rows(r).Cells(4).Value
            End If

            If IsNull(dgv_plates.Rows(r).Cells(5).Value) Then
                plates = Nothing
            Else
                plates = dgv_plates.Rows(r).Cells(5).Value
            End If

            frm_main.dgv_plates.Rows.Add(pn, loc, desc, check, plates)
        Next
        frm_main.export_plates()
        sql.update("plate_options", "quant", frm_main.nud_plate_boxes.Value, "box_size", "=", "'" & frm_main.cbo_plate_boxsize.Text & _
                   "' AND plant=" & s_plant)

        sql.update("plate_options", "col", frm_main.nud_plate_columns.Value, "box_size", "=", "'" & frm_main.cbo_plate_boxsize.Text & _
                   "' AND plant=" & s_plant)

        sql.update("plate_options", "row", frm_main.nud_plate_rows.Value, "box_size", "=", "'" & frm_main.cbo_plate_boxsize.Text & _
                   "' AND plant=" & s_plant)

        frm_main.dgv_plates.ClearSelection()
        frm_main.dgv_plates.CurrentCell = Nothing

        Me.Close()
    End Sub
End Class
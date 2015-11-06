Public Class frm_plate_search

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim dt As DataTable = sql.get_table("db_plate_location", "plant", "=", s_plant & " AND plate_num IS NOT NULL")
        DataGridView1.Rows.Clear()
        If TextBox1.Text = Nothing Then Exit Sub
        For i As Integer = 0 To dt.Rows.Count - 1
            If Not IsNull(dt(i)("plates")) Then
                Dim plates() As String = Split(dt(i)("plates"), ",")
                For Each plate As String In plates
                    If plate.Contains(TextBox1.Text) Then
                        Dim loc As String = dt(i)("box_size").ToString & dt(i)("col").ToString.PadLeft(3, "0") & dt(i)("row").ToString

                        DataGridView1.Rows.Add(plate, dt(i)("plate_num"), loc, dt(i)("description"))
                    End If
                Next
            End If
        Next
    End Sub

    Private Sub frm_plate_search_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        TextBox1.Clear()
    End Sub
End Class
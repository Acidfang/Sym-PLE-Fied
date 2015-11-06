Public Class frm_reprint

    Private Sub frm_reprint_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Button1.Text = "Check All"
        With frm_main
            If Not .tb_prod_num.Text = Nothing Then
                With .dgv_production
                    Dim dp As Integer = Strings.Right(.Columns("dgc_production_km_out").DefaultCellStyle.Format, 1)
                    Dim str() As String = Split(.Columns("dgc_production_km_out").HeaderText, " ")
                    dgv_reprint.Columns(2).HeaderText = str(0)
                    dgv_reprint.Rows.Clear()
                    For i As Integer = 0 To .RowCount - 1
                        If Not IsNull(.Rows(i).Cells("dgc_production_kg_out").Value) And Not IsNull(.Rows(i).Cells("dgc_production_km_out").Value) Then
                            dgv_reprint.Rows.Add(.Rows(i).Cells("dgc_production_num_out").Value, _
                                                   FormatNumber(.Rows(i).Cells("dgc_production_kg_out").Value, 1, , , TriState.False), _
                                                   FormatNumber(.Rows(i).Cells("dgc_production_km_out").Value, dp, , , TriState.False))
                        End If
                    Next
                End With
            End If
        End With

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim c As Boolean
        If Button1.Text = "Check All" Then
            Button1.Text = "Uncheck All"
            c = True
        Else
            Button1.Text = "Check All"
            c = False
        End If
        With dgv_reprint
            For i As Integer = 0 To .RowCount - 1
                .Rows(i).Cells(3).Value = c
            Next
        End With

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim str() As String = Split(frm_main.dgv_production.Columns("dgc_production_km_out").HeaderText, " ")

        Dim m_info As materialInfo = get_material_info()
        Dim section As String = "Contents"
        Dim quant_out As Decimal = 0

        With dgv_reprint
            For i As Integer = 0 To .RowCount - 1

                If .Rows(i).Cells("dgc_reprint_reprint").Value = True Then
                    .Rows(i).Cells("dgc_reprint_reprint").Value = False
                    Select Case s_selected_job(0)("req_uom_1")
                        Case "KM"
                            quant_out = .Rows(i).Cells("dgc_reprint_km_out").Value
                        Case "KG"
                            quant_out = .Rows(i).Cells("dgc_reprint_kg_out").Value
                        Case Else
                            quant_out = quant_out
                    End Select
                    Dim meters As String = .Rows(i).Cells("dgc_reprint_km_out").Value
                    If Strings.Left(frm_main.dgv_production.Columns("dgc_production_km_out").HeaderText, 2) = "KM" Then
                        meters = meters * 1000
                    End If
                    Select Case m_info.material_type
                        Case material_type.barrier
                            acsis.print_ticket(frm_main.tb_prod_num.Text, quant_out, meters, s_selected_job(0)("req_uom_1"), _
                                               Nothing, Nothing, .Rows(i).Cells("dgc_reprint_num_out").Value, 0, 0)
                        Case material_type.film, material_type.laminate
                            If Strings.Left(frm_main.tb_mat_num_out.Text, 1) = "2" Then
                                If s_department = "printing" Then
                                    meters = frm_main.tb_label_desc.Text & " (" & CInt(meters) & "m)"
                                Else
                                    meters = CInt(meters) & "m"
                                End If
                                acsis.print_ticket(frm_main.tb_prod_num.Text, quant_out, meters, s_selected_job(0)("req_uom_1"), _
                                                          section, meters, .Rows(i).Cells("dgc_reprint_num_out").Value, 0, 0)
                            Else
                                If meters = m_info.length Then
                                    acsis.print_ticket(frm_main.tb_prod_num.Text, quant_out, meters, s_selected_job(0)("req_uom_1"), _
                                                       Nothing, Nothing, .Rows(i).Cells("dgc_reprint_num_out").Value, 0, 0)
                                Else
                                    section = "Length"
                                    acsis.print_ticket(frm_main.tb_prod_num.Text, quant_out, meters, s_selected_job(0)("req_uom_1"), _
                                                              section, meters, .Rows(i).Cells("dgc_reprint_num_out").Value, 0, 0)
                                End If

                            End If
                        Case Else
                            meters = meters
                    End Select
                End If
            Next
        End With
        Button1.Text = "Uncheck All"

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_reprint.CellClick
        With dgv_reprint

            If e.ColumnIndex = 3 Then
                If .CurrentCell.Value = True Then
                    .CurrentCell.Value = False
                Else
                    .CurrentCell.Value = True
                End If
            End If
            .ClearSelection()
            .CurrentCell = Nothing
        End With
    End Sub
End Class
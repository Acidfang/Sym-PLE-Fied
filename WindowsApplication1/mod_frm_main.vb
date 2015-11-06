Module mod_frm_main

    Sub get_production_info(ByVal row As Integer)
        frm_main.l_production_estimation.Text = "Estimated Quantity:"
        frm_main.l_production_remaining.Text = "Remaining:"
        frm_main.l_roll_info_length.Text = "Roll length IN (Estimate):"
        frm_main.l_roll_info_speed.Text = " Minute Roll:"

        With frm_main.dgv_production
            Dim calculated As Boolean = True
            Dim i_info As item_info = get_item_info(.CurrentRow.Index, True)
            Dim m_info As materialInfo = get_material_info(False)
            If .SelectedCells.Count > 1 Then Exit Sub
            Dim i As Integer = 0
            Select Case s_department
                Case "printing", "filmsline", "bagplant"
                    If Not s_extruder Then
                        Dim w_in, w_out As Integer
                        w_in = get_material_info(True).width

                        w_out = get_material_info().width
                        If Not w_in = 0 And w_out = 0 And Not w_in = w_out Then
                            i = get_roll_length(i_info.kg_in, s_selected_job(0)("mat_desc_in"))
                        ElseIf s_department = "bagplant" Then
                            If i_info.km_in = 0 Then
                                i = get_roll_length(i_info.kg_in, s_selected_job(0)("mat_desc_in"))
                                i_info.km_in = i / 1000
                            Else
                                i = i_info.km_in * 1000
                                calculated = False
                            End If
                        Else
                            If IsNull(s_selected_job(0)("mat_desc_semi")) And Not IsNull(s_selected_job(0)("mat_desc_fin")) Then
                                i = get_roll_length(i_info.kg_in, s_selected_job(0)("mat_desc_fin"))
                            Else
                                i = get_roll_length(i_info.kg_in, s_selected_job(0)("mat_desc_semi"))
                            End If
                        End If
                    End If
                    frm_main.l_production_remaining.Text = "Remaining:"
                    Dim lbl As Label = frm_main.l_roll_info_length
                    If s_department = "bagplant" Then
                        lbl = frm_main.l_production_estimation
                    End If
                    If calculated Then
                        lbl.Text = "Roll length IN (Estimate): " & i & " M"
                    Else
                        lbl.Text = "Roll length IN: " & i & " M"
                    End If
                    frm_main.l_roll_info_speed.Visible = True
                    frm_main.nud_min_per_roll.Visible = True
                    frm_main.l_roll_info_speed.Location = New Point(47, 79)
                    Select Case s_machine
                        Case "COMEXI 3"

                            frm_main.nud_min_per_roll.Visible = False
                            frm_main.l_roll_info_speed.Location = New Point(6, 79)

                            If i > 2400 Then
                                frm_main.l_roll_info_speed.Text = "Recommended Speed: 200 MPM"
                            ElseIf i > 2200 Then
                                frm_main.l_roll_info_speed.Text = "Recommended Speed: " & CInt((i / 12) / 5) * 5 & " MPM"
                            ElseIf i > 1750 Then
                                frm_main.l_roll_info_speed.Text = "Recommended Speed: 180 MPM"
                            Else
                                frm_main.l_roll_info_speed.Text = "Recommended Speed: " & CInt((i / 10) / 5) * 5 - 5 & " MPM"
                            End If

                        Case Else
                            If s_department = "bagplant" Then
                                frm_main.nud_min_per_roll.Visible = False
                                frm_main.l_roll_info_speed.Location = New Point(6, 79)
                            Else
                                frm_main.l_roll_info_speed.Text = " Minute Roll: " & CInt((i / frm_main.nud_min_per_roll.Value) / 5) * 5 - 5 & " MPM"
                            End If

                    End Select
                    Dim est As String = Nothing
                    If s_department = "bagplant" Then
                        'do total
                        est = FormatNumber(i / (m_info.length / 1000) / frm_main.nud_items_out.Value, 2, , , TriState.False) & " Cartons"
                        frm_main.l_roll_info_speed.Text = "Estimated: " & est
                        frm_main.l_production_remaining.Text = "Bags on Roll: " & FormatNumber(i / (m_info.length / 1000), 0, , , TriState.False)
                    Else
                        Select Case .Columns(.CurrentCell.ColumnIndex).Name
                            Case "dgc_production_kg_out"
                                If Not IsNull(.CurrentRow.Cells("dgc_production_km_out").Value) Then

                                    If get_conversion_factor(True, frm_main.tb_mat_num_out.Text, True) = 0 Then
                                        est = FormatNumber(get_roll_length(.CurrentRow.Cells("dgc_production_km_out").Value, s_selected_job(0)("mat_desc_semi")) / 1000, 1, , , TriState.False) & " KG (Not AVG)"
                                    Else
                                        est = FormatNumber(get_conversion_factor(True, frm_main.tb_mat_num_out.Text, True) * .CurrentRow.Cells("dgc_production_km_out").Value, 1, , , TriState.False) & " KG (AVG)"
                                    End If
                                End If

                            Case "dgc_production_km_out"
                                Dim str() As String = Split(.Columns("dgc_production_km_out").HeaderText, " ")
                                Dim dp As Integer = Strings.Right(.Columns("dgc_production_km_out").DefaultCellStyle.Format, 1)

                                If Not IsNull(.CurrentRow.Cells("dgc_production_kg_out").Value) Then
                                    If get_conversion_factor(True, frm_main.tb_mat_num_out.Text, False) = 0 Then
                                        est = FormatNumber(get_roll_length(.CurrentRow.Cells("dgc_production_kg_out").Value, s_selected_job(0)("mat_desc_semi")) / 1000, dp, , , TriState.False) & " " & str(0) & " (Not AVG)"
                                    Else
                                        est = FormatNumber(get_conversion_factor(True, frm_main.tb_mat_num_out.Text, False) * .CurrentRow.Cells("dgc_production_kg_out").Value, dp, , , TriState.False) & " " & str(0) & " (AVG)"
                                    End If

                                End If

                            Case Else
                                i = i
                        End Select
                        frm_main.l_production_estimation.Text = "Estimated Quantity: " & est

                    End If
                    If IsNumeric(.CurrentRow.Cells("dgc_production_kg_in").Value) Then
                        Dim kg_i, kg_o, rts, reject As Decimal
                        kg_i = .CurrentRow.Cells("dgc_production_kg_in").Value
                        If Not IsNull(.CurrentRow.Cells("dgc_production_kg_out").Value) Then kg_o = .CurrentRow.Cells("dgc_production_kg_out").Value
                        If Not IsNull(.CurrentRow.Cells("dgc_production_rts").Value) Then rts = .CurrentRow.Cells("dgc_production_rts").Value
                        If Not IsNull(.CurrentRow.Cells("dgc_production_reject").Value) Then reject = .CurrentRow.Cells("dgc_production_reject").Value

                    End If
                    'End If
                    'If kg_i - kg_o - rts - reject > 0 Then
                    If s_department = "bagplant" Then
                        i = i_info.km_in * 1000 - i_info.used_km
                        frm_main.l_roll_info_length.Text = "Remaining: " & _
                            FormatNumber(i, 0, , , TriState.False) & " M (" & FormatNumber(i / (m_info.length / 1000), 0, , , TriState.False) & " Bags)"

                    Else
                        frm_main.l_production_remaining.Text = "Remaining: " & _
                            FormatNumber(i_info.remaining, 1, , , TriState.False) & " KG"

                    End If
                    'FormatNumber(kg_i - kg_o - rts - reject, 1, , , TriState.False) & " KG"
                    'Else
                    'frm_main.l_production_remaining.Text = "Remaining:"
                    'End If


                Case "filmsline"
                Case Else
                    i = i
            End Select
        End With
    End Sub
    Function has_info(row As Integer) As Boolean
        With frm_main.dgv_production
            If IsNull(.Rows(row).Cells("dgc_production_kg_out").Value) And _
                       IsNull(.Rows(row).Cells("dgc_production_km_out").Value) And _
                       IsNull(.Rows(row).Cells("dgc_production_num_out").Value) And _
                       IsNull(.Rows(row).Cells("dgc_production_rts").Value) And _
                       IsNull(.Rows(row).Cells("dgc_production_reject").Value) And _
                       IsNull(.Rows(row).Cells("dgc_production_scrap").Value) And _
                       IsNull(.Rows(row).Cells("dgc_production_barcode").Value) And _
                       IsNull(.Rows(row).Cells("dgc_production_rho_left").Value) And _
                       IsNull(.Rows(row).Cells("dgc_production_rho_right").Value) And _
                       IsNull(.Rows(row).Cells("dgc_production_comments").Value) Then
                Return False
            Else
                Return True
            End If
        End With
    End Function
    Function can_delete(row As Integer, Optional ByVal ignore_in As Boolean = False) As Boolean
        With frm_main.dgv_production
            Dim num_in, item_id, kg_in, kg_out, km_out, gauge, width_, barcode, rts, reject, comments, user_ As String
            num_in = .Rows(row).Cells("dgc_production_num_in").Value
            item_id = .Rows(row).Cells("dgc_production_item_id").Value
            kg_in = .Rows(row).Cells("dgc_production_kg_in").Value
            kg_out = .Rows(row).Cells("dgc_production_kg_out").Value
            km_out = .Rows(row).Cells("dgc_production_km_out").Value
            gauge = .Rows(row).Cells("dgc_production_gauge").Value
            width_ = .Rows(row).Cells("dgc_production_width_").Value
            barcode = .Rows(row).Cells("dgc_production_barcode").Value
            rts = .Rows(row).Cells("dgc_production_rts").Value
            reject = .Rows(row).Cells("dgc_production_reject").Value
            comments = .Rows(row).Cells("dgc_production_comments").Value
            user_ = .Rows(row).Cells("dgc_production_user_").Value

            If num_in = "-" And item_id = "-" And kg_in = "-" And IsNull(kg_out) And IsNull(km_out) And _
                IsNull(gauge) And IsNull(width_) And IsNull(barcode) And IsNull(rts) And IsNull(reject) Then
                If IsNull(comments) Then
                    If IsNull(user_) Then
                        Return True
                    Else
                        If user_ = s_user Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                Else
                    If IsNull(user_) Then
                        Return True
                    Else
                        If user_ = s_user Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                End If
            ElseIf ignore_in And _
                    IsNull(kg_out) And IsNull(km_out) And IsNull(gauge) And IsNull(width_) And _
                    IsNull(barcode) And IsNull(kg_out) And IsNull(km_out) And IsNull(rts) And IsNull(reject) Then

                If IsNull(comments) Then
                    If IsNull(user_) Then
                        Return True
                    Else
                        If user_ = s_user Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                Else
                    If IsNull(user_) Then
                        Return True
                    Else
                        If user_ = s_user Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                End If
            Else
                Return False
            End If

        End With

    End Function
    Function get_item_info(ByVal row As Integer, ByVal reset As Boolean) As item_info
        Dim str_rid() As String = Nothing
        Dim str_rn() As String = Nothing
        Dim str_kg_in() As String = Nothing
        Dim str_prod_num() As String = Nothing
        Dim str_mat_num() As String = Nothing
        Dim str() As String = Nothing
        Dim r As Integer = -1
        If s_loading Then
            Return s_roll_info
        End If
        If reset Then
            s_roll_index = -1
            s_roll_info = Nothing
        End If
        If s_extruder Then Return Nothing
        With frm_main.dgv_production
            If .Rows(row).Cells("dgc_production_item_id").Value.ToString = "-" And .Rows(row).Cells("dgc_production_num_in").Value.ToString = "-" Then
                For r = .CurrentRow.Index To 0 Step -1
                    If Not .Rows(r).Cells("dgc_production_item_id").Value.ToString = "-" Then
                        If s_roll_index = r Then
                            Return s_roll_info
                        ElseIf s_roll_index = -1 Then
                            s_roll_index = r
                        End If
                        If .Rows(r).Cells("dgc_production_num_in").Value.ToString.Contains("&") Then
                            GoTo multi
                            Exit For
                        Else
                            GoTo one
                        End If

                    End If
                Next
            Else
                If s_roll_index = row Then
                    Return s_roll_info
                ElseIf s_roll_index = -1 Then
                    s_roll_index = row
                End If
                s_roll_index = row
                If .Rows(row).Cells("dgc_production_num_in").Value.ToString.Contains("&") Then
                    r = row
                    GoTo multi
                Else
                    r = row
                    GoTo one
                End If
            End If


        End With
one:

        With frm_main.dgv_production
            s_roll_info.row_start = r
            s_roll_info.row_end = r
            s_roll_info.row_last_item_row = r
            s_roll_info.used_kg = 0
            s_roll_info.used_km = 0
            s_roll_info.scrap = 0
            s_roll_info.rts = 0
            s_roll_info.reject = 0
            s_roll_info.total_rolls = 1
            s_roll_info.rts_count = 0
            s_roll_info.reject_count = 0
            s_roll_info.km_in = 0
            If Not .Rows(r).Cells("dgc_production_item_id").Value.ToString = "-" Then s_roll_info.item_id = .Rows(r).Cells("dgc_production_item_id").Value
            If IsNumeric(.Rows(r).Cells("dgc_production_num_in").Value) Then s_roll_info.num_in = .Rows(r).Cells("dgc_production_num_in").Value
            If Not IsNull(.Rows(r).Cells("dgc_production_prod_num_orig").Value) Then s_roll_info.prod_num = .Rows(r).Cells("dgc_production_prod_num_orig").Value
            If Not IsNull(.Rows(r).Cells("dgc_production_mat_num").Value) Then s_roll_info.mat_num = .Rows(r).Cells("dgc_production_mat_num").Value
            If Not IsNull(.Rows(r).Cells("dgc_production_km_in").Value) Then s_roll_info.km_in = .Rows(r).Cells("dgc_production_km_in").Value
            If Not IsNull(.Rows(r).Cells("dgc_production_rts_info").Value) Then s_roll_info.rts_count = 1
            If Not IsNull(.Rows(r).Cells("dgc_production_reject_info").Value) Then s_roll_info.reject_count = 1
            s_roll_info.kg_in = .Rows(s_roll_info.row_start).Cells("dgc_production_kg_in_orig").Value

            For i As Integer = s_roll_info.row_start To .RowCount - 1
                If .Rows(i).Visible Then
                    If .Rows(i).Cells(0).Value.ToString = "-" Or i = s_roll_info.row_start Then
                        If IsNumeric(.Rows(i).Cells("dgc_production_kg_out").Value) Then
                            s_roll_info.used_kg = s_roll_info.used_kg + .Rows(i).Cells("dgc_production_kg_out").Value
                        End If
                        If IsNumeric(.Rows(i).Cells("dgc_production_km_out").Value) Then
                            s_roll_info.used_km = s_roll_info.used_km + .Rows(i).Cells("dgc_production_km_out").Value
                        End If
                        If IsNumeric(.Rows(i).Cells("dgc_production_rts").Value) Then
                            s_roll_info.rts = s_roll_info.rts + .Rows(i).Cells("dgc_production_rts").Value
                        End If
                        If IsNumeric(.Rows(i).Cells("dgc_production_scrap").Value) Then
                            s_roll_info.scrap = s_roll_info.scrap + .Rows(i).Cells("dgc_production_scrap").Value
                        End If

                        If IsNumeric(.Rows(i).Cells("dgc_production_reject").Value) Then
                            s_roll_info.reject = s_roll_info.reject + .Rows(i).Cells("dgc_production_reject").Value
                        End If
                    Else
                        s_roll_info.row_end = i - 1
                        Exit For
                    End If

                    If IsNumeric(.Rows(i).Cells("dgc_production_num_out").Value) Then
                        s_roll_info.row_last_item_row = i
                    End If

                    If i = .RowCount - 1 Then
                        s_roll_info.row_end = i
                    End If
                End If
            Next
            If s_department = "bagplant" Then s_roll_info.used_kg = get_conversion_factor(True, frm_main.tb_mat_num_out.Text, False) * s_roll_info.used_km / 1000

            If s_roll_info.used_kg = 0 Then
                For r = s_roll_info.row_start To s_roll_info.row_end
                    If r = s_roll_info.row_start Then
                        .Rows(r).Cells("dgc_production_kg_in").Value = s_roll_info.kg_in
                    Else
                        .Rows(r).Cells("dgc_production_kg_in").Value = "-"
                    End If
                Next
            End If

        End With
        If s_department = "bagplant" Then
            s_roll_info.remaining = s_roll_info.kg_in - s_roll_info.used_kg - s_roll_info.rts - s_roll_info.reject - s_roll_info.scrap
        Else
            s_roll_info.remaining = s_roll_info.kg_in - s_roll_info.used_kg - s_roll_info.rts - s_roll_info.reject - s_roll_info.scrap
        End If
        Return s_roll_info
multi:
        With frm_main.dgv_production
            s_roll_info.row_start = r
            s_roll_info.row_end = r
            s_roll_info.row_last_item_row = r
            s_roll_info.used_kg = 0
            s_roll_info.used_km = 0
            s_roll_info.scrap = 0
            s_roll_info.rts = 0
            s_roll_info.reject = 0
            s_roll_info.rts_count = 0
            s_roll_info.reject_count = 0
            s_roll_info.kg_in = 0

            str_rid = Split(.Rows(r).Cells("dgc_production_item_id").Value, " & ")
            str_rn = Split(.Rows(r).Cells("dgc_production_num_in").Value, " & ")
            str_kg_in = Split(.Rows(r).Cells("dgc_production_kg_in_orig").Value, ",")
            str_prod_num = Split(.Rows(r).Cells("dgc_production_prod_num_orig").Value, ",")
            str_mat_num = Split(.Rows(r).Cells("dgc_production_mat_num").Value, ",")
            s_roll_info.total_rolls = str_rn.Length

            For i As Integer = 0 To UBound(str_kg_in)
                s_roll_info.kg_in = s_roll_info.kg_in + CDec(str_kg_in(i))
            Next

            For i As Integer = s_roll_info.row_start To .RowCount - 1
                If .Rows(i).Visible Then
                    If .Rows(i).Cells(0).Value.ToString = "-" Or i = s_roll_info.row_start Then
                        If IsNumeric(.Rows(i).Cells("dgc_production_kg_out").Value) Then
                            s_roll_info.used_kg = s_roll_info.used_kg + .Rows(i).Cells("dgc_production_kg_out").Value
                        End If
                        If IsNumeric(.Rows(i).Cells("dgc_production_km_out").Value) Then
                            s_roll_info.used_km = s_roll_info.used_km + .Rows(i).Cells("dgc_production_km_out").Value
                        End If
                        If IsNumeric(.Rows(i).Cells("dgc_production_rts").Value) Then
                            s_roll_info.rts = s_roll_info.rts + .Rows(i).Cells("dgc_production_rts").Value
                        End If
                        If IsNumeric(.Rows(i).Cells("dgc_production_num_out").Value) Then
                            s_roll_info.row_last_item_row = i
                        End If
                        If IsNumeric(.Rows(i).Cells("dgc_production_scrap").Value) Then
                            s_roll_info.scrap = s_roll_info.scrap + .Rows(i).Cells("dgc_production_scrap").Value
                        End If

                        If IsNumeric(.Rows(i).Cells("dgc_production_reject").Value) Then
                            s_roll_info.reject = s_roll_info.reject + .Rows(i).Cells("dgc_production_reject").Value
                        End If
                    Else
                        s_roll_info.row_end = i - 1
                        Exit For
                    End If
                    If i = .RowCount - 1 Then
                        s_roll_info.row_end = i
                    End If
                End If
            Next

        End With

        With frm_item_id
            .StartPosition = FormStartPosition.CenterParent
            .dgv_itemlist.Rows.Clear()
            .b_ok.Enabled = False
            Dim add_count As Boolean = True
            Dim c_name As String
            For i As Integer = 0 To str_rn.Length - 1
                Dim add As Boolean = True
                c_name = frm_main.dgv_production.Columns(frm_main.dgv_production.CurrentCell.ColumnIndex).Name
                Select Case c_name
                    Case "dgc_production_rts", "dgc_production_reject"
                        For r = s_roll_info.row_start To s_roll_info.row_end
                            If Not IsNull(frm_main.dgv_production.Rows(r).Cells(c_name & "_info").Value) Then
                                str = Split(frm_main.dgv_production.Rows(r).Cells(c_name & "_info").Value, ",")
                                If str.Length = 4 Then
                                    If frm_main.dgv_production.CurrentRow.Index = r Then
                                        s_roll_info.prod_num = str(0)
                                        s_roll_info.item_id = str(1)
                                        s_roll_info.num_in = str(2)
                                        s_roll_info.mat_num = str(3)
                                        For n As Integer = 0 To UBound(str_rid)
                                            If str_rid(n) = str(1) Then
                                                s_roll_info.kg_in = str_kg_in(n)
                                                Exit For
                                            End If
                                        Next
                                        add = False
                                    Else
                                        If str(1) = str_rid(i) Then
                                            add = False
                                        End If
                                    End If
                                    If add_count And Not r = frm_main.dgv_production.CurrentRow.Index Then

                                        If c_name = "dgc_production_rts" Then
                                            s_roll_info.rts_count = s_roll_info.rts_count + 1
                                        Else
                                            s_roll_info.reject_count = s_roll_info.reject_count + 1
                                        End If
                                    End If
                                End If
                            End If

                        Next
                        add_count = False

                    Case Else
                        i = i
                End Select
                If add Then
                    .dgv_itemlist.Rows.Add(str_rn(i), str_rid(i), str_kg_in(i), str_prod_num(i), str_mat_num(i))
                End If
            Next
            .dgv_itemlist.ClearSelection()
            c_name = frm_main.dgv_production.Columns(frm_main.dgv_production.CurrentCell.ColumnIndex).Name
            Select Case c_name
                Case "dgc_production_rts", "dgc_production_reject"
                    If .dgv_itemlist.RowCount > 1 And Not IsNull(frm_main.dgv_production.CurrentCell.Value) Then
                        .ShowDialog()
                    ElseIf .dgv_itemlist.RowCount = 1 Then
                        s_roll_info.num_in = .dgv_itemlist.Rows(0).Cells("dgc_itemlist_num_in").Value
                        s_roll_info.item_id = .dgv_itemlist.Rows(0).Cells("dgc_itemlist_item_id").Value
                        s_roll_info.kg_in = .dgv_itemlist.Rows(0).Cells("dgc_itemlist_kg_in").Value
                        s_roll_info.prod_num = .dgv_itemlist.Rows(0).Cells("dgc_itemlist_prod_num").Value
                        s_roll_info.mat_num = .dgv_itemlist.Rows(0).Cells("dgc_itemlist_mat_num").Value
                    End If
            End Select
            ' If total_weight Then
            s_roll_info.remaining = s_roll_info.kg_in - s_roll_info.used_kg - s_roll_info.rts - s_roll_info.reject - s_roll_info.scrap
            Return s_roll_info
            'End If

        End With

    End Function


    Sub report_kg_out(sender As Object)
        Dim dgv As DataGridView = CType(sender, DataGridView)
        With dgv
            s_loading = False
            Dim i_info As item_info = get_item_info(.CurrentRow.Index, True)
            If s_department = "bagplant" Then
                i_info.remaining = i_info.remaining + (get_conversion_factor(True, frm_main.tb_mat_num_out.Text, False) * .CurrentRow.Cells("dgc_production_km_out").Value / 1000)
            Else
                i_info.remaining = i_info.remaining + .CurrentRow.Cells("dgc_production_kg_out").Value
            End If
            s_loading = True
            If .CurrentRow.Index + 1 <= .RowCount And IsNumeric(.CurrentCell.Value) Then
                For r As Integer = 0 To i_info.row_end
                    If Not r = .CurrentRow.Index Then
                        If can_delete(r) And r < i_info.row_start Then
                            .Rows(r).Visible = False
                            s_report_export_all = True
                        ElseIf i_info.row_start < .CurrentRow.Index And IsNumeric(.Rows(r).Cells("dgc_production_kg_out").Value) And r >= i_info.row_start Or _
                            s_department = "bagplant" And i_info.row_start < .CurrentRow.Index And IsNumeric(.Rows(r).Cells("dgc_production_km_out").Value) And _
                            r >= i_info.row_start Then

                            If i_info.remaining > 0 Then
                                If s_department = "bagplant" Then
                                    .Rows(r).Cells("dgc_production_kg_in").Value = get_conversion_factor(True, frm_main.tb_mat_num_out.Text, False) * .Rows(r).Cells("dgc_production_km_out").Value / 1000
                                Else
                                    .Rows(r).Cells("dgc_production_kg_in").Value = .Rows(r).Cells("dgc_production_kg_out").Value
                                End If
                            End If

                        End If
                    ElseIf Not r = i_info.row_start Then
                        If i_info.remaining > 0 Then
                            .CurrentRow.Cells("dgc_production_kg_in").Value = FormatNumber(i_info.remaining, 1, , , TriState.False)
                        End If
                        Exit For
                    End If
                Next

            End If
        End With
    End Sub

    Sub report_km_out(sender As Object, e As DataGridViewCellEventArgs)
        Dim dgv As DataGridView = CType(sender, DataGridView)
        With dgv
            Dim dp As Integer = Strings.Right(.Columns(.CurrentCell.ColumnIndex).DefaultCellStyle.Format, 1)
            If IsNumeric(.CurrentCell.Value) Then .CurrentCell.Value = FormatNumber(.CurrentCell.Value, dp, , , TriState.False)
            If s_department = "bagplant" And s_symple_access Then
                report_kg_out(sender)
                If Not .CurrentCell.Value = frm_main.nud_items_out.Value And Not IsNull(.CurrentCell.Value) Then
                    Dim resp As MsgBoxResult
                    resp = MsgBox("Reprint the Carton label?" & vbCrLf & "Carton: " & .CurrentRow.Cells("dgc_production_num_out").Value & vbCrLf & "Quantity: " & .CurrentCell.Value, MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Reprint Label?")
                    If resp = MsgBoxResult.Yes Then
                        Dim str() As String = Split(frm_main.dgv_production.Columns("dgc_production_km_out").HeaderText, " ")
                        'Dim meters As Decimal = .Rows(i).Cells(2).Value
                        If str(0) = "KM" Then
                            ' meters = meters * 1000
                        End If
                        'Label!
                        acsis.print_ticket(frm_main.tb_prod_num.Text, .CurrentRow.Cells("dgc_production_kg_out").Value, _
                                           .CurrentRow.Cells("dgc_production_km_out").Value, s_selected_job(0)("req_uom_1"), "Carton Quantity", _
                                           .CurrentRow.Cells("dgc_production_km_out").Value, .CurrentRow.Cells("dgc_production_num_out").Value, _
                                           .CurrentRow.Cells("dgc_production_num_out").Value, .CurrentRow.Cells("dgc_production_num_out").Value)
                    End If
                End If
            End If
        End With

    End Sub

    Sub carton_out()
        Dim found As Boolean = False
        Dim row As DataGridViewRow = Nothing
        Dim rn As Integer = 0
        With frm_main.dgv_production
            For r = 0 To .RowCount - 1
                Dim n_out As Integer = 0
                If IsNumeric(.Rows(r).Cells("dgc_production_num_out").Value) Then
                    n_out = .Rows(r).Cells("dgc_production_num_out").Value
                    If n_out > rn Then rn = n_out
                End If
            Next
            rn = rn + 1
            For i As Integer = 0 To .RowCount - 1
                If IsNull(.Rows(i).Cells("dgc_production_km_out").Value) Then
                    row = .Rows(i)
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                s_loading = True
                .Rows.Add()
                row = .Rows(.RowCount - 1)
                row.Cells("dgc_production_num_in").Value = "-"
                row.Cells("dgc_production_item_id").Value = "-"
                row.Cells("dgc_production_kg_in").Value = "-"
                row.Cells("dgc_production_prod_num").Value = s_selected_job(0)("prod_num") '.Rows(.RowCount - 1).Cells("dgc_production_prod_num").Value
                .CurrentCell = row.Cells("dgc_production_km_out")
                .Focus()
                s_loading = False
            End If

            row.Cells("dgc_production_num_out").Value = rn
            row.Cells("dgc_production_km_out").Value = FormatNumber(frm_main.nud_items_out.Value, 0, , , TriState.False)
            row.Cells("dgc_production_date_").Value = frm_main.tb_start_date.Text
            row.Cells("dgc_production_shift").Value = s_shift
            row.Cells("dgc_production_user_").Value = s_user
            row.Cells("dgc_production_machine").Value = s_machine
        End With
        report_kg_out(frm_main.dgv_production)
        sql.update("db_schedule", "status", 2, "prod_num", "=", frm_main.tb_prod_num.Text)
        s_selected_job(0)("status") = 2

        frm_main.export_report()
        calc_totals()
    End Sub
    Sub end_prep()
        With frm_main
            If Not s_current_job = Nothing Then
                If s_selected_job(0)("status") < 2 Then
                    s_selected_job(0)("status") = 3
                    sql.update("db_schedule", "status", 3, "prod_num", "=", s_selected_job(0)("prod_num") & " AND plant=" & s_plant)
                End If
                s_selected_job = dt_schedules.Select("prod_num='" & s_current_job & "'")
                .display_job()
                .import_report()
            End If
            If s_preparing Then
                s_preparing = False
                s_user = Nothing
                s_pass = Nothing
                s_user_name = Nothing
                w_sessions(0).process.Kill()
            End If
            .b_report_part_roll.Text = "Part Roll"
            .l_prep_warning.Visible = False
        End With

    End Sub

    Sub items_out(kg As Decimal, meters As Decimal, ByVal export As Boolean)
        Dim rn As Integer, entered As Boolean = False
        'Dim km As Decimal = get_material_info(False).length / 1000


        With frm_main
            Dim r As Integer = 0
            With .dgv_production
                For r = 0 To .RowCount - 1
                    Dim n_out As Integer = 0
                    If IsNumeric(.Rows(r).Cells("dgc_production_num_out").Value) Then
                        n_out = .Rows(r).Cells("dgc_production_num_out").Value
                        If n_out > rn Then rn = n_out
                    End If
                Next
                rn = rn + 1
                For r = 0 To .RowCount - 1
                    If IsNull(.Rows(r).Cells("dgc_production_num_out").Value) And .Rows(r).Visible = True Then
                        entered = True
                        Exit For
                    End If
                Next
                If Not entered Then
                    Dim resp As MsgBoxResult
                    If s_extruder Then
                        resp = MsgBoxResult.Yes
                    Else
                        resp = MsgBox("Enter a new row?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "No spare lines!")
                    End If
                    If resp = MsgBoxResult.Yes Then
                        r = .RowCount
                        insert_row(r - 1, False)
                        entered = True
                    End If
                End If
                If entered Then
                    Dim dp As Integer = Strings.Right(.Columns("dgc_production_km_out").DefaultCellStyle.Format, 1)

                    .CurrentCell = .Rows(r).Cells("dgc_production_kg_out")
                    .Rows(r).Cells("dgc_production_num_out").Value = rn
                    .Rows(r).Cells("dgc_production_kg_out").Value = kg
                    If meters > 0 Then
                        If s_department = "printing" Then
                            .Rows(r).Cells("dgc_production_km_out").Value = FormatNumber(meters / 1000, dp, , , TriState.False)
                        Else
                            .Rows(r).Cells("dgc_production_km_out").Value = FormatNumber(meters, dp, , , TriState.False)
                        End If
                    End If

                    If IsNull(.CurrentRow.Cells("dgc_production_user_").Value) Then
                        .CurrentRow.Cells("dgc_production_date_").Value = frm_main.tb_start_date.Text
                        .CurrentRow.Cells("dgc_production_shift").Value = s_shift
                        .CurrentRow.Cells("dgc_production_user_").Value = s_user
                        .CurrentRow.Cells("dgc_production_machine").Value = frm_main.cbo_machine.SelectedItem("machine")
                    End If

                    report_kg_out(frm_main.dgv_production)
                    sql.update("db_schedule", "status", 2, "prod_num", "=", frm_main.tb_prod_num.Text)
                    s_selected_job(0)("status") = 2
                    If frm_main.cb_report_auto_print.Checked Then
                        ticket_print(meters)
                    End If
                End If
            End With
            If export Then
                .export_report()
                calc_totals()
            End If
        End With

    End Sub

    Sub export_report_rows(ByVal rows As String)
        Dim r() As String = Split(rows, ",")
        For i As Integer = 0 To UBound(r)
            frm_main.export_report(r(i))
        Next
    End Sub
    Sub ticket_print(ByVal meters)
        With frm_main.dgv_production
            Dim m_info As materialInfo = get_material_info()
            Dim section As String = "Contents"
            Dim quant_out As Decimal = 0
            Select Case s_selected_job(0)("req_uom_1")
                Case "KM"
                    quant_out = .CurrentRow.Cells("dgc_production_km_out").Value
                Case "KG"
                    quant_out = .CurrentRow.Cells("dgc_production_kg_out").Value
                Case Else
                    quant_out = quant_out
            End Select
            Select Case m_info.material_type
                Case material_type.barrier
                    acsis.print_ticket(frm_main.tb_prod_num.Text, quant_out, meters, s_selected_job(0)("req_uom_1"), _
                                       Nothing, Nothing, .CurrentRow.Cells("dgc_production_num_out").Value, 0, 0)
                    Exit Sub
                Case material_type.film, material_type.laminate
                    If Strings.Left(frm_main.tb_mat_num_out.Text, 1) = "2" Then
                        If s_department = "printing" Then
                            meters = frm_main.tb_label_desc.Text & " (" & CInt(meters) & "m)"
                        Else
                            meters = CInt(meters) & "m"
                        End If
                    Else
                        If meters = m_info.length Then
                            acsis.print_ticket(frm_main.tb_prod_num.Text, quant_out, meters, s_selected_job(0)("req_uom_1"), _
                                               Nothing, Nothing, .CurrentRow.Cells("dgc_production_num_out").Value, 0, 0)
                            Exit Sub
                        Else
                            section = "Length"
                        End If

                    End If
                Case Else
                    meters = meters
            End Select
            If IsNumeric(meters) Then meters = CInt(meters)
            acsis.print_ticket(frm_main.tb_prod_num.Text, quant_out, meters, s_selected_job(0)("req_uom_1"), _
                                       section, meters, .CurrentRow.Cells("dgc_production_num_out").Value, 0, 0)

        End With
    End Sub
    Sub clear_inks()
        With frm_main
            .tb_station_1.Clear()
            .tb_station_2.Clear()
            .tb_station_3.Clear()
            .tb_station_4.Clear()
            .tb_station_5.Clear()
            .tb_station_6.Clear()
            .tb_station_7.Clear()
            .tb_station_8.Clear()
            .tb_station_9.Clear()
            .tb_station_10.Clear()
        End With
    End Sub
    Sub calc_totals()
        Dim prod_num As Long = frm_main.tb_prod_num.Text, uom, total, unused, report As String
        uom = s_selected_job(0)("req_uom_1")
        unused = 0
        report = 0
        'If IsNull(dt_produced) Then
        If frm_main.tc_main.SelectedTab Is frm_main.tp_scheduling Then
            dt_produced = sql.get_table("db_produced", "prod_num", "=", "'" & frm_main.tb_prod_num.Text & "'")
        Else
            dt_produced = dgv_to_dt(frm_main.dgv_production)
        End If
        With frm_main.dgv_job_requirements
            For r As Integer = 0 To .RowCount - 1
                Dim dp As Integer = 0
                total = 0
                Dim str() As String = Split(frm_main.dgv_production.Columns("dgc_production_km_out").HeaderText, " ")

                Select Case .Rows(r).Cells(0).Value

                    Case "ROL"
                        ' total = CInt(sql.count("db_produced", "prod_num", "=", prod_num & " AND kg_out<>'0'"))
                        If dt_produced Is Nothing Then
                            total = 0
                        Else
                            total = CInt(dt_produced.Compute("count(kg_out)", "kg_out<>'0'"))
                        End If
                    Case "KG"
                        'total = FormatNumber(sql.sum("db_produced", "kg_out", "prod_num", "=", prod_num), 1, , , TriState.False)
                        If dt_produced Is Nothing Then
                            total = 0
                        ElseIf s_department = "bagplant" Then
                            If dt_produced.Compute("count(km_out)", "km_out<>'0'") > 0 Then
                                total = FormatNumber(get_conversion_factor(True, frm_main.tb_mat_num_out.Text, False) * dt_produced.Compute("sum(km_out)", "") / 1000, 1, , , TriState.False)
                            End If

                        Else
                            If dt_produced.Compute("count(kg_out)", "kg_out<>'0'") > 0 Then
                                total = FormatNumber(dt_produced.Compute("sum(kg_out)", ""), 1, , , TriState.False)
                            End If
                        End If
                        dp = 1
                    Case "KM"
                        If dt_produced Is Nothing Then
                            total = 0
                        Else
                            If str(0) = "KM" Then
                                If dt_produced.Compute("count(km_out)", "kg_out<>'0'") > 0 Then
                                    total = FormatNumber(dt_produced.Compute("sum(km_out)", ""), 2, , , TriState.False)
                                End If
                            Else
                                total = FormatNumber(dt_produced.Compute("sum(km_out)", "") / 1000, 2, , , TriState.False)
                            End If
                        End If
                        dp = 2
                    Case "M"
                        If dt_produced Is Nothing Then
                            total = 0
                        Else
                            If dt_produced.Compute("count(km_out)", "kg_out<>'0'") > 0 Then
                                If str(0) = "KM" Then
                                    total = FormatNumber(dt_produced.Compute("sum(km_out)", "") * 1000, 1, , , False)
                                Else
                                    total = FormatNumber(dt_produced.Compute("sum(km_out)", ""), 1, , , False)
                                End If
                            End If
                        End If
                        dp = 1
                    Case "MU"
                        If dt_produced Is Nothing Then
                            total = 0
                        Else
                            If dt_produced.Compute("count(km_out)", "km_out<>'0'") > 0 Then
                                total = FormatNumber(dt_produced.Compute("sum(km_out)", "") / 1000, 1, , , TriState.False)
                            End If
                        End If
                        dp = 1
                    Case "IMP"
                        If dt_produced Is Nothing Then
                            total = 0
                        Else

                            total = (dt_produced.Compute("sum(km_out)", "") * 1000) / (frm_main.tb_cyl_size.Text / 1000)
                        End If
                    Case Else
                        r = r
                        If dt_produced Is Nothing Then
                            unused = 0
                        Else
                            unused = sql.sum("db_produced", "kg_in", "prod_num", "=", prod_num & " AND kg_out IS NULL AND km_out IS NULL AND rts IS NULL AND reject IS NULL")
                        End If
                End Select

                '  End If
                .Rows(r).Cells("dgc_job_requirements_produced").Value = total
                .Rows(r).Cells("dgc_job_requirements_remaining").Value = _
                    FormatNumber(.Rows(r).Cells("dgc_job_requirements_requirement").Value - .Rows(r).Cells("dgc_job_requirements_produced").Value, dp, , , TriState.False)

                .Rows(r).Cells("dgc_job_requirements_percent").Value = _
                    Replace(FormatPercent(.Rows(r).Cells("dgc_job_requirements_produced").Value / .Rows(r).Cells("dgc_job_requirements_requirement").Value, 1), "%", Nothing)
                If .Rows(r).Cells("dgc_job_requirements_percent").Value > 200 Then
                    .Rows(r).Cells("dgc_job_requirements_produced").Value = total / 1000
                    .Rows(r).Cells("dgc_job_requirements_remaining").Value = _
                        FormatNumber(.Rows(r).Cells("dgc_job_requirements_requirement").Value - .Rows(r).Cells("dgc_job_requirements_produced").Value, dp, , , TriState.False)
                    .Rows(r).Cells("dgc_job_requirements_percent").Value = _
                    Replace(FormatPercent(.Rows(r).Cells("dgc_job_requirements_produced").Value / .Rows(r).Cells("dgc_job_requirements_requirement").Value, 1), "%", Nothing)
                End If
                r = r
            Next

            If dt_produced Is Nothing Then
                report = 0
            Else
                report = sql.sum("db_produced", "kg_in", "prod_num", "=", prod_num)
            End If
            If Replace(.Rows(0).Cells("dgc_job_requirements_percent").Value, "%", Nothing) > 100 Then
                frm_main.tspb_progress.Value = 100
            Else
                frm_main.tspb_progress.Value = Replace(.Rows(0).Cells("dgc_job_requirements_percent").Value, "%", Nothing)
            End If
            If Not IsNull(dt_produced) Then
                unused = dt_produced.Compute("sum(kg_in)", "kg_out IS NULL AND km_out IS NULL AND rts IS NULL AND reject IS NULL AND scrap IS NULL").ToString
            End If
            If Not IsNull(unused) Then
                If Not s_extruder Then frm_main.l_production_unused.Visible = True
                'If frm_main.l_production_remaining.Visible Then
                '    Dim r As String = Replace(Replace(frm_main.l_production_remaining.Text, "KG", Nothing), "Remaining:", Nothing)
                '    If IsNumeric(r) Then unused = unused + CDec(r)
                'End If
                frm_main.l_production_unused.Text = FormatNumber(unused, 1, , , TriState.False) & " Unused KG"

                If .Rows(0).Cells("dgc_job_requirements_uom").Value = "KG" Then
                    If .Rows(0).Cells("dgc_job_requirements_remaining").Value > CDec(unused) Then
                        frm_main.l_production_unused.ForeColor = Color.Red
                    Else
                        frm_main.l_production_unused.ForeColor = Color.Green
                    End If
                    frm_main.l_production_unused.Text = frm_main.l_production_unused.Text & vbCr & _
                        FormatNumber(.Rows(0).Cells("dgc_job_requirements_remaining").Value, 1, , , TriState.False) & " remaining on the job"

                ElseIf .RowCount > 1 Then
                    If .Rows(1).Cells("dgc_job_requirements_remaining").Value > CDec(unused) Then
                        frm_main.l_production_unused.ForeColor = Color.Red
                    Else
                        frm_main.l_production_unused.ForeColor = Color.Green
                    End If
                    frm_main.l_production_unused.Text = frm_main.l_production_unused.Text & vbCr & _
                        FormatNumber(.Rows(1).Cells("dgc_job_requirements_remaining").Value, 1, , , TriState.False) & " remaining on the job"

                End If
            Else
                frm_main.l_production_unused.Visible = False
            End If
            If Not report = 0 Then
                If Not s_extruder Then
                    frm_main.l_production_kg_in.Visible = True
                    frm_main.l_production_kg_in.Text = """KG IN"" on Report: " & FormatNumber(report, 1, , , TriState.False)

                End If
                If .Rows(0).Cells("dgc_job_requirements_uom").Value = "KG" Then
                    If .Rows(0).Cells("dgc_job_requirements_requirement").Value > CDec(report) Then
                        frm_main.l_production_kg_in.ForeColor = Color.Red
                    Else
                        frm_main.l_production_kg_in.ForeColor = Color.Green
                    End If
                ElseIf .RowCount > 1 Then
                    If .Rows(1).Cells("dgc_job_requirements_requirement").Value > CDec(report) Then
                        frm_main.l_production_kg_in.ForeColor = Color.Red
                    Else
                        frm_main.l_production_kg_in.ForeColor = Color.Green
                    End If
                End If
            Else
                If Not s_extruder Then frm_main.l_production_kg_in.Visible = False
            End If
        End With

    End Sub
    Sub insert_row(ByVal r As Integer, ByVal export As Boolean)
        With frm_main.dgv_production
            s_loading = True
            r = r + 1
            If r = .RowCount Then
                .Rows.Add()
            Else
                .Rows.Insert(r, 1)
            End If

            .Rows(r).Cells(0).Value = "-"
            .Rows(r).Cells(1).Value = "-"
            .Rows(r).Cells(2).Value = "-"
            .Rows(r).Cells("dgc_production_prod_num").Value = s_selected_job(0)("prod_num")
            If Not r = 0 Then
                .Rows(r).Cells("dgc_production_assistant").Value = .Rows(r - 1).Cells("dgc_production_assistant").Value
            End If
            s_loading = False
        End With
        If export Then frm_main.export_report(-1)
    End Sub
    Sub part_roll()

        Dim entered As Boolean = False

        With frm_main.dgv_production
            .ClearSelection()
            For r As Integer = 0 To .RowCount - 1
                If IsNull(.Rows(r).Cells(3).Value) Then
                    s_loading = True
                    .Rows.Insert(r, 1)
                    .Rows(r).Cells(0).Value = "-"
                    .Rows(r).Cells(1).Value = "-"
                    .Rows(r).Cells(2).Value = "-"
                    .Rows(r).Cells("dgc_production_prod_num").Value = .Rows(r - 1).Cells("dgc_production_prod_num").Value
                    .Rows(r).Cells("dgc_production_assistant").Value = .Rows(r - 1).Cells("dgc_production_assistant").Value
                    .Rows(r).Cells(4).Selected = True
                    .CurrentCell = .Rows(r).Cells("dgc_production_kg_out")
                    .Focus()
                    s_loading = False
                    frm_main.export_report()
                    entered = True
                    Exit For
                End If

            Next
            If Not entered Then
                s_loading = True
                .Rows.Add()
                .Rows(.RowCount - 1).Cells("dgc_production_num_in").Value = "-"
                .Rows(.RowCount - 1).Cells("dgc_production_item_id").Value = "-"
                .Rows(.RowCount - 1).Cells("dgc_production_kg_in").Value = "-"
                .Rows(.RowCount - 1).Cells("dgc_production_prod_num").Value = .Rows(.RowCount - 2).Cells("dgc_production_prod_num").Value
                .Rows(.RowCount - 1).Cells("dgc_production_assistant").Value = .Rows(.RowCount - 2).Cells("dgc_production_assistant").Value
                .Rows(.RowCount - 1).Cells("dgc_production_kg_out").Selected = True
                .CurrentCell = .Rows(.RowCount - 1).Cells("dgc_production_kg_out")
                .Focus()
                s_loading = False
                frm_main.export_report()

            End If
        End With
    End Sub
    Sub item_entered(ByVal uom As String, quant As Decimal)

        With frm_main.dgv_production

            If frm_main.cb_report_auto_print.Checked And Not IsNull(.CurrentRow.Cells("dgc_production_kg_out").Value) Then
                If Not .CurrentRow.Cells("dgc_production_kg_out").Value = 0 Then

                    Dim resp As MsgBoxResult
                    Dim ask As Boolean = False
                    Select Case get_material_info.material_type
                        Case material_type.barrier
                            If .Columns(.CurrentCell.ColumnIndex).Name = "dgc_production_kg_out" Then
                                ask = True
                                resp = MsgBox("Print the Roll ticket?" & vbCrLf & "Roll: " & _
                                              .CurrentRow.Cells("dgc_production_num_out").Value & vbCrLf & "Quantity: " & _
                                              FormatNumber(.CurrentRow.Cells(.CurrentCell.ColumnIndex).Value, 1) & " " & _
                                              uom, MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Print Ticket?")
                            End If
                        Case material_type.laminate, material_type.film
                            If Not IsNull(.CurrentRow.Cells("dgc_production_kg_out").Value) And Not IsNull(.CurrentRow.Cells("dgc_production_km_out").Value) Then
                                ask = True
                                If uom = "ROL" Then
                                    resp = MsgBox("Print the Roll ticket?" & vbCrLf & "Roll: " & _
                                                  .CurrentRow.Cells("dgc_production_num_out").Value & vbCrLf & "Quantity: 1 " & _
                                                  uom, MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Print Ticket?")
                                ElseIf uom = "IMP" Then
                                    Do While CInt(frm_main.tb_no_up.Text) = 0
                                        frm_main.tb_no_up.Text = InputBox("How many impressions up is this job? This is needed for the calculation to be correct.")
                                        If IsNumeric(frm_main.tb_no_up.Text) And Not frm_main.tb_no_up.Text = 0 Then
                                            sql.update("db_schedule", "no_up", 2, "prod_num", "=", frm_main.tb_no_up.Text)
                                        Else
                                            frm_main.tb_no_up.Text = 0
                                        End If
                                    Loop
                                    quant = FormatNumber((.Rows(.CurrentCell.RowIndex).Cells("dgc_production_km_out").Value * 1000) / _
                                                         ((frm_main.tb_cyl_size.Text * frm_main.tb_no_up.Text) / 1000), 2)
                                    resp = MsgBox("Print the Roll ticket?" & vbCrLf & "Roll: " & _
                                                  .CurrentRow.Cells("dgc_production_num_out").Value & vbCrLf & "Quantity: " & _
                                                  quant & " IMP", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Print Ticket?")
                                Else
                                    Dim uom1() As String = Split(.Columns("dgc_production_kg_out").HeaderText, " ")
                                    Dim uom2() As String = Split(.Columns("dgc_production_km_out").HeaderText, " ")
                                    resp = MsgBox("Print the Roll ticket?" & vbCrLf & "Roll: " & _
                                                  .CurrentRow.Cells("dgc_production_num_out").Value & vbCrLf & "Quantity: " & _
                                                  .CurrentRow.Cells("dgc_production_kg_out").Value & " " & uom1(0) & vbCrLf & "Length: " & _
                                                  .CurrentRow.Cells("dgc_production_km_out").Value & " " & uom2(0), MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Print Ticket?")
                                End If
                            End If
                        Case Else
                    End Select

                    If ask And resp = MsgBoxResult.Yes Then
                        Dim str() As String = Split(frm_main.dgv_production.Columns("dgc_production_km_out").HeaderText, " ")
                        Dim meters As Decimal = .CurrentRow.Cells("dgc_production_km_out").Value
                        If str(0) = "KM" Then
                            meters = meters * 1000
                        End If
                        ticket_print(meters)
                    End If

                End If
            End If
        End With

    End Sub
End Module

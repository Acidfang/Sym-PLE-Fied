Public Class frm_pallet

    Private Sub frm_pallet_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        populate()
    End Sub
    Sub populate()
        dt_produced = dgv_to_dt(frm_main.dgv_production)
        Dim dr() As DataRow = dt_produced.Select("pallet IS NULL AND num_out IS NOT NULL")
        GroupBox1.Enabled = False
        s_loading = True
        ListBox1.Items.Clear()
        Label2.Text = "Qty. (" & s_selected_job(0)("req_uom_1") & ")"
        Label6.Text = Label2.Text
        TextBox1.Text = 0
        TextBox2.Text = 0
        TextBox3.Text = 0
        TextBox4.Text = 0
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        If dr.Length > 0 Then

            With frm_main.dgv_production
                For r As Integer = 0 To .RowCount - 1
                    If Not IsNull(.Rows(r).Cells("dgc_production_num_out").Value) And _
                        Not IsNull(.Rows(r).Cells("dgc_production_kg_out").Value) And Not _
                        IsNull(.Rows(r).Cells("dgc_production_km_out").Value) Then

                        If Not .Rows(r).Cells("dgc_production_num_out").Value.ToString = "-" Then
                            If IsNull(.Rows(r).Cells("dgc_production_pallet").Value) Then
                                With DataGridView1
                                    .Rows.Add()
                                    .Rows(.RowCount - 1).Cells(0).Value = frm_main.dgv_production.Rows(r).Cells("dgc_production_num_out").Value
                                    Select Case s_selected_job(0)("req_uom_1")
                                        Case "KG"
                                            .Rows(.RowCount - 1).Cells(1).Value = frm_main.dgv_production.Rows(r).Cells("dgc_production_kg_out").Value
                                        Case "KM"
                                            .Rows(.RowCount - 1).Cells(1).Value = frm_main.dgv_production.Rows(r).Cells("dgc_production_km_out").Value
                                        Case "ROL"
                                            .Rows(.RowCount - 1).Cells(1).Value = frm_main.dgv_production.Rows(r).Cells("dgc_production_kg_out").Value
                                    End Select
                                    If .RowCount < 33 Then
                                        .Rows(.RowCount - 1).Cells(2).Value = True
                                    End If
                                    .Rows(.RowCount - 1).Cells(3).Value = r
                                End With

                            End If
                        End If
                    End If
                Next
                '  Dim rv As DataRowView, fieldinfo As System.Reflection.FieldInfo

                Dim dt As DataTable = sql.get_table("DP_Print_Templates", "Label_Location", "=", "'Full' AND Label_Type='Pallet' ORDER BY List_Order", "symple")
                ComboBox1.DataSource = dt
                ComboBox1.DisplayMember = "Template_Code"
                select_cbo_item(ComboBox1, s_pallet, "Template_Code")
                'For Each rv In ComboBox1.Items
                '    If rv(0) = s_pallet Then
                '        fieldinfo = rv.Row.GetType.GetField("_rowID", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
                '        ComboBox1.SelectedIndex = fieldinfo.GetValue(rv.Row) - 1
                '    End If
                'Next
            End With

        End If
        dr = dt_produced.Select("pallet IS NOT NULL")
        If Not dr Is Nothing Then
            For r As Integer = 0 To dr.Length - 1
                Dim pallet As String = dr(r)("pallet")
                If ListBox1.Items.IndexOf(pallet) = -1 Then
                    ListBox1.Items.Add(pallet)
                End If
            Next


            For r As Integer = 0 To ListBox1.Items.Count - 1
                dr = dt_produced.Select("pallet='" & ListBox1.Items(r) & "'")
                Dim pid As String = dr(0)("pallet_id")
                ListBox1.Items(r) = pid
                'If Not IsNull(r("pallet_id")) Then
                '    ListBox1.Items(r("pallet") - 1) = r("pallet_id")
                'End If
            Next

            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        End If
        If ListBox1.Items.Count = 0 Then
            GroupBox2.Enabled = False
        Else
            GroupBox2.Enabled = True
        End If

        calc()
        With DataGridView1
            If CInt(TextBox1.Text > 0) Then
                GroupBox1.Enabled = True
                .Sort(.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
                .PerformLayout()
                .Rows(.RowCount - 1).Cells(0).Selected = True

            End If
        End With


        s_loading = False
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim response As MsgBoxResult = MsgBox("Are you sure you want to delete this pallet?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Delete pallet")
        If response = MsgBoxResult.Yes Then
            If Not ListBox1.SelectedItem = "NO PALLET" Then
                show_status("Deleting pallet")

                If ListBox1.SelectedItem.ToString.Length > 4 Then
                    acsis.open_window(cls_ACSIS.eWindowTarget.ModifyPalletLabel, 0)
                    acsis.insert_string(ListBox1.SelectedItem, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("modify_pallet_id")).handle)
                    acsis.click_button(acsis.get_acsis_option("modify_pallet_get_info"), False, 0)
                    If s_popup = 3 Then
                        s_popup = 0
                    Else
                        acsis.click_button(acsis.get_acsis_option("modify_pallet_delete"), False, 0)
                    End If
                End If
            End If

            With frm_main.dgv_production

                For r As Integer = 0 To .RowCount - 1
                    If ListBox1.SelectedItem.ToString.Length > 4 Then
                        If Not IsNull(.Rows(r).Cells("dgc_production_pallet_id").Value) Then
                            If .Rows(r).Cells("dgc_production_pallet_id").Value.ToString = ListBox1.SelectedItem Then
                                .Rows(r).Cells("dgc_production_pallet").Value = Nothing
                                .Rows(r).Cells("dgc_production_pallet_id").Value = Nothing
                            End If
                        End If
                    Else
                        If Not IsNull(.Rows(r).Cells("dgc_production_pallet").Value) Then
                            If .Rows(r).Cells("dgc_production_pallet").Value.ToString = ListBox1.SelectedItem Then
                                .Rows(r).Cells("dgc_production_pallet").Value = Nothing
                            End If
                        End If
                    End If


                Next
            End With
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            If ListBox1.Items.Count = 0 Then
                GroupBox2.Enabled = False
            Else
                GroupBox2.Enabled = True
            End If
        End If
        frm_main.export_report()
        dt_produced = dgv_to_dt(frm_main.dgv_production)
        'frm_main.import_report()
        frm_status.Close()
        populate()
        SetForegroundWindow(Me.Handle)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim movement As Integer = 107
        Select Case s_department
            Case "printing"
                movement = 107
            Case "bagplant"
                movement = 105
            Case "filmsline"
                If s_extruder Then
                    movement = 113
                Else
                    movement = 111
                End If
        End Select
        acsis.open_window(cls_ACSIS.eWindowTarget.PrintPalletLabels, 0)

        show_status("Printing pallet")

        SendMessage(w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_movement")).handle, WM_SETFOCUS, 0, 0)
        acsis.insert_string(movement, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_movement")).handle)
        w_sessions(0).process.WaitForInputIdle()

        SendMessage(w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_prod_num")).handle, WM_SETFOCUS, 0, 0)
        acsis.insert_string(frm_main.tb_prod_num.Text, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_prod_num")).handle)
        w_sessions(0).process.WaitForInputIdle()

        SendMessage(w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_mat_num")).handle, WM_SETFOCUS, 0, 0)
        acsis.insert_string(frm_main.tb_mat_num_out.Text, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_mat_num")).handle)
        w_sessions(0).process.WaitForInputIdle()


        SendMessage(w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_seq")).handle, WM_SETFOCUS, 0, 0)
        w_sessions(0).process.WaitForInputIdle()

        acsis.popup_check(True, 0)
        Dim i As Integer = 0, uom As String = s_selected_job(0)("req_uom_1")

        acsis.get_controls(0, True, False, True)

        Dim pallet_number As Long = 0

        For r As Integer = 0 To dt_produced.Rows.Count - 1
            If IsNumeric(dt_produced(r)("pallet")) Then
                If dt_produced(r)("pallet") > pallet_number Then
                    pallet_number = dt_produced(r)("pallet")
                End If
            End If
        Next
        pallet_number = pallet_number + 1
        'If IsNull(dt_produced.Compute("MAX(pallet)", "")) Then
        '    pallet_number = 1
        'Else
        '    pallet_number = dt_produced.Compute("MAX(pallet)", "") + 1
        'End If
        show_status("Printing pallet: Entering Items")
        With DataGridView1
            Dim c As Integer = 1
            For r As Integer = 0 To .RowCount - 1
                If .Rows(r).Cells(2).Value = True Then
                    Dim row As Integer = .Rows(r).Cells(3).Value
                    frm_main.dgv_production.Rows(row).Cells("dgc_production_pallet").Value = pallet_number
                    SendMessage(w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_seq")).handle, WM_SETFOCUS, 0, 0)
                    w_sessions(0).process.WaitForInputIdle()
                    acsis.insert_string(.Rows(r).Cells(0).Value, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_seq")).handle) ' roll number
                    SendMessage(w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_quant")).handle, WM_SETFOCUS, 0, 0)
                    w_sessions(0).process.WaitForInputIdle()
                    If uom = "ROL" Then
                        acsis.insert_string(1, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_quant")).handle) ' 1 roll
                    ElseIf uom = "IMP" Then
                        Dim quant As Decimal = FormatNumber((frm_main.dgv_production.Rows(row).Cells("km_out").Value * 1000) / (frm_main.tb_cyl_size.Text / 1000), 2)
                        acsis.insert_string(quant, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_quant")).handle) ' 1 roll
                    Else
                        acsis.insert_string(.Rows(r).Cells(1).Value, w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_quant")).handle) ' roll weight
                    End If
                    SendMessage(w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_seq")).handle, WM_SETFOCUS, 0, 0)
                    Dim count As Integer = SendMessage(w_sessions(0).winWnd(0).lbWnd(0).handle, LB_GETCOUNT, 0, 0)
                    While Not count = c ' no of items in the list box
                        If i = acsis.get_acsis_option("pallet_screen_quant") Then
                            i = acsis.get_acsis_option("pallet_screen_seq")
                        Else
                            i = acsis.get_acsis_option("pallet_screen_quant")
                        End If
                        SendMessage(w_sessions(0).winWnd(0).tbWnd(i).handle, WM_SETFOCUS, 0, 0)
                        w_sessions(0).process.WaitForInputIdle()
                        count = SendMessage(w_sessions(0).winWnd(0).lbWnd(0).handle, LB_GETCOUNT, 0, 0)
                    End While
                    c = c + 1
                End If

            Next
            If RadioButton1.Checked Then
                acsis.select_cb_item("ROL", 0)
            Else
                acsis.select_cb_item("CAR", 0)
            End If

            acsis.select_cb_item(s_pallet, 0)
            acsis.select_cb_item(s_printer, 0)
            If uom = "ROL" Then
                acsis.set_check(acsis.get_control_handle(acsis.get_acsis_option("pallet_screen_label_data"), _
                                                         cls_ACSIS.eControlType.CheckBox, 0), True)

                acsis.select_grid_item("Pallet Quantity UOM", False)
                If Not Clipboard.GetText = "FAILED" Then

                    If CInt(TextBox1.Text) = 1 Then
                        Clipboard.SetText("ROLL")
                    Else
                        Clipboard.SetText("ROLLS")
                    End If
                    SendKeys.SendWait("^(v)")
                End If
            End If

        End With

        acsis.click_button(acsis.get_acsis_option("pallet_screen_print"), False, 0)
        w_sessions(0).process.WaitForInputIdle()
        Dim pid As String = Nothing
        show_status("Printing pallet: Waiting for pallet ID")
        While pid = Nothing
            pid = acsis.get_text(w_sessions(0).winWnd(0).tbWnd(acsis.get_acsis_option("pallet_screen_id")).handle, False)
            System.Threading.Thread.Sleep(500)
        End While

        With frm_main.dgv_production
            For r As Integer = 0 To .RowCount - 1
                If Not IsNull(.Rows(r).Cells("dgc_production_pallet").Value) Then
                    If .Rows(r).Cells("dgc_production_pallet").Value.ToString = pallet_number.ToString Then
                        .Rows(r).Cells("dgc_production_pallet_id").Value = pid
                    End If
                End If
            Next
        End With

        acsis.click_button(acsis.get_button_back(0), False, 0)
        frm_main.export_report()
        frm_status.Close()
        SetForegroundWindow(frm_main.Handle)
        Me.Close()
    End Sub

    Sub calc()
        ' If s_loading = True Then Exit Sub
        TextBox1.Text = 0
        TextBox4.Text = 0
        With DataGridView1
            For r As Integer = 0 To .RowCount - 1
                If .Rows(r).Cells(2).Value = True And Not IsNull(.Rows(r).Cells(1).Value) Then
                    TextBox1.Text = CInt(TextBox1.Text) + 1
                    TextBox4.Text = CDec(TextBox4.Text) + CDec(.Rows(r).Cells(1).Value)
                End If
            Next
        End With
        Select Case s_selected_job(0)("req_uom_1")
            Case "ROL"
                TextBox4.Text = TextBox1.Text
        End Select

    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        With DataGridView1
            If .CurrentCell.ColumnIndex = 2 Then
                If .CurrentCell.Value = True Then
                    .CurrentCell.Value = False
                Else
                    If CInt(TextBox1.Text) < 32 Then
                        .CurrentCell.Value = True
                    Else
                        MsgBox("You need to remove an item before adding this.", MsgBoxStyle.OkOnly, "32 Items max!")
                    End If
                End If
            End If
        End With
        calc()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If Not ListBox1.SelectedItem = Nothing Then
            Dim dr() As DataRow = dt_produced.Select("pallet_id='" & ListBox1.SelectedItem & "'")
            With DataGridView2
                .Rows.Clear()
                Dim w As Decimal = 0
                For i As Integer = 0 To dr.Length - 1
                    .Rows.Add()
                    Dim r As Integer = .RowCount - 1
                    .Rows(r).Cells(0).Value = CInt(dr(i)("num_out"))
                    Select Case s_selected_job(0)("req_uom_1")
                        Case "KG", "ROL"
                            .Rows(r).Cells(1).Value = dr(i)("kg_out")
                            w = w + dr(i)("kg_out")
                        Case "KM"
                            .Rows(r).Cells(1).Value = dr(i)("km_out")
                            w = w + dr(i)("km_out")
                    End Select

                    .Rows(r).Cells(1).Value = dr(i)("kg_out")
                Next
                TextBox2.Text = .RowCount
                TextBox3.Text = w
                .Sort(.Columns(0), 0)
                .ClearSelection()
                .PerformLayout()
            End With
        End If

    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim pallet_number As Integer = -1

        With DataGridView1
            Dim c As Integer = 1
            For r As Integer = 0 To .RowCount - 1
                If .Rows(r).Cells(2).Value = True Then
                    Dim row As Integer = .Rows(r).Cells(3).Value
                    frm_main.dgv_production.Rows(row).Cells("dgc_production_pallet").Value = pallet_number
                End If
            Next

        End With
        With frm_main.dgv_production
            For r As Integer = 0 To .RowCount - 1
                If Not IsNull(.Rows(r).Cells("dgc_production_pallet").Value) Then
                    If CLng(.Rows(r).Cells("dgc_production_pallet").Value) = pallet_number Then
                        .Rows(r).Cells("dgc_production_pallet_id").Value = "NO PALLET"
                    End If
                End If

            Next
        End With
        populate()
    End Sub
End Class
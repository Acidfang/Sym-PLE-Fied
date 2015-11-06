Imports System.Deployment.Application
Imports System.IO
Imports System.IO.Ports
Imports System.Text
Module mod_common
    ' Private WithEvents moRS232 As Rs232

    Function add_to_string(ByVal previous_string As String, ByVal string_to_add As String, Optional ByVal delim As Object = ",", Optional ByVal allow_double As Boolean = True) As String
        If string_to_add = Nothing Then
            Return previous_string
        Else
            If previous_string = Nothing Then
                Return string_to_add
            Else
                If allow_double Then
                    Return previous_string & delim & string_to_add
                Else
                    Dim s() As String = Split(previous_string, delim)
                    If s.ToList.IndexOf(string_to_add) > -1 Then
                        Return previous_string
                    Else
                        Return previous_string & delim & string_to_add
                    End If
                End If
            End If

        End If

    End Function

    Function get_label(ByVal mat_num As Integer, mat_desc As String, Optional pallet As Boolean = False, Optional ByVal rts As Boolean = False) As String
        Dim dr() As DataRow
        Dim s As String
        Dim m_info As materialInfo = get_material_info(mat_desc:=mat_desc)
        If pallet Then
            s = "pallet_"
        Else
            s = "label_"
        End If
        If rts Then
            If Left(mat_num, 1) = 1 Then
                dr = dt_acsis_options.Select("item='" & s & "semi_finished'")
            Else
                dr = dt_acsis_options.Select("item='" & s & "semi_finished'")
            End If
        Else
            Select Case m_info.material_type
                Case material_type.barrier
                    dr = dt_acsis_options.Select("item='" & s & "barrier'")
                Case material_type.film
                    If Left(mat_num, 1) = 1 Then
                        dr = dt_acsis_options.Select("item='" & s & "film'")
                    Else
                        dr = dt_acsis_options.Select("item='" & s & "semi_finished'")
                    End If
                Case material_type.laminate
                    If Left(mat_num, 1) = 1 Then
                        dr = dt_acsis_options.Select("item='" & s & "laminate'")
                    Else
                        dr = dt_acsis_options.Select("item='" & s & "semi_finished'")
                    End If
                Case material_type.rollbag
                    dr = dt_acsis_options.Select("item='" & s & "rollbag'")
                Case Else
                    dr = dt_acsis_options.Select("item='" & s & "carton'")
            End Select
        End If

        If dr Is Nothing Then
            Return Nothing
        ElseIf dr.Length > 0 Then
            Return dr(0)("val")
        Else
            Return Nothing
        End If

    End Function

    Function get_job_value(ByVal str As String, ByVal number As Boolean) As Object
        Dim n As String = Nothing
        Dim s As String = Nothing
        For i As Integer = 1 To str.Length
            Dim c As String = Mid(str, i, 1)
            If number Then
                Select Case True
                    Case IsNumeric(c)
                        n = add_to_string(n, c, Nothing)
                    Case c = "."
                        If Not n = Nothing Then
                            If Not n.Contains(".") Then n = add_to_string(n, c, Nothing)
                        End If
                End Select
            ElseIf Not IsNumeric(c) And Not c = "." And Not c = "," Then
                s = add_to_string(s, c, Nothing)
                If Not n = Nothing And number Then
                    Return n
                End If
            End If
        Next
        If number Then
            If n = Nothing Then
                Return 0
            Else
                Return CDec(n)
            End If
        Else
            Return s
        End If
    End Function

    Function get_string_difference(ByVal string_1 As String, string_2 As String) As String
        Dim s As String = Nothing
        If string_1.Length = string_2.Length Then
            For i As Integer = 1 To string_1.Length
                Dim c As String = Mid(string_1, i, 1)
                If Not c = Mid(string_2, i, 1) Then
                    s = add_to_string(s, c)
                End If
            Next
            Return s
        Else
            Return string_2
        End If
    End Function

    Function get_extruder_kg_hr(ByVal m_info As materialInfo) As Integer
        Select Case m_info.formulation
            Case "D321"
                If m_info.width > 1250 Then
                    Return 300
                Else
                    Return 220
                End If
            Case "A203"
                If m_info.width > 1200 Then
                    Return 329
                Else
                    Return 250
                End If
            Case Else
                Return 0
        End Select
    End Function

    Function get_finish_time() As Date
        With frm_main.dgv_summary
            If IsNull(.Rows(.RowCount - 1).Cells("dgc_summary_finish").Value) Then
                Return Now
            Else
                Dim end_time As String = .Rows(.RowCount - 1).Cells("dgc_summary_finish").Value.ToString
                Return frm_main.tb_start_date.Text & " " & Strings.Left(end_time, 2) & ":" & Strings.Right(end_time, 2)
            End If
        End With
    End Function

    Function get_conversion_factor(ByVal on_prod As Boolean, ByVal mat_num As Integer, ByVal length As Boolean) As Decimal
        Dim kg, km As Decimal
        get_conversion_factor = 0


        If on_prod Then

            With frm_main.dgv_production

                For i As Integer = 0 To .RowCount - 1
                    If Not IsNull(.Rows(i).Cells("dgc_production_kg_out").Value) And Not IsNull(.Rows(i).Cells("dgc_production_km_out").Value) Then
                        kg = kg + .Rows(i).Cells("dgc_production_kg_out").Value
                        km = km + .Rows(i).Cells("dgc_production_km_out").Value
                    End If
                Next

            End With
        End If

        If kg = 0 Or km = 0 Or Not on_prod Then

            Dim dr() As DataRow = Nothing
            dr = dt_uom.Select("mat_num=" & mat_num)

            If Not dr Is Nothing Then
                If dr.Length > 0 Then
                    kg = dr(0)("numer")
                    km = dr(0)("denom")

                End If
            End If
            If km = 0 Or kg = 0 Then
                With frm_main.dgv_job_requirements
                    For i As Integer = 0 To .RowCount - 1
                        Select Case .Rows(i).Cells(0).Value
                            Case "KG"
                                kg = .Rows(i).Cells(1).Value
                            Case "M"
                                km = .Rows(i).Cells(1).Value / 1000
                            Case "MU"
                                i = i
                            Case "ROL"
                            Case Else
                                i = i
                        End Select
                    Next
                End With
            End If
        End If

        If length Then
            If kg > 0 Then Return kg / km
        Else
            If km > 0 Then Return km / kg
        End If
    End Function

    Sub CellKeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        number_only(sender, e, allow_decimal)
    End Sub

    Function get_material_info(Optional ByVal mat_in As Boolean = False, Optional ByVal mat_desc As String = Nothing) As materialInfo
        get_material_info = Nothing
        Dim s As String = Nothing, md As String = Nothing

        With get_material_info
            If s_selected_job Is Nothing And frm_main.tb_prod_num.Text = Nothing And mat_desc = Nothing Then
                Exit Function
            End If
            If mat_desc Is Nothing Then
                If s_selected_job Is Nothing Then
                    s_selected_job = dt_schedules.Select("prod_num=" & frm_main.tb_prod_num.Text & " AND plant=" & s_plant)
                ElseIf s_selected_job.Length = 0 Then
                    s_selected_job = dt_schedules.Select("prod_num=" & frm_main.tb_prod_num.Text & " AND plant=" & s_plant)
                End If
                If mat_in Then
                    md = s_selected_job(0)("mat_desc_in").ToString
                Else
                    md = s_selected_job(0)("mat_desc_semi").ToString
                    If md = Nothing Then md = s_selected_job(0)("mat_desc_fin").ToString
                End If
            Else
                md = mat_desc
            End If
            If md = Nothing Then Exit Function
            .formulation = Trim(Mid(md, 1, 8))
            ' MsgBox(Trim(Mid(md, 9, 1)))
            Select Case Trim(Mid(md, 9, 1)) '9th
                Case "A", "B", "C", "E", "F", "O", "P", "R", "U", "W", "Y"
                    Select Case Trim(Mid(md, 10, 1)) '10th
                        Case "C", "E", "S", "T", "W", "Y", "Z"
                            Select Case Trim(Mid(md, 11, 1)) '11th
                                Case "F", "W"
                                    Select Case Trim(Mid(md, 9, 1))
                                        Case "P", "F"
                                            .material_type = material_type.film
                                        Case Else
                                            md = md
                                    End Select
                                Case Else
                                    Select Case True
                                        Case Len(Trim(Mid(md, 28, 2))) = 2 'seal type string
                                            .material_type = material_type.bag
                                        Case IsNumeric(Trim(Mid(md, 31, 1))) And Not IsNumeric(Trim(Mid(md, 30, 1))) And Not Mid(md, 9, 1) = "C"
                                            .material_type = material_type.laminate
                                        Case Else
                                            .material_type = material_type.barrier
                                    End Select
                            End Select
                        Case "B", "Q"
                            .material_type = material_type.bag
                        Case "R"
                            .material_type = material_type.rollbag
                        Case "P"
                            .material_type = material_type.laminate
                        Case Else
                            s = Trim(Mid(md, 10, 1))
                            md = md
                    End Select
                Case "L", "M"
                    Select Case Trim(Mid(md, 10, 1)) '10th
                        Case "B", "P", "X", "1", "2"
                            .material_type = material_type.laminate
                        Case "C", "D", "J", "M", "N", "S"
                            .material_type = material_type.film
                        Case Else
                            s = Trim(Mid(md, 10, 1))
                            md = md
                    End Select
                Case Else
                    md = md
            End Select

            Select Case .material_type
                Case material_type.barrier, material_type.bag, material_type.rollbag
                    .colour = Mid(md, 9, 1)
                    .form = Mid(md, 10, 1)

                    s = Trim(Mid(md, 11, 6))
                    If IsNumeric(s) Then .width = CDec(s)

                    s = Trim(Mid(md, 18, 6))
                    If IsNumeric(s) Then .length = CDec(s)

                    If .material_type = material_type.bag Then
                        .gauge = Trim(Mid(md, 24, 3))
                        .seal_type = Trim(Mid(md, 28, 2))
                        .bundling_type = Mid(md, 30, 2)
                    Else
                        .gauge = Trim(Replace(LCase(Mid(md, 24, 4)), "m", Nothing))
                    End If
                    .print_type = Trim(Mid(md, 32, 2))
                    s = Trim(Right(md, 5))
                    If s.Length = 5 Then .zaw = s


                Case material_type.film, material_type.laminate
                    .form = Mid(md, 9, 1)
                    s = Replace(Mid(md, 12, 6), "X", Nothing)
                    If IsNumeric(s) Then .width = CDec(s)

                    s = Replace(Mid(md, 19, 6), "X", Nothing)
                    If IsNumeric(s) Then .length = CDec(s)

                    s = Replace(LCase(Mid(md, 25, 4)), "m", Nothing)
                    If IsNumeric(s) Then .gauge = s
                    If .material_type = material_type.film Then

                        s = Trim(Mid(md, 36, 5))
                        If s.Length = 5 And Not s = "PRESS" Then .zaw = s
                        .folding_type = Mid(md, 10, 2)
                        .core_type = Mid(md, 30, 1)
                        .core_size = Mid(md, 31, 1)
                        .print_type = Mid(md, 32, 1)

                    ElseIf .material_type = material_type.laminate Then
                        s = Mid(md, 36, 4)
                        If s = " PLT" Then
                            s = Replace(LCase(Mid(md, 19, 4)), "m", Nothing)
                            If IsNumeric(s) Then .gauge = s

                            .core_type = Mid(md, 24, 1)
                            .core_size = Mid(md, 25, 1)
                            .print_type = Mid(md, 26, 1)
                            s = Trim(Mid(md, 30, 5))
                            If s.Length = 5 Then .zaw = s
                        Else
                            .print_type = Trim(Mid(md, 32, 2))
                            s = Mid(md, 19, 6)
                            If IsNumeric(s) Then .length = CDec(s)
                            .core_type = Mid(md, 30, 1)
                            .core_size = Mid(md, 31, 1)

                            s = Trim(Right(md, 5))
                            If s.Length = 5 Then .zaw = s
                        End If
                    Else
                        s = s
                    End If


                Case Else
                    s = s
            End Select

            's = Trim(Mid(md, 36, 5))
            'If s.Length = 5 Then
            '    .zaw = Mid(md, 36, 5)
            'Else
            '    md = s_selected_job(0)("mat_desc_fin").ToString
            '    s = Trim(Mid(md, 36, 5))
            '    If s.Length = 5 Then
            '        .zaw = Mid(md, 35, 5)
            '    Else
            '        .zaw = Nothing
            '    End If
            'End If

        End With
    End Function

    Sub move_lb_item(ByVal lb As ListBox, Optional ByVal down As Boolean = False)
        Dim i As Integer = lb.SelectedIndex
        Dim s As String = lb.SelectedItem
        If down Then
            i = i + 2
        Else
            i = i - 1
        End If
        lb.Items.Insert(i, lb.SelectedItem)
        lb.Items.RemoveAt(lb.SelectedIndex)
        If down Then
            lb.SelectedIndex = i - 1
        Else
            lb.SelectedIndex = i
        End If
    End Sub

    Function get_department(ByVal str As String) As String
        Select Case Strings.Left(str, 4)
            Case "PPGE"
                Return "printing"
            Case "BMES", "RWGE"
                Return "bagplant"
            Case "SLFL", "EXFL"
                Return "filmsline"
            Case Else
                Return Nothing
        End Select
    End Function
    Function get_item_desity(ByVal md As String) As Decimal
        Dim mat_info As materialInfo = get_material_info(False, md)
        Select Case True
            Case mat_info.formulation.Contains("A203")
                Return 0.932
            Case mat_info.formulation.Contains("DL19")
                Return 0.92
            Case mat_info.formulation.Contains("DOGG")
                Return 0.92
            Case Else
                Return 1
        End Select

    End Function
    Function get_roll_length(ByVal kg As Decimal, ByVal md As String) As Integer
        Dim mat_info As materialInfo = get_material_info(False, md)
        Select Case mat_info.material_type
            Case material_type.film, material_type.laminate
                Dim d As Decimal = get_item_desity(md)
                Dim m As Integer = 1
                If mat_info.width < 400 Then
                    m = 2
                End If

                Return CInt((kg / (mat_info.width * _
                                  mat_info.gauge) / (0.000001 * m) / d)) '/ 50) * 50
            Case Else
                Return CInt((kg / (mat_info.width * _
                                   mat_info.gauge) / 0.000002)) '/ 50) * 50
        End Select
    End Function

    Function ConvertToLetter(iCol As Integer) As String
        ConvertToLetter = Nothing
        Dim iAlpha As Integer
        Dim iRemainder As Integer
        iAlpha = Int(iCol / 27)
        iRemainder = iCol - (iAlpha * 26)
        If iAlpha > 0 Then
            ConvertToLetter = Chr(iAlpha + 64)
        End If
        If iRemainder > 0 Then
            ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
        End If
    End Function
    Function get_print_type(ByVal print_type As String) As Print_type
        With get_print_type
            Select Case print_type
                Case "P"
                    .registered = False
                    .web_paths = 1
                    .style = Print_style.plain
                Case Else
                    .registered = False
                    .web_paths = 1
                    .style = Print_style.plain
            End Select
        End With

    End Function

    Function get_file_path(ByVal file As String) As String
        Return sql.read("settings_file_locations", "name", "=", "'" & file & "' AND plant=" & s_plant, "location")
    End Function

    Function get_handle(ByVal windowname As String) As IntPtr
        get_handle = FindWindow(vbNullString, windowname)
    End Function

    Sub KeyPress(ByVal key As Keys, ByVal handle As IntPtr, Optional ByVal up As Boolean = True)
        PostMessage(handle, WM_KEYDOWN, key, 0)
        If up Then PostMessage(handle, WM_KEYUP, key, 0)
    End Sub

    Sub close_excel()
        Dim xlHWND As Integer = xlApp.hwnd
        'this will have the process ID after call to GetWindowThreadProcessId
        Dim ProcIdXL As Integer = 0
        'get the process ID
        GetWindowThreadProcessId(xlHWND, ProcIdXL)
        'get the process
        Dim xproc As Process = Process.GetProcessById(ProcIdXL)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
        'set to nothing
        xlApp = Nothing
        'kill it with glee
        If Not xproc.HasExited Then
            xproc.Kill()
        End If
    End Sub

    Sub show_status(ByVal message As String, Optional ByVal next_line As String = Nothing)
        With frm_status
            If .Visible Then .Refresh()

            .l_status.Text = message
            .StartPosition = FormStartPosition.CenterScreen
            .ShowInTaskbar = False

            If Not .Visible Then
                'If next_line = Nothing Then
                '    .Label2.Visible = False
                '    .Height = 42
                'Else
                '    .Height = 65
                '    .Label2.Text = next_line
                'End If
                .Show()
            End If
            .Refresh()
        End With
    End Sub

    Function ConvertTime(ByVal time As String, Optional FromMin As Boolean = False) As String
        If IsNumeric(time) Then

            If FromMin Then
                If time < 0 Or time = Nothing Then
                    Return Nothing
                Else
                    Dim hours As Integer = time \ 60
                    Dim minutes As Integer = time - (hours * 60)
                    Return hours.ToString.PadLeft(2, "0") & minutes.ToString.PadLeft(2, "0")

                End If
            Else
                If time < 0 Or time = Nothing Then
                    Return Nothing
                Else
                    time = time * 60
                    Dim t As New TimeSpan(0, time, 0)
                    Return Format(t.Hours, "00") & Format(t.Minutes, "00")
                End If
            End If
        Else
            Return Nothing
        End If
    End Function

    Sub number_only(sender As Object, e As KeyPressEventArgs, ByVal dec As Boolean)

        Dim tb As Object = Nothing
        If TypeOf sender Is DataGridViewTextBoxEditingControl Then
            tb = CType(sender, DataGridViewTextBoxEditingControl)
        ElseIf TypeOf sender Is TextBox Then
            tb = CType(sender, TextBox)
        Else
            tb = tb
        End If

        If dec Then
            If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "."c Then
                e.Handled = True
            End If

            ' only allow one decimal point
            If tb.Text.Contains(".") And e.KeyChar = "."c Then
                tb.Text = "."
                tb.Select(tb.TextLength + 1, 1)
                e.Handled = True
            End If
        Else
            If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
                e.Handled = True
            End If

        End If

    End Sub

    Sub open_zaw(ByVal zaw As String)
        Dim s As String = get_file_path("zaw_doc") & zaw & ".pdf"
        If File.Exists(s) Then
            System.Diagnostics.Process.Start(s)
        Else
            MsgBox("The file for " & zaw & " was not found, please check FileCM" & vbCr & "If the file is not in FileCM, then see graphics.")
        End If

    End Sub

    Function GetMessage(ByVal message As String, ByVal language As String) As String
        Return sql.read("db_language", "message", "=", "'" & message & "'", language)
    End Function

    Function IsNull(ByVal item As Object) As Boolean
        If item Is Nothing Then Return True
        Dim t = item.GetType.Name
        If t = "String" Then
            If item = Nothing Then
                Return True
            Else
                Return False
            End If
        ElseIf t = "DBNull" Then
            Return True
        Else
            Return False
        End If
    End Function

    Sub select_cbo_item(ByVal cbo As ComboBox, ByVal match As String, Optional ByVal item As String = "-1")
        For i As Integer = 0 To cbo.Items.Count - 1
            If item = "-1" Then
                If LCase(cbo.Items(i)) = LCase(match) Then
                    cbo.SelectedIndex = i
                End If
            Else
                If LCase(cbo.Items(i)(item)) = LCase(match) Then
                    cbo.SelectedIndex = i
                End If
            End If
        Next
    End Sub

    Sub check_for_update()
        Dim info As UpdateCheckInfo = Nothing

        If (ApplicationDeployment.IsNetworkDeployed) Then
            Dim AD As ApplicationDeployment = ApplicationDeployment.CurrentDeployment

            Try
                info = AD.CheckForDetailedUpdate()
            Catch dde As DeploymentDownloadException
                MessageBox.Show("The new version of the application cannot be downloaded at this time. " + ControlChars.Lf & ControlChars.Lf & "Please check your network connection, or try again later. Error: " + dde.Message)
                Return
            Catch ioe As InvalidOperationException
                MessageBox.Show("This application cannot be updated. It is likely not a ClickOnce application. Error: " & ioe.Message)
                Return
            End Try

            If (info.UpdateAvailable) Then
                Dim doUpdate As Boolean = True

                If (Not info.IsUpdateRequired) Then
                    frm_main.Timer1.Enabled = False
                    Select Case s_department
                        Case "graphics", "planning"
                            s_user = "non_prod"
                            s_shift = "1"
                            s_machine = "non_prod"
                        Case Else
                            If Not s_user = Nothing Then
                                Dim s As Object = sql.read("db_restore", "user_", "=", "'" & s_user & "'", "restore_")
                                If IsNull(s) Then
                                    sql.insert("db_restore", "user_,date_,restore_,machine,shift,tab_page", "'" & s_user & "','" & frm_main.tb_start_date.Text & "',1,'" & s_machine & "','" & s_shift & "'," & frm_main.tc_main.SelectedIndex)
                                Else
                                    sql.update("db_restore", "restore_", 1 & ",date_='" & frm_main.tb_start_date.Text & "',machine='" & s_machine & "',shift='" & s_shift & "',tab_page=" & frm_main.tc_main.SelectedIndex, "user_", "=", "'" & s_user & "'")
                                End If
                            End If
                    End Select
                    Dim dr As DialogResult = MessageBox.Show("An update has been deteced and will be installed", "Update Detected", MessageBoxButtons.OK)
                    'If (Not System.Windows.Forms.DialogResult.OK = dr) Then
                    '    doUpdate = False
                    'End If
                Else
                    ' Display a message that the app MUST reboot. Display the minimum required version.
                    MessageBox.Show("This application has detected a mandatory update from your current " & _
                        "version to version " & info.MinimumRequiredVersion.ToString() & _
                        ". The application will now install the update and restart.", _
                        "Update Available", MessageBoxButtons.OK, _
                        MessageBoxIcon.Information)
                End If

                If (doUpdate) Then
                    Select Case s_department
                        Case "planning", "graphics"
                        Case Else
                            With frm_main
                                .export_report()
                                .export_summary()
                                .export_rts()
                                .export_reject()
                            End With
                    End Select

                    Try
                        AD.Update()
                        'MessageBox.Show("The application has been upgraded, and will now restart.")
                        Application.Restart()
                    Catch dde As DeploymentDownloadException
                        MessageBox.Show("Cannot install the latest version of the application. " & ControlChars.Lf & ControlChars.Lf & "Please check your network connection, or try again later.")
                        Return
                    End Try
                End If
            End If
        End If
    End Sub

    Function get_printer_name() As String
        Dim oPS As New System.Drawing.Printing.PrinterSettings

        Try
            Dim str() As String = Split(oPS.PrinterName, "\")
            Return str(str.Length - 1)
        Catch ex As System.Exception
            Return ""
        Finally
            oPS = Nothing
        End Try
    End Function

    Function dr_to_dt(ByVal dr() As DataRow) As DataTable
        If dr.Length > 0 Then
            Return dr.CopyToDataTable
        Else
            Return Nothing
        End If
    End Function

    Public Function IfNullObj(ByVal o As Object, Optional ByVal DefaultValue As String = "") As String
        Dim ret As String = ""
        Try
            If o Is DBNull.Value Then
                ret = DefaultValue
            Else
                ret = o.ToString
            End If
            Return ret
        Catch ex As Exception
            Return ret
        End Try
    End Function

    Public Function dgv_to_dt(ByVal dtg As DataGridView, Optional ByVal DataTableName As String = "myDataTable") As DataTable
        Try
            Dim dt As New DataTable(DataTableName)
            Dim row As DataRow
            Dim TotalDatagridviewColumns As Integer = dtg.ColumnCount - 1
            'Add Datacolumn
            For Each c As DataGridViewColumn In dtg.Columns
                Dim idColumn As DataColumn = New DataColumn()
                Dim c_name As String = Replace(dtg.Name, "dgv", "dgc")
                idColumn.ColumnName = Replace(c.Name, c_name & "_", Nothing)
                dt.Columns.Add(idColumn)
                If c.DefaultCellStyle.Format.ToString.Contains("N") Then
                    dt.Columns(idColumn.ColumnName).DataType = GetType(Decimal)
                End If
            Next
            'Now Iterate thru Datagrid and create the data row
            For Each dr As DataGridViewRow In dtg.Rows
                'Iterate thru datagrid
                row = dt.NewRow 'Create new row
                'Iterate thru Column 1 up to the total number of columns
                For cn As Integer = 0 To TotalDatagridviewColumns
                    If IsNumeric(dr.Cells(cn).Value) Then
                        row.Item(cn) = CDec(dr.Cells(cn).Value) ' This Will handle error datagridviewcell on NULL Values
                    ElseIf dt.Columns(cn).DataType = GetType(Decimal) Then
                        row.Item(cn) = DBNull.Value ' This Will handle error datagridviewcell on NULL Values
                    Else
                        row.Item(cn) = dr.Cells(cn).Value ' This Will handle error datagridviewcell on NULL Values
                    End If
                Next
                'Now add the row to Datarow Collection
                dt.Rows.Add(row)
            Next
            'Now return the data table
            Return dt
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Sub GetSerialPortNames()
        ' Show all available COM ports. 
        frm_main.ListBox1.Items.Clear()
        For Each sp As String In My.Computer.Ports.SerialPortNames
            frm_main.ListBox1.Items.Add(sp)
        Next
    End Sub

    Function get_status(ByVal status As Integer) As String
        get_status = Nothing
        Select Case status
            Case 0
                Return "Pending"
            Case 1
                Return "Mounted"
            Case 2
                Return "Running"
            Case 3
                Return "Prepared"
            Case 9
                Return "Last Completed Job"
            Case 10
                Return "Finished"
            Case -3
                Return "Label Printed"
            Case -2
                Return "Label Required"
            Case -1
                Return "New Job"
        End Select
    End Function
    'Function ReceiveSerialData(ByVal number_of_bytes As Integer) As String
    '    moRS232 = New Rs232()
    '    Dim returnStr As String = ""
    '    Dim cp As String
    '    If frm_main.ListBox1.SelectedItem = Nothing Then
    '        Exit Function
    '    Else
    '        cp = Replace(frm_main.ListBox1.SelectedItem, "COM", Nothing)
    '    End If
    '    Try
    '        '// Setup parameters
    '        With moRS232
    '            .Port = cp
    '            .BaudRate = 9600
    '            .DataBit = 1
    '            .Parity = Rs232.DataParity.Parity_None
    '            .StopBit = Rs232.DataStopBit.StopBit_1

    '            .Timeout = 500
    '        End With
    '        '// Initializes port
    '        moRS232.Open()
    '        '// Set state of RTS / DTS
    '        moRS232.Dtr = False
    '        moRS232.Rts = True
    '        'If chkEvents.Checked Then moRS232.EnableEvents()
    '        'chkEvents.Enabled = True
    '    Catch Ex As Exception
    '        returnStr = Ex.Message
    '    Finally
    '        moRS232.Read(number_of_bytes)
    '        returnStr = moRS232.InputStreamString
    '        moRS232.PurgeBuffer(Rs232.PurgeBuffers.RXClear)
    '        moRS232.Write("ATZ")
    '        moRS232.Close()
    '    End Try     ' Receive strings from a serial port. 


    '    Return returnStr
    'End Function

End Module
Imports System.Threading.Thread
Imports System.Text
Public Class cls_ACSIS
    Public logon_name As String
    Enum eWindowTarget
        ProductionConfirmation
        PrintPalletLabels
        StockOverview
        WarehouseReceiptsOther
        ReLabeling
        ModifyPalletLabel
    End Enum

    Enum eControlType
        Button
        ComboBox
        CheckBox
        TextBox
        ListBox
        OptionButton
    End Enum

    Enum eProdbox
        Date_
        Production_Number
        User
    End Enum

    Sub set_check(ByVal handle As IntPtr, ByVal checkstate As Boolean)
        If checkstate Then
            SendMessage(handle, BM_SETCHECK, BST_CHECKED, 0)
        Else
            SendMessage(handle, BM_SETCHECK, BST_UNCHECKED, 0)
        End If
    End Sub

    Sub select_item(ByVal handle As IntPtr, ByVal index As Integer, ByVal list As Boolean)

        Dim sel, cur
        If list Then
            sel = LB_SETCURSEL
            cur = LB_GETCURSEL
        Else
            sel = CB_SETCURSEL
            cur = CB_GETCURSEL
        End If
        Dim i As Integer = SendMessage(handle, cur, 0, 0)
        While Not i = index
            popup_check(True, 0)
            If s_popup = 4 Then
                Exit Sub
            End If
            i = SendMessage(handle, cur, 0, 0)
            If list Then
                If i < index Then
                    KeyPress(Keys.Down, handle)
                ElseIf i > index Then
                    KeyPress(Keys.Up, handle)
                End If
            Else
                If index = 0 Then
                    SendMessage(handle, sel, index + 1, 0)
                    KeyPress(Keys.Up, handle)
                Else
                    SendMessage(handle, sel, index - 1, 0)
                    KeyPress(Keys.Down, handle)
                End If
            End If

        End While
    End Sub

    Sub select_cb_item(ByVal item_name As String, ByVal session As Integer)

        For i As Integer = 0 To w_sessions(session).winWnd(0).cbWnd.Length - 1
            For l As Integer = 0 To w_sessions(session).winWnd(0).cbWnd(i).count
                If w_sessions(session).winWnd(0).cbWnd(i).text(l).Contains(item_name) Then
                    select_item(w_sessions(session).winWnd(0).cbWnd(i).handle, l, False)
                    Exit Sub
                End If
            Next
        Next
    End Sub

    Function get_button_back(ByVal session As Integer) As String
        dr_navigation = dt_navigation.Select("window_name='" & w_sessions(session).winWnd(0).name & "'")
        Return dr_navigation(0)("button_exit")
    End Function

    Function get_button_access(ByVal item As String, ByVal session As Integer) As String

        If item = Nothing Then
            dr_navigation = dt_navigation.Select("window_name='" & w_sessions(session).winWnd(0).name & "'")
            Return dr_navigation(0)("button_access")
        Else
            get_button_access = Nothing
            For i As Integer = 0 To dt_navigation.Rows.Count - 1
                If dt_navigation(i)("target") = item Then
                    Return dt_navigation(i)("button_access")
                End If
            Next
        End If
    End Function

    Function get_acsis_option(ByVal item As String) As String
        get_acsis_option = Nothing
        For i As Integer = 0 To dt_acsis_options.Rows.Count - 1
            If dt_acsis_options(i)("item") = item Then
                Return dt_acsis_options(i)("val")
            End If
        Next
        If get_acsis_option = Nothing Then
            get_acsis_option = get_acsis_option
        End If
    End Function

    Function get_view_item(ByVal item As String, ByVal session As Integer, ByVal type As Boolean) As String
        get_view_item = Nothing
        Dim s() As String = Nothing
        For i As Integer = 0 To w_sessions(session).winWnd(0).lbWnd(0).count
            s = Split(w_sessions(session).winWnd(0).lbWnd(0).text(i), "  ")
            If s(1) = item Then
                If type Then
                    If s.Length = 2 Then
                        Return s(3)
                    Else
                        Return s(3)
                    End If
                Else
                    Return s(2)
                End If
            End If
        Next
        If w_sessions(session).winWnd(0).lbWnd(0).count >= 0 And get_view_item = Nothing Then
            s = s
        End If
    End Function
    Function running(ByVal session As Integer) As Boolean
        'acsis.get_controls(0, True, False, True)
        Try
            If w_sessions(session).process.HasExited Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
        'If w_sessions(session).process.MainWindowTitle Is Nothing Then
        '    Return False
        'End If
        'If w_sessions(session).process.HasExited Then
        '    Return False
        'Else
        '    Return true
        'End If
        'Dim twnd As IntPtr, s As String
        'For i As Integer = 0 To 1
        '    s = Replace(dt_symple_session(0)("logon"), "#", i + 1)
        '    twnd = FindWindowByCaption(0, s)
        '    If twnd = 0 Then
        '        s = Replace(dt_symple_session(0)("main"), "#", i + 1)
        '        twnd = FindWindowByCaption(0, s)
        '    Else
        '        Dim pid As Integer = 0
        '        GetWindowThreadProcessId(twnd, pid)
        '        Dim aProcess As System.Diagnostics.Process
        '        aProcess = System.Diagnostics.Process.GetProcessById(pid)
        '        aProcess.Kill()
        '    End If

        '    If Not twnd = 0 Then
        '        Return True

        '    End If

        'Next

        '   Return False
    End Function

    Sub enter_minute(ByVal prod_num As Long)
        Dim min_entered As Boolean = False
        While Not min_entered
            If prod_num = 0 Then prod_num = w_sessions(0).winWnd(0).tbWnd(get_textbox("prod_num", 0, 0)).text(0)
            If Not IsNull(s_selected_job(0)("min_entered")) Then
                min_entered = s_selected_job(0)("min_entered")
            End If
            If Not min_entered Then

                open_job_screen(prod_num, True)
                popup_check(True, 0)
                If s_failed_job Then Exit Sub
                If has_prep() = True Then

                    show_status("Checking for 1 minute in Pre-Press")
                    acsis.get_controls(0, True, False, True)
                    Do While w_sessions(0).winWnd(0).lbWnd Is Nothing
                        click_button(get_button_access("View Production Order", 0), False, 0)
                    Loop

                    If get_view_item("Setup", 0, False) = Nothing Then min_entered = False

                    click_button(get_button_back(0), False, 0)

                    If Not min_entered Then

                        show_status("Entering 1 minute into Pre-Press")

                        click_button(get_button_access("Setup Confirmation", 0), False, 0)

                        insert_string("0000", w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_setup_start")).handle)
                        insert_string("0001", w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_setup_end")).handle)

                        click_button(get_acsis_option("production_setup_confirm"), False, 0)
                        sql.update("db_schedule", "min_entered", 1, "prod_num", "=", prod_num)
                        s_selected_job(0)("min_entered") = True
                        s_popup = 0
                        open_job_screen(prod_num, False)

                    End If
                    s_prep_minute_count = 1
                Else
                    min_entered = True
                End If
            Else
                If s_prep_minute_count > 10 Then
                    sql.update("db_schedule", "min_entered", 0, "prod_num", "=", prod_num)
                    s_selected_job(0)("min_entered") = False
                    min_entered = False
                Else
                    show_status("Waiting for minute to process attempt " & s_prep_minute_count & " of 10")

                    s_prep_minute_count = s_prep_minute_count + 1
                End If

            End If
            System.Threading.Thread.Sleep(1000)
        End While

        frm_status.Close()

    End Sub

    Sub popup_check(ByVal skip_user As Boolean, ByVal session As Integer)
        Dim btn As String, tWnd As IntPtr, prod_num As Long
        If Not s_symple_access Then Exit Sub
        tWnd = get_handle(w_sessions(session).process.MainWindowTitle)
        If Not frm_main.tb_prod_num.Text = Nothing Then prod_num = frm_main.tb_prod_num.Text
        If Not tWnd = 0 And s_popup = 0 Then
            ' get_controls(get_handle(symple.MainWindowTitle))
            acsis.get_controls(0, True, False, skip_user)
            If Not w_sessions(session).winWnd Is Nothing Then

                Dim result() As DataRow

                For i As Integer = 0 To w_sessions(session).winWnd.Length - 1
                    If w_sessions(session).winWnd(i).popup Then

                        result = dt_popups.Select("message='" & Replace(w_sessions(session).winWnd(i).name, "'", "''") & "'")
                        If result.Length = 0 Then
                            sql.delete("acsis_popups", "message", "=", "'" & Replace(w_sessions(session).winWnd(i).staticWnd(i).name, "'", "''") & "' AND plant=" & s_plant)
                            sql.insert("acsis_popups", "message,button,plant", "'" & Replace(w_sessions(session).winWnd(i).staticWnd(i).name, "'", "''") & "','" & _
                                       w_sessions(session).winWnd(i).btnWnd(0).text(0) & "'," & s_plant)

                            dt_popups = sql.get_table("acsis_popups", "plant", "=", s_plant)
                            result = dt_popups.Select("message='" & Replace(w_sessions(session).winWnd(i).staticWnd(i).name, "'", "''") & "'")
                        End If
                        If Not result.Length = 0 Then
                            Do While w_sessions(session).winWnd(i).btnWnd Is Nothing
                                acsis.get_controls(0, True, False, skip_user)
                            Loop
                            btn = result(0)("button")
                            For b As Integer = 0 To w_sessions(session).winWnd(i).btnWnd.Length - 1
                                If w_sessions(session).winWnd(i).btnWnd(b).text(0) = btn Then
                                    SendMessage(w_sessions(session).winWnd(i).btnWnd(b).handle, WM_ACTIVATE, 1, 0)
                                    PostMessage(w_sessions(session).winWnd(i).btnWnd(b).handle, BM_CLICK, 0, 0)
                                    w_sessions(session).process.WaitForInputIdle()

                                    btn = Nothing
                                    If Not IsDBNull(result(0)("type")) Then
                                        s_popup = result(0)("type")
                                    Else
                                        s_popup = 0
                                    End If
                                    Exit For
                                    popup_check(skip_user, session)
                                End If
                            Next
                            If btn = Nothing Then Exit For
                        End If
                    End If
                Next

                Select Case s_popup
                    Case 1
                        ' login error - no longer in use
                    Case 2
                        show_status("Setup time for PMNT has not been detected yet.")
                        enter_minute(prod_num)
                        s_popup = 0
                    Case 3
                        MsgBox("That Pallet does not exist!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Pallet not found")
                        s_popup = 0
                    Case 4
                        'external code
                        s_failed_job = True
                    Case 5
                        'shutdown
                        s_popup = 0
                    Case 7
                        s_popup = 0
                        'MsgBox("not done")
                    Case 8
                        frm_main.cb_sql_symple.Checked = False
                        If sw_symple.IsRunning Then
                            sw_symple.Reset()
                        Else
                            sw_symple.Start()
                        End If
                        s_popup = 0
                    Case 9
                        ' too much has been rts
                        'MsgBox("not done")
                    Case 10
                        'prod_num is missing
                        acsis.get_controls(0, True, False, skip_user)
                        insert_string(frm_main.tb_prod_num.Text, w_sessions(session).winWnd(0).tbWnd(get_acsis_option("rts_screen_prod_num")).handle)
                        click_button(get_acsis_option("rts_screen_put_away"), False, session)
                        s_popup = 0
                    Case 12
                        If s_production Then s_popup = 0
                        If frm_symple_times.Visible Then
                            SetForegroundWindow(frm_symple_times.Handle)
                            s_final_done = True
                        Else
                            SetForegroundWindow(frm_main.Handle)
                            s_final_done = False
                        End If
                        If s_production Then
                            frm_status.Close()
                            MsgBox("This job has been final confirmed, you need to reverse this to enter data.")
                        End If
                        sql.update("db_schedule", "status", 10, "prod_num", "=", s_selected_job(0)("prod_num"))
                        s_selected_job(0)("status") = 10
                    Case 13
                        frm_status.Close()
                        MsgBox("Error Logging in, please close all instances of the Sym-PLE client")
                    Case 14
                        s_popup = 0
                        logon()
                End Select

                'End If
                '    Next
            End If

        End If
        acsis.get_controls(0, True, False, skip_user)
        'While w_sessions Is Nothing
        '    get_controls(0, True)
        'End While

    End Sub
    Sub get_symple_controls(ByVal hParent As Long, ByVal start As Boolean, Optional ByVal dBug As Boolean = False, Optional popup As Boolean = False)
        Dim tWnd As IntPtr
        Dim strText, lngRet, type, session As String
        If Not w_sessions Is Nothing Then
            session = frm_main.lb_symple_sesions.SelectedItem
        Else
            session = Nothing
        End If
        Dim i, c As Integer


        tWnd = FindWindowEx(hParent, 0, vbNullString, vbNullString)
        While Not tWnd = 0

            strText = StrDup(100, Chr(0))
            lngRet = GetClassName(tWnd, strText, 100)
            type = Left(strText, lngRet)
            If Not tWnd = 0 Then
                'If port_reader Then
                '     insert_string("asd", tWnd)
                '    i = i
                'End If

                For ss As Integer = 0 To w_sessions.Length - 1
                    If w_sessions(ss).session = s_current_session Then
                        For w As Integer = 0 To w_sessions(ss).winWnd.Length - 1
                            If w_sessions(ss).winWnd(w).name = s_current_window Then
                                Select Case type
                                    Case "ThunderRT6TextBox", "MSMaskWndClass"
                                        If w_sessions Is Nothing Then Exit Select

                                        If w_sessions(ss).winWnd(w).tbWnd Is Nothing Then
                                            ReDim w_sessions(ss).winWnd(w).tbWnd(0)
                                        Else
                                            ReDim Preserve w_sessions(ss).winWnd(w).tbWnd(w_sessions(ss).winWnd(w).tbWnd.Length)
                                        End If
                                        i = w_sessions(ss).winWnd(w).tbWnd.Length - 1
                                        ReDim w_sessions(ss).winWnd(w).tbWnd(i).text(0)
                                        w_sessions(ss).winWnd(w).tbWnd(i).handle = tWnd
                                        w_sessions(ss).winWnd(w).tbWnd(i).text(0) = get_text(tWnd, False)
                                        w_sessions(ss).winWnd(w).tbWnd(i).enabled = IsWindowEnabled(tWnd)

                                        If dBug Then
                                            insert_string(i, w_sessions(ss).winWnd(w).tbWnd(i).handle)
                                            i = i
                                        End If
                                        '                    End If
                                        '        Next

                                        '    End If
                                        'Next

                                    Case "ThunderRT6ComboBox"
                                        If w_sessions Is Nothing Then Exit Select
                                        'For ss As Integer = 0 To w_sessions.Length - 1
                                        '    If w_sessions(ss).session = s_current_session Then
                                        '        For w As Integer = 0 To w_sessions(ss).winWnd.Length - 1
                                        '            If w_sessions(ss).winWnd(w).name = s_current_window Then

                                        If w_sessions(ss).winWnd(w).cbWnd Is Nothing Then
                                            ReDim w_sessions(ss).winWnd(w).cbWnd(0)
                                        Else
                                            ReDim Preserve w_sessions(ss).winWnd(w).cbWnd(w_sessions(ss).winWnd(w).cbWnd.Length)
                                        End If
                                        i = w_sessions(ss).winWnd(w).cbWnd.Length - 1
                                        w_sessions(ss).winWnd(w).cbWnd(i).count = SendMessage(tWnd, CB_GETCOUNT, 0, 0) - 1
                                        w_sessions(ss).winWnd(w).cbWnd(i).handle = tWnd
                                        ReDim w_sessions(ss).winWnd(w).cbWnd(i).text(w_sessions(ss).winWnd(w).cbWnd(i).count)
                                        For c = 0 To w_sessions(ss).winWnd(w).cbWnd(i).count
                                            w_sessions(ss).winWnd(w).cbWnd(i).text(c) = get_text(tWnd, False, c)
                                        Next
                                        '            End If
                                        '        Next

                                        '    End If
                                        'Next

                                    Case "ThunderRT6ListBox"
                                        If w_sessions Is Nothing Then Exit Select
                                        'For ss As Integer = 0 To w_sessions.Length - 1
                                        '    If w_sessions(ss).session = s_current_session Then
                                        '        For w As Integer = 0 To w_sessions(ss).winWnd.Length - 1
                                        '            If w_sessions(ss).winWnd(w).name = s_current_window Then

                                        If w_sessions(ss).winWnd(w).lbWnd Is Nothing Then
                                            ReDim w_sessions(ss).winWnd(w).lbWnd(0)
                                        Else
                                            ReDim Preserve w_sessions(ss).winWnd(w).lbWnd(w_sessions(ss).winWnd(w).lbWnd.Length)
                                        End If
                                        i = w_sessions(ss).winWnd(w).lbWnd.Length - 1
                                        w_sessions(ss).winWnd(w).lbWnd(i).count = SendMessage(tWnd, LB_GETCOUNT, 0, 0) - 1
                                        w_sessions(ss).winWnd(w).lbWnd(i).handle = tWnd
                                        ReDim w_sessions(ss).winWnd(w).lbWnd(i).text(w_sessions(ss).winWnd(w).lbWnd(i).count)
                                        For c = 0 To w_sessions(ss).winWnd(w).lbWnd(i).count
                                            w_sessions(ss).winWnd(w).lbWnd(i).text(c) = get_text(tWnd, True, c)
                                        Next
                                        '            End If
                                        '        Next
                                        '    End If
                                        'Next


                                    Case "ThunderRT6OptionButton"
                                        If w_sessions Is Nothing Then Exit Select
                                        'For ss As Integer = 0 To w_sessions.Length - 1
                                        '    If w_sessions(ss).session = s_current_session Then
                                        '        For w As Integer = 0 To w_sessions(ss).winWnd.Length - 1
                                        '            If w_sessions(ss).winWnd(w).name = s_current_window Then

                                        If Not get_text(tWnd, False) = Nothing Then
                                            If IsWindowVisible(tWnd) Then
                                                If w_sessions(ss).winWnd(w).optWnd Is Nothing Then
                                                    ReDim w_sessions(ss).winWnd(w).optWnd(0)
                                                Else
                                                    ReDim Preserve w_sessions(ss).winWnd(w).optWnd(w_sessions(ss).winWnd(w).optWnd.Length)
                                                End If
                                                i = w_sessions(ss).winWnd(w).optWnd.Length - 1
                                                ReDim w_sessions(ss).winWnd(w).optWnd(i).text(0)
                                                w_sessions(ss).winWnd(w).optWnd(i).enabled = IsWindowEnabled(tWnd)
                                                w_sessions(ss).winWnd(w).optWnd(i).checked = SendMessage(tWnd, BM_GETCHECK, 0, 0)
                                                w_sessions(ss).winWnd(w).optWnd(i).text(0) = get_text(tWnd, False)
                                                w_sessions(ss).winWnd(w).optWnd(i).handle = tWnd
                                            End If
                                        End If
                                        '            End If
                                        '        Next

                                        '    End If
                                        'Next
                                    Case "ThunderRT6CheckBox"
                                        If w_sessions Is Nothing Then Exit Select
                                        'For ss As Integer = 0 To w_sessions.Length - 1
                                        '    If w_sessions(ss).session = s_current_session Then
                                        '        For w As Integer = 0 To w_sessions(ss).winWnd.Length - 1
                                        '            If w_sessions(ss).winWnd(w).name = s_current_window Then

                                        If IsWindowVisible(tWnd) Then
                                            If w_sessions(ss).winWnd(w).cbxWnd Is Nothing Then
                                                ReDim w_sessions(ss).winWnd(w).cbxWnd(0)
                                            Else
                                                ReDim Preserve w_sessions(ss).winWnd(w).cbxWnd(w_sessions(ss).winWnd(w).cbxWnd.Length)
                                            End If
                                            i = w_sessions(ss).winWnd(w).cbxWnd.Length - 1
                                            ReDim w_sessions(ss).winWnd(w).cbxWnd(i).text(0)
                                            w_sessions(ss).winWnd(w).cbxWnd(i).checked = SendMessage(tWnd, BM_GETCHECK, 0, 0)
                                            w_sessions(ss).winWnd(w).cbxWnd(i).handle = tWnd
                                            w_sessions(ss).winWnd(w).cbxWnd(i).text(0) = get_text(tWnd, False)
                                        End If
                                        '            End If
                                        '        Next

                                        '    End If
                                        'Next
                                    Case "ThunderRT6CommandButton", "Button"
                                        If w_sessions Is Nothing Then Exit Select
                                        'For ss As Integer = 0 To w_sessions.Length - 1
                                        '    If w_sessions(ss).session = s_current_session Then
                                        '        For w As Integer = 0 To w_sessions(ss).winWnd.Length - 1
                                        '            If w_sessions(ss).winWnd(w).name = s_current_window Then
                                        'If popup Then
                                        '    If w_sessions(ss).staticWnd(w).btnWnd Is Nothing Then
                                        '        ReDim w_sessions(ss).staticWnd(w).btnWnd(0)
                                        '    Else
                                        '        ReDim Preserve w_sessions(ss).staticWnd(w).btnWnd(w_sessions(ss).staticWnd(w).btnWnd.Length)
                                        '    End If
                                        '    i = w_sessions(ss).staticWnd(w).btnWnd.Length - 1
                                        '    ReDim w_sessions(ss).staticWnd(w).btnWnd(i).text(0)
                                        '    w_sessions(ss).staticWnd(w).btnWnd(i).handle = tWnd
                                        '    w_sessions(ss).staticWnd(w).btnWnd(i).text(0) = get_text(tWnd, False)
                                        '    w_sessions(ss).staticWnd(w).btnWnd(i).enabled = IsWindowEnabled(tWnd)
                                        'Else

                                        If w_sessions(ss).winWnd(w).btnWnd Is Nothing Then
                                            ReDim w_sessions(ss).winWnd(w).btnWnd(0)
                                        Else
                                            ReDim Preserve w_sessions(ss).winWnd(w).btnWnd(w_sessions(ss).winWnd(w).btnWnd.Length)
                                        End If
                                        i = w_sessions(ss).winWnd(w).btnWnd.Length - 1
                                        ReDim w_sessions(ss).winWnd(w).btnWnd(i).text(0)
                                        w_sessions(ss).winWnd(w).btnWnd(i).handle = tWnd
                                        w_sessions(ss).winWnd(w).btnWnd(i).text(0) = get_text(tWnd, False)
                                        w_sessions(ss).winWnd(w).btnWnd(i).enabled = IsWindowEnabled(tWnd)
                                        ' If popup Then
                                        If type = "Button" Then
                                            popup = popup
                                        End If
                                        'End If
                                        'End If
                                        '        End If
                                        '    Next

                                        'End If
                                        'Next
                                    Case "ApexGrid.19"
                                        If w_sessions Is Nothing Then Exit Select
                                        'For ss As Integer = 0 To w_sessions.Length - 1
                                        '    If w_sessions(ss).session = s_current_session Then
                                        '        For w As Integer = 0 To w_sessions(ss).winWnd.Length - 1
                                        '            If w_sessions(ss).winWnd(w).name = s_current_window Then

                                        If w_sessions(ss).winWnd(w).gridWnd Is Nothing Then
                                            ReDim w_sessions(ss).winWnd(w).gridWnd(0)
                                        Else
                                            ReDim Preserve w_sessions(ss).winWnd(w).gridWnd(w_sessions(ss).winWnd(w).gridWnd.Length)
                                        End If
                                        i = w_sessions(ss).winWnd(w).gridWnd.Length - 1
                                        w_sessions(ss).winWnd(w).gridWnd(i).handle = tWnd
                                        '            End If
                                        '        Next

                                        '    End If
                                        'Next


                                    Case "Edit", "StatusBar20WndClass", "ThunderRT6Frame", _
                                        "ThunderRT6Timer", "MSFlexGridWndClass", "VBFocusRT6"

                                        'do nothing

                                        'TrayClockWClass = time in symple

                                End Select
                            End If
                        Next
                    End If
                Next
            End If

            get_symple_controls(tWnd, False, dBug)
            tWnd = FindWindowEx(hParent, tWnd, vbNullString, vbNullString)

        End While
    End Sub
    Sub assign_process(ByVal session As Integer, ByVal tWnd As IntPtr)
        Dim pid As Integer
        GetWindowThreadProcessId(tWnd, pid)
        'symple(i) = System.Diagnostics.Process.GetProcessById(pid)
        Try
            If w_sessions(session).process.HasExited Then
                w_sessions(session).user = Nothing
                'get_controls(0, True, False, False)
                w_sessions(session).process = System.Diagnostics.Process.GetProcessById(pid)
            Else
                ' frm_main.lb_symple_sesions.ClearSelected()
                'kill logon
            End If
        Catch ex As Exception
            'no process, assign it
            w_sessions(session).user = Nothing
            ' get_controls(0, True, False, False)
            w_sessions(session).process = System.Diagnostics.Process.GetProcessById(pid)
        End Try

    End Sub
    Sub get_controls(ByVal hParent As Long, ByVal Start As Boolean, ByVal dBug As Boolean, ByVal skip_user As Boolean)

        Dim tWnd As IntPtr
        Dim strText, lngRet, type, s As String ', session As String
        'If frm_main.lb_symple_sesions.SelectedIndex = -1 Then
        '    session = frm_main.lb_symple_sesions.SelectedItem
        'Else
        '    session = Nothing
        'End If
        Dim i As Integer
        If Start Then
            'Erase tbWnd, btnWnd, cbWnd, optWnd, gridWnd, lbWnd, cbxWnd, staticWnd
            'Dim l_on As Boolean = False
            If hParent = 0 Then
                'Erase w_sessions
                s_current_session = Nothing
                'Start = False
                'frm_main.lb_symple_sesions.Visible = False
                For i = 0 To 1
                    Dim pid As Integer = 0
                    s = Replace(dt_symple_session(0)("logon"), "#", i + 1)
                    tWnd = FindWindowByCaption(0, s)
                    If tWnd = 0 Then
                        s = Replace(dt_symple_session(0)("main"), "#", i + 1)
                        tWnd = FindWindowByCaption(0, s)
                    Else
                        If Not w_sessions(i).session = Nothing And Not w_sessions(i).session = s Then
                            If w_sessions(i).process.HasExited Then
                                '  l_on = True
                                s = Replace(dt_symple_session(0)("main"), "#", i + 1)
                                tWnd = FindWindowByCaption(0, s)
                            Else
                                GetWindowThreadProcessId(tWnd, pid)
                                System.Diagnostics.Process.GetProcessById(pid).Kill()
                                MsgBox("An instance of this session is already running!")
                                s = Replace(dt_symple_session(0)("main"), "#", i + 1)
                                tWnd = FindWindowByCaption(0, s)
                            End If
                        End If

                    End If

                    s_current_process = 0
                    GetWindowThreadProcessId(tWnd, s_current_process)
                    If Not tWnd = 0 Then
                        If Not running(i) Then assign_process(i, tWnd)

                        w_sessions(i).winWnd = Nothing
                        w_sessions(i).session = s
                        w_sessions(i).checked = False

                        frm_main.lb_symple_sesions.Items(i) = s
                        If frm_main.lb_symple_sesions.SelectedIndex = -1 And Not s = Nothing Then
                            If i = 0 Then frm_main.lb_symple_sesions.SelectedIndex = 0
                            'session = s
                        End If

                        s_current_session = s
                        Do While Not w_sessions(i).checked 'And Not s_department = "planning"
                            get_controls(0, False, dBug, skip_user)
                            'If w_sessions Is Nothing Then Exit Do
                        Loop

                    Else
                        w_sessions(i) = Nothing
                        ' frm_main.lb_symple_users.SelectedIndex = -1
                        frm_main.lb_symple_sesions.Items(i) = " "
                        frm_main.lb_symple_users.Items(i) = " "
                    End If


                Next

            End If
            Exit Sub
        End If

        tWnd = FindWindowEx(hParent, 0, vbNullString, vbNullString)
        While Not tWnd = 0
            Dim pid As Integer = 0
            GetWindowThreadProcessId(tWnd, pid)
            If pid = s_current_process Then

                strText = StrDup(100, Chr(0))
                lngRet = GetClassName(tWnd, strText, 100)
                type = Left(strText, lngRet)
                If Not tWnd = 0 Then
                    'If port_reader Then
                    '     insert_string("asd", tWnd)
                    '    i = i
                    'End If

                    Select Case type
                        Case "ThunderRT6FormDC", "Static"
                            s = get_text(tWnd, False)
                            'If Start Then s_current_session = s
                            If s = "" Then Exit Select
                            Dim s_name As String = dt_symple_session(0)("main")
                            If IsWindowVisible(tWnd) And Not s.Contains("WFT") Then
                                If IsNumeric(get_string_difference(s, s_name)) Then
                                    s_current_session = s
                                End If

                                If frm_main.lb_symple_sesions.Items.Count > 0 Then
                                    For ss As Integer = 0 To 1
                                        If get_handle(w_sessions(ss).session) = 0 Then w_sessions(ss).checked = True
                                        If w_sessions(ss).session = s_current_session Then
                                            s_current_window = s
                                            Dim add As Boolean = True
                                            If w_sessions(ss).winWnd Is Nothing Then
                                                ReDim w_sessions(ss).winWnd(0)
                                                add = False
                                            End If
                                            For i = 0 To w_sessions(ss).winWnd.Length - 1
                                                If w_sessions(ss).winWnd(0).name = s Then
                                                    add = False
                                                End If
                                            Next
                                            If add Then
                                                ReDim Preserve w_sessions(ss).winWnd(w_sessions(ss).winWnd.Length)
                                            End If
                                            i = w_sessions(ss).winWnd.Length - 1
                                            w_sessions(ss).winWnd(i).name = s
                                            w_sessions(ss).winWnd(i).handle = tWnd
                                            If type = "Static" Then
                                                w_sessions(ss).winWnd(i).popup = True
                                                get_symple_controls(get_handle(w_sessions(ss).process.MainWindowTitle), True, dBug, True)
                                            Else
                                                get_symple_controls(tWnd, True, dBug)
                                            End If
                                            w_sessions(ss).checked = True

                                        End If
                                    Next
                                End If
                            End If
                            'Case "Static"
                            '    If w_sessions Is Nothing Then Exit Select
                            '    For ss As Integer = 0 To w_sessions.Length - 1
                            '        If w_sessions(ss).session = s_current_session Then
                            '            Dim add As Boolean = True
                            '            If w_sessions(ss).winWnd Is Nothing Then
                            '                ReDim w_sessions(ss).winWnd(0)
                            '                add = False
                            '            End If
                            '            For i = 0 To w_sessions(ss).winWnd.Length - 1
                            '                If w_sessions(ss).winWnd(0).name = s Then
                            '                    add = False
                            '                End If
                            '            Next
                            '            If add Then
                            '                ReDim Preserve w_sessions(ss).winWnd(w_sessions(ss).winWnd.Length)
                            '            End If
                            '            i = w_sessions(ss).winWnd.Length - 1
                            '            w_sessions(ss).winWnd(i).name = s
                            '            w_sessions(ss).winWnd(i).handle = tWnd
                            '            w_sessions(ss).checked = True
                            '            get_symple_controls(tWnd, True, dBug)

                            '        End If
                            '    Next
                        Case Else
                            If type = type Then

                            End If
                            Dim sd As String = get_text(tWnd, False)
                            ' MsgBox(type)
                            If Not sd = Nothing Then

                                sd = sd
                            End If


                    End Select
                End If

                get_controls(tWnd, False, dBug, skip_user)
                'Else
                '    Exit Sub
            End If
            tWnd = FindWindowEx(hParent, tWnd, vbNullString, vbNullString)
        End While
    End Sub

    Function get_text(ByVal hWnd As IntPtr, ByVal list As Boolean, Optional ByVal i As Integer = -1) As String
        Dim nIndex As Long, GetTrim As Integer, GetString, TrimSpace As String
        If list Then

            While nIndex < 0 Or GetTrim < 0
                nIndex = SendMessage(hWnd, LB_GETCURSEL, 0&, 0&)
                GetTrim = SendMessage(hWnd, LB_GETTEXTLEN, nIndex, 0&)
                Sleep(100)
            End While
            If i > -1 Then
                nIndex = i
            Else
                nIndex = SendMessage(hWnd, LB_GETCURSEL, 0&, 0&)
            End If
            If nIndex > -1 Then
                GetTrim = SendMessage(hWnd, LB_GETTEXTLEN, nIndex, 0&)
                TrimSpace$ = Space$(GetTrim)
                GetString = SendMessage(hWnd, LB_GETTEXT, nIndex, TrimSpace$)
                get_text = TrimSpace$
            Else
                get_text = Nothing
            End If
        Else
            If i = -1 Then
                GetTrim = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0)
                TrimSpace$ = Space$(GetTrim) '6
                GetString = SendMessage(hWnd, WM_GETTEXT, GetTrim + 1, TrimSpace$)
                get_text = TrimSpace$
            Else
                GetTrim = SendMessage(hWnd, CB_GETLBTEXTLEN, i, 0)
                GetString = StrDup(GetTrim, Chr(0))
                GetTrim = SendMessage(hWnd, CB_GETLBTEXT, i, GetString)
                get_text = GetString
            End If
        End If

    End Function

    Function get_control_handle(ByVal name As String, ByVal type As eControlType, ByVal session As Integer) As IntPtr
        Select Case type
            Case eControlType.CheckBox
                For i As Integer = 0 To w_sessions(session).winWnd(0).cbxWnd.Length - 1
                    If w_sessions(session).winWnd(0).cbxWnd(i).text(0) = name Then
                        Return w_sessions(session).winWnd(0).cbxWnd(i).handle
                    End If
                Next
            Case eControlType.Button
                For i As Integer = 0 To w_sessions(session).winWnd(0).btnWnd.Length - 1
                    If w_sessions(session).winWnd(0).btnWnd(i).text(0) = name Then
                        Return w_sessions(session).winWnd(0).btnWnd(i).handle
                    End If
                Next
            Case eControlType.OptionButton
                For i As Integer = 0 To w_sessions(session).winWnd(0).optWnd.Length - 1
                    If w_sessions(session).winWnd(0).optWnd(i).text(0) = name Then
                        Return w_sessions(session).winWnd(0).optWnd(i).handle
                    End If
                Next
            Case Else
                get_control_handle = get_control_handle
        End Select
    End Function

    Sub select_grid_item(ByVal item As String, ByVal to_clipboard As Boolean, Optional item_count As Integer = 1)
        'This can be improved, currently needs focus and uses sendkeys

        show_status("Selecting " & item & " ")

        Clipboard.Clear()

        Dim cb As String = Nothing, last_cb As String = Nothing, key As String = "{DOWN}"
        'Dim cb As String = Nothing, last_cb As String = Nothing, key As Windows.Forms.Keys = Keys.Down
        dr_navigation = dt_navigation.Select("plant=" & s_plant & " AND window_name='Print Pallet labels' and target_reached=1")

        While Clipboard.GetText = Nothing


            If w_sessions(0).winWnd(0).name = dr_navigation(0)("window_name") Then
                SendMessage(get_control_handle(get_acsis_option("pallet_screen_last_pallet"), eControlType.CheckBox, 0), WM_SETFOCUS, 0, 0)
            Else
                SendMessage(w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_label_to")).handle, WM_SETFOCUS, 0, 0)
            End If
            If Not GetForegroundWindow() = w_sessions(0).winWnd(0).handle Then
                SetForegroundWindow(w_sessions(0).winWnd(0).handle)
                'AppActivate(w_sessions(0).process.Id)
            End If
            SendKeys.SendWait("{TAB}")
            SendKeys.SendWait("{LEFT}")
            SendKeys.SendWait("^(c)")
            'KeyPress(Keys.Tab, gridWnd(0).handle)
            'KeyPress(Keys.Left, gridWnd(0).handle)
            'KeyPress(Keys.ControlKey, gridWnd(0).handle, False)
            'KeyPress(Keys.C, gridWnd(0).handle)
            'KeyPress(Keys.ControlKey, gridWnd(0).handle)
            SendKeys.SendWait("^(c)")
        End While
        Dim atempts As Integer = 0
        Dim c As Integer = 0
        Do While Not cb = item Or Not c = item_count
            If Not cb = Nothing Then
                If last_cb = cb Then
                    If atempts = 20 Then
                        Clipboard.SetText("FAILED")
                        Exit Sub
                    End If
                    If key = "{DOWN}" Then
                        key = "{UP}"
                    Else
                        key = "{DOWN}"
                    End If
                    atempts = atempts + 1
                    'If key = Keys.Down Then
                    '    key = Keys.Up
                    'Else
                    '    key = Keys.Down
                    'End If
                End If

                last_cb = cb
                Clipboard.Clear()
                If Not GetForegroundWindow() = w_sessions(0).winWnd(0).handle Then
                    SetForegroundWindow(w_sessions(0).winWnd(0).handle)
                    'AppActivate(w_sessions(0).process.Id)
                End If
                SendKeys.SendWait(key)
                'KeyPress(key, gridWnd(0).handle)

            End If

            cb = Clipboard.GetText

            If Not GetForegroundWindow() = w_sessions(0).winWnd(0).handle Then
                SetForegroundWindow(w_sessions(0).winWnd(0).handle)
                'AppActivate(w_sessions(0).process.Id)
            End If

            SendKeys.SendWait("^(c)")
            'KeyPress(Keys.ControlKey, gridWnd(0).handle, False)
            'KeyPress(Keys.C, gridWnd(0).handle)
            'KeyPress(Keys.ControlKey, gridWnd(0).handle)

            Dim count As Integer = 0
            Do While Clipboard.GetText = Nothing
                If count > 10 Then
                    If w_sessions(0).winWnd(0).name = dr_navigation(0)("window_name") Then
                        SendMessage(get_control_handle(get_acsis_option("pallet_screen_last_pallet"), eControlType.CheckBox, 0), WM_SETFOCUS, 0, 0)
                    Else
                        SendMessage(w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_label_to")).handle, WM_SETFOCUS, 0, 0)
                    End If
                    If Not GetForegroundWindow() = w_sessions(0).winWnd(0).handle Then
                        SetForegroundWindow(w_sessions(0).winWnd(0).handle)
                        ' AppActivate(w_sessions(0).process.Id)
                    End If
                    SendKeys.SendWait("{TAB}")
                    If Not GetForegroundWindow() = w_sessions(0).winWnd(0).handle Then
                        SetForegroundWindow(w_sessions(0).winWnd(0).handle)
                        'AppActivate(w_sessions(0).process.Id)
                    End If
                    SendKeys.SendWait("{LEFT}")
                    'KeyPress(Keys.Tab, gridWnd(0).handle)
                    'KeyPress(Keys.Left, gridWnd(0).handle)

                    count = 0
                End If
                'KeyPress(Keys.ControlKey, gridWnd(0).handle, False)
                'KeyPress(Keys.C, gridWnd(0).handle)
                'KeyPress(Keys.ControlKey, gridWnd(0).handle)

                If Not GetForegroundWindow() = w_sessions(0).winWnd(0).handle Then
                    SetForegroundWindow(w_sessions(0).winWnd(0).handle)
                    ' AppActivate(w_sessions(0).process.Id)
                End If
                SendKeys.SendWait("^(c)")
                count = count + 1
            Loop
            cb = Clipboard.GetText
            If cb = item Then c = c + 1
        Loop
        If Not GetForegroundWindow() = w_sessions(0).winWnd(0).handle Then
            SetForegroundWindow(w_sessions(0).winWnd(0).handle)
            'AppActivate(w_sessions(0).process.Id)
        End If
        SendKeys.SendWait("{RIGHT}")
        ' KeyPress(Keys.Right, gridWnd(0).handle)
        If to_clipboard Then
            If Not GetForegroundWindow() = w_sessions(0).winWnd(0).handle Then
                SetForegroundWindow(w_sessions(0).winWnd(0).handle)
                'AppActivate(w_sessions(0).process.Id)
            End If

            SendKeys.SendWait("^(c)")
        End If
        frm_status.Close()

    End Sub

    Sub insert_string(ByVal s As String, ByVal handle As IntPtr)
        If s = s_user Then
            s = s
        End If

        SendMessage(handle, WM_SETTEXT, 0&, s)
    End Sub

    Sub send_downtime(ByVal time As String, ByVal type As String)
        select_cb_item(type, 0)

        insert_string("0000", w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_downtime_start")).handle)
        w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_downtime_start")).text(0) = "0000"
        insert_string(time, w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_downtime_end")).handle)
        Do While IsWindowEnabled(get_control_handle(get_acsis_option("production_downtime_confirm"), eControlType.Button, 0)) = False
            SendMessage(w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_downtime_start")).handle, WM_SETFOCUS, 0, 0)
            SendMessage(w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_downtime_end")).handle, WM_SETFOCUS, 0, 0)
        Loop
        Do While Not w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_downtime_start")).text(0) = Nothing
            get_controls(0, True, False, True)
            Sleep(500)
            click_button(get_acsis_option("production_downtime_confirm"), False, 0)
        Loop
        click_button(get_acsis_option("production_downtime_confirm"), False, 0)
        click_button(get_button_back(0), False, 0)
    End Sub

    Sub click_button(ByVal name As String, ByVal logon As Boolean, ByVal session As Integer)

        Dim tWnd As IntPtr, enabled As Boolean
        get_controls(0, False, False, True)
        'get_controls(get_window_controls.handle)
        For Each btn In w_sessions(session).winWnd(0).btnWnd
            If btn.text(0) = name Then
                tWnd = btn.handle
                enabled = btn.enabled
            End If
        Next

        If enabled Then
            SendMessage(tWnd, WM_ACTIVATE, 0, 0)
            SendMessage(tWnd, WM_ACTIVATE, 1, 0)
            PostMessage(tWnd, BM_CLICK, 0, 0)

            w_sessions(session).process.WaitForInputIdle()
            If logon Then
                Dim found As Boolean = True
                While found
                    For i = 1 To 2
                        Dim s As String = Replace(dt_symple_session(0)("logon"), "#", 1)
                        tWnd = FindWindowByCaption(0, s)
                        If tWnd = 0 Then
                            s = Replace(dt_symple_session(0)("logon"), "#", 2)
                            If tWnd = 0 Then
                                found = False
                            Else
                                found = True
                            End If
                        Else
                            found = True
                        End If
                        s_current_process = 0
                    Next
                End While
            End If
            get_controls(0, False, False, True)

        End If

        popup_check(True, 0)
        get_controls(0, False, False, True)

    End Sub

    Function get_material_desc(ByVal material_number As Long)

        s_check_pop = True
        open_window(eWindowTarget.WarehouseReceiptsOther, 0)

        show_status("Retrieving Material information")

        insert_string(material_number, w_sessions(0).winWnd(0).tbWnd(get_acsis_option("rts_screen_mat_num")).handle)
        While get_text(w_sessions(0).winWnd(0).tbWnd(get_acsis_option("rts_screen_desc")).handle, False) = Nothing
            popup_check(True, 0)

            SendMessage(w_sessions(0).winWnd(0).tbWnd(get_acsis_option("rts_screen_mat_num")).handle, WM_SETFOCUS, 0, 0)
            SendMessage(w_sessions(0).winWnd(0).tbWnd(get_acsis_option("rts_screen_desc")).handle, WM_SETFOCUS, 0, 0)

            If Not acsis.running(0) Then
                logon()
                open_window(eWindowTarget.WarehouseReceiptsOther, 0)
                insert_string(material_number, w_sessions(0).winWnd(0).tbWnd(get_acsis_option("rts_screen_mat_num")).handle)
            End If
            w_sessions(0).process.WaitForInputIdle()
        End While
        popup_check(True, 0)

        get_material_desc = get_text(w_sessions(0).winWnd(0).tbWnd(get_acsis_option("rts_screen_desc")).handle, False)
        click_button(get_button_back(0), False, 0)
        s_check_pop = False
        frm_status.Close()
    End Function

    Function get_twnd(ByVal w As String) As IntPtr
        get_twnd = FindWindow(vbNullString, w)
    End Function

    Sub logon()
        get_controls(0, True, False, True)
        If Not running(0) Then
            w_sessions(0).process = New Process
            With w_sessions(0).process
                Dim s() As String, file As String = get_file_path("symple")
                s = Split(file, "\")
                Dim n As String = s(s.Length - 1)
                .StartInfo.FileName = file
                .StartInfo.WorkingDirectory = Replace(file, n, Nothing)

                .StartInfo.CreateNoWindow = True
                .Start()
                .WaitForInputIdle()

            End With
            get_controls(0, True, False, True)
            ' get_session(True)
            insert_string(s_user, w_sessions(0).winWnd(0).tbWnd(get_option("login_id")).handle)
            insert_string(s_pass, w_sessions(0).winWnd(0).tbWnd(get_option("login_pass")).handle)
            w_sessions(0).user = s_user
            w_sessions(0).user_name = s_user_name
            w_sessions(0).pass = s_pass
            frm_main.lb_symple_users.Items(0) = s_user
            frm_main.lb_symple_users.SelectedIndex = 0
            click_button(get_option("login_button"), True, 0)
            w_sessions(0).process.WaitForInputIdle()
            get_controls(0, True, False, True)
            SetForegroundWindow(frm_main.Handle)
        End If
    End Sub


    Function get_textbox(ByVal item As eProdbox, ByVal session As Integer, ByVal window_index As Integer) As IntPtr
        'get_controls(0)

        For i As Integer = 0 To w_sessions(session).winWnd(window_index).tbWnd.Length - 1
            If w_sessions(session).winWnd(window_index).tbWnd(i).text(0).Contains("/") Then
                If item = eProdbox.Date_ Then
                    Return i
                End If
            ElseIf w_sessions(session).winWnd(window_index).tbWnd(i).text(0) = Nothing And item = eProdbox.Production_Number _
                Or IsNumeric(w_sessions(session).winWnd(window_index).tbWnd(i).text(0)) And _
                w_sessions(session).winWnd(window_index).tbWnd(i).text(0).Length = 10 Then
                If item = eProdbox.Production_Number Then
                    Return i
                End If
            ElseIf item = eProdbox.User Then
                Return i
            End If
        Next
    End Function
    Function get_option(ByVal item As String) As Object
        get_option = Nothing
        For i As Integer = 0 To dt_acsis_options.Rows.Count - 1
            If dt_acsis_options(i)("item") = item Then
                get_option = dt_acsis_options(i)("val")
                Exit For
            End If
        Next
    End Function
    Function get_selected_option(ByVal session As Integer) As String
        get_selected_option = Nothing
        get_controls(0, False, False, True)
        For i As Integer = 0 To w_sessions(session).winWnd(w_sessions(session).winWnd.Length - 1).optWnd.Length - 1
            If w_sessions(session).winWnd(w_sessions(session).winWnd.Length - 1).optWnd(i).checked Then
                get_selected_option = w_sessions(session).winWnd(w_sessions(session).winWnd.Length - 1).optWnd(i).text(0)
                Exit For
            End If
        Next
        'get_controls(get_window_controls.handle)

    End Function

    Function get_lisbox(ByVal work_center As Boolean, ByVal session As Integer, Optional ByVal win_index As Integer = 0) As Integer
        Dim str() As String = Nothing
        get_lisbox = -1
        For i As Integer = 0 To w_sessions(session).winWnd(win_index).lbWnd.Length - 1
            If w_sessions(session).winWnd(win_index).lbWnd(i).text.Length > 0 Then

                str = Split(w_sessions(session).winWnd(win_index).lbWnd(i).text(0), "   ")
                If str.Length > 1 Then
                    If work_center Then
                        Return i
                    End If
                ElseIf Not work_center Then
                    Return i
                End If
            End If
        Next
    End Function
    Sub get_logon_details(ByVal session As Integer)

        popup_check(True, session)

        dr_navigation = dt_navigation.Select("plant=" & s_plant & _
                                             " AND sub_target='Production Confirmation' and target_reached=1")
        Dim found As Boolean = False
        Dim i As Integer = 0
        For i = w_sessions(session).winWnd.Length - 1 To 0 Step -1
            If w_sessions(session).winWnd(i).name = dr_navigation(0)("window_name") Then
                get_controls(w_sessions(session).winWnd(i).handle, True, False, False)
                found = True
                Exit For
            End If
        Next
        If Not found Then
            open_window(eWindowTarget.ProductionConfirmation, session)
            ' get_controls(0, True, False, False)
            i = 0
        End If
        w_sessions(session).user = UCase(w_sessions(session).winWnd(i).tbWnd(get_textbox(eProdbox.User, session, i)).text(0)) ' wrong fix the export to do pod user
        w_sessions(session).user_name = sql.read("DP_Logins", "Login_Id", "=", "'" & w_sessions(session).user & "'", "User_Name", "symple")
        w_sessions(session).pass = sql.read("DP_Logins", "Login_Id", "=", "'" & w_sessions(session).user & "'", "Password", "symple")
        frm_main.lb_symple_users.Items(session) = w_sessions(session).user
        If session = 0 Then
            s_user = w_sessions(session).user
            s_pass = w_sessions(session).pass
            s_user_name = w_sessions(session).user_name
            frm_main.lb_symple_users.SelectedIndex = 0
        End If
    End Sub
    Sub open_job_screen(ByVal prod_num As Long, ByVal prep As Boolean)

        s_final_done = False
        If s_failed_job And s_popup = 4 Then
            s_popup = 0
            s_failed_job = False
            'Exit Sub
        End If
        popup_check(True, 0)

        Dim min_entered As Boolean
        If Not IsDBNull(sql.read("db_schedule", "prod_num", "=", prod_num, "min_entered")) Then
            min_entered = sql.read("db_schedule", "prod_num", "=", prod_num, "min_entered")
        End If
        'check for window already open
        dr_navigation = dt_navigation.Select("target='Production Entry' AND plant=" & s_plant)
        Dim window_name As String = dr_navigation(0)("window_name")
        dr_navigation = dt_navigation.Select("target='Production Confirmation' AND target_reached=1 AND plant=" & s_plant)
        Dim window_target As String = dr_navigation(0)("window_name")
        get_controls(0, False, False, True)
        Dim prod_entry As Boolean = False

        For i As Integer = 0 To w_sessions(0).winWnd.Length - 1
            If w_sessions(0).winWnd(i).name = window_name Then

                For w As Integer = 0 To w_sessions(0).winWnd.Length - 1
                    If w_sessions(0).winWnd(w).name = window_target Then
                        If w_sessions(0).winWnd(w).tbWnd(get_textbox(eProdbox.Production_Number, 0, w)).text(0) = prod_num Then
                            'check for correct WC and shift
                            prod_entry = True
                            If i > 0 Then
                                click_button(get_button_back(0), False, 0)
                                Exit Sub
                            ElseIf Not get_text(w_sessions(0).winWnd(w).lbWnd(get_lisbox(False, 0, w)).handle, True) = s_shift Then
                                click_button(get_button_back(0), False, 0)
                                Exit For
                            ElseIf Not prep Then
                                If Not SendMessage(w_sessions(0).winWnd(w).lbWnd(get_lisbox(True, 0, w)).handle, LB_GETCURSEL, 0&, 0&) = _
                                    w_sessions(0).winWnd(w).lbWnd(get_lisbox(True, 0, w)).count Then

                                    click_button(get_button_back(0), False, 0)
                                    Exit For
                                Else
                                    While Not w_sessions(0).winWnd(0).name = window_name
                                        click_button(get_button_back(0), False, 0)
                                    End While

                                    Exit Sub
                                End If
                            Else
                                If has_prep() Then
                                    popup_check(True, 0)
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                Next

                Exit For
            ElseIf w_sessions(0).winWnd(i).name = window_target Then
                popup_check(True, 0)
                click_button(get_button_back(0), False, 0)
                'If Not get_window_controls.tbWnd(get_textbox(eProdbox.Production_Number)).text(0) = prod_num Then
                '    open_window(eWindowTarget.ProductionConfirmation)

                'End If
                'prod_entry = True
                Exit For
            End If
        Next

        If Not prod_entry Then
            open_window(eWindowTarget.ProductionConfirmation, 0)
        End If
        show_status("Opening Job Screen for: " & prod_num & " PMNT: " & prep.ToString)
        SendMessage(w_sessions(0).winWnd(0).tbWnd(get_textbox(eProdbox.Production_Number, 0, 0)).handle, WM_SETFOCUS, 0, 0)
        insert_string(prod_num, w_sessions(0).winWnd(0).tbWnd(get_textbox(eProdbox.Production_Number, 0, 0)).handle)

        'SetForegroundWindow(symple.MainWindowHandle)

        SendMessage(w_sessions(0).winWnd(0).tbWnd(get_textbox(eProdbox.User, 0, 0)).handle, WM_SETFOCUS, 0, 0)

        w_sessions(0).process.WaitForInputIdle()

        popup_check(True, 0)

        show_status("Selecting shift: " & s_shift)

        While w_sessions(0).winWnd(0).lbWnd(0).count = -1
            get_controls(0, True, False, True)
            popup_check(True, 0)
            If s_failed_job Then
                frm_status.Close()
                MsgBox("Failed to open the job in Sym-PLE")
                Exit Sub
            End If
        End While

        If prep Then
            select_item(w_sessions(0).winWnd(0).lbWnd(get_lisbox(True, 0)).handle, 0, True)
        Else
            select_item(w_sessions(0).winWnd(0).lbWnd(get_lisbox(True, 0)).handle, w_sessions(0).winWnd(0).lbWnd(get_lisbox(True, 0)).count, True)
        End If

        show_status("Selecting work center")

        For i As Integer = 0 To w_sessions(0).winWnd(0).lbWnd(get_lisbox(False, 0)).count
            If w_sessions(0).winWnd(0).lbWnd(get_lisbox(False, 0)).text(i) = s_shift Then
                select_item(w_sessions(0).winWnd(0).lbWnd(get_lisbox(False, 0)).handle, i, True)
                Exit For
            End If
        Next

        click_button(get_button_access(Nothing, 0), False, 0)
        'SetForegroundWindow(symple.MainWindowHandle)
        Do While Not w_sessions(0).winWnd(0).name = window_name
            If s_popup = 4 Then
                MsgBox("That is an invalid job!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ERROR!")
                s_popup = 0
                Exit Sub
            End If
            open_job_screen(prod_num, prep)
            If s_failed_job Then Exit Sub

        Loop

        frm_status.Close()
    End Sub
    Function has_prep() As Boolean
        has_prep = False
        get_controls(0, False, False, True)
        dr_navigation = dt_navigation.Select("sub_target='Production Confirmation' AND target_reached=1")
        If dr_navigation.Length = 0 Then
            Exit Function
        End If
        For i As Integer = 0 To w_sessions(0).winWnd.Length - 1
            If w_sessions(0).winWnd(i).name = dr_navigation(0)("window_name") Then
                For l As Integer = 0 To w_sessions(0).winWnd(i).lbWnd.Length - 1
                    Dim s() As String = Split(w_sessions(0).winWnd(i).lbWnd(l).text(0), "   ")
                    If s.Length > 1 Then
                        If w_sessions(0).winWnd(i).lbWnd(l).text.Length = 1 Then
                            Return False
                        Else
                            Return True
                        End If

                    End If
                Next
            End If
        Next


    End Function

    Sub print_ticket(ByVal prod_num, ByVal quantity, ByVal length, ByVal uom, ByVal custom_section, ByVal custom_quantity, ByVal item, ByVal print_from, ByVal print_to)

        open_job_screen(prod_num, False)
        If s_failed_job Then Exit Sub
        show_status("Printing ticket")

        click_button(get_button_access("Label Reprint", 0), False, 0)

        select_cb_item(s_label, 0)
        select_cb_item(s_printer, 0)
        Select Case uom
            Case "ROL"
                insert_string(1, w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_label_quant")).handle)
                select_grid_item("Carton (item) quantity uom", False)
                If Not Clipboard.GetText = "FAILED" Then
                    Clipboard.SetText("ROLL")
                    If Not GetForegroundWindow() = w_sessions(0).winWnd(0).handle Then
                        SetForegroundWindow(w_sessions(0).winWnd(0).handle)
                        'AppActivate(w_sessions(0).process.Id)
                    End If
                    SendKeys.SendWait("^(v)")
                    System.Threading.Thread.Sleep(500)
                End If
            Case "KG", "KM", "IMP", "MU"
                insert_string(quantity, w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_label_quant")).handle)
            Case Else
                insert_string(quantity, w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_label_quant")).handle)

        End Select

        If Not custom_section = Nothing Then

            ' If get_material_info.print_type = "S" Or get_material_info.print_type = "P" Or get_material_info.print_type = "R" Then
            Dim item_count As Integer = 1
            Dim m_info As materialInfo = get_material_info()
            If m_info.material_type = material_type.laminate Then item_count = 2

            select_grid_item(custom_section, False, item_count)
            If Not Clipboard.GetText = "FAILED" Then

                If Not GetForegroundWindow() = w_sessions(0).winWnd(0).handle Then
                    SetForegroundWindow(w_sessions(0).winWnd(0).handle)
                    'AppActivate(w_sessions(0).process.Id)
                End If
                Clipboard.SetText(custom_quantity.ToString)
                SendKeys.SendWait("^(v)")
            End If

            'End If
        End If

        If print_from = 0 Then print_from = item
        If print_to = 0 Then print_to = item

        SetForegroundWindow(frm_main.Handle)
        insert_string(print_from, w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_label_from")).handle)
        insert_string(print_to, w_sessions(0).winWnd(0).tbWnd(get_acsis_option("production_label_to")).handle)
        click_button(get_acsis_option("production_label_confirm"), False, 0)

        frm_status.Close()

    End Sub


    Sub open_window(ByVal window_index As eWindowTarget, ByVal session As Integer)


        Dim window_name As String, target_window As String
        Select Case window_index
            Case 0
                window_name = "Production Confirmation"
            Case 1
                window_name = "Print Pallet Labels"
            Case 2
                window_name = "Stock Overview"
            Case 3
                window_name = "Warehouse Reciepts - Other"
            Case 4
                window_name = "Re-Labeling"
            Case 5
                window_name = "Modify Pallet Label"
            Case Else
                window_name = "s"
        End Select

        dr_navigation = dt_navigation.Select("plant=" & s_plant & " AND sub_target='" & window_name & "' and target_reached=1")
        target_window = dr_navigation(0)("window_name")

        'if  target has handle then exit
        show_status("Opening window: " & target_window)

        get_controls(0, True, False, True)
        ' get_controls(get_window_controls.handle)


        If w_sessions(session).winWnd(0).name = target_window Then Exit Sub
        Dim s As String = Nothing, target As String = Nothing

        dr_navigation = dt_navigation.Select("plant=" & s_plant & " AND window_count=" & w_sessions(session).winWnd.Length)

        Do While w_sessions(session).winWnd.Length > 1 ' close open windows
            If w_sessions(session).winWnd(0).name = window_name Then Exit Sub
            For i As Integer = 0 To dr_navigation.Length - 1
                Dim dr() As DataRow = dt_navigation.Select("sub_target='" & dr_navigation(i)("sub_target") & "'", "step ASC")
            Next
            get_controls(w_sessions(session).winWnd(0).handle, True, False, True)
            dr_navigation = dt_navigation.Select("window_count=" & w_sessions(session).winWnd.Length & _
                         " AND opt_target='" & get_selected_option(session) & "'")

            While dr_navigation.Length = 0
                frm_status.Close()

                MsgBox("You are in a screen that has not been recognised." & vbCr & _
                       "Please exit the current screen in Sym-PLE screen.")
                SetForegroundWindow(w_sessions(session).process.MainWindowHandle)
                MsgBox("Click OK to proceed")

                'get_controls(0, skip_user:=True)
                'get_controls(w_sessions(session).winWnd(0).handle, skip_user:=True)
                If w_sessions(session).winWnd.Length = 1 Then
                    For i As Integer = 0 To w_sessions(session).winWnd(0).optWnd.Length - 1
                        dr_navigation = dt_navigation.Select("window_count=" & w_sessions(session).winWnd.Length & _
                                    " AND opt_target='" & w_sessions(session).winWnd(0).optWnd(i).text(0) & "'")
                        If dr_navigation.Length > 0 Then Exit For
                    Next
                Else
                    dr_navigation = dt_navigation.Select("window_count=" & w_sessions(session).winWnd.Length & _
                                 " AND opt_target='" & get_selected_option(session) & "'")

                End If


            End While

            show_status("Closing current window...")

            click_button(dr_navigation(0)("button_exit"), False, session)

            dr_navigation = dt_navigation.Select("window_count=1 AND opt_target='" & get_selected_option(session) & "'")
            For i As Integer = 0 To dr_navigation.Length - 1
                If Not target = window_name Then
                    target = dr_navigation(i)("sub_target")
                End If
            Next

        Loop
        dr_navigation = dt_navigation.Select("plant=" & s_plant & " AND window_count=1")

        Do While Not target = window_name
            'check if window can be open from this menu
            s = Nothing
            For o As Integer = 0 To w_sessions(session).winWnd(0).optWnd.Length - 1
                If s = Nothing Then
                    s = w_sessions(session).winWnd(0).optWnd(o).text(0)
                Else
                    s = s & "," & w_sessions(session).winWnd(0).optWnd(o).text(0)
                End If
            Next

            For r As Integer = 0 To dr_navigation.Length - 1

                If s.Contains(dr_navigation(r)("opt_1")) And s.Contains(dr_navigation(r)("opt_2")) And _
                   s.Contains(dr_navigation(r)("opt_3")) And window_name = dr_navigation(r)("sub_target") Then

                    target = dr_navigation(r)("sub_target")
                    Exit Do
                End If
            Next
            dr_navigation = dt_navigation.Select("window_count=" & w_sessions(session).winWnd.Length & _
                         " AND opt_target='" & get_selected_option(session) & "'")

            If dr_navigation.Length > 0 Then click_button(dr_navigation(0)("button_exit"), False, session)

        Loop
        dr_navigation = dt_navigation.Select("sub_target='" & window_name & "' AND window_count=1", "step ASC")

        For i As Integer = 0 To dr_navigation.Length - 1
            Dim str As String = dr_navigation(i)("opt_target")
            s = Nothing
            For o As Integer = 0 To w_sessions(session).winWnd(0).optWnd.Length - 1
                If s = Nothing Then
                    s = w_sessions(session).winWnd(0).optWnd(o).text(0)
                Else
                    s = s & "," & w_sessions(session).winWnd(0).optWnd(o).text(0)
                End If
            Next

            If s.Contains(dr_navigation(i)("opt_1")) And s.Contains(dr_navigation(i)("opt_2")) And _
                s.Contains(dr_navigation(i)("opt_3")) Then
                For o As Integer = 0 To w_sessions(session).winWnd(0).optWnd.Length - 1
                    If w_sessions(session).winWnd(0).optWnd(o).text(0) = dr_navigation(i)("opt_target") Then
                        set_check(w_sessions(session).winWnd(0).optWnd(o).handle, True)
                        click_button(dr_navigation(i)("button_access"), False, session)
                        Exit For
                    End If
                Next
            End If

        Next
        frm_status.Close()
        popup_check(True, session)
    End Sub

End Class

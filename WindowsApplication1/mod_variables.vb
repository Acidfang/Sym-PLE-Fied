Imports System.IO
Module mod_variables
    Public port_reader As New Process 'symple(1) As Process,
    Public sw_symple, sw_sql, sw_tables, sw_schedules As New Stopwatch
    Public sql As New cSQL
    Public acsis As New cls_ACSIS
    Public dtie As New DataTable_Export
    Public jobs() As JobInfo
    'Public schedule As DataSet1.printing_scheduleDataTable = New DataSet1TableAdapters.printing_scheduleTableAdapter().GetData
    Public dt_schedules, dt_machines, dt_acsis_options, dt_produced, dt_popups, dt_shifts, _
        dt_downtime_reasons, dt_plates, dt_boxes, dt_restore, dt_navigation, dt_temp, dt_run_data, _
        dt_department, dt_uom, dt_general_options, dt_num_up, dt_symple_session, dt_labels As New DataTable

    Public dr_navigation(), dr_burst(), s_selected_job() As DataRow

    Public s_loading As Boolean = True, s_sql_access As Boolean = False, _
        s_symple_access As Boolean = False, s_downtime_needed As Boolean = False, _
        s_check_pop As Boolean = False, s_extruder As Boolean = False, _
        s_failed_job As Boolean = False, s_dt_edit As Boolean = False, _
        s_restored As Boolean = False, s_final_done As Boolean = False, _
        s_preparing As Boolean = False, s_report_export_all As Boolean = False, _
        s_start_up As Boolean = False, allow_decimal As Boolean = False, s_production As Boolean = False

    Public s_prep_minute_count As Integer = 0

    Public s_old_department, s_department, s_user, s_assistant, s_pass, s_machine, s_shift, s_shift_name, _
        s_label, s_label_rts, s_pallet, s_printer, s_user_name, s_downtime_type, _
        s_pass_plate, s_version, s_port_reader, s_missing_num_up, s_missing_uom, _
        s_machines, s_departments, s_resource_path, s_current_session, s_current_window, _
        s_pdf() As String

    Public s_plant, s_popup, s_roll_index, s_current_process As Integer
    'Public w_sessions(frm_main.cbo_symple_sessions.SelectedIndex).tbWnd(), w_sessions(frm_main.cbo_symple_sessions.SelectedIndex).btnWnd(), w_sessions(frm_main.cbo_symple_sessions.SelectedIndex).cbWnd(), w_sessions(frm_main.cbo_symple_sessions.SelectedIndex).optWnd(), gridWnd(), w_sessions(frm_main.cbo_symple_sessions.SelectedIndex).lbWnd(), w_sessions(frm_main.cbo_symple_sessions.SelectedIndex).cbxWnd(), staticWnd(), w_sessions(frm_main.cbo_symple_sessions.SelectedIndex).winWnd() As acsisControl
    Public w_sessions(1), w_current_session As symple_session
    Public hWnd, cWnd As IntPtr
    Public s_current_job As Long
    Public s_roll_info As item_info
    Public symple_values() As symple_time
    Public xlApp As Object
    Public xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = Nothing
    Public xlWorkBooks As Microsoft.Office.Interop.Excel.Workbooks = Nothing
    Public xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Public xlWorkSheets As Microsoft.Office.Interop.Excel.Sheets = Nothing

End Module

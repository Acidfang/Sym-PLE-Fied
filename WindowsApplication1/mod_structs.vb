Public Module mod_structs

    Public Structure symple_time
        Dim section As String
        Dim quantity As Decimal
        Dim unit As String
        Dim type As String
    End Structure

    Public Structure quantInfo
        Dim uom As String
        Dim quant As Decimal
    End Structure
    Public Structure symple_window
        Dim name As String
        Dim handle As IntPtr
        Dim popup As Boolean
        Dim tbWnd() As acsisControl
        Dim btnWnd() As acsisControl
        Dim cbWnd() As acsisControl
        Dim optWnd() As acsisControl
        Dim gridWnd() As acsisControl
        Dim lbWnd() As acsisControl
        Dim cbxWnd() As acsisControl
        Dim staticWnd() As symple_window
    End Structure
    Public Structure symple_session
        Dim session As String
        Dim checked As Boolean
        Dim user As String
        Dim user_name As String
        Dim pass As String
        Dim process As Process
        'Dim running As Boolean
        Dim winWnd() As symple_window
    End Structure
    Public Structure item_info
        Dim item_id As String
        Dim kg_in As Decimal
        Dim km_in As Decimal
        Dim prod_num As Long
        Dim mat_num As Long
        Dim row_start As Integer
        Dim row_end As Integer
        Dim row_last_item_row As Integer
        Dim num_in As Integer
        Dim used_kg As Decimal
        Dim used_km As Decimal
        Dim remaining As Decimal
        Dim rts As Decimal
        Dim reject As Decimal
        Dim scrap As Decimal
        Dim total_rolls As Integer
        Dim rts_count As Integer
        Dim reject_count As Integer
    End Structure

    Public Structure acsisControl
        Dim handle As IntPtr
        Dim text() As String
        Dim count As Integer
        Dim checked As Boolean
        Dim enabled As Boolean
    End Structure

    Public Structure windowControl
        Dim name As String
        Dim section As String
        Dim window As Boolean
        Dim button As String
        Dim start As Boolean
    End Structure

    Public Structure materialInfo
        Dim bundling_type As String
        Dim colour As String
        Dim folding_type As String
        Dim form As String
        Dim formulation As String
        Dim gauge As String
        Dim length As Decimal
        Dim material_type As material_type
        Dim print_type As String
        Dim seal_type As String
        Dim width As Decimal
        Dim zaw As String
        Dim core_type As String
        Dim core_size As Integer
    End Structure

    Public Structure JobInfo
        Dim prod_num As Long
        Dim mat_num_semi As String
        Dim mat_num_fin As String
        Dim mat_num_in As String
        Dim mat_num_in_1 As String
        Dim mat_num_in_2 As String
        Dim mat_desc_semi As String
        Dim mat_desc_fin As String
        Dim mat_desc_in As String
        Dim mat_desc_in_1 As String
        Dim mat_desc_in_2 As String
        Dim label_desc As String
        Dim cyl_size As String
        Dim inksys As String
        Dim inks As String
        Dim issue_quant As Decimal
        Dim issue_uom As String
        Dim req_quant_1 As Decimal
        Dim req_quant_2 As Decimal
        Dim req_uom_1 As String
        Dim req_uom_2 As String
        Dim machine As String
        Dim date_added As Date
        Dim date_due As Date
        Dim no_up As String
        Dim time_setup As String
        Dim time_run As String
        Dim customer As String
        Dim status As String
        Dim index As String
        Dim MaterialAvail As String
        Dim design As String
        Dim sched_text As String
        Dim BCOMP As String
        Dim time As String
        Dim dept As String
        Dim plant As String
        Dim sales_doc As String
        Dim sales_line As String
        Dim bag_info As String
        Dim stamp_type As String
        Dim issued As Boolean
        Dim del As Boolean
        Dim logo_colour As String
        Dim moved As Boolean
        Dim new_job As Boolean
    End Structure

    Public Structure Print_type
        Dim style As Print_style
        Dim web_paths As Integer
        Dim registered As Boolean
    End Structure
End Module

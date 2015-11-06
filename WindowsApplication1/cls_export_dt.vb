Imports System.Data
Imports System.IO

Class DataTable_Export
    ''' <summary>
    ''' Export DataTable Columns , Rows to file
    ''' </summary>
    ''' <param name="datatable">The datatable to exported from.</param>
    ''' <param name="delimited">string for delimited exported row items</param>
    ''' <param name="exportcolumnsheader">Including columns header with exporting</param>
    ''' <param name="file">The file path to export to.</param>
    Sub export(datatable As DataTable, delimited As String, exportcolumnsheader As Boolean, file As String)
        Dim strFile As New StreamWriter(Environ("USERPROFILE") & "\" & file & ".csv", False, System.Text.Encoding.[Default])
        If exportcolumnsheader Then
            Dim Columns As String = String.Empty
            For Each column As DataColumn In datatable.Columns
                Columns += column.ColumnName + delimited
            Next
            strFile.WriteLine(Columns.Remove(Columns.Length - 1, 1))
        End If

        For Each datarow As DataRow In datatable.Rows
            Dim row As String = String.Empty
            For Each items As Object In datarow.ItemArray
                If IsDBNull(items) Then items = "<NULL>"
                items = Replace(items, "'", Nothing)
                If items = "True" Then items = "1"
                If items = "False" Then items = "0"
                If items = Nothing Then items = "<NULL>"
                If items.ToString.Contains(",") Then
                    row += """" & items.ToString() & """" & delimited
                Else
                    row += items.ToString() & delimited
                End If

            Next
            strFile.WriteLine(row.Remove(row.Length - 1, 1))
        Next
        strFile.Flush()
        strFile.Close()
    End Sub

End Class

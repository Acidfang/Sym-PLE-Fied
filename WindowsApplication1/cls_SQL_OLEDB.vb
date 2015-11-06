Imports System.Data.OleDb
Imports System.IO
Public Class cSQL
    '  Implements IDisposable
    Dim adapter As New OleDbDataAdapter
    Dim cmd As OleDbCommand
    Dim cmdstring As String
    Dim conn As OleDbConnection
    Dim reader As OleDbDataReader
  
    Sub export_tables()
        Exit Sub
        'If frm_status.Visible Then Exit Sub
        'show_status("Creating a backup of database...")
        'Dim local_file As String = Environ("USERPROFILE") & "\Sym-PLE-Fied.mdb"
        'Dim cat As ADOX.Catalog = New ADOX.Catalog()

        'If File.Exists(local_file) Then
        '    File.Delete(local_file)
        '    cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        '                "Data Source=" & local_file & ";" & _
        '                "Jet OLEDB:Engine Type=5")

        '    local_file = Environ("USERPROFILE") & "\Sym-PLE.mdb"
        '    File.Delete(local_file)
        '    cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        '                "Data Source=" & local_file & ";" & _
        '                "Jet OLEDB:Engine Type=5")
        '    'Console.WriteLine("Database Created Successfully")

        '    cat = Nothing
        '    local_file = Environ("USERPROFILE") & "\Sym-PLE-Fied.mdb"
        'End If

        'local_file = local_file
        ''Dim conn2 As OleDbConnection


        'If connect("lean") Then
        '    cmdstring = "Select * from Information_Schema.Tables where table_type = 'base table' and table_name <> 'dtproperties'"
        '    Dim dt As New DataTable
        '    Using conn
        '        adapter = New OleDbDataAdapter(cmdstring, conn)
        '        conn.Open()

        '        adapter.Fill(dt)

        '    End Using

        '    For Each d As DataRow In dt.Rows

        '        Dim t1 As New DataTable
        '        connect("lean")
        '        cmdstring = "SELECT * FROM " & d("TABLE_NAME")

        '        Using conn
        '            adapter = New OleDbDataAdapter(cmdstring, conn)
        '            conn.Open()
        '            'Dim columns As String = Nothing
        '            adapter.AcceptChangesDuringFill = False
        '            adapter.Fill(t1)
        '            'Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & local_file)
        '            'Using conn As New sqlConnection("Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Common;Data Source=localhost")
        '            'Dim cmdSelect As SqlCommand = New SqlCommand("Select FirstName, LastName, BirthDate From Person")
        '            'cmdSelect.Connection = cnn
        '            'Dim ad As New SqlDataAdapter(cmdSelect)
        '            'ad.AcceptChangesDuringFill = False
        '            'ad.Fill(dt)
        '            'End Using

        '            'Insert Data from DataTable into Access database  
        '            Using cnn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & local_file & ";Persist Security Info=False")
        '                'Dim cmdSelect As OleDbCommand = New OleDbCommand("Select FirstName, LastName,BirthDate From Person")
        '                'cmdSelect.Connection = cnn
        '                Dim ad As New OleDbDataAdapter(cmdstring, cnn)
        '                Dim cmdBuilder As New OleDbCommandBuilder(ad)
        '                Dim cmd As OleDbCommand = cmdBuilder.GetInsertCommand()
        '                cmd.Connection = cnn
        '                ad.InsertCommand = cmd
        '                ad.Update(t1)
        '            End Using
        '            'Dim AccessCommand As New System.Data.OleDb.OleDbCommand("SELECT * INTO " & d("TABLE_NAME") & " FROM [" & d("TABLE_NAME") & "] IN '' Provider=SQLOLEDB.1;Password=lmuser;Persist Security Info=True;User ID=lmuser;Initial Catalog=LeanManufacturing;Data Source=WLGSQL01;", AccessConn)

        '            'AccessCommand.ExecuteNonQuery()
        '            'AccessConn.Close()

        '            'conn2 = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        '            '                "Data Source=" & local_file & ";" & _
        '            '                "Jet OLEDB:Engine Type=5")
        '            'Using conn2
        '            '    Dim adapter2 As New OleDbDataAdapter
        '            '    cmdstring = "SELECT * INTO " & d("TABLE_NAME") & " FROM [" & d("TABLE_NAME") & "] IN '' Provider=SQLOLEDB.1;Password=lmuser;Persist Security Info=True;User ID=lmuser;Initial Catalog=LeanManufacturing;Data Source=WLGSQL01;Connect Timeout=5;"
        '            '    adapter2 = New OleDbDataAdapter(cmdstring, conn2)
        '            '    conn2.Open()

        '            '    adapter2.Update(t1)
        '            'End Using
        '        End Using




        '        show_status("Creating a backup of database..." & d("TABLE_NAME"))

        '        dtie.export(t1, ",", True, d("TABLE_NAME"))
        '    Next

        '    dt = New DataTable

        '    connect("symple")
        '    cmdstring = "SELECT * FROM DP_Logins"
        '    Using conn
        '        adapter = New OleDbDataAdapter(cmdstring, conn)
        '        conn.Open()
        '        adapter.Fill(dt)
        '    End Using

        '    show_status("Creating a backup of database...Logins")

        '    dtie.export(dt, ",", True, "DP_Logins")

        '    dt = New DataTable

        '    connect("symple")
        '    cmdstring = "SELECT * FROM DP_Print_Templates"
        '    Using conn
        '        adapter = New OleDbDataAdapter(cmdstring, conn)
        '        conn.Open()
        '        adapter.Fill(dt)
        '    End Using

        '    show_status("Creating a backup of database...Print Templates")

        '    dtie.export(dt, ",", True, "DP_Print_Templates")

        '    dt = New DataTable

        '    connect("symple")
        '    cmdstring = "SELECT * FROM DP_Printers_Installed"
        '    Using conn
        '        adapter = New OleDbDataAdapter(cmdstring, conn)
        '        conn.Open()
        '        adapter.Fill(dt)
        '    End Using

        '    show_status("Creating a backup of database...Sym-PLE Printers")

        '    dtie.export(dt, ",", True, "DP_Printers_Installed")

        '    If sw_tables.IsRunning Then
        '        sw_tables.Reset()
        '    Else
        '        sw_tables.Start()
        '    End If

        'End If
        'frm_status.Close()
    End Sub

    Function get_table_full(ByVal table As String, ByVal online As Boolean, Optional ByVal server As String = "lean") As DataTable
        Dim dt As New DataTable
        If Not connect(server) Then
            table = "[" & table & ".csv]"
        End If

        cmdstring = "SELECT * FROM " & table
        Using conn
            adapter = New OleDbDataAdapter(cmdstring, conn)
            conn.Open()
            adapter.Fill(dt)
        End Using

        If dt.Rows.Count > 0 Then
            Return dt
        Else
            Return Nothing
        End If

    End Function

    Function get_table(ByVal table As String, ByVal column As String, ByVal how As String, ByVal value As String, Optional ByVal server As String = "lean") As DataTable
        If s_loading Then show_status("Retrieving Database information (" & table & ")...")
        ' show_status("Retrieving Database information (" & table & ")...")

        Dim dt As New DataTable
        If Not connect(server) Then
            table = "[" & table & ".csv]"
        End If

        cmdstring = "SELECT * FROM " & table & " WHERE " & column & " " & how & " " & value

        Using conn
            adapter = New OleDbDataAdapter(cmdstring, conn)
            conn.Open()
            adapter.Fill(dt)
        End Using
        If Not s_loading Then frm_status.Close()
        If dt.Rows.Count > 0 Then
            Return dt
        Else
            Return Nothing
        End If

    End Function

    Function connect(ByVal server As String, Optional ByVal bypass As Boolean = False) As Boolean
        Dim hostname As String = Nothing, online As Boolean = False, try_ping As Boolean = False, local_con As Boolean = False
        Select Case server
            Case "symple"
                hostname = "WLGACSIS01"
                conn = New OleDbConnection("Provider=SQLOLEDB.1;Password=UserDP;Persist Security Info=True;User ID=DPUser;Initial Catalog=DataPass;Data Source=" & hostname & ";Connect Timeout=5;")
            Case "lean"
                hostname = "WLGSQL01"
                conn = New OleDbConnection("Provider=SQLOLEDB.1;Password=lmuser;Persist Security Info=True;User ID=lmuser;Initial Catalog=LeanManufacturing;Data Source=" & hostname & ";Connect Timeout=5;")
                'Case "sap"
                '    conn = New OleDbConnection("Driver={IBM DB2 OleDb DRIVER};Database=SAPR3;DBALIAS=SACNBPL;DSN=SACNBPL; UID=dpassnbp; Pwd=nbp1pass;")
        End Select
        'If server = "lean" And s_sql_access Or server = "symple" And s_symple_access Or bypass Then
        Dim output As System.Net.NetworkInformation.PingReply
        Dim pinger As New System.Net.NetworkInformation.Ping

        Select Case server
            Case "lean"
                If sw_sql.IsRunning And sw_sql.Elapsed.Minutes > 30 Then
                    try_ping = True
                ElseIf sw_sql.IsRunning Then
                    local_con = True
                Else
                    try_ping = True
                End If
            Case "symple"

                If sw_symple.IsRunning And sw_symple.Elapsed.Minutes > 30 Then
                    try_ping = True
                ElseIf sw_symple.IsRunning Then
                    local_con = True
                Else
                    try_ping = True
                    'try_ping = False
                End If
        End Select
        If try_ping Then

            Try
                output = pinger.Send(hostname)
                online = True
            Catch ex As Exception
                local_con = True
            End Try
        End If

        If local_con Then conn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Environ("USERPROFILE") & "\;text;HDR=Yes;FMT=Delimited")

        Select Case server
            Case "lean"
                s_sql_access = online
                If sw_sql.IsRunning And online Then
                    sw_sql.Reset()
                ElseIf Not sw_sql.IsRunning And Not online Then
                    sw_sql.Start()
                End If
            Case "symple"
                s_symple_access = online
                If sw_symple.IsRunning And online Then
                    sw_symple.Reset()
                ElseIf Not sw_sql.IsRunning And Not online Then
                    sw_symple.Start()
                End If
        End Select

        'End If
        Return online
    End Function

    Function version(Optional ByVal server As String = "lean") As String

        If connect(server) Then

            cmdstring = "Select @@version"

            Using conn
                cmd = New OleDbCommand(cmdstring, conn)
                conn.Open()
                reader = cmd.ExecuteReader
                version = reader.GetString(0)
            End Using

        Else
            Return Nothing
        End If
    End Function

    Function convert_value(ByVal str As Object) As String

        If IsDBNull(str) Then
            convert_value = str.ToString
        Else
            convert_value = str
        End If

        If Not IsNumeric(convert_value) Then
            convert_value = Replace(str, "'", "''")

            If convert_value = Nothing Then
                convert_value = "NULL"
            ElseIf IsDate(convert_value) Then
                Dim dt As Date = convert_value
                convert_value = "'" & Format(dt, "yyyyMMdd HH:mm") & "'"
                dt = dt
            ElseIf convert_value = "-" Then
                convert_value = "0"
            ElseIf Not Strings.Left(convert_value, 1) = "'" And Not Strings.Right(convert_value, 1) = "'" Then
                convert_value = "'" & convert_value & "'"
            End If
        ElseIf convert_value.Contains(",") Then
            convert_value = "'" & convert_value & "'"
        End If
        Select Case convert_value
            Case "'False'"
                convert_value = 0
            Case "'True'"
                convert_value = 1
        End Select
    End Function

    Sub insert(ByVal table As String, ByVal column As String, ByVal item As String, Optional ByVal server As String = "lean")

        If Not connect(server) Then
            table = "[" & table & ".csv]"
        End If

        cmdstring = "INSERT INTO " & table & " (" & column & ") VALUES (" & item & ")"
        Using conn
            cmd = New OleDbCommand(cmdstring, conn)
            conn.Open()
            cmd.ExecuteNonQuery()
        End Using

    End Sub

    Sub update(ByVal table As String, ByVal update_column As String, ByVal update_item As String, ByVal find_column As String, ByVal how As String, ByVal find_value As String, Optional ByVal server As String = "lean")

        If Not connect(server) Then
            table = "[" & table & ".csv]"
        End If

        cmdstring = "UPDATE " & table & " SET " & update_column & "=" & update_item & " WHERE " & find_column & " " & how & " " & find_value
        Using conn
            cmd = New OleDbCommand(cmdstring, conn)
            conn.Open()
            cmd.ExecuteNonQuery()
        End Using

    End Sub

    Function exists(ByVal table As String, ByVal column As String, ByVal how As String, ByVal item As String, Optional ByVal server As String = "lean") As Boolean

        If Not connect(server) Then
            table = "[" & table & ".csv]"
        End If

        cmdstring = "SELECT * FROM " & table & " WHERE " & column & " " & how & " " & item
        Using conn
            cmd = New OleDbCommand(cmdstring, conn)
            conn.Open()
            reader = cmd.ExecuteReader
            If reader.HasRows Then
                exists = True
            Else
                exists = False
            End If
        End Using
    End Function

    Function sum(ByVal table As String, ByVal tcolumn As String, ByVal column As String, ByVal how As String, ByVal value As String, Optional ByVal server As String = "lean") As Double

        If Not connect(server) Then
            table = "[" & table & ".csv]"
        End If

        cmdstring = "SELECT SUM (" & tcolumn & ") FROM " & table & " WHERE " & column & " " & how & " " & value
        Using conn
            cmd = New OleDbCommand(cmdstring, conn)
            conn.Open()
            Dim str = cmd.ExecuteScalar
            If IsNull(str) Then
                sum = 0
            Else
                sum = str
            End If

        End Using

    End Function

    Function read(ByVal table As String, ByVal column As String, ByVal how As String, ByVal item As String, ByVal target As String, Optional ByVal server As String = "lean")

        If Not connect(server) Then
            table = "[" & table & ".csv]"
        End If
        cmdstring = "SELECT * FROM " & table & " WHERE " & column & " " & how & " " & item
        Using conn
            cmd = New OleDbCommand(cmdstring, conn)
            conn.Open()
            reader = cmd.ExecuteReader
            If reader.HasRows Then
                reader.Read()
                If IsNull(reader.Item(target)) Then
                    read = Nothing
                Else
                    read = reader.Item(target)
                End If
            Else
                read = Nothing
            End If

        End Using

    End Function

    Sub delete(ByVal table As String, ByVal column As String, ByVal how As String, ByVal item As String, Optional ByVal server As String = "lean")

        If Not connect(server) Then
            table = "[" & table & ".csv]"
        End If

        cmdstring = "DELETE FROM " & table & " WHERE " & column & " " & how & " " & item
        Using conn
            cmd = New OleDbCommand(cmdstring, conn)
            conn.Open()
            cmd.ExecuteNonQuery()
        End Using

    End Sub

    Function count(ByVal table As String, ByVal column As String, ByVal how As String, ByVal value As String, Optional ByVal server As String = "lean") As Integer

        If Not connect(server) Then
            table = "[" & table & ".csv]"
        End If

        cmdstring = "SELECT COUNT (*) FROM " & table & " WHERE " & column & " " & how & " " & value
        Using conn
            cmd = New OleDbCommand(cmdstring, conn)
            conn.Open()
            Dim str = cmd.ExecuteScalar
            If IsNull(str) Then
                count = 0
            Else
                count = str
            End If
        End Using

    End Function

    Function get_run_data() As DataTable
        get_run_data = Nothing
        Dim s As String
        Dim ExcelCon As New OleDbConnection
        Dim ExcelAdp As OleDbDataAdapter
        Dim ExcelComm As OleDbCommand
        Try
            ExcelCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & Environ("USERPROFILE") & "\specification master sheet with Aussie.xls;Extended Properties=""Excel 8.0;"""
            ExcelCon.Open()

            ExcelComm = New OleDbCommand("Select * From [Specification Master Sheet$]", ExcelCon)
            ExcelAdp = New OleDbDataAdapter(ExcelComm)
            Dim dt As New DataTable
            ExcelAdp.Fill(dt)

            ExcelCon.Close()
            Return dt
        Catch ex As Exception
            s = ex.Message
        Finally
            ExcelCon = Nothing
            ExcelAdp = Nothing
            ExcelComm = Nothing
        End Try
    End Function
    Function GetDataTypeEnum(ByVal val, ByVal siz)
        GetDataTypeEnum = Nothing
        Select Case val
            'Case 0
            '    Return "adEmpty"
            'Case 2
            '    Return "adSmallInt"
            'Case 3
            '    Return "INT" '"adInteger"
            'Case 4
            '    Return "SINGLE" '"adSingle"
            'Case 5
            '    Return "DOUBLE" '"adDouble"
            'Case 6
            '    Return "CURRENCY" '"adCurrency"
            'Case 7
            '    Return "DATETIME" '"adDate"
            'Case 8
            '    Return "adBSTR"
            'Case 9
            '    Return "adIDispatch"
            'Case 10
            '    Return "adError"
            'Case 11
            '    Return "YESNO" '"adBoolean"
            'Case 12
            '    Return "adVariant"
            'Case 13
            '    Return "adIUnknown"
            'Case 14
            '    Return "adDecimal"
            'Case 16
            '    Return "adTinyInt"
            'Case 17
            '    Return "adUnsignedTinyInt"
            'Case 18
            '    Return "adUnsignedSmallInt"
            'Case 19
            '    Return "adUnsignedInt"
            'Case 20
            '    Return "adBigInt"
            'Case 21
            '    Return "adUnsignedBigInt"
            'Case 64
            '    Return "DATETIME" '"adFileTime"
            'Case 72
            '    Return "GUID" '"adGUID"
            'Case 128
            '    Return "BINARY(" & siz & ")" '"adBinary"
            'Case 129
            '    Return "VARCHAR(" & siz & ")" '"adChar"
            'Case 130
            '    Return "adWChar"
            'Case 131
            '    Return "LONG" '"adNumeric"
            'Case 132
            '    Return "adUserDefined"
            'Case 133
            '    Return "DATETIME" '"adDBDate"
            'Case 134
            '    Return "DATETIME" '"adDBTime"
            'Case 135
            '    Return "DATETIME" '"adDBTimeStamp"
            'Case 136
            '    Return "adChapter"
            'Case 137
            '    Return "adDBFileTime"
            'Case 138
            '    Return "adPropVariant"
            'Case 139
            '    Return "adVarNumeric"
            'Case 200  '"adVarChar"
            '    If siz < 255 Then
            '        Return "VARCHAR(" & siz & ")"
            '    Else
            '        Return "MEMO"
            '    End If
            'Case 201
            '    Return "MEMO DEFAULT ''" '"adLongVarChar"
            'Case 202
            '    Return "adVarWChar"
            'Case 203
            '    Return "LONGBINARY" '"adLongVarWChar"
            'Case 204
            '    Return "LONGBINARY" '"adVarBinary"
            'Case 205
            '    Return "LONGBINARY" '"adLongVarBinary"
            Case "String"

            Case Else
                GetDataTypeEnum = 1
        End Select
    End Function


End Class

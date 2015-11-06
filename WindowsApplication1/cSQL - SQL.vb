'Imports System.Data.SqlClient
'Public Class cSQL

'    Dim conn As Sqlconnection
'    Dim reader As SqlDataReader
'    Dim cmdstring As String
'    Dim cmd As SqlCommand
'    Dim adapter As New SqlDataAdapter

'    Sub export_tables(Optional ByVal server As String = "lean")
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then
'            cmdstring = "Select * from Information_Schema.Tables where table_type = 'base table' and table_name <> 'dtproperties'"
'            Dim dt As New DataTable
'            Using conn
'                adapter = New SqlDataAdapter(cmdstring, conn)
'                conn.Open()
'                adapter.Fill(dt)
'            End Using

'            For Each d As DataRow In dt.Rows
'                Dim t1 As New DataTable
'                connect(server)
'                cmdstring = "SELECT * FROM " & DirectCast(d, DataRow).Item("TABLE_NAME")

'                Using conn
'                    adapter = New SqlDataAdapter(cmdstring, conn)
'                    conn.Open()
'                    adapter.Fill(t1)
'                End Using
'                dtie.export(t1, ",", True, DirectCast(d, DataRow).Item("TABLE_NAME"))
'            Next

'            connect("symple")
'            cmdstring = "SELECT * FROM DP_Logins"
'            Using conn
'                adapter = New SqlDataAdapter(cmdstring, conn)
'                conn.Open()
'                adapter.Fill(dt)
'            End Using
'            dtie.export(dt, ",", True, "DP_Logins")

'            connect("symple")
'            cmdstring = "SELECT * FROM DP_Print_Templates"
'            Using conn
'                adapter = New SqlDataAdapter(cmdstring, conn)
'                conn.Open()
'                adapter.Fill(dt)
'            End Using
'            dtie.export(dt, ",", True, "DP_Print_Templates")

'            connect("symple")
'            cmdstring = "SELECT * FROM DP_Printers_Installed"
'            Using conn
'                adapter = New SqlDataAdapter(cmdstring, conn)
'                conn.Open()
'                adapter.Fill(dt)
'            End Using
'            dtie.export(dt, ",", True, "DP_Printers_Installed")

'            If sw_tables.IsRunning Then
'                sw_tables.Reset()
'            Else
'                sw_tables.Start()
'            End If

'        End If

'    End Sub

'    Function get_table_full(ByVal table As String, ByVal online As Boolean, Optional ByVal server As String = "lean") As DataTable
'        Dim dt As New DataTable
'        If online Then
'            connect(server)
'            CmdString = "SELECT * FROM " & table
'            Using conn
'                adapter = New SqlDataAdapter(cmdstring, conn)
'                conn.Open()
'                adapter.Fill(dt)
'            End Using

'            If dt.Rows.Count > 0 Then
'                Return dt
'            Else
'                Return Nothing
'            End If
'        Else
'            dt = dtie.import(".\" & table, ",", True)
'            If dt.Rows.Count > 0 Then
'                Return dt
'            Else
'                Return Nothing
'            End If
'        End If

'    End Function

'    Function get_table(ByVal table As String, ByVal column As String, ByVal how As String, ByVal value As String, Optional ByVal server As String = "lean") As DataTable
'        Dim dt As New DataTable
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then

'            cmdstring = "SELECT * FROM " & table & " WHERE " & column & how & value

'            Using conn
'                adapter = New SqlDataAdapter(cmdstring, conn)
'                conn.Open()
'                adapter.Fill(dt)
'            End Using

'            If dt.Rows.Count > 0 Then
'                Return dt
'            Else
'                Return Nothing
'            End If
'        Else
'            dt = dtie.import(".\" & table, ",", True)
'            Dim dr() As DataRow = dt.Select(column & how & value)
'            dt = New DataTable
'            If dr.Length > 0 Then
'                dt = dr.CopyToDataTable
'                Return dt
'            Else
'                Return Nothing
'            End If

'            'Return dt
'        End If
'    End Function

'    Sub connect(ByVal server As String, Optional ByVal bypass As Boolean = False)

'        Select Case server
'            Case "symple"
'                conn = New SqlConnection("Server=wlgacsis01;Database=DataPass; User Id=DPUser; Password=UserDP;")
'            Case "lean"
'                conn = New SqlConnection("Server=WLGSQL01;Database=LeanManufacturing; User Id=lmuser; Password=lmuser;")
'                'Case "sap"
'                '    conn = New OdbcConnection("Driver={IBM DB2 ODBC DRIVER};Database=SAPR3;DBALIAS=SACNBPL;DSN=SACNBPL; UID=dpassnbp; Pwd=nbp1pass;")
'        End Select
'        If server = "lean" And sql_access Or server = "symple" And symple_access Or bypass Then
'            Try
'                conn.Open()
'                conn.Close()
'                Select Case server
'                    Case "lean"
'                        sql_access = True
'                        If sw_sql.IsRunning Then
'                            sw_sql.Stop()
'                        End If
'                        Form1.Label27.Text = "SQL OK"
'                    Case "symple"
'                        symple_access = True
'                        If sw_symple.IsRunning Then
'                            sw_symple.Stop()
'                        End If
'                        Form1.Label28.Text = "SYMPLE OK"
'                End Select
'            Catch ex As Exception
'                Select Case server
'                    Case "lean"
'                        sql_access = False
'                        If Not sw_sql.IsRunning Then
'                            sw_sql.Start()
'                        End If
'                        Form1.Label27.Text = ex.ToString
'                    Case "symple"
'                        symple_access = False
'                        If Not sw_symple.IsRunning Then
'                            sw_symple.Start()
'                        End If
'                        Form1.Label28.Text = ex.ToString
'                End Select
'            End Try
'        End If

'    End Sub

'    Function version(Optional ByVal server As String = "lean") As String
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then

'            cmdstring = "Select @@version"

'            Using conn
'                cmd = New SqlCommand(cmdstring, conn)
'                conn.Open()
'                reader = cmd.ExecuteReader
'                version = reader.GetString(0)
'            End Using

'        Else
'            Return Nothing
'        End If
'    End Function

'    Sub insert(ByVal table As String, ByVal column As String, ByVal item As String, Optional ByVal server As String = "lean")
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then

'            cmdstring = "INSERT INTO " & table & " (" & column & ") VALUES (" & item & ")"
'            Using conn
'                cmd = New SqlCommand(cmdstring, conn)
'                conn.Open()
'                cmd.ExecuteNonQuery()
'            End Using

'        Else
'            Dim dt As DataTable = dtie.import(table, ",", True)
'            Dim nr As DataRow = dt.NewRow
'            item = Replace(item, "NULL", Nothing)
'            If column.ToString.Contains(",") Then
'                Dim col() As String = Split(column, ",")
'                Dim itm() As String = Split(item, ",")
'                For i = 0 To col.Length - 1
'                    If itm(i) = "" Then
'                        nr(col(i)) = DBNull.Value
'                    Else
'                        nr(col(i)) = itm(i)
'                    End If
'                Next
'                dt.Rows.Add(nr)
'            Else
'                nr(column) = item
'                dt.Rows.Add(nr)
'            End If
'            dtie.export(dt, ",", True, table)
'        End If
'    End Sub

'    Sub update(ByVal table As String, ByVal column As String, ByVal item As String, ByVal Section As String, ByVal SectionValue As String, Optional ByVal server As String = "lean")
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then

'            cmdstring = "UPDATE " & table & " SET " & column & "=" & item & " WHERE " & Section & "=" & SectionValue
'            Using conn
'                cmd = New SqlCommand(cmdstring, conn)
'                conn.Open()
'                cmd.ExecuteNonQuery()
'            End Using

'        Else
'            Dim dt As DataTable = dtie.import(table, ",", True)
'            dt.Select(Section & "=" & SectionValue)
'            dtie.export(dt, ",", True, table)
'        End If
'    End Sub

'    Function exists(ByVal table As String, ByVal column As String, ByVal item As String, Optional ByVal server As String = "lean") As Boolean
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then

'            cmdstring = "SELECT * FROM " & table & " WHERE " & column & "=" & item
'            Using conn
'                cmd = New SqlCommand(cmdstring, conn)
'                conn.Open()
'                reader = cmd.ExecuteReader
'                If reader.HasRows Then
'                    exists = True
'                Else
'                    exists = False
'                End If
'            End Using


'        Else
'            Dim dt As DataTable = dtie.import(table, ",", True)
'            dt.Select(column & "=" & item)
'        End If
'    End Function

'    Function sum(ByVal table As String, ByVal tcolumn As String, ByVal column As String, ByVal how As String, ByVal value As String, Optional ByVal server As String = "lean") As Double
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then


'            cmdstring = "SELECT SUM (" & tcolumn & ") FROM " & table & " WHERE " & column & how & value
'            Using conn
'                cmd = New SqlCommand(cmdstring, conn)
'                conn.Open()
'                Dim str = cmd.ExecuteScalar
'                If IsNull(str) Then
'                    sum = 0
'                Else
'                    sum = str
'                End If

'            End Using

'        Else
'            Dim s As Double
'            Dim dt As DataTable = dtie.import(table, ",", True)
'            Dim dr() As DataRow = dt.Select(column & how & value)
'            If dr.Length > 0 Then
'                For r As Integer = 1 To dr.Length - 1
'                    If Not IsNull(dr(r).Item(tcolumn)) Then s += dr(r).Item(tcolumn)

'                Next
'            End If

'            Return s
'        End If
'    End Function

'    Function read(ByVal table As String, ByVal column As String, ByVal item As String, ByVal target As String, Optional ByVal server As String = "lean")
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then


'            cmdstring = "SELECT * FROM " & table & " WHERE " & column & "=" & item
'            Using conn
'                cmd = New SqlCommand(cmdstring, conn)
'                conn.Open()
'                reader = cmd.ExecuteReader
'                If reader.HasRows Then
'                    reader.Read()
'                    read = reader(target)
'                Else
'                    read = Nothing

'                End If
'            End Using

'        Else
'            Dim dt As DataTable = dtie.import(table, ",", True)
'            Dim dr() As DataRow = dt.Select(column & "=" & item)
'            If dr.Length > 0 Then
'                Return DirectCast(dr(0), DataRow).Item((target))
'            Else
'                Return Nothing
'            End If
'        End If
'    End Function

'    Sub delete(ByVal table As String, ByVal column As String, ByVal item As String, Optional ByVal how As String = "=", Optional ByVal server As String = "lean")
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then
'            cmdstring = "DELETE FROM " & table & " WHERE " & column & how & item
'            Using conn
'                cmd = New SqlCommand(cmdstring, conn)
'                conn.Open()
'                cmd.ExecuteNonQuery()
'            End Using

'        Else
'            Dim dt As DataTable = dtie.import(table, ",", True)
'            Dim dr() As DataRow = dt.Select(column & how & item)
'            For Each row As DataRow In dr
'                For Each r As DataRow In dt.Rows
'                    If r Is row Then
'                        dt.Rows(dt.Rows.IndexOf(r)).Delete()
'                        Exit For
'                    End If
'                Next

'            Next
'            dtie.export(dt, ",", True, table)
'        End If
'    End Sub

'    Function count(ByVal table As String, ByVal column As String, ByVal how As String, ByVal value As String, Optional ByVal server As String = "lean") As Integer
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then


'            cmdstring = "SELECT COUNT (*) FROM " & table & " WHERE " & column & how & value
'            Using conn
'                cmd = New SqlCommand(cmdstring, conn)
'                conn.Open()
'                Dim str = cmd.ExecuteScalar
'                If IsNull(str) Then
'                    count = 0
'                Else
'                    count = str
'                End If
'            End Using

'        Else
'            Dim dt As DataTable = dtie.import(table, ",", True)
'            Dim dr() As DataRow = dt.Select(column & how & value)
'            Return dr.Length
'        End If
'    End Function

'End Class

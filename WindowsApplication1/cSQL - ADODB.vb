'Public Class cSQL

'    Dim conn As ADODB.Connection
'    Dim rs As New ADODB.Recordset
'    Dim cmdstring As String
'    Dim adapter As New OleDb.OleDbDataAdapter

'    Sub export_tables(Optional ByVal server As String = "lean")
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then
'            cmdstring = "Select * from Information_Schema.Tables where table_type = 'base table' and table_name <> 'dtproperties'"
'            Dim dt As New DataTable

'            conn.Open()
'            rs.Open(cmdstring, conn)
'            adapter.Fill(dt, rs)
'            conn.close()

'            For Each d As DataRow In dt.Rows
'                Dim t1 As New DataTable
'                connect(server)
'                cmdstring = "SELECT * FROM " & DirectCast(d, DataRow).Item("TABLE_NAME")


'                conn.Open()
'                rs.Open(cmdstring, conn)
'                adapter.Fill(t1, rs)
'                conn.close()
'                dtie.export(t1, ",", True, DirectCast(d, DataRow).Item("TABLE_NAME"))
'            Next
'            If server = "symple" And symple_access Then
'                connect("symple")
'                cmdstring = "SELECT * FROM DP_Logins"

'                conn.Open()
'                rs.Open(cmdstring, conn)
'                adapter.Fill(dt, rs)
'                conn.Close()
'                dtie.export(dt, ",", True, "DP_Logins")

'                connect("symple")
'                cmdstring = "SELECT * FROM DP_Print_Templates"

'                conn.Open()
'                rs.Open(cmdstring, conn)
'                adapter.Fill(dt, rs)
'                conn.Close()
'                dtie.export(dt, ",", True, "DP_Print_Templates")

'                connect("symple")
'                cmdstring = "SELECT * FROM DP_Printers_Installed"

'                conn.Open()
'                rs.Open(cmdstring, conn)
'                adapter.Fill(dt, rs)
'                conn.Close()
'                dtie.export(dt, ",", True, "DP_Printers_Installed")
'            End If


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
'            cmdstring = "SELECT * FROM " & table

'            conn.Open()
'            rs.Open(cmdstring, conn)
'            adapter.Fill(dt, rs)
'            conn.close()

'            If dt.Rows.Count > 0 Then
'                Return dt
'            Else
'                Return Nothing
'            End If
'        Else
'            dt = dtie.import(table, ",", True)
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


'            conn.Open()
'            rs.Open(cmdstring, conn)
'            adapter.Fill(dt, rs)
'            conn.close()

'            If dt.Rows.Count > 0 Then
'                Return dt
'            Else
'                Return Nothing
'            End If
'        Else
'            dt = dtie.import(table, ",", True)
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
'        conn = New ADODB.Connection
'        Dim uid, pwd As String
'        Select Case server
'            Case "symple"
'                'conn.ConnectionString = "Driver={SQL Server};Server=WLGACSIS01;Database=DataPass;User ID=DPUser;Password=UserDP;Trusted_Connection=no;"
'                conn.ConnectionString = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=DPUser;Initial Catalog=DataPass;Data Source=WLGACSIS01;"
'                uid = "DPUser"
'                pwd = "UserDP"
'            Case "lean"
'                'conn.ConnectionString = "Driver={SQL Server};Server=WLGSQL01;Database=LeanManufacturing;User ID=lmuser;Password=lmuser;Trusted_Connection=no;"
'                conn.ConnectionString = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=lmuser;Initial Catalog=LeanManufacturing;Data Source=WLGSQL01;"
'                uid = "lmuser"
'                pwd = "lmuser"
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


'            conn.Open()
'            rs.Open(cmdstring, conn)
'            version = rs.GetString
'            conn.Close()

'        Else
'            Return Nothing
'        End If
'    End Function

'    Sub insert(ByVal table As String, ByVal column As String, ByVal item As String, Optional ByVal server As String = "lean")
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then

'            cmdstring = "INSERT INTO " & table & " (" & column & ") VALUES (" & item & ")"

'            conn.Open()
'            rs.Open(cmdstring, conn)
'            conn.close()

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

'            conn.Open()
'            rs.Open(cmdstring, conn)
'            conn.close()

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

'            conn.Open()
'            rs.Open(cmdstring, conn)
'            If rs.EOF Or rs.BOF Then
'                exists = True
'            Else
'                exists = False
'            End If
'            conn.close()


'        Else
'            Dim dt As DataTable = dtie.import(table, ",", True)
'            dt.Select(column & "=" & item)
'        End If
'    End Function

'    Function sum(ByVal table As String, ByVal tcolumn As String, ByVal column As String, ByVal how As String, ByVal value As String, Optional ByVal server As String = "lean") As Double
'        connect(server)
'        If server = "lean" And sql_access Or server = "symple" And symple_access Then
'            rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
'            rs.CursorType = ADODB.CursorTypeEnum.adOpenStatic
'            rs.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic

'            cmdstring = "SELECT SUM (" & tcolumn & ") AS MYTOTAL FROM " & table & " WHERE " & column & how & value

'            conn.Open()
'            rs.Open(cmdstring, conn)
'            If IsNull(rs.Fields("MYTOTAL").Value) Then
'                sum = 0
'            Else
'                sum = rs.Fields("MYTOTAL").Value
'            End If
'            conn.Close()
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
'            rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
'            rs.CursorType = ADODB.CursorTypeEnum.adOpenStatic
'            rs.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic

'            cmdstring = "SELECT * FROM " & table & " WHERE " & column & "=" & item

'            conn.Open()
'            rs.Open(cmdstring, conn)
'            If rs.EOF Or rs.BOF Then
'                read = Nothing
'            Else
'                read = rs.Fields(target).Value
'            End If
'            conn.Close()

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

'            conn.Open()
'            rs.Open(cmdstring, conn)
'            conn.close()

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
'            rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
'            rs.CursorType = ADODB.CursorTypeEnum.adOpenStatic
'            rs.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic

'            cmdstring = "SELECT COUNT (*) FROM " & table & " WHERE " & column & how & value

'            conn.Open()
'            rs.Open(cmdstring, conn)

'            count = rs.GetString
'            conn.close()

'        Else
'            Dim dt As DataTable = dtie.import(table, ",", True)
'            Dim dr() As DataRow = dt.Select(column & how & value)
'            Return dr.Length
'        End If
'    End Function

'End Class

Public Class frm_plate_contents

    Private Sub frm_plate_contents_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        With frm_main.dgv_plates
            TextBox1.Text = .CurrentRow.Cells(0).Value
            TextBox2.Text = .CurrentRow.Cells(1).Value
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            TextBox3.Clear()
            DataGridView1.Rows.Clear()
            NumericUpDown1.Value = 1
            CheckBox1.Checked = False
            If Not IsNull(.CurrentRow.Cells(2).Value) Then
                Me.Text = "Box Content: " & .CurrentRow.Cells(2).Value
            Else
                Me.Text = "Box Content"
            End If

            If Not IsNull(.CurrentRow.Cells(5).Value) Then
                Dim plates() As String = Split(.CurrentRow.Cells(5).Value, ",")
                With DataGridView1
                    If plates.Length > 0 Then
                        For Each plate As String In plates
                            .Rows.Add(plate)
                        Next
                    End If
                    .ClearSelection()
                    .CurrentCell = Nothing
                End With

            End If

        End With
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        With DataGridView1
            If Not .CurrentCell Is Nothing Then
                Button1.Enabled = True
                Button3.Enabled = True
                TextBox3.Text = .CurrentCell.Value
            Else
                Button1.Enabled = False
                Button3.Enabled = False
            End If
        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With DataGridView1
            .Rows.RemoveAt(.CurrentRow.Index)
            .ClearSelection()
            .CurrentCell = Nothing
        End With
        update_plates()
        Button1.Enabled = False
        Button3.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With DataGridView1
            If CheckBox1.Checked Then
                For i As Integer = 0 To NumericUpDown1.Value
                    .Rows.Add()
                    .Rows(.RowCount - 1).Cells(0).Value = CInt(TextBox3.Text) + i
                Next
            Else
                .Rows.Add()
                .Rows(.RowCount - 1).Cells(0).Value = TextBox3.Text
            End If
        End With
        update_plates()
    End Sub

    Sub update_plates()
        With DataGridView1
            Dim str As String = "NULL"
            If .RowCount > 0 Then
                For i = 0 To .RowCount - 1
                    If str = "NULL" Then
                        str = .Rows(i).Cells(0).Value
                    Else
                        str = str & "," & .Rows(i).Cells(0).Value
                    End If
                Next
                str = "'" & str & "'"
            End If
            sql.update("db_plate_location", "plates", str, "plate_num", "=", "'" & TextBox1.Text & "'")
        End With
        frm_main.plate_search()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        NumericUpDown1.Visible = CheckBox1.Checked
        NumericUpDown1.Enabled = CheckBox1.Checked
        Label1.Visible = CheckBox1.Checked
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        CheckBox1.Enabled = False
        If IsNumeric(TextBox3.Text) And Not TextBox3.Text = Nothing Then
            CheckBox1.Enabled = True
        End If
        If TextBox3.Text = Nothing Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        With DataGridView1
            .Rows(.CurrentRow.Index).Cells(0).Value = TextBox3.Text
        End With
        update_plates()
        Button1.Enabled = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub

    Private Sub Label_Click(sender As Object, e As EventArgs) Handles l_print_1.Click, l_print_2.Click, l_print_3.Click, l_print_4.Click
        Dim lbl As Label = CType(sender, Label)
        Button5.Enabled = False
        For Each l In Me.Controls
            If TypeOf l Is Label Then
                If l.Name.Contains("l_print_") And Not l Is lbl Then
                    l.BackColor = Color.White
                End If
            End If
        Next
        If lbl.BackColor = Color.White Then
            lbl.BackColor = Color.Green
            Button5.Enabled = True
        Else
            lbl.BackColor = Color.White
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim c As Integer = 0
        Dim r As Integer = 0
        xlApp = New Microsoft.Office.Interop.Excel.Application
        With xlApp
            .UserControl = True
            .Visible = True
        End With
        xlWorkBook = xlApp.Workbooks.Open(s_resource_path & "box labels.xlsx", [ReadOnly]:=True)
        xlWorkSheets = xlWorkBook.Worksheets
        xlWorkSheet = xlWorkSheets("Contents")

        For Each l In Me.Controls
            If TypeOf l Is Label Then
                If l.Name.Contains("l_print_") And l.BackColor = Color.Green Then
                    Select Case Replace(l.name, "l_print_", Nothing)
                        Case 1
                            '3
                            c = 1
                            r = 13
                        Case 2
                            '1
                            c = 1
                            r = 1
                        Case 3
                            '4
                            c = 6
                            r = 13
                        Case 4
                            '2
                            c = 6
                            r = 1
                    End Select
                    Dim s As String = ConvertToLetter(c) & r & ":" & ConvertToLetter(c + 3) & r + 10
                    With xlWorkSheet.Range(s)
                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    End With

                    Exit For
                End If
            End If
        Next
        xlWorkSheet.Cells(r, c).value = Replace(Me.Text, "Box Content: ", Nothing)
        Dim n As Integer = 1
        Dim row As Integer = r + 1
        With DataGridView1

            For i As Integer = 0 To .RowCount - 1
                If n = 11 Then
                    n = 1
                    c = c + 1
                    row = r + 1
                End If
                xlWorkSheet.Cells(row, c).value = .Rows(i).Cells(0).Value
                row = row + 1
                n = n + 1
            Next
        End With
        xlWorkSheet.PrintOutEx(Copies:=1)
        xlWorkBook.Close(False)
        close_excel()
    End Sub
End Class
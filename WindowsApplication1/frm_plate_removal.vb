
Public Class frm_plate_removal

    Private Sub frm_plate_removal_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' If s_loading Then Exit Sub
        frm_main.plate_search(True)
        dgv_plates.Rows.Clear()
        DataGridView1.Rows.Clear()
        For col As Integer = 1 To frm_main.nud_plate_columns.Value
            For row As Integer = 1 To frm_main.nud_plate_rows.Value
                For boxes As Integer = 1 To frm_main.nud_plate_boxes.Value
                    dgv_plates.Rows.Add(Nothing, Strings.Left(frm_main.cbo_plate_boxsize.Text, 1) & col.ToString.PadLeft(3, "0") & Convert.ToChar(row + 64))
                Next
            Next
        Next
        For r As Integer = 0 To frm_main.dgv_plates.RowCount - 1
            Dim added As Boolean = False
            Dim count As Integer = 0
            Dim plate_num As String = frm_main.dgv_plates.Rows(r).Cells(0).Value.ToString

            If Not IsNull(plate_num) Then
                For i As Integer = 0 To frm_main.dgv_plates.RowCount - 1
                    If plate_num = frm_main.dgv_plates.Rows(i).Cells(0).Value Then
                        count = count + 1
                    End If
                Next
            End If
            If count > 1 Then
                Dim n As Integer = 0
                For new_r As Integer = 0 To dgv_plates.RowCount - 1
                    If dgv_plates.Rows(new_r).Cells(1).Value = frm_main.dgv_plates.Rows(r).Cells(1).Value Then
                        For c As Integer = 0 To frm_main.dgv_plates.Columns.Count - 1
                            dgv_plates.Rows(new_r).Cells(c).Value = frm_main.dgv_plates.Rows(r).Cells(c).Value
                        Next
                        n = n + 1
                        If n = count Then Exit For
                    End If
                Next
            End If
        Next
        For r As Integer = 0 To frm_main.dgv_plates.RowCount - 1
            Dim added As Boolean = False
            Dim plate_num As String = frm_main.dgv_plates.Rows(r).Cells(0).Value.ToString

            If Not IsNull(plate_num) Then
                For new_r As Integer = 0 To dgv_plates.RowCount - 1
                    If IsNull(dgv_plates.Rows(new_r).Cells(0).Value) And frm_main.dgv_plates.Rows(r).Cells(1).Value = dgv_plates.Rows(new_r).Cells(1).Value And Not added Then

                        For c As Integer = 0 To frm_main.dgv_plates.Columns.Count - 1
                            dgv_plates.Rows(new_r).Cells(c).Value = frm_main.dgv_plates.Rows(r).Cells(c).Value
                        Next
                        added = True
                        Exit For
                    ElseIf dgv_plates.Rows(new_r).Cells(0).Value = plate_num Then
                        added = True
                    End If
                Next
                If Not added Then
                    DataGridView1.Rows.Add(frm_main.dgv_plates.Rows(r).Cells(0).Value, frm_main.dgv_plates.Rows(r).Cells(1).Value)
                End If
            End If

        Next

        dgv_plates.ClearSelection()

    End Sub

    ''' <summary>
    ''' structire to hold printed page details
    ''' </summary>
    ''' <remarks></remarks>
    Private Structure pageDetails
        Dim columns As Integer
        Dim rows As Integer
        Dim startCol As Integer
        Dim startRow As Integer
    End Structure
    ''' <summary>
    ''' dictionary to hold printed page details, with index key
    ''' </summary>
    ''' <remarks></remarks>
    Private pages As Dictionary(Of Integer, pageDetails)

    Dim maxPagesWide As Integer
    Dim maxPagesTall As Integer




    ''' <summary>
    ''' shows a PrintPreviewDialog
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ppd As New PrintPreviewDialog
        ppd.Document = PrintDocument1
        ppd.WindowState = FormWindowState.Maximized
        ppd.ShowDialog()
    End Sub

    ''' <summary>
    ''' starts print job
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click

        PrintDocument1.Print()
    End Sub

    ''' <summary>
    ''' the majority of this Sub is calculating printed page ranges
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub PrintDocument1_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        ''this removes the printed page margins
        PrintDocument1.OriginAtMargins = True
        PrintDocument1.DefaultPageSettings.Margins = New Drawing.Printing.Margins(0, 0, 0, 0)

        pages = New Dictionary(Of Integer, pageDetails)

        Dim maxWidth As Integer = CInt(PrintDocument1.DefaultPageSettings.PrintableArea.Width) - 40
        Dim maxHeight As Integer = CInt(PrintDocument1.DefaultPageSettings.PrintableArea.Height) - 40

        Dim pageCounter As Integer = 0
        pages.Add(pageCounter, New pageDetails)

        Dim columnCounter As Integer = 0

        Dim columnSum As Integer = DataGridView1.RowHeadersWidth

        For c As Integer = 0 To DataGridView1.Columns.Count - 1
            If columnSum + DataGridView1.Columns(c).Width < maxWidth Then
                columnSum += DataGridView1.Columns(c).Width
                columnCounter += 1
            Else
                pages(pageCounter) = New pageDetails With {.columns = columnCounter, .rows = 0, .startCol = pages(pageCounter).startCol}
                columnSum = DataGridView1.RowHeadersWidth + DataGridView1.Columns(c).Width
                columnCounter = 1
                pageCounter += 1
                pages.Add(pageCounter, New pageDetails With {.startCol = c})
            End If
            If c = DataGridView1.Columns.Count - 1 Then
                If pages(pageCounter).columns = 0 Then
                    pages(pageCounter) = New pageDetails With {.columns = columnCounter, .rows = 0, .startCol = pages(pageCounter).startCol}
                End If
            End If
        Next

        maxPagesWide = pages.Keys.Max + 1

        pageCounter = 0

        Dim rowCounter As Integer = 0

        Dim rowSum As Integer = DataGridView1.ColumnHeadersHeight

        For r As Integer = 0 To DataGridView1.Rows.Count - 2
            If rowSum + DataGridView1.Rows(r).Height < maxHeight Then
                rowSum += DataGridView1.Rows(r).Height
                rowCounter += 1
            Else
                pages(pageCounter) = New pageDetails With {.columns = pages(pageCounter).columns, .rows = rowCounter, .startCol = pages(pageCounter).startCol, .startRow = pages(pageCounter).startRow}
                For x As Integer = 1 To maxPagesWide - 1
                    pages(pageCounter + x) = New pageDetails With {.columns = pages(pageCounter + x).columns, .rows = rowCounter, .startCol = pages(pageCounter + x).startCol, .startRow = pages(pageCounter).startRow}
                Next

                pageCounter += maxPagesWide
                For x As Integer = 0 To maxPagesWide - 1
                    pages.Add(pageCounter + x, New pageDetails With {.columns = pages(x).columns, .rows = 0, .startCol = pages(x).startCol, .startRow = r})
                Next

                rowSum = DataGridView1.ColumnHeadersHeight + DataGridView1.Rows(r).Height
                rowCounter = 1
            End If
            If r = DataGridView1.Rows.Count - 2 Then
                For x As Integer = 0 To maxPagesWide - 1
                    If pages(pageCounter + x).rows = 0 Then
                        pages(pageCounter + x) = New pageDetails With {.columns = pages(pageCounter + x).columns, .rows = rowCounter, .startCol = pages(pageCounter + x).startCol, .startRow = pages(pageCounter + x).startRow}
                    End If
                Next
            End If
        Next

        maxPagesTall = pages.Count \ maxPagesWide

    End Sub

    ''' <summary>
    ''' this is the actual printing routine.
    ''' using the pagedetails i calculated earlier, it prints a title,
    ''' + as much of the datagridview as will fit on 1 page, then moves to the next page.
    ''' this is setup to be dynamic. try resizing the dgv columns or rows
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' Dim rect As New Rectangle(20, 20, CInt(PrintDocument1.DefaultPageSettings.PrintableArea.Width), 0)
        Dim sf As New StringFormat
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center

        'e.Graphics.DrawString(Label1.Text, Label1.Font, Brushes.Black, rect, sf)

        sf.Alignment = StringAlignment.Near

        Dim startX As Integer = 50
        Dim startY As Integer = 50 'rect.Bottom

        Static startPage As Integer = 0

        For p As Integer = startPage To pages.Count - 1
            Dim cell As New Rectangle(startX, startY, DataGridView1.RowHeadersWidth, DataGridView1.ColumnHeadersHeight)
            ' e.Graphics.FillRectangle(New SolidBrush(SystemColors.ControlLight), cell)
            '  e.Graphics.DrawRectangle(Pens.Black, cell)

            startY += DataGridView1.ColumnHeadersHeight

            For r As Integer = pages(p).startRow To pages(p).startRow + pages(p).rows - 1
                cell = New Rectangle(startX, startY, DataGridView1.RowHeadersWidth, DataGridView1.Rows(r).Height)
                ' e.Graphics.FillRectangle(New SolidBrush(SystemColors.ControlLight), cell)
                'e.Graphics.DrawRectangle(Pens.Black, cell)
                '            e.Graphics.DrawString(DataGridView1.Rows(r).HeaderCell.Value.ToString, DataGridView1.Font, Brushes.Black, cell, sf)
                startY += DataGridView1.Rows(r).Height
            Next

            startX += cell.Width
            startY = 0 'rect.Bottom

            For c As Integer = pages(p).startCol To pages(p).startCol + pages(p).columns - 1
                cell = New Rectangle(startX, startY, DataGridView1.Columns(c).Width, DataGridView1.ColumnHeadersHeight)
                ' e.Graphics.FillRectangle(New SolidBrush(SystemColors.ControlLight), cell)
                ' e.Graphics.DrawRectangle(Pens.Black, cell)
                e.Graphics.DrawString(DataGridView1.Columns(c).HeaderCell.Value.ToString, DataGridView1.Font, Brushes.Black, cell, sf)
                startX += DataGridView1.Columns(c).Width
            Next

            startY = DataGridView1.ColumnHeadersHeight

            For r As Integer = pages(p).startRow To pages(p).startRow + pages(p).rows - 1
                startX = 50 + DataGridView1.RowHeadersWidth
                For c As Integer = pages(p).startCol To pages(p).startCol + pages(p).columns - 1
                    cell = New Rectangle(startX, startY, DataGridView1.Columns(c).Width, DataGridView1.Rows(r).Height)
                    e.Graphics.DrawRectangle(Pens.Black, cell)
                    e.Graphics.DrawString(DataGridView1(c, r).Value.ToString, DataGridView1.Font, Brushes.Black, cell, sf)
                    startX += DataGridView1.Columns(c).Width
                Next
                startY += DataGridView1.Rows(r).Height
            Next

            If p <> pages.Count - 1 Then
                startPage = p + 1
                e.HasMorePages = True
                Return
            Else
                startPage = 0
            End If

        Next

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frm_main.dgv_plates.Rows.Clear()
        For r As Integer = 0 To dgv_plates.RowCount - 1
            Dim pn, loc, desc, check, plates, com As String
            If IsNull(dgv_plates.Rows(r).Cells(0).Value) Then
                pn = Nothing
            Else
                pn = dgv_plates.Rows(r).Cells(0).Value
            End If
            If IsNull(dgv_plates.Rows(r).Cells(1).Value) Then
                loc = Nothing
            Else
                loc = dgv_plates.Rows(r).Cells(1).Value
            End If

            If IsNull(dgv_plates.Rows(r).Cells(2).Value) Then
                desc = Nothing
            Else
                desc = dgv_plates.Rows(r).Cells(2).Value
            End If

            If IsNull(dgv_plates.Rows(r).Cells(3).Value) Then
                com = Nothing
            Else
                com = dgv_plates.Rows(r).Cells(3).Value
            End If

            If IsNull(dgv_plates.Rows(r).Cells(4).Value) Then
                check = Nothing
            Else
                check = dgv_plates.Rows(r).Cells(4).Value
            End If

            If IsNull(dgv_plates.Rows(r).Cells(5).Value) Then
                plates = Nothing
            Else
                plates = dgv_plates.Rows(r).Cells(5).Value
            End If

            frm_main.dgv_plates.Rows.Add(pn, loc, desc, check, plates)
        Next
        frm_main.export_plates()
        sql.update("plate_options", "quant", frm_main.nud_plate_boxes.Value, "box_size", "=", "'" & frm_main.cbo_plate_boxsize.Text & _
                   "' AND plant=" & s_plant)

        sql.update("plate_options", "col", frm_main.nud_plate_columns.Value, "box_size", "=", "'" & frm_main.cbo_plate_boxsize.Text & _
                   "' AND plant=" & s_plant)

        sql.update("plate_options", "row", frm_main.nud_plate_rows.Value, "box_size", "=", "'" & frm_main.cbo_plate_boxsize.Text & _
                   "' AND plant=" & s_plant)

        frm_main.dgv_plates.ClearSelection()
        frm_main.dgv_plates.CurrentCell = Nothing
        Me.Close()
    End Sub
End Class
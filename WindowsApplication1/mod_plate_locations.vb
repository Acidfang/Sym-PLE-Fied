Module mod_plate_locations
    Function find_empty_slot(ByVal dgv As DataGridView, ByVal slots As Integer, ByVal allow_double As Boolean) As String
        Dim loc(0) As String
        Dim found As Boolean = False
        Dim count As Integer = 0
        With dgv
            'get all locations
            For r As Integer = 0 To .RowCount - 1
                If loc(0) = Nothing Then
                    loc(0) = .Rows(r).Cells(1).Value
                Else
                    Dim f As Integer = Array.IndexOf(loc, .Rows(r).Cells(1).Value)
                    If f = -1 Then
                        ReDim Preserve loc(loc.Length)
                        loc(loc.Length - 1) = .Rows(r).Cells(1).Value
                    End If
                End If

            Next
            'get empty slots of locs, find doubles

            For i As Integer = 0 To loc.Length - 1
                count = 0
                For r As Integer = 0 To .RowCount - 1
                    If IsNull(.Rows(r).Cells(0).Value) And .Rows(r).Cells(1).Value = loc(i) Then
                        count = count + 1
                        If count >= slots Then
                            If allow_double Then
                                'return first empty location with slots
                                Return loc(i)
                            Else
                                'find empty slot with no double up plates
                                Dim plates(0) As String
                                Dim double_up As Boolean = False
                                For dr As Integer = 0 To .RowCount - 1
                                    If .Rows(dr).Cells(1).Value = loc(i) And Not IsNull(.Rows(dr).Cells(0).Value) Then
                                        'get list of plates in location
                                        If plates(0) = Nothing Then
                                            plates(0) = .Rows(dr).Cells(0).Value
                                        Else
                                            Dim f As Integer = Array.IndexOf(plates, .Rows(dr).Cells(0).Value)
                                            If f = -1 Then
                                                ReDim Preserve plates(plates.Length)
                                                plates(plates.Length - 1) = .Rows(dr).Cells(0).Value
                                            Else
                                                double_up = True
                                                Exit For
                                            End If

                                        End If
                                    End If
                                Next
                                If Not double_up Then
                                    Return loc(i)
                                End If
                            End If
                        End If
                    End If
                Next
            Next
            Return Nothing

        End With

    End Function

End Module

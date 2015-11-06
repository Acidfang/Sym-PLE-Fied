Public Class frm_number_pad

    'Sub keypress(sender As Object, e As KeyPressEventArgs)


    'End Sub

    Private Sub frm_number_pad_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is Button Then

            End If
        Next
    End Sub
End Class
Imports System
Imports System.Windows.Forms

Public Class MyTextBox
    Inherits TextBox

    Private alreadyFocused As Boolean

    Protected Overrides Sub OnLeave(ByVal e As EventArgs)
        MyBase.OnLeave(e)

        Me.alreadyFocused = False

    End Sub

    Protected Overrides Sub OnGotFocus(ByVal e As EventArgs)
        MyBase.OnGotFocus(e)

        ' Select all text only if the mouse isn't down.
        ' This makes tabbing to the textbox give focus.
        If MouseButtons = MouseButtons.None Then

            Me.SelectAll()
            Me.alreadyFocused = True

        End If

    End Sub

    Protected Overrides Sub OnMouseUp(ByVal mevent As MouseEventArgs)
        MyBase.OnMouseUp(mevent)

        ' Web browsers like Google Chrome select the text on mouse up.
        ' They only do it if the textbox isn't already focused,
        ' and if the s_user hasn't selected all text.
        If Not Me.alreadyFocused AndAlso Me.SelectionLength = 0 Then

            Me.alreadyFocused = True
            Me.SelectAll()

        End If

    End Sub

End Class
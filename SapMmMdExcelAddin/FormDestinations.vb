' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class FormDestinations
    Private Sub Button_OK_Click(sender As Object, e As EventArgs) Handles Button_OK.Click
        If Me.ListBoxDest.SelectedItem Is Nothing Then
            MsgBox("Select a Destination", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
        Else
            DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub ListBoxDest_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxDest.DoubleClick
        DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
End Class
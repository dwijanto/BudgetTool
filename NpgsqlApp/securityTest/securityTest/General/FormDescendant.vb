Public Class FormDescendant
    Inherits FormBaseTest


    Private WithEvents timer1 As New Timer


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Progress1.Value = 0

        timer1.Interval = 100
        timer1.Enabled = True
    End Sub


    Private Sub timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles timer1.Tick
        Progress1.PerformStep()
        If Progress1.Value = Progress1.Maximum Then
            timer1.Enabled = False
        End If
    End Sub
End Class
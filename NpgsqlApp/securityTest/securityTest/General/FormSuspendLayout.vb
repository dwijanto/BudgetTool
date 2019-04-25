Public Class FormSuspendLayout
    Dim myid As Integer
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'FlowLayoutPanel1.SuspendLayout()
        Dim mybutton1 As New Button
        With mybutton1
            .Size = New Size(100, 25)
            myid += 1
            .Text = myid
        End With

        Dim mybutton2 As New Button
        With mybutton2
            .Size = New Size(100, 25)
            myid += 1
            .Text = myid
        End With

        FlowLayoutPanel1.Controls.AddRange(New Control() {mybutton1, mybutton2})        
        'FlowLayoutPanel1.ResumeLayout()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        FlowLayoutPanel1.SuspendLayout()
        For i = FlowLayoutPanel1.Controls.Count - 1 To 0 Step -1
            FlowLayoutPanel1.Controls.Remove(FlowLayoutPanel1.Controls(i))
        Next
        FlowLayoutPanel1.ResumeLayout()
    End Sub


    Private Sub FlowLayoutPanel1_Layout(ByVal sender As Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles FlowLayoutPanel1.Layout

    End Sub
End Class
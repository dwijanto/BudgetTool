Public Class Progress
    Inherits System.Windows.Forms.UserControl


    Public Property Value As Integer
        Get
            Return ProgressBar1.Value
        End Get
        Set(ByVal value As Integer)
            ProgressBar1.Value = value
        End Set
    End Property

    Public Property Maximum As Integer

        Get
            Return ProgressBar1.Maximum
        End Get
        Set(ByVal value As Integer)
            ProgressBar1.Maximum = value
        End Set
    End Property

    Public Property [Step] As Integer

        Get
            Return ProgressBar1.Step
        End Get
        Set(ByVal value As Integer)
            ProgressBar1.Step = value
        End Set
    End Property

    Public Sub PerformStep()
        ProgressBar1.PerformStep()
        updateLabel()
    End Sub

    Private Sub updateLabel()
        Label1.Text = Math.Round(CDec(ProgressBar1.Value * 100) / ProgressBar1.Maximum).ToString
        Label1.Text += "% Done"
    End Sub

End Class

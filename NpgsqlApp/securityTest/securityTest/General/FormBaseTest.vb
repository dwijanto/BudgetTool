Public Class FormBaseTest
    Public Property ButtonText As String
        Get
            Return Me.Button1.Text
        End Get
        Set(ByVal value As String)
            Me.Button1.Text = value
        End Set
    End Property


End Class

Public Class UCHeader
    Dim hideToolbar As HideToolbarDelegate


    Public Sub New(ByVal HideToolbar As HideToolbarDelegate)
        ' This call is required by the designer.
        InitializeComponent()
        Me.hideToolbar = HideToolbar
        Me.BackColor = DJLib.My.MySettings.Default.ToolStripPanelGradientBegin
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        'Me.hideToolbar.Invoke(True)
        Me.hideToolbar(True)
    End Sub
End Class

Public Class ClassRowHeaderImage
    Inherits DataGridViewRowHeaderCell

    Private ic As System.Drawing.Image
    Private gp As System.Drawing.Graphics
    Public Enum CurrentRowState
        Normal
        Modified
        Deleted
        [New]
        Undefine
    End Enum

    Public _currentRowState As CurrentRowState

    Public Sub New()
        _currentRowState = CurrentRowState.Normal
    End Sub

    Public Sub New(ByVal rowstate As CurrentRowState)
        _currentRowState = rowstate
    End Sub

    Protected Overrides Sub Paint(ByVal graphics As System.Drawing.Graphics, ByVal clipBounds As System.Drawing.Rectangle, ByVal cellBounds As System.Drawing.Rectangle, ByVal rowIndex As Integer, ByVal cellState As System.Windows.Forms.DataGridViewElementStates, ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, ByVal cellStyle As System.Windows.Forms.DataGridViewCellStyle, ByVal advancedBorderStyle As System.Windows.Forms.DataGridViewAdvancedBorderStyle, ByVal paintParts As System.Windows.Forms.DataGridViewPaintParts)
        Dim myrect As Rectangle
        Try
            myrect = New Rectangle
            'Don't paint background
            If Not _currentRowState = CurrentRowState.Normal Then
                paintParts = paintParts And Not DataGridViewPaintParts.ContentBackground
            End If
            MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts)

            myrect.X = cellBounds.X + 7
            myrect.Y = cellBounds.Y + 2
            Dim resources = New System.ComponentModel.ComponentResourceManager(GetType(MeHost))
            Select Case _currentRowState
                Case CurrentRowState.Deleted
                    ic = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
                    myrect.Size = ic.Size
                    graphics.DrawImage(ic, myrect)

                Case CurrentRowState.New
                    ic = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
                    ic = New Bitmap(Application.StartupPath & "\images\stock_autofilter.png")
                    ic = My.Resources.Clear_Green_Button16
                    myrect.Size = ic.Size
                    graphics.DrawImage(ic, myrect)
                Case CurrentRowState.Modified
                    ic = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
                    ic = My.Resources.Clear_Green_Button16
                    myrect.Size = ic.Size
                    graphics.DrawImage(ic, myrect)
                Case CurrentRowState.Normal
                    ic = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
                    myrect.Size = ic.Size
                    graphics = Nothing
            End Select
        Catch ex As Exception
            'MessageBox.Show("Error")
        End Try
    End Sub
End Class

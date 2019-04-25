Public Class FormGDI3
    Dim hitrect As Rectangle
    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)

        hitrect = New Rectangle(300, 100, 150, 100)
        e.Graphics.DrawRectangle(Pens.Black, hitrect)

        Dim rect As Rectangle = New Rectangle(10, 30, 250, 50)
        e.Graphics.DrawRectangle(Pens.Black, rect)

        Dim ClippingRegion As Region = New Region(rect)
        e.Graphics.Clip = ClippingRegion
        If Not CheckBox1.Checked Then
            e.Graphics.ResetClip()
        End If
        e.Graphics.DrawString("Clippedadfaadaffsfsf", New Font("Verdana", 36, FontStyle.Bold), Brushes.Black, 10, 30)
        e.Graphics.ResetClip()


        Dim mypath As Drawing2D.GraphicsPath = New Drawing2D.GraphicsPath
        Dim rect2 As Rectangle = New Rectangle(10, 100, 250, 50)
        mypath.AddEllipse(rect2)

        e.Graphics.DrawPath(Pens.Red, mypath)
        Dim clipreg As Region = New Region(mypath)
        e.Graphics.Clip = clipreg

        If Not CheckBox1.Checked Then
            e.Graphics.ResetClip()
        End If
        e.Graphics.DrawString("Clipped", New Font("Arial", 36, FontStyle.Bold), Brushes.Black, 10, 100)
        e.Graphics.ResetClip()

        Dim mypath2 As Drawing2D.GraphicsPath = New Drawing2D.GraphicsPath
        mypath2.AddString("Clipped", New FontFamily("Verdana"), 0, 70, New Point(10, 200), New StringFormat)
        e.Graphics.DrawPath(Pens.Blue, mypath2)

        Dim regclip As Region = New Region(mypath2)
        e.Graphics.Clip = regclip
        If Not CheckBox1.Checked Then
            e.Graphics.ResetClip()
        End If
        For i = 0 To 39
            e.Graphics.DrawEllipse(Pens.Red, 180 - i * 3, 250 - i * 3, i * 6, i * 6)
        Next

        regclip.Dispose()
        mypath2.Dispose()
        mypath.Dispose()
        clipreg.Dispose()
        ClippingRegion.Dispose()


        'e.Graphics.TranslateTransform(180, 60)
        drawrectangle(e.Graphics)
        e.Graphics.TranslateTransform(180, 60)
        drawrectangle(e.Graphics)
        e.Graphics.TranslateTransform(-50, 80)
        drawrectangle(e.Graphics)
        e.Graphics.TranslateTransform(-100, 50)
        drawrectangle(e.Graphics)


    End Sub
    'Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
    '    'MyBase.OnPaintBackground(e)
    '    Dim renderer = New VisualStyles.VisualStyleRenderer(VisualStyles.VisualStyleElement.ExplorerBar.NormalGroupCollapse.Normal)
    '    renderer.DrawBackground(e.Graphics, New Rectangle(0, 0, ClientSize.Width, ClientSize.Height))

    'End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Me.Invalidate()
    End Sub

    Private Sub drawrectangle(ByVal g As Graphics)
        Dim drawingpen As Pen = New Pen(Color.Red, 30)
        g.DrawRectangle(drawingpen, 20, 20, 20, 20)
        drawingpen.Dispose()
    End Sub

    Private Sub FormGDI3_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        MsgBox(e.X & " " & e.Y)
        If e.Button = Windows.Forms.MouseButtons.Left Then

            If hitrect.Contains(e.X, e.Y) Then
                MessageBox.Show("Point inside rect")
            End If
        End If
    End Sub
End Class
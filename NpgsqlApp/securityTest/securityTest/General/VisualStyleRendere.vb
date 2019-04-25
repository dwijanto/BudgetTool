Imports System.Windows.Forms.VisualStyles

Public Class FormVisualStyleRenderer
    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)
        If Application.RenderWithVisualStyles Then
            Dim flags As TextFormatFlags = TextFormatFlags.Bottom Or TextFormatFlags.EndEllipsis
            TextRenderer.DrawText(e.Graphics, "This text drawn with GDI", New Font("Verdana", 20), New Rectangle(10, 10, 300, 50), Color.Black, flags)
            TextRenderer.DrawText(e.Graphics, "►", New Font("Verdana", 14), New Rectangle(100, 100, 50, 50), Color.Black, flags)
            ControlPaint.DrawButton(e.Graphics, New Rectangle(10, 100, 100, 30), CType(ComboBox1.SelectedValue, ButtonState))
            ControlPaint.DrawCheckBox(e.Graphics, New Rectangle(20, 100, 20, 20), CType(ComboBox1.SelectedValue, ButtonState))
            ControlPaint.DrawScrollButton(e.Graphics, New Rectangle(200, 100, 16, 20), ScrollButton.Down, CType(ComboBox1.SelectedValue, ButtonState))
            ControlPaint.DrawMenuGlyph(e.Graphics, New Rectangle(50, 100, 20, 20), CType(ComboBox2.SelectedValue, MenuGlyph))
            CheckBoxRenderer.DrawCheckBox(e.Graphics, New Point(80, 100), CType(ComboBox3.SelectedValue, VisualStyles.CheckBoxState))
            ScrollBarRenderer.DrawArrowButton(e.Graphics, New Rectangle(100, 100, 20, 20), CType(ComboBox4.SelectedValue, VisualStyles.ScrollBarArrowButtonState))
            e.Graphics.DrawPolygon(Pens.Black, New Point() {New Point(10, 10), New Point(10, 18), New Point(14, 14)})
            e.Graphics.FillPolygon(Brushes.Black, New Point() {New Point(10, 10), New Point(10, 18), New Point(14, 14)})
            Dim myimage As Image
            myimage = My.Resources.stock_advanced_filter
            e.Graphics.DrawImage(myimage, 20, 10)
            DrawVisualStyleElementExplorerBarHeaderClose1(e)
            myimage.Dispose()
        Else

        End If
        
    End Sub
    Public Sub DrawVisualStyleElementExplorerBarHeaderClose1(ByVal e As PaintEventArgs)
        If Application.RenderWithVisualStyles Then
            If (VisualStyleRenderer.IsElementDefined(VisualStyleElement.ExplorerBar.HeaderClose.Normal)) Then
                Dim renderer As New VisualStyleRenderer _
                  (VisualStyleElement.ExplorerBar.HeaderClose.Normal)
                Dim rectangle1 As New Rectangle(400, 250, 50, 50)
                renderer.DrawBackground(e.Graphics, rectangle1)
                e.Graphics.DrawString("VisualStyleElement.ExplorerBar.HeaderClose.Normal", Me.Font, Brushes.Black, New Point(400, 200))
                'Else
                '    e.Graphics.DrawString("This element is not defined in the current visual style.", _
                '      Me.Font, Brushes.Black, New Point(400, 200))
            End If
        Else
            e.Graphics.DrawString("This element is not defined in the current visual style.", _
                 Me.Font, Brushes.Black, New Point(400, 200))
        End If

    End Sub

    Private Sub FormVisualStyleRenderer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ComboBox1.DataSource = System.Enum.GetValues(GetType(System.Windows.Forms.ButtonState))
        ComboBox2.DataSource = System.Enum.GetValues(GetType(System.Windows.Forms.MenuGlyph))
        ComboBox3.DataSource = System.Enum.GetValues(GetType(VisualStyles.CheckBoxState))
        ComboBox4.DataSource = System.Enum.GetValues(GetType(VisualStyles.ScrollBarArrowButtonState))
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged, ComboBox2.SelectedIndexChanged, ComboBox4.SelectedIndexChanged
        Me.Invalidate()
    End Sub

    Private Sub HeaderButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles HeaderButton1.Click
        MsgBox("hello")
    End Sub
End Class
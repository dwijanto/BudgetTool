Public Class FormGDI2

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)

        MyBase.OnPaint(e)
        Dim myrect As Rectangle = New Rectangle(0, 0, ClientSize.Width, ClientSize.Height)
        Dim bgbrush As Brush = New Drawing2D.HatchBrush(Drawing2D.HatchStyle.DiagonalBrick, Color.Yellow)
        e.Graphics.FillRectangle(bgbrush, myrect)
        Dim txt = "C:\abcn\daldfj\ldfj\sadkjf\aslkfj\klsdfjl\sdlfkj\lskfj"
        Dim rect As Rectangle = New Rectangle(10, 10, 400, 70)
        Dim rect2 As Rectangle = New Rectangle(10, 100, 100, 50)

        Dim string_format As New StringFormat
        string_format.Trimming = CType(ComboBox1.SelectedValue, StringTrimming) 'StringTrimming.EllipsisCharacter
        string_format.FormatFlags = CType(ComboBox2.SelectedValue, StringFormatFlags)
        Dim semitransparent As Color = Color.FromArgb(CType(TrackBar1.Value, Integer), Color.Blue)
        Dim mybrush As Brush = New SolidBrush(semitransparent)

        'e.Graphics.DrawRectangle(Pens.Blue, rect2)
        e.Graphics.FillRectangle(mybrush, rect2)
        'e.Graphics.DrawString(txt, New Font("Tahoma", 20), Brushes.Brown, rect, string_format)
        e.Graphics.DrawString(txt, New Font("Tahoma", 40), mybrush, rect, string_format)
        string_format.Dispose()
        mybrush.Dispose()
        bgbrush.Dispose()
    End Sub

    Private Sub FormGDI2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.ResizeRedraw = True
        ComboBox1.DataSource = System.Enum.GetValues(GetType(System.Drawing.StringTrimming))
        ComboBox2.DataSource = System.Enum.GetValues(GetType(System.Drawing.StringFormatFlags))
    End Sub

    Private Sub objectChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged, ComboBox2.SelectedIndexChanged, TrackBar1.ValueChanged
        Me.Invalidate()
    End Sub


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        UpdateStyles()
        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class
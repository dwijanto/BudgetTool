Imports System.Drawing.Text
Imports System.ComponentModel

Public Class FormGDI
    Private mysize As Integer
    'Private Sub FormGDI_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
    '    Dim drawingpen As Pen = New Pen(Color.Red, 15)
    '    e.Graphics.DrawArc(drawingpen, 50, 20, 100, 200, 40, 210)
    '    drawingpen.Dispose()

    'End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        'Dim drawingpen As Pen = New Pen(Color.Red, 15)
        'e.Graphics.DrawArc(drawingpen, 50, 20, 100, 200, 40, 210)
        'drawingpen.Dispose()
        Try
            e.Graphics.TextRenderingHint = CType(ComboBox3.SelectedIndex, TextRenderingHint)
            e.Graphics.SmoothingMode = CType(ComboBox2.SelectedIndex, Drawing2D.SmoothingMode)
        Catch ex As Exception
            ToolStripStatusLabel1.Text = ex.Message
        End Try
        Dim myrect As Rectangle = New Rectangle(250, 150, 100, 10)
        If ComboBox1.SelectedIndex <> -1 Then
            Try
                'Int.parse(TextBox1.Text)
                e.Graphics.DrawString(ComboBox1.Text, New Font(ComboBox1.Text, Integer.Parse(TextBox1.Text)), Brushes.Black, 250, 150)
                'e.Graphics.DrawString(ComboBox1.Text, New Font(ComboBox1.Text, Integer.Parse(TextBox1.Text)), Brushes.Black, myrect)
                ToolStripStatusLabel1.Text = ""
            Catch ex As Exception
                ToolStripStatusLabel1.Text = ex.Message
            End Try
        End If

        'Dim DrawingPen As Pen = New Pen(Color.Red, 15)
        'Dim rect As Rectangle = New Rectangle(New Point(0, 0), New Size(mysize, mysize))
        'e.Graphics.DrawRectangle(DrawingPen, rect)
        'DrawingPen.Dispose()
        'System.Threading.Thread.Sleep(10)

        Dim drawingpen As Pen = New Pen(Color.Red, 15)
        drawingpen.Alignment = CType(ComboBox5.SelectedValue, Drawing2D.PenAlignment)
        drawingpen.DashStyle = CType(ComboBox4.SelectedValue, Drawing2D.DashStyle)
        drawingpen.LineJoin = CType(ComboBox6.SelectedValue, Drawing2D.LineJoin)
        Dim singlePen As Pen = New Pen(Color.Black, 15)

        Try
            singlePen.StartCap = CType(ComboBox7.SelectedValue, Drawing2D.LineCap)
            singlePen.EndCap = CType(ComboBox7.SelectedValue, Drawing2D.LineCap)
            singlePen.DashStyle = CType(ComboBox4.SelectedValue, Drawing2D.DashStyle)
        Catch ex As Exception
            ToolStripStatusLabel1.Text = ex.Message
        End Try
        e.Graphics.DrawLine(singlePen, New Point(300, 150), New Point(500, 150))

        e.Graphics.DrawEllipse(drawingpen, New Rectangle(New Point(0, 0), Me.ClientSize))
        Dim rect As Rectangle = New Rectangle(250, 200, 110, 110)
        Dim mybrush As Brush = New Drawing2D.HatchBrush(CType(ComboBox8.SelectedValue, Drawing2D.HatchStyle), Color.Blue)

        'e.Graphics.DrawRectangle(drawingpen, rect)
        'e.Graphics.FillRectangle(mybrush, rect)

        Dim rect2 As Rectangle = New Rectangle(400, 200, 110, 110)
        Dim myGradientBrush As Brush = New Drawing2D.LinearGradientBrush(rect2, Color.Coral, Color.Black, CType(ComboBox9.SelectedValue, Drawing2D.LinearGradientMode))
        'e.Graphics.DrawRectangle(drawingpen, rect2)
        'e.Graphics.FillRectangle(myGradientBrush, rect2)
        Dim rect3 As Rectangle = New Rectangle(500, 200, 110, 110)

        Dim mypath As Drawing2D.GraphicsPath = New Drawing2D.GraphicsPath
        mypath.AddEllipse(500, 200, 110, 110)
        Dim pathbrush As Drawing2D.PathGradientBrush = New Drawing2D.PathGradientBrush(mypath)
        pathbrush.SurroundColors = New Color() {Color.Red}
        pathbrush.CenterColor = Color.Violet
        e.Graphics.FillEllipse(pathbrush, 500, 200, 110, 110)


        mypath.Dispose()
        mybrush.Dispose()
        myGradientBrush.Dispose()
        drawingpen.Dispose()
        singlePen.Dispose()
        MyBase.OnPaint(e)
    End Sub

    Private Sub FormGDI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim fonts As InstalledFontCollection = New InstalledFontCollection
        Dim fname = From myfont In fonts.Families
                    Select myfont.Name
        For Each myname In fname
            ComboBox1.Items.Add(myname)
        Next
        ComboBox1.SelectedIndex = 0

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'For i = 0 To 50
        '    mysize = i
        '    Invalidate()
        '    Update()
        'Next

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        Me.ResizeRedraw = True
        'Dim enumtype As Type = GetType(System.Drawing.Drawing2D.SmoothingMode)
        ComboBox2.DataSource = System.Enum.GetValues(GetType(System.Drawing.Drawing2D.SmoothingMode))
        ComboBox3.DataSource = System.Enum.GetValues(GetType(System.Drawing.Text.TextRenderingHint))
        ComboBox4.DataSource = System.Enum.GetValues(GetType(System.Drawing.Drawing2D.DashStyle))
        ComboBox5.DataSource = System.Enum.GetValues(GetType(System.Drawing.Drawing2D.PenAlignment))
        ComboBox6.DataSource = System.Enum.GetValues(GetType(System.Drawing.Drawing2D.LineJoin))
        ComboBox7.DataSource = System.Enum.GetValues(GetType(System.Drawing.Drawing2D.LineCap))
        Dim abc As System.Array = System.Enum.GetValues(GetType(System.Drawing.Drawing2D.HatchStyle))
        Array.Sort(abc)
        ComboBox8.DataSource = abc 'System.Enum.GetValues(GetType(System.Drawing.Drawing2D.HatchStyle))
  
        ComboBox9.DataSource = System.Enum.GetValues(GetType(System.Drawing.Drawing2D.LinearGradientMode))
        ' Add any initialization after the InitializeComponent() call.

    End Sub



    Private Sub InvalidateMe(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged, ComboBox6.SelectedIndexChanged, ComboBox4.SelectedIndexChanged, ComboBox3.SelectedIndexChanged, ComboBox2.SelectedIndexChanged, ComboBox1.SelectedIndexChanged, TextBox1.TextChanged, ComboBox7.SelectedIndexChanged, ComboBox8.SelectedIndexChanged, ComboBox9.SelectedIndexChanged
        Me.Invalidate()
    End Sub
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If interceptMessage(m) Then

        End If
        MyBase.WndProc(m)
    End Sub

    Private Function interceptMessage(ByVal m As Message) As Boolean
        Return False
    End Function
End Class

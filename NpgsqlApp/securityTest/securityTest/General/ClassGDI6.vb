
Imports System
Imports System.Text
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles



Class ClassGDI6
    Inherits Form

    'Private renderer As VisualStyleRenderer = Nothing
    'Private element As VisualStyleElement = VisualStyleElement.TreeView.Glyph.Closed
    'VisualStyleElement.StartPanel.LogOffButtons.Normal

    Public Sub New()
        'Me.Location = New Point(50, 50)
        'Me.Size = New Size(200, 200)
        'Me.BackColor = SystemColors.ActiveBorder

        'If Application.RenderWithVisualStyles Then
        '    If VisualStyleRenderer.IsElementDefined(element) Then
        '        renderer = New VisualStyleRenderer(element)
        '    End If
        'End If

        Me.ResizeRedraw = True
    End Sub
    'Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
    '    If Application.RenderWithVisualStyles Then
    '        ' Draw the element if the renderer has been set.
    '        If (renderer IsNot Nothing) Then
    '            renderer.DrawBackground(e.Graphics, Me.ClientRectangle)

    '            ' Visual styles are disabled or the element is undefined, 
    '            ' so just draw a message.
    '        Else

    '        End If
    '    Else
    '        Me.Text = "Visual styles are disabled."
    '        TextRenderer.DrawText(e.Graphics, Me.Text, Me.Font, _
    '            New Point(0, 0), Me.ForeColor)
    '    End If

    'End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'ClassGDI6
        '
        Me.ClientSize = New System.Drawing.Size(490, 262)
        Me.Name = "ClassGDI6"
        Me.ResumeLayout(False)

    End Sub

    Private Sub ClassGDI6_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim dirtree As DirectoryTree = New DirectoryTree
        dirtree.Size = New Size(Me.Width - 30, Me.Height - 60)
        dirtree.Location = New Point(5, 5)
        dirtree.Drive = Char.Parse("C")
        Me.Controls.Add(dirtree)
        AddHandler dirtree.DirectorySelected, AddressOf directoryselected
    End Sub

    Private Sub directoryselected(ByVal sender As Object, ByVal e As DirectorySelectedEventArgs)
        MsgBox(e.DirectoryName)
    End Sub

End Class


<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormDescendant
    Inherits HR.FormBaseTest

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Progress1 = New HR.Progress()
        Me.ToolStripContainer1.ContentPanel.SuspendLayout()
        Me.ToolStripContainer1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStripContainer1
        '
        '
        'ToolStripContainer1.ContentPanel
        '
        Me.ToolStripContainer1.ContentPanel.Size = New System.Drawing.Size(685, 408)
        Me.ToolStripContainer1.Size = New System.Drawing.Size(685, 458)
        Me.ToolStripContainer1.TopToolStripPanelVisible = True
        '
        'SplitContainer1
        '
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Progress1)
        Me.SplitContainer1.Size = New System.Drawing.Size(685, 408)
        Me.SplitContainer1.SplitterDistance = 227
        '
        'Progress1
        '
        Me.Progress1.AutoSize = True
        Me.Progress1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Progress1.Location = New System.Drawing.Point(34, 109)
        Me.Progress1.Maximum = 100
        Me.Progress1.Name = "Progress1"
        Me.Progress1.Size = New System.Drawing.Size(222, 42)
        Me.Progress1.Step = 1
        Me.Progress1.TabIndex = 1
        Me.Progress1.Value = 0
        '
        'FormDescendant
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ButtonText = "Hello"
        Me.ClientSize = New System.Drawing.Size(685, 458)
        Me.Name = "FormDescendant"
        Me.Text = "FormDescendant"
        Me.ToolStripContainer1.ContentPanel.ResumeLayout(False)
        Me.ToolStripContainer1.ResumeLayout(False)
        Me.ToolStripContainer1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Progress1 As HR.Progress
End Class

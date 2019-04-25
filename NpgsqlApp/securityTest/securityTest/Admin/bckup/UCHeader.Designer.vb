<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UCHeader
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.ToolStripCustom1 = New DJLib.ToolStripCustom()
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ToolStripCustom1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStripCustom1
        '
        Me.ToolStripCustom1.AutoSize = False
        Me.ToolStripCustom1.ForeColor = System.Drawing.Color.Black
        Me.ToolStripCustom1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabel1, Me.ToolStripButton1})
        Me.ToolStripCustom1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStripCustom1.Name = "ToolStripCustom1"
        Me.ToolStripCustom1.Size = New System.Drawing.Size(170, 28)
        Me.ToolStripCustom1.TabIndex = 1
        Me.ToolStripCustom1.Text = "ToolStripCustom1"
        Me.ToolStripCustom1.ToolStripBorder = System.Drawing.Color.Black
        Me.ToolStripCustom1.ToolStripContentPanelGradientBegin = System.Drawing.Color.Silver
        Me.ToolStripCustom1.ToolStripContentPanelGradientEnd = System.Drawing.Color.LightGray
        Me.ToolStripCustom1.ToolStripDropDownBackground = System.Drawing.Color.LightGray
        Me.ToolStripCustom1.ToolStripForeColor = System.Drawing.Color.Black
        Me.ToolStripCustom1.ToolStripGradientBegin = System.Drawing.Color.Silver
        Me.ToolStripCustom1.ToolStripGradientEnd = System.Drawing.Color.White
        Me.ToolStripCustom1.ToolStripGradientMiddle = System.Drawing.Color.LightGray
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.AutoSize = False
        Me.ToolStripLabel1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(135, 22)
        Me.ToolStripLabel1.Text = "Put Label Here"
        Me.ToolStripLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.AutoSize = False
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = Global.securityTest.My.Resources.Resources.dialog_close
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(22, 20)
        Me.ToolStripButton1.Text = "ToolStripButton2"
        '
        'Panel1
        '
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 28)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(170, 4)
        Me.Panel1.TabIndex = 2
        '
        'UCHeader
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ToolStripCustom1)
        Me.Name = "UCHeader"
        Me.Size = New System.Drawing.Size(170, 62)
        Me.ToolStripCustom1.ResumeLayout(False)
        Me.ToolStripCustom1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolStripCustom1 As DJLib.ToolStripCustom
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents Panel1 As System.Windows.Forms.Panel

End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UCCollapsiblePanel
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
        Me.CollapsiblePanel1 = New HR.CollapsiblePanel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.CollapsiblePanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'CollapsiblePanel1
        '
        Me.CollapsiblePanel1.Collapsed = False
        Me.CollapsiblePanel1.Controls.Add(Me.Button4)
        Me.CollapsiblePanel1.Controls.Add(Me.Button3)
        Me.CollapsiblePanel1.Controls.Add(Me.Button2)
        Me.CollapsiblePanel1.Controls.Add(Me.Button1)
        Me.CollapsiblePanel1.Location = New System.Drawing.Point(3, 3)
        Me.CollapsiblePanel1.Name = "CollapsiblePanel1"
        Me.CollapsiblePanel1.Size = New System.Drawing.Size(209, 146)
        Me.CollapsiblePanel1.TabIndex = 0
        Me.CollapsiblePanel1.Text = "CollapsiblePanel1"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(6, 32)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(198, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(6, 61)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(198, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(6, 90)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(198, 23)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(6, 119)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(198, 23)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "Button4"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'UCCollapsiblePanel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Controls.Add(Me.CollapsiblePanel1)
        Me.Name = "UCCollapsiblePanel"
        Me.Size = New System.Drawing.Size(215, 152)
        Me.CollapsiblePanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CollapsiblePanel1 As HR.CollapsiblePanel
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button

End Class

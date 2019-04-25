<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormVisualStyleRenderer
    Inherits System.Windows.Forms.Form

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
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.ComboBox3 = New System.Windows.Forms.ComboBox()
        Me.ComboBox4 = New System.Windows.Forms.ComboBox()
        Me.CollapsibleHeaderButton1 = New HR.CollapsibleHeaderButton()
        Me.HeaderButton1 = New HR.CollapsibleHeaderButton()
        Me.CollapsiblePanel1 = New HR.CollapsiblePanel()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.CollapsiblePanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(391, 13)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox1.TabIndex = 0
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(391, 40)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox2.TabIndex = 1
        '
        'ComboBox3
        '
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Location = New System.Drawing.Point(391, 67)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox3.TabIndex = 2
        '
        'ComboBox4
        '
        Me.ComboBox4.FormattingEnabled = True
        Me.ComboBox4.Location = New System.Drawing.Point(391, 94)
        Me.ComboBox4.Name = "ComboBox4"
        Me.ComboBox4.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox4.TabIndex = 3
        '
        'CollapsibleHeaderButton1
        '
        Me.CollapsibleHeaderButton1.DialogResult = System.Windows.Forms.DialogResult.None
        Me.CollapsibleHeaderButton1.Location = New System.Drawing.Point(12, 11)
        Me.CollapsibleHeaderButton1.Name = "CollapsibleHeaderButton1"
        Me.CollapsibleHeaderButton1.Size = New System.Drawing.Size(233, 23)
        Me.CollapsibleHeaderButton1.TabIndex = 6
        Me.CollapsibleHeaderButton1.Text = "CollapsibleHeaderButton1"
        '
        'HeaderButton1
        '
        Me.HeaderButton1.DialogResult = System.Windows.Forms.DialogResult.None
        Me.HeaderButton1.Location = New System.Drawing.Point(18, 187)
        Me.HeaderButton1.Name = "HeaderButton1"
        Me.HeaderButton1.Size = New System.Drawing.Size(227, 23)
        Me.HeaderButton1.TabIndex = 5
        Me.HeaderButton1.Text = "HeaderButton1"
        '
        'CollapsiblePanel1
        '
        Me.CollapsiblePanel1.Collapsed = False
        Me.CollapsiblePanel1.Controls.Add(Me.Button4)
        Me.CollapsiblePanel1.Controls.Add(Me.Button3)
        Me.CollapsiblePanel1.Controls.Add(Me.Button2)
        Me.CollapsiblePanel1.Controls.Add(Me.Button1)
        Me.CollapsiblePanel1.Location = New System.Drawing.Point(12, 40)
        Me.CollapsiblePanel1.Name = "CollapsiblePanel1"
        Me.CollapsiblePanel1.Size = New System.Drawing.Size(234, 139)
        Me.CollapsiblePanel1.TabIndex = 4
        Me.CollapsiblePanel1.Text = "CollapsiblePanel1"
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(3, 108)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(230, 23)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "Button4"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(3, 79)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(230, 23)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(3, 52)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(230, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(3, 27)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(230, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FormVisualStyleRenderer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.ClientSize = New System.Drawing.Size(556, 262)
        Me.Controls.Add(Me.CollapsibleHeaderButton1)
        Me.Controls.Add(Me.HeaderButton1)
        Me.Controls.Add(Me.CollapsiblePanel1)
        Me.Controls.Add(Me.ComboBox4)
        Me.Controls.Add(Me.ComboBox3)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.ComboBox1)
        Me.Name = "FormVisualStyleRenderer"
        Me.Text = "VisualStyleRenderer"
        Me.CollapsiblePanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox4 As System.Windows.Forms.ComboBox
    Friend WithEvents CollapsiblePanel1 As HR.CollapsiblePanel
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents HeaderButton1 As HR.CollapsibleHeaderButton
    Friend WithEvents CollapsibleHeaderButton1 As HR.CollapsibleHeaderButton
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BaseSortFilter
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(BaseSortFilter))
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.ToolStripContainerCustom1 = New DJLib.ToolStripContainerCustom()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.ToolStripCustom1 = New DJLib.ToolStripCustom()
        Me.BindingNavigator1 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorAddNewItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel()
        Me.BindingNavigatorDeleteItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox()
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.SaveToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.RefreshToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.SortFilterToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.VerticalHorizontalToolStripButton = New System.Windows.Forms.ToolStripButton()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStripContainerCustom1.ContentPanel.SuspendLayout()
        Me.ToolStripContainerCustom1.LeftToolStripPanel.SuspendLayout()
        Me.ToolStripContainerCustom1.TopToolStripPanel.SuspendLayout()
        Me.ToolStripContainerCustom1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStripContainerCustom1
        '
        '
        'ToolStripContainerCustom1.ContentPanel
        '
        Me.ToolStripContainerCustom1.ContentPanel.Controls.Add(Me.SplitContainer1)
        Me.ToolStripContainerCustom1.ContentPanel.Size = New System.Drawing.Size(586, 456)
        Me.ToolStripContainerCustom1.Dock = System.Windows.Forms.DockStyle.Fill
        '
        'ToolStripContainerCustom1.LeftToolStripPanel
        '
        Me.ToolStripContainerCustom1.LeftToolStripPanel.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.ToolStripContainerCustom1.LeftToolStripPanel.Controls.Add(Me.ToolStripCustom1)
        Me.ToolStripContainerCustom1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStripContainerCustom1.Name = "ToolStripContainerCustom1"
        Me.ToolStripContainerCustom1.Size = New System.Drawing.Size(612, 481)
        Me.ToolStripContainerCustom1.TabIndex = 0
        Me.ToolStripContainerCustom1.Text = "ToolStripContainerCustom1"
        Me.ToolStripContainerCustom1.ToolStripPanelGradientBegin = System.Drawing.SystemColors.ActiveCaption
        Me.ToolStripContainerCustom1.ToolStripPanelGradientEnd = System.Drawing.SystemColors.GradientInactiveCaption
        '
        'ToolStripContainerCustom1.TopToolStripPanel
        '
        Me.ToolStripContainerCustom1.TopToolStripPanel.Controls.Add(Me.BindingNavigator1)
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.SplitContainer1.Size = New System.Drawing.Size(586, 456)
        Me.SplitContainer1.SplitterDistance = 155
        Me.SplitContainer1.TabIndex = 0
        '
        'ToolStripCustom1
        '
        Me.ToolStripCustom1.Dock = System.Windows.Forms.DockStyle.None
        Me.ToolStripCustom1.ForeColor = System.Drawing.Color.Black
        Me.ToolStripCustom1.Location = New System.Drawing.Point(0, 3)
        Me.ToolStripCustom1.Name = "ToolStripCustom1"
        Me.ToolStripCustom1.Size = New System.Drawing.Size(26, 111)
        Me.ToolStripCustom1.TabIndex = 0
        Me.ToolStripCustom1.ToolStripBorder = System.Drawing.Color.Black
        Me.ToolStripCustom1.ToolStripContentPanelGradientBegin = System.Drawing.Color.Silver
        Me.ToolStripCustom1.ToolStripContentPanelGradientEnd = System.Drawing.Color.LightGray
        Me.ToolStripCustom1.ToolStripDropDownBackground = System.Drawing.Color.LightGray
        Me.ToolStripCustom1.ToolStripForeColor = System.Drawing.Color.Black
        Me.ToolStripCustom1.ToolStripGradientBegin = System.Drawing.SystemColors.GradientActiveCaption
        Me.ToolStripCustom1.ToolStripGradientEnd = System.Drawing.SystemColors.GradientInactiveCaption
        Me.ToolStripCustom1.ToolStripGradientMiddle = System.Drawing.SystemColors.GradientActiveCaption
        '
        'BindingNavigator1
        '
        Me.BindingNavigator1.AddNewItem = Nothing
        Me.BindingNavigator1.CountItem = Me.BindingNavigatorCountItem
        Me.BindingNavigator1.DeleteItem = Nothing
        Me.BindingNavigator1.Dock = System.Windows.Forms.DockStyle.None
        Me.BindingNavigator1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2, Me.BindingNavigatorAddNewItem, Me.BindingNavigatorDeleteItem, Me.SaveToolStripButton, Me.RefreshToolStripButton, Me.SortFilterToolStripButton, Me.VerticalHorizontalToolStripButton})
        Me.BindingNavigator1.Location = New System.Drawing.Point(3, 0)
        Me.BindingNavigator1.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.BindingNavigator1.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.BindingNavigator1.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.BindingNavigator1.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.BindingNavigator1.Name = "BindingNavigator1"
        Me.BindingNavigator1.PositionItem = Me.BindingNavigatorPositionItem
        Me.BindingNavigator1.Size = New System.Drawing.Size(378, 25)
        Me.BindingNavigator1.TabIndex = 0
        '
        'BindingNavigatorAddNewItem
        '
        Me.BindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorAddNewItem.Image = CType(resources.GetObject("BindingNavigatorAddNewItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorAddNewItem.Name = "BindingNavigatorAddNewItem"
        Me.BindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorAddNewItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorAddNewItem.Text = "Add new"
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(35, 22)
        Me.BindingNavigatorCountItem.Text = "of {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Total number of items"
        '
        'BindingNavigatorDeleteItem
        '
        Me.BindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorDeleteItem.Image = CType(resources.GetObject("BindingNavigatorDeleteItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorDeleteItem.Name = "BindingNavigatorDeleteItem"
        Me.BindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorDeleteItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorDeleteItem.Text = "Delete"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveFirstItem.Text = "Move first"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMovePreviousItem.Text = "Move previous"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(50, 23)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Current position"
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator1"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveNextItem.Text = "Move next"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveLastItem.Text = "Move last"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator2"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'SaveToolStripButton
        '
        Me.SaveToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.SaveToolStripButton.Image = CType(resources.GetObject("SaveToolStripButton.Image"), System.Drawing.Image)
        Me.SaveToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.SaveToolStripButton.Name = "SaveToolStripButton"
        Me.SaveToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.SaveToolStripButton.Text = "&Save"
        '
        'RefreshToolStripButton
        '
        Me.RefreshToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.RefreshToolStripButton.Image = Global.HR.My.Resources.Resources.view_refresh
        Me.RefreshToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.RefreshToolStripButton.Name = "RefreshToolStripButton"
        Me.RefreshToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.RefreshToolStripButton.Text = "ToolStripButton1"
        '
        'SortFilterToolStripButton
        '
        Me.SortFilterToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.SortFilterToolStripButton.Image = Global.HR.My.Resources.Resources.stock_advanced_filter
        Me.SortFilterToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.SortFilterToolStripButton.Name = "SortFilterToolStripButton"
        Me.SortFilterToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.SortFilterToolStripButton.Text = "ToolStripButton1"
        '
        'VerticalHorizontalToolStripButton
        '
        Me.VerticalHorizontalToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.VerticalHorizontalToolStripButton.Image = Global.HR.My.Resources.Resources.object_flip_vertical
        Me.VerticalHorizontalToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.VerticalHorizontalToolStripButton.Name = "VerticalHorizontalToolStripButton"
        Me.VerticalHorizontalToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.VerticalHorizontalToolStripButton.Text = "ToolStripButton2"
        '
        'BaseSortFilter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(612, 481)
        Me.Controls.Add(Me.ToolStripContainerCustom1)
        Me.Name = "BaseSortFilter"
        Me.Text = "BaseSortFilter"
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStripContainerCustom1.ContentPanel.ResumeLayout(False)
        Me.ToolStripContainerCustom1.LeftToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainerCustom1.LeftToolStripPanel.PerformLayout()
        Me.ToolStripContainerCustom1.TopToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainerCustom1.TopToolStripPanel.PerformLayout()
        Me.ToolStripContainerCustom1.ResumeLayout(False)
        Me.ToolStripContainerCustom1.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator1.ResumeLayout(False)
        Me.BindingNavigator1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolStripContainerCustom1 As DJLib.ToolStripContainerCustom
    Friend WithEvents BindingNavigator1 As System.Windows.Forms.BindingNavigator
    Friend WithEvents BindingNavigatorAddNewItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorCountItem As System.Windows.Forms.ToolStripLabel
    Friend WithEvents BindingNavigatorDeleteItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveFirstItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SaveToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents RefreshToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
    Friend WithEvents SortFilterToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents VerticalHorizontalToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripCustom1 As DJLib.ToolStripCustom
    Protected WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
End Class

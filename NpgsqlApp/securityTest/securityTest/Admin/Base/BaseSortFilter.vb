Imports HR.HelperClass
Public Class BaseSortFilter
    Protected CM As CurrencyManager
    Protected DataSet As New DataSet
    Protected mypanel1 As UCSortTx
    Protected mypanel As UCFilterTx
    Dim WithEvents _DGV As New DataGridView
    'Friend WithEvents DataGridView1 As New DataGridView
    'Friend WithEvents ColAnimalType As System.Windows.Forms.DataGridViewComboBoxColumn
    'Friend WithEvents ColFirstName As System.Windows.Forms.DataGridViewTextBoxColumn
    'Friend WithEvents ColLastName As System.Windows.Forms.DataGridViewTextBoxColumn

    Public Property DGV As DataGridView
        Get
            Return _DGV
        End Get
        Set(ByVal value As DataGridView)
            _DGV = value
        End Set
    End Property


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub LoadData()
        InitObject()
        FillData()
        BindDataSource()
        BindingObject()
        CM = CType(BindingContext(BindingSource1), CurrencyManager)
    End Sub

    Public Overridable Sub FillData()
        'Sample Snippet
        'Dim sqlstr = "select * from expet order by expetid;select exanimalid,exanimaltype from exanimal order by exanimaltype"
        'DbAdapter1.TbgetDataSet(sqlstr, DataSet)
        'DataSet.Tables(0).TableName = "TBExPet"
        'DataSet.Tables(1).TableName = "TBExAnimal"
    End Sub

    Public Overridable Sub InitObject()
        InitDataGrid()
        BindingSource1 = New BindingSource
        'BindingSource2 = New BindingSource
        DataSet = New DataSet
        _DGV.DataSource = BindingSource1
        BindingNavigator1.BindingSource = BindingSource1


    End Sub
    Public Overridable Sub InitDataGrid()
        'Sample Snippet
        'Me.ColAnimalType = New System.Windows.Forms.DataGridViewComboBoxColumn()
        'Me.ColFirstName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        'Me.ColLastName = New System.Windows.Forms.DataGridViewTextBoxColumn()

        ''DataGridViewComboBoxColumn1
        ''
        'With ColAnimalType
        '    .DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.[Nothing]
        '    .HeaderText = "Animal Type"
        '    .Name = "DataGridViewComboBoxColumn1"
        '    .ReadOnly = True
        '    .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '    .Width = 90
        'End With

        ''
        ''DataGridViewTextBoxColumn1
        ''
        'With ColFirstName
        '    .HeaderText = "First Name"
        '    .Name = "DataGridViewTextBoxColumn1"
        '    .ReadOnly = True
        '    .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        '    .Width = 82
        'End With

        ''
        ''DataGridViewTextBoxColumn2
        ''
        'With ColLastName
        '    .AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        '    .HeaderText = "Last Name"
        '    .Name = "DataGridViewTextBoxColumn2"
        '    .ReadOnly = True
        '    .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        'End With

        'DataGridView1.Columns.Add(ColAnimalType)
        'DataGridView1.Columns.Add(ColFirstName)
        'DataGridView1.Columns.Add(ColLastName)

    End Sub
    Public Overridable Sub BindDataSource()
        _DGV.DataSource = BindingSource1
        BindingNavigator1.BindingSource = BindingSource1

        'Sample snippet
        'BindingSource1.DataSource = DataSet.Tables("TBExPet")
        'BindingSource2.DataSource = DataSet.Tables("TBExAnimal")

    End Sub

    Public Overridable Sub BindingObject()
        'Sample Snippet
        'bind DataGridView

        'With ColAnimalType
        '    .DataPropertyName = "exanimalid"
        '    .DataSource = BindingSource2
        '    .DisplayMember = "exanimaltype"
        '    .ValueMember = "exanimalid"
        'End With
        'With ColFirstName
        '    .DataPropertyName = "firstname"
        'End With
        'With ColLastName
        '    .DataPropertyName = "lastname"
        'End With

        'Bind Combobox
        'With ComboBox1
        '    .DataSource = BindingSource2
        '    .DisplayMember = "exanimaltype"
        '    .ValueMember = "exanimalid"
        'End With

        ''DataBinding
        'ComboBox1.DataBindings.Clear()
        'TextBox1.DataBindings.Clear()
        'TextBox2.DataBindings.Clear()
        'TextBox1.DataBindings.Add("Text", BindingSource1, "firstname", True)
        'TextBox2.DataBindings.Add("Text", BindingSource1, "lastname", True)
        'ComboBox1.DataBindings.Add("SelectedValue", BindingSource1, "exanimalid", True)

    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)
        MessageBox.Show(e.Exception.Message)
    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        UpdateRecord()
    End Sub

    Private Sub RefreshToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripButton.Click
        LoadData()
    End Sub

    Private Sub BindingNavigatorDeleteItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorDeleteItem.Click
        DeleteRecordHelper(_DGV, BindingSource1, CM)
    End Sub

    Private Sub UpdateRecord()
        Me.Validate()
        BindingSource1.EndEdit()
        DoUpdate()
    End Sub

    Public Overridable Sub DoUpdate()
        'If DataSet.HasChanges Then
        '    Dim TBPetSaveChanges As SaveChangesRecordDelegate = AddressOf DbAdapter1.TBPetSaveChanges
        '    Dim Reload As ReloadDelegate = AddressOf LoadData
        '    UpdateRecordHelper(DataSet, TBPetSaveChanges, Reload, CM)
        '    BindingSource1.Sort = mypanel1.Sort
        '    BindingSource1.Filter = mypanel.myFilter
        'End If
    End Sub
    Private Sub FormComboBoxManual_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim UpdateRecord As UpdateRecordDelegate = AddressOf Me.UpdateRecord
        CheckFormClosingHelper(DataSet, UpdateRecord, e)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SortFilterToolStripButton.Click
        Call toolstripvisible(ToolStripCustom1.Visible)
    End Sub

    Private Sub LoadToolstrip()
        Dim myaction As HideToolbarDelegate = AddressOf toolstripvisible
        LoadToolstripFilterSort(myaction, _DGV, mypanel1, ToolStripCustom1, mypanel)
    End Sub

    Private Sub toolstripvisible(ByVal p1 As Boolean)
        ToolStripCustom1.Visible = Not (p1)
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VerticalHorizontalToolStripButton.Click
        If SplitContainer1.Orientation = Orientation.Vertical Then
            SplitContainer1.Orientation = Orientation.Horizontal
            VerticalHorizontalToolStripButton.Image = My.Resources.object_flip_vertical
            VerticalHorizontalToolStripButton.ToolTipText = "Vertical View"
        Else
            SplitContainer1.Orientation = Orientation.Vertical
            VerticalHorizontalToolStripButton.Image = My.Resources.object_flip_horizontal
            VerticalHorizontalToolStripButton.ToolTipText = "Horizontal View"
        End If
    End Sub

    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs)
        BindingSource1.EndEdit()
    End Sub

    Private Sub _DGV_CellPainting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles _DGV.CellPainting
        If e.RowIndex = -1 Then
            'erase the cell
            e.Graphics.FillRectangle(New SolidBrush(e.CellStyle.BackColor), e.CellBounds)
            If isfilter(e) Then

                Dim image As Image = My.Resources.stock_advanced_filter
                'Dim rectIcon As Rectangle = New Rectangle(e.CellBounds.X + e.CellBounds.Width - 20, e.CellBounds.Y + 1, 16, e.CellBounds.Height)
                Dim rectIcon As Rectangle = New Rectangle(e.CellBounds.X + e.CellBounds.Width - 20, e.CellBounds.Y + 1, image.Width, image.Height)


                'Dim rectString As Rectangle = New Rectangle(e.CellBounds.X, e.CellBounds.Y + 4, e.CellBounds.Width - 16, e.CellBounds.Height)
                Dim rectString As Rectangle = New Rectangle(e.CellBounds.X, e.CellBounds.Y + 4, e.CellBounds.Width - image.Width, e.CellBounds.Height)
                'Dim gridLinePen As New Pen(Pens.Black)
                e.Graphics.DrawLine(Pens.DarkGray, e.CellBounds.X, e.CellBounds.Height, e.CellBounds.X + e.CellBounds.Width, e.CellBounds.Height)
                'e.Graphics.DrawLine(Pens.Black, e.CellBounds.Width, e.CellBounds.X, e.CellBounds.Width, e.CellBounds.Height)

                e.Graphics.DrawString(e.Value, New Font(e.CellStyle.Font.FontFamily.Name, e.CellStyle.Font.Size, e.CellStyle.Font.Style, e.CellStyle.Font.Unit), New SolidBrush(e.CellStyle.ForeColor), rectString)
                e.Graphics.DrawImage(image, rectIcon)

                e.Handled = True
            Else
                Me.Invalidate()
            End If
        End If
    End Sub
    Private Function isfilter(ByVal e As DataGridViewCellPaintingEventArgs) As Boolean
        Dim myfilter As String = BindingSource1.Filter
        If IsNothing(myfilter) Or e.ColumnIndex < 0 Then
            Return False
        Else
            Dim mycolname As String = e.Value
            If mypanel.ColumnIndexFiltered.Contains(e.ColumnIndex) Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Protected Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles _DGV.ColumnHeaderMouseClick
        columnheadermouseclickHelper(sender, e, BindingSource1)
        mypanel1.Sort = BindingSource1.Sort
    End Sub

    Private Sub DataGridView1_DataError1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles _DGV.DataError
        MessageBox.Show(e.Exception.Message)
    End Sub

    Private Sub FormBaseSortFilter_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadData()
        LoadToolstrip()
    End Sub

    Private Sub BindingNavigatorAddNewItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem.Click
        BindingSource1.AddNew()
    End Sub

  
End Class
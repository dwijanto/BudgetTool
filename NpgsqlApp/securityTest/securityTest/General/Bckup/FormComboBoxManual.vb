Imports securityTest.HelperClass
Public Class FormComboBoxManual
    Private CM As CurrencyManager
    Private DataSet As DataSet
    Private mypanel1 As UCSortTx
    Private mypanel As UCFilterTx
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        LoadData()
        LoadToolstrip()
    End Sub

    Private Sub LoadData()
        InitObject()
        Dim sqlstr = "select * from expet order by expetid;select exanimalid,exanimaltype from exanimal order by exanimaltype"
        DbAdapter1.TbgetDataSet(sqlstr, DataSet)
        DataSet.Tables(0).TableName = "TBExPet"
        DataSet.Tables(1).TableName = "TBExAnimal"
        BindDataSource()
        BindingObject()
        CM = CType(BindingContext(BindingSource1), CurrencyManager)
    End Sub

    Private Sub InitObject()
        BindingSource1 = New BindingSource
        BindingSource2 = New BindingSource
        DataSet = New DataSet
        DataGridView1.DataSource = BindingSource1
        BindingNavigator1.BindingSource = BindingSource1
    End Sub

    Private Sub BindDataSource()
        BindingSource1.DataSource = DataSet.Tables("TBExPet")
        BindingSource2.DataSource = DataSet.Tables("TBExAnimal")
    End Sub

    Private Sub BindingObject()
        'bind DataGridView
        With ColAnimalType
            .DataPropertyName = "exanimalid"
            .DataSource = BindingSource2
            .DisplayMember = "exanimaltype"
            .ValueMember = "exanimalid"
        End With
        With ColFirstName
            .DataPropertyName = "firstname"
        End With
        With ColLastName
            .DataPropertyName = "lastname"
        End With

        'Bind Combobox
        With ComboBox1
            .DataSource = BindingSource2
            .DisplayMember = "exanimaltype"
            .ValueMember = "exanimalid"
        End With

        'DataBinding
        ComboBox1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox1.DataBindings.Add("Text", BindingSource1, "firstname", True)
        TextBox2.DataBindings.Add("Text", BindingSource1, "lastname", True)
        ComboBox1.DataBindings.Add("SelectedValue", BindingSource1, "exanimalid", True)
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
        DeleteRecordHelper(DataGridView1, BindingSource1, CM)
    End Sub

    Private Sub UpdateRecord()
        Me.Validate()
        BindingSource1.EndEdit()

        If DataSet.HasChanges Then
            Dim TBPetSaveChanges As SaveChangesRecordDelegate = AddressOf DbAdapter1.TBPetSaveChanges
            Dim Reload As ReloadDelegate = AddressOf LoadData
            UpdateRecordHelper(DataSet, TBPetSaveChanges, Reload, CM)
            BindingSource1.Sort = mypanel1.Sort
        End If
    End Sub

    Private Sub FormComboBoxManual_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim UpdateRecord As UpdateRecordDelegate = AddressOf Me.UpdateRecord
        CheckFormClosingHelper(DataSet, UpdateRecord, e)
    End Sub

    Private Sub BindingNavigatorAddNewItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem.Click
        BindingSource1.AddNew()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Call toolstripvisible(ToolStrip1.Visible)
    End Sub
    
    Private Sub LoadToolstrip()
        Dim myaction As HideToolbarDelegate = AddressOf toolstripvisible
        LoadToolstripFilterSort(myaction, DataGridView1, mypanel1, ToolStrip1, mypanel)
    End Sub

    Private Sub toolstripvisible(ByVal p1 As Boolean)
        ToolStrip1.Visible = Not (p1)
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If SplitContainer1.Orientation = Orientation.Vertical Then
            SplitContainer1.Orientation = Orientation.Horizontal
            ToolStripButton1.Image = My.Resources.object_flip_vertical
            ToolStripButton1.ToolTipText = "Vertical View"
        Else
            SplitContainer1.Orientation = Orientation.Vertical
            ToolStripButton1.Image = My.Resources.object_flip_horizontal
            ToolStripButton1.ToolTipText = "Horizontal View"
        End If
    End Sub

    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated, TextBox2.Validated, ComboBox1.Validated
        BindingSource1.EndEdit()
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        columnheadermouseclickHelper(sender, e, BindingSource1)
    End Sub

End Class
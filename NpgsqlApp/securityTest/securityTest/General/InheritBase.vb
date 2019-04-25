Imports HR.HelperClass
Public Class InheritBase
    Inherits BaseSortFilter

    Public Overrides Sub InitObject()
        MyBase.InitObject()
        BindingSource2 = New BindingSource
    End Sub


    Public Overrides Sub FillData()
        'MyBase.FillData()
        Dim sqlstr = "select * from expet order by expetid;select exanimalid,exanimaltype from exanimal order by exanimaltype"
        DbAdapter1.TbgetDataSet(sqlstr, DataSet)
        DataSet.Tables(0).TableName = "TBExPet"
        DataSet.Tables(1).TableName = "TBExAnimal"
    End Sub

    Public Overrides Sub InitDataGrid()

        'MyBase.InitDataGrid()
        'DataGridViewComboBoxColumn1
        '
        With ColAnimalType
            .DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.[Nothing]
            .HeaderText = "Animal Type"
            .Name = "DataGridViewComboBoxColumn1"
            .ReadOnly = True
            .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
            .Width = 90
        End With

        '
        'DataGridViewTextBoxColumn1
        '
        With ColFirstName
            .HeaderText = "First Name"
            .Name = "DataGridViewTextBoxColumn1"
            .ReadOnly = True
            .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
            .Width = 82
        End With

        '
        'DataGridViewTextBoxColumn2
        '
        With ColLastName
            .AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
            .HeaderText = "Last Name"
            .Name = "DataGridViewTextBoxColumn2"
            .ReadOnly = True
            .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        End With
    End Sub

    Public Overrides Sub BindingObject()
        'MyBase.BindingObject()
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


        With ComboBox1
            .DataSource = BindingSource2
            .DisplayMember = "exanimaltype"
            .ValueMember = "exanimalid"
        End With

        'DataBinding
        ComboBox1.DataBindings.Clear()
        Me.TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox1.DataBindings.Add("Text", BindingSource1, "firstname", True)
        TextBox2.DataBindings.Add("Text", BindingSource1, "lastname", True)
        ComboBox1.DataBindings.Add("SelectedValue", BindingSource1, "exanimalid", True)
    End Sub

    Public Overrides Sub BindDataSource()
        MyBase.BindDataSource()
        BindingSource1.DataSource = DataSet.Tables("TBExPet")
        BindingSource2.DataSource = DataSet.Tables("TBExAnimal")
    End Sub

    Public Overrides Sub DoUpdate()
        'MyBase.DoUpdate()
        If DataSet.HasChanges Then
            Dim TBPetSaveChanges As SaveChangesRecordDelegate = AddressOf DbAdapter1.TBPetSaveChanges
            Dim Reload As ReloadDelegate = AddressOf MyBase.LoadData
            UpdateRecordHelper(DataSet, TBPetSaveChanges, Reload, CM)
            BindingSource1.Sort = mypanel1.Sort
            BindingSource1.Filter = mypanel.myFilter
        End If
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        MyBase.new()
        InitializeComponent()
        Me.DataGridViewCustom1.AutoGenerateColumns = False
        MyBase.DGV = Me.DataGridViewCustom1
        ' Add any initialization after the InitializeComponent() call.

    End Sub

End Class
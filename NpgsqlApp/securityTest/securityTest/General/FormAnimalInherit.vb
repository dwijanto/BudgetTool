Imports HR.HelperClass

Public Class FormAnimalInherit
    Inherits FormBaseSortFilterNoDGV

    Public Overrides Sub InitObject()
        MyBase.InitObject()
        'BindingSource2 = New BindingSource
    End Sub


    Public Overrides Sub FillData()
        'MyBase.FillData()
        Dim sqlstr = "select * from exanimal order by exanimaltype"
        DbAdapter1.TbgetDataSet(sqlstr, DataSet)
        DataSet.Tables(0).TableName = "TBExAnimal"
    End Sub

    Public Overrides Sub InitDataGrid()

        'MyBase.InitDataGrid()
        'DataGridViewComboBoxColumn1
        '
        With ColAnimalType            
            .HeaderText = "Animal Type"
            .Name = "DataGridViewTextBoxColumn1"
            .ReadOnly = True
            .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
            .Width = 190
        End With
    End Sub

    Public Overrides Sub BindingObject()
        'MyBase.BindingObject()
        With ColAnimalType
            .DataPropertyName = "exanimaltype"          
        End With
        
        'DataBinding
        Me.TextBox1.DataBindings.Clear()
        TextBox1.DataBindings.Add("Text", BindingSource1, "exanimaltype", True)
        
    End Sub

    Public Overrides Sub BindDataSource()
        MyBase.BindDataSource()
        BindingSource1.DataSource = DataSet.Tables("TBExAnimal")
    End Sub

    Public Overrides Sub DoUpdate()
        'MyBase.DoUpdate()
        If DataSet.HasChanges Then
            Dim TBAnimalSaveChanges As SaveChangesRecordDelegate = AddressOf DbAdapter1.TBAnimalSaveChanges
            Dim Reload As ReloadDelegate = AddressOf MyBase.LoadData
            UpdateRecordHelper(DataSet, TBAnimalSaveChanges, Reload, CM)
            BindingSource1.Sort = mypanel1.Sort
            BindingSource1.Filter = mypanel.myFilter
        End If
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        MyBase.new()
        InitializeComponent()
        MyBase.DGV = Me.DataGridView1

        ' Add any initialization after the InitializeComponent() call.

    End Sub

   

End Class
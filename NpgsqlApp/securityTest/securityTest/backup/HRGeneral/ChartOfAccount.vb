﻿Imports HR.HelperClass
Imports DJLib
Public Class ChartOfAccount
    Protected CM As CurrencyManager
    Protected mypanel1 As UCSortTx
    Protected mypanel As UCFilterTx
    Dim Dataset As DataSet
    Dim sqlstr As String = String.Empty
    Dim WithEvents DT As New DateTimePicker
    Private Sub FormCOA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'AddDateTimePickerToBindingNavigator()
        LoadData()
        LoadToolstrip()
        ToolStrip1.Visible = False
    End Sub
    Public Sub LoadData()
        InitObject()
        FillData()
        BindDataSource()
        BindingObject()
        CM = CType(BindingContext(BindingSource1), CurrencyManager)
    End Sub
    Public Overridable Sub InitObject()
        InitDataGrid()
        BindingSource1 = New BindingSource

        'BindingSource2 = New BindingSource
        Dataset = New DataSet
        With DataGridView1
            .DataSource = BindingSource1
            .RowsDefaultCellStyle.BackColor = System.Drawing.Color.White
            .AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.WhiteSmoke
        End With
        'DataGridView1.AutoGenerateColumns = False
        BindingNavigator1.BindingSource = BindingSource1


    End Sub
    Public Overridable Sub FillData()
        'Sample Snippet
        'Dim sqlstr = "select * from expet order by expetid;select exanimalid,exanimaltype from exanimal order by exanimaltype"
        'DbAdapter1.TbgetDataSet(sqlstr, DataSet)
        'DataSet.Tables(0).TableName = "TBExPet"
        'DataSet.Tables(1).TableName = "TBExAnimal"

        sqlstr = "select ac.sapaccountf as ""Account"",cc.sapccf as ""SAPCC"",id.sapindexf as ""SAP Account ID"",an.sapaccnamef as ""Account Name"" from sapindexaccnamef idac" & _
                 " left join sapindexf id on id.sapindexfid = idac.sapindexfid" & _
                 " left join sapccf cc on cc.sapccfid = id.sapccfid" & _
                 " left join sapaccountf ac on ac.sapaccountfid = id.sapaccountfid" & _
                 " left join sapaccnamef an on an.sapaccnamefid = idac.sapaccnamefid"
        Dim message As String = String.Empty

        If Not DbAdapter1.TbgetDataSet(sqlstr, Dataset, message) Then
            MessageBox.Show(message)
            Exit Sub
        End If
        Dataset.Tables(0).TableName = "coa"

    End Sub
    Public Overridable Sub BindDataSource()
        BindingSource1.DataSource = Dataset.Tables("coa")
        DataGridView1.DataSource = BindingSource1

        'BindingNavigator1.BindingSource = BindingSource1

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
    Private Sub LoadToolstrip()
        Dim myaction As HideToolbarDelegate = AddressOf toolstripvisible
        LoadToolstripFilterSort(myaction, DataGridView1, mypanel1, ToolStrip1, mypanel)
    End Sub

    Private Sub toolstripvisible(ByVal toolstripvisible As Boolean)
        ToolStrip1.Visible = Not (toolstripvisible)
        'Button3.Visible = toolstripvisible
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


    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoadData()
    End Sub


    Private Sub AddDateTimePickerToBindingNavigator()
        DT.Format = DateTimePickerFormat.Custom
        DT.CustomFormat = "yyyy"
        DT.Value = CDate(Today.Year + 1 & "-1-1")
        BindingNavigator1.Items.Add(New ToolStripControlHost(DT))
    End Sub

    Private Sub DT_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DT.ValueChanged
        LoadData()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim myform = New FormImportCAO
        myform.ShowDialog()
        LoadData()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim filename As String = "ChartOfAccount-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        ExcelStuff.ExportToExcelAskDirectory(filename, DataGridView1)
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Call toolstripvisible(ToolStrip1.Visible)
    End Sub
End Class
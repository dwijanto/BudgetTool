Imports HR.HelperClass
Imports DJLib
Public Class GroupCategory
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
        BindingNavigator2.BindingSource = BindingSource1


    End Sub
    Public Overridable Sub FillData()
        'Sample Snippet
        'Dim sqlstr = "select * from expet order by expetid;select exanimalid,exanimaltype from exanimal order by exanimaltype"
        'DbAdapter1.TbgetDataSet(sqlstr, DataSet)
        'DataSet.Tables(0).TableName = "TBExPet"
        'DataSet.Tables(1).TableName = "TBExAnimal"

        'sqlstr = "select category as ""Category"",sapaccname as ""Account Name"",sapaccount as ""Account"" from groupingtable order by category"
        'Dim message As String = String.Empty

        'If Not DbAdapter1.TbgetDataSet(sqlstr, Dataset, message) Then
        '    MessageBox.Show(message)
        '    Exit Sub
        'End If
        'Dataset.Tables(0).TableName = "coa"

        Dim message As String = String.Empty
        Dim ra As Integer = 0

        If DbAdapter1.GroupCategoryAdapter(Dataset, message, ra) Then
            Dataset.Tables(0).TableName = "coa"
            'BindingSource1.DataSource = Dataset.Tables("paramdt")
            'bindingDataGridView()
            'CM = CType(BindingContext(BindingSource1), CurrencyManager)
            'applyfilter()
        Else
            MessageBox.Show(message)
        End If


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

        'DataGridView1.Columns(3).Visible = False
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

        DataGridView1.Columns.Clear()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = BindingSource1

        Dim Col1 As New DataGridViewTextBoxColumn()
        With Col1
            .DataPropertyName = "category"
            .Name = "col1"
            .HeaderText = "Category"
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            .Visible = True
        End With

        Dim Col2 As New DataGridViewTextBoxColumn()
        With Col2
            .DataPropertyName = "sapaccname"
            .Name = "col2"
            .HeaderText = "SAP Account Name"
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            .Visible = True
        End With

        Dim Col3 As New DataGridViewTextBoxColumn()
        With Col3
            .DataPropertyName = "sapaccount"
            .Name = "Col3"
            .HeaderText = "SAP Account"
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            .Visible = True
        End With

        
        With DataGridView1
            .Columns.Insert(0, Col1)
            .Columns.Insert(1, col2)
            .Columns.Insert(2, Col3)
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End With

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
        BindingNavigator2.Items.Add(New ToolStripControlHost(DT))
    End Sub

    Private Sub DT_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DT.ValueChanged
        LoadData()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim myform = New ImportGroupingCategory
        myform.ShowDialog()
        LoadData()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim filename As String = "GroupCategory-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        ExcelStuff.ExportToExcelAskDirectory(filename, DataGridView1)
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call toolstripvisible(ToolStrip1.Visible)
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Dim myform = New ImportGroupingCategory
        myform.ShowDialog()
        LoadData()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Dim filename As String = "GroupCategory-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        ExcelStuff.ExportToExcelAskDirectory(filename, DataGridView1)
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Call toolstripvisible(ToolStrip1.Visible)
    End Sub


    Private Sub BindingNavigatorAddNewItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem.Click
        BindingSource1.AddNew()
    End Sub

    Private Sub BindingNavigatorDeleteItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorDeleteItem.Click
        If MessageBox.Show("Delete selected record?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                If DataGridView1.SelectedRows.Count = 0 Then
                    BindingSource1.RemoveAt(CM.Position)
                Else
                    For Each a As DataGridViewRow In DataGridView1.SelectedRows
                        BindingSource1.RemoveAt(a.Index)
                    Next
                End If
                UpdateRecord()
            Catch ex As Exception
            End Try
        End If
    End Sub
    Public Sub UpdateRecord()

        Dim ra As Integer
        Dim message As String = String.Empty
        Dim sb As New System.Text.StringBuilder
        Try
            BindingSource1.EndEdit()
            Dim ds2 = Dataset.GetChanges
            If Not IsNothing(ds2) Then
                'If DbAdapter1.TBRolesSaveChanges(ds2, message, ra) Then
                If DbAdapter1.GroupCategoryAdapter(ds2, message, ra) Then
                    'sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected.")
                    Dataset.Merge(ds2)
                    Dataset.AcceptChanges()
                End If
                If Dataset.HasErrors Then
                    sb.Append("Some Record(s) has been modified/deleted by other user. Records will refresh shortly.")
                    'sb.Append(message)
                    MessageBox.Show(sb.ToString)
                    loadData()
                Else
                    If sb.ToString <> "" Then
                        MessageBox.Show(sb.ToString)
                    End If
                    loadData()
                    sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected.")
                    MessageBox.Show(sb.ToString)
                End If

            Else
                MessageBox.Show("Nothing to save.")
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub SaveToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton1.Click
        UpdateRecord()
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        refreshrecord()
    End Sub
    Private Sub refreshrecord()
        If Dataset.HasChanges Then
            Dim datasetchanges As DataSet
            datasetchanges = Dataset.GetChanges()
            Dim response As Windows.Forms.DialogResult
            response = MessageBox.Show(datasetchanges.Tables(0).Rows.Count & " unsaved data. Do you want to store to the database?", "Unsaved data", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case response
                Case Windows.Forms.DialogResult.Yes
                    UpdateRecord()
                    loadData()
                Case Windows.Forms.DialogResult.Cancel

                Case Windows.Forms.DialogResult.No
                    loadData()
            End Select
        Else
            loadData()
        End If
    End Sub
End Class
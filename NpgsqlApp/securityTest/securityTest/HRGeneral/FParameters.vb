Imports HR.HelperClass

Public Class FParameters
    Dim myyear As Integer
    Dim Dataset As DataSet
    Dim sqlstr As String
    Dim applicationname = DJLib.AppConfig.RoleAttribute.ApplicationName
    Dim CurrentRole As String
    Dim mycheck As New ArrayList
    Dim CM As CurrencyManager
    Dim CurrentPosition As Integer = 0
    Dim paramhdid As Integer = 0
    Private Sub FParameters_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SuspendLayout()
        BindingSource1 = New BindingSource
        Dim ds As New DataSet
        Dim bindingsource2 As New BindingSource
        Dim sqlstr As String = "select paramname,paramhdid from paramhd ph where ph.cvalue = '" & dbtools1.Region & "' order by paramname;"
        DbAdapter1.TbgetDataSet(sqlstr, ds)
        bindingsource2.DataSource = ds.Tables(0)
        ToolStripComboBox1.ComboBox.DataSource = bindingsource2
        ToolStripComboBox1.ComboBox.DisplayMember = "paramname"
        ToolStripComboBox1.ComboBox.ValueMember = "paramhdid"
        loadData()
        AddHandler Dataset.Tables(0).TableNewRow, AddressOf onTableNewRow
        CM = CType(BindingContext(BindingSource1), CurrencyManager)
        ResumeLayout()
    End Sub


    Private Sub loadData()      
        Dataset = New DataSet
        DataGridView1.DataSource = BindingSource1
        DataGridView1.AutoGenerateColumns = True

        BindingNavigator1.BindingSource = BindingSource1
        'sqlstr = "select * from roles where applicationname = '" & applicationname & "' order by rolename;" & _
        '         "select * from users  where applicationname='" & applicationname & "' order by username;" & _
        '         "select * from tbprogram  where applicationname='" & applicationname & "' order by parentid,myorder;"

        'sqlstr = "select pd.paramname,pd.dvalue,pd.nvalue,pd.paramdtid,pd.paramhdid from paramdt pd " & _
        '         " where pd.paramhdid in (select paramhdid from paramhd where paramname =  'Exchange Rate') " & _
        '         " order by dvalue desc,pd.paramname "

        'DbAdapter1.TbgetDataSet(sqlstr, Dataset)
        'Dataset.Tables(0).TableName = "paramdt"
        Dataset = New DataSet
        DataGridView1.DataSource = BindingSource1

        Dim message As String = String.Empty
        Dim ra As Integer = 0

        If DbAdapter1.TBParamDataAdapter(Dataset, paramhdid, message, ra) Then
            Dataset.Tables(0).TableName = "paramdt"
            BindingSource1.DataSource = Dataset.Tables("paramdt")
            bindingDataGridView()
            CM = CType(BindingContext(BindingSource1), CurrencyManager)
            applyfilter()
        Else
            MessageBox.Show(message)
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
                If DbAdapter1.TBParamDataAdapter(ds2, paramhdid, message, ra) Then
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



    Private Sub RefreshToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripButton.Click
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

    Private Sub FormUser_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs)
        If Dataset.HasChanges Then
            Dim datasetchanges As DataSet
            datasetchanges = Dataset.GetChanges()
            Dim response As Windows.Forms.DialogResult
            response = MessageBox.Show(datasetchanges.Tables(0).Rows.Count & " unsaved data. Do you want to store to the database?", "Unsaved data", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case response
                Case Windows.Forms.DialogResult.Yes
                    UpdateRecord()
                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
                Case Windows.Forms.DialogResult.No
            End Select
        End If
    End Sub

    Private Sub BindingNavigatorAddNewItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem.Click
        BindingSource1.AddNew()
    End Sub
    Private Sub onTableNewRow(ByVal sender As Object, ByVal e As DataTableNewRowEventArgs)
        e.Row(6) = paramhdid
    End Sub
    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)
        MessageBox.Show(e.Exception.Message.ToString)
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim Userinroles() As String
        'Try
        '    Dim dg As DataGridView = CType(sender, DataGridView)
        '    CurrentRole = DataGridView1.Rows(CM.Position).Cells(0).Value

        '    'assign checkbox
        '    Userinroles = DJLib.AppConfig.RoleAttribute.GetUsersInRole(CurrentRole)
        '    'Dim chk As CheckBox

        'Catch ex As Exception
        'End Try


    End Sub

    Private Sub bindingDataGridView()
        DataGridView1.Columns.Clear()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = BindingSource1

        'Dim Col0 As New DataGridViewComboBoxColumn

        'With Col0
        '    .DataPropertyName = "paramname"
        '    .Name = "col0"
        '    .HeaderText = "Parameter Name"
        '    .DropDownWidth = 160
        '    .Width = 150
        '    .MaxDropDownItems = 7
        '    .FlatStyle = FlatStyle.Flat
        '    .Visible = True
        '    .DataSource = {"10", "15", "Amount A", "Amount B", "Amount C", "Expat Rate", "General Rate"}
        '    '.ValueMember = "paramname"
        '    '.DisplayMember = "paramname"
        'End With

        Dim Col1 As New DataGridViewTextBoxColumn()
        With Col1
            .DataPropertyName = "paramname"
            .Name = "col1"
            .HeaderText = "Param Name"
            .Width = 100
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            .Visible = True
        End With

        Dim col2 As New DJLib.DataGridViewCalendarColumn()
        With col2
            .DataPropertyName = "dvalue"
            .Name = "col2"
            .HeaderText = "Valid From"
            .Width = 200
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            .Visible = True
        End With

        'Dim Col2 As New DataGridViewTextBoxColumn()
        'With Col2
        '    .DataPropertyName = "email"
        '    .Name = "col2"
        '    .HeaderText = "Email"
        '    .Width = 200
        '    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
        '    .Visible = True
        'End With

        Dim Col3 As New DataGridViewTextBoxColumn()
        With Col3
            .DataPropertyName = "nvalue"
            .Name = "Col3"
            .HeaderText = "Value"
            .Width = 100
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopRight
            .Visible = True
        End With
        Dim Col4 As New DataGridViewTextBoxColumn()
        With Col4
            .DataPropertyName = "cvalue"
            .Name = "Col4"
            .HeaderText = "Description"
            .Width = 200
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            .Visible = True
        End With
        Dim Col5 As New DataGridViewTextBoxColumn()
        With Col5
            .DataPropertyName = "ivalue"
            .Name = "Col5"
            .HeaderText = "Month"
            .Width = 60
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopRight
            .Visible = True
        End With

        'Dim col4 As New DataGridViewCheckBoxColumn()
        'With col4
        '    .HeaderText = "Is Approved"
        '    .Name = "Col4"
        '    .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        '    .FlatStyle = FlatStyle.Standard
        '    '.CellTemplate = New DataGridViewCheckBoxCell()
        '    '.CellTemplate.Style.BackColor = Color.Beige
        '    .DataPropertyName = "isapproved"
        'End With

        'Dim col5 As DataGridViewComboBoxColumn
        'col5 = CreateComboBoxColumn()
        'SetAlternateChoicesUsingDataSource(col5)
        ''col5.HeaderText = "Region Name"

        With DataGridView1
            '.Columns.Insert(0, Col0)
            .Columns.Insert(0, Col1)
            .Columns.Insert(1, col2)
            .Columns.Insert(2, Col3)
            .Columns.Insert(3, Col4)
            .Columns.Insert(4, Col5)

        End With


    End Sub



    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        Me.Validate()
        'If DJLib.AppConfig.Principal.IsInRole("superuser") Then
        'Dim mycombovalue = ToolStripComboBox1.Text 'DirectCast(ToolStripComboBox1.ComboBox.SelectedItem, DataRowView).Row.Item(1).ToString
        'CurrentPosition = CM.Position
        UpdateRecord()
        'ToolStripComboBox1.Text = mycombovalue
        'CM.Position = CurrentPosition
        'End If
    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        applyfilter()
    End Sub

    Private Sub applyfilter()
        Try
            paramhdid = DirectCast(ToolStripComboBox1.ComboBox.SelectedItem, DataRowView).Row.Item(1).ToString
            BindingSource1.Filter = "paramhdid = '" & paramhdid & "'"

        Catch ex As Exception

        End Try
    End Sub


End Class
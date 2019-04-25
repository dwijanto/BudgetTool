﻿Imports HR.HelperClass
Public Class FCurrency
    Dim Dataset As DataSet
    Dim sqlstr As String
    Dim applicationname = DJLib.AppConfig.RoleAttribute.ApplicationName
    Dim CurrentRole As String
    Dim mycheck As New ArrayList
    Dim CM As CurrencyManager

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
       
    End Sub

    Private Sub FCurrency_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SuspendLayout()

        BindingSource1 = New BindingSource
        loadData()
        bindingText()
        AddHandler Dataset.Tables(0).TableNewRow, AddressOf onTableNewRow


        CM = CType(BindingContext(BindingSource1), CurrencyManager)

        ResumeLayout()
    End Sub

    Private Sub bindingText()
        'TextBox1.DataBindings.Add("Text", BindingSource1, "rolename", True)
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
        Dim message As String = String.Empty
        Dim ra As Integer = 0

        If DbAdapter1.TBParamDtSaveChanges(Dataset, message, ra) Then
            Dataset.Tables(0).TableName = "paramdt"
            BindingSource1.DataSource = Dataset.Tables("paramdt")
            bindingDataGridView()
            CM = CType(BindingContext(BindingSource1), CurrencyManager)
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
                If DbAdapter1.TBParamDtSaveChanges(ds2, message, ra) Then
                    'sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected.")
                    Dataset.Merge(ds2)
                    Dataset.AcceptChanges()
                End If
                If Dataset.HasErrors Then
                    sb.Append("Some Record(s) has been modified/deleted by other user. Records will refresh shortly.")
                    MessageBox.Show(sb.ToString)
                    'loadData()
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

    Private Sub FormUser_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
        'e.Row(1) = applicationname
    End Sub
    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message.ToString)
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
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


        Dim Col1 As New DataGridViewTextBoxColumn()
        With Col1
            .DataPropertyName = "paramname"
            .Name = "col1"
            .HeaderText = "Currency"
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
            .Columns.Insert(0, Col1)
            .Columns.Insert(1, Col2)
            .Columns.Insert(2, Col3)
            '.Columns.Insert(3, col4)
            '.Columns.Insert(4, col5)
            '.Columns.Insert(5, Col6)
        End With


    End Sub



    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click

        'If DJLib.AppConfig.Principal.IsInRole("superuser") Then
        UpdateRecord()
        'End If
    End Sub



End Class
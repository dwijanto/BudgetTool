'Imports DJLib.AppConfig
Imports HR.HelperClass
Public Class FormProgram
    Dim Dataset As DataSet    
    Dim sqlstr As String

    Dim CM As CurrencyManager

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        BindingSource1 = New BindingSource
        loadData()
        bindingText()
        bindingCombobox()
    End Sub
    Private Sub bindingCombobox()
        ComboBox1.DataSource = Dataset.Tables(1)
        ComboBox1.DisplayMember = "rolename"
        ComboBox1.SelectedItem = "rolename"
    End Sub
    Private Sub bindingText()
        TextBox1.DataBindings.Add("Text", BindingSource1, "programid", True)
        TextBox2.DataBindings.Add("Text", BindingSource1, "parentid", True)
        TextBox3.DataBindings.Add("Text", BindingSource1, "myorder", True)
        TextBox4.DataBindings.Add("Text", BindingSource1, "description", True)
        TextBox5.DataBindings.Add("Text", BindingSource1, "programname", True)
        CheckBox1.DataBindings.Add("Checked", BindingSource1, "isactive", True)
        TextBox6.DataBindings.Add("Text", BindingSource1, "icon", True)
        TextBox7.DataBindings.Add("Text", BindingSource1, "iconindex", True)
    End Sub
    Private Sub loadData()
        Dataset = New DataSet
        DataGridView1.DataSource = BindingSource1
        DataGridView1.AutoGenerateColumns = True
        BindingNavigator1.BindingSource = BindingSource1        
        sqlstr = "select * from tbprogram order by programid;" & _
                 "select '' as rolename, '' as applicationname union all select * from roles order by rolename;"
        DbAdapter1.TbgetDataSet(sqlstr, Dataset)
        Dataset.Tables(0).TableName = "tbprogram"
        BindingSource1.DataSource = Dataset.Tables("tbprogram")
        DataGridView1.Columns(9).Visible = False
        CM = CType(BindingContext(BindingSource1), CurrencyManager)
    End Sub

    Public Sub UpdateRecord()
        Dim ra As Integer
        Dim message As String = String.Empty
        Dim sb As New System.Text.StringBuilder

        BindingSource1.EndEdit()
        Dim ds2 = Dataset.GetChanges
        If DbAdapter1.TBProgramSaveChanges(ds2, message, ra) Then
            sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected.")
            Dataset.Merge(ds2)
            Dataset.AcceptChanges()
        End If
        If Dataset.HasErrors Then
            sb.Append("Some Record(s) has been modified/deleted by other user. Records will refresh shortly.")
            MessageBox.Show(sb.ToString)
            loadData()
        Else
            MessageBox.Show(sb.ToString)
        End If
    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        UpdateRecord()
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)
        MessageBox.Show(e.Exception.Message & "Some Records has been deleted by other user.Refreshing record in progress")
    End Sub

    Private Sub RefreshToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripButton.Click
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

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        Dim dg As DataGridView = CType(sender, DataGridView)
        ListBox1.Items.Clear()
        If dg.CurrentRow IsNot Nothing Then
            If Not IsDBNull(dg.CurrentRow.Cells(8).Value) Then
                Dim myvalue As String = dg.CurrentRow.Cells(8).Value
                Dim mylist() As String = myvalue.Split(" ")
                For i = 0 To mylist.Count - 1
                    ListBox1.Items.Add(mylist(i).ToString)
                Next
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If DJLib.AppConfig.Principal.IsInRole("superuser") Then
            If Not ListBox1.Items.Contains(ComboBox1.Text.ToString) AndAlso ComboBox1.Text <> "" Then
                ListBox1.Items.Add(ComboBox1.Text)
                getlistboxlist()
            End If
        End If
    End Sub

    Private Sub getlistboxlist()
        Dim sb As New System.Text.StringBuilder
        For i = 0 To ListBox1.Items.Count - 1
            If sb.Length > 0 Then
                sb.Append(" ")
            End If
            sb.Append(ListBox1.Items.Item(i))
        Next
        DataGridView1.CurrentRow.Cells(8).Value = sb.ToString
        BindingSource1.EndEdit()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If DJLib.AppConfig.Principal.IsInRole("superuser") Then
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            getlistboxlist()
        End If

    End Sub
End Class
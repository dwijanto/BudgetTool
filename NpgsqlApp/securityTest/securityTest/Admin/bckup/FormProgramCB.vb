Imports securityTest.HelperClass
Public Class FormProgramCB
    Dim Dataset As DataSet
    Dim sqlstr As String
    Dim myCheck As New ArrayList
    Dim CM As CurrencyManager
    Dim applicationname As String = DJLib.AppConfig.RoleAttribute.ApplicationName
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        BindingSource1 = New BindingSource
        loadData()
        bindingText()   
        bindingCheckbox()

        AddHandler Dataset.Tables(0).TableNewRow, AddressOf onNewRow

    End Sub
    Private Sub onNewRow(ByVal sender As Object, ByVal e As DataTableNewRowEventArgs)
        e.Row(10) = applicationname
    End Sub
    Private Sub bindingCheckbox()

        For i = 0 To Dataset.Tables(1).Rows.Count - 1
            Dim dr As DataRow = Dataset.Tables(1).Rows(i)
            Dim mycheckbox As New CheckBox
            mycheckbox.Text = dr.Item(0).ToString
            myCheck.Add(mycheckbox)

            FlowLayoutPanel1.Controls.Add(myCheck(i))
            AddHandler CType(myCheck(i), CheckBox).Click, AddressOf onCheckboxClicked
        Next
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
        sqlstr = "select * from tbprogram  where applicationname='" & applicationname & "' order by programid;" & _
                 "select * from roles  where applicationname='" & applicationname & "' order by rolename;"
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
        CheckBox1.Checked = False
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        Dim dg As DataGridView = CType(sender, DataGridView)

        Try
            Dim myvalue As String = DataGridView1.Rows(CM.Position).Cells(8).Value.ToString 'dg.CurrentRow.Cells(8).Value
            Dim mylist() As String = myvalue.Split(" ")
            Dim chk As CheckBox
            For Each mycontrol In FlowLayoutPanel1.Controls
                If TypeOf (mycontrol) Is CheckBox Then
                    chk = CType(mycontrol, CheckBox)
                    chk.Checked = mylist.Contains(chk.Text)
                End If
            Next
        Catch ex As Exception
        End Try        
    End Sub

    Private Sub onCheckboxClicked(ByVal sender As Object, ByVal e As EventArgs)
        If DJLib.AppConfig.Principal.IsInRole("administrators") Then
            Dim sb As New System.Text.StringBuilder
            For i = 0 To FlowLayoutPanel1.Controls.Count - 1                
                If CType(FlowLayoutPanel1.Controls(i), CheckBox).Checked Then
                    If sb.Length > 0 Then
                        sb.Append(" ")
                    End If
                    sb.Append(CType(FlowLayoutPanel1.Controls(i), CheckBox).Text)
                End If
            Next
            'DataGridView1.CurrentRow.Cells(8).Value = sb.ToString
            DataGridView1.Rows(CM.Position).Cells(8).Value = sb.ToString
            BindingSource1.EndEdit()
        End If
    End Sub





End Class
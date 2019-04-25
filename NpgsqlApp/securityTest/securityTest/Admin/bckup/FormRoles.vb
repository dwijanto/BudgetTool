Imports securityTest.HelperClass

Public Class FormRoles
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
        SuspendLayout()

        BindingSource1 = New BindingSource
        loadData()
        bindingText()
        AddHandler Dataset.Tables(0).TableNewRow, AddressOf onTableNewRow

        loaduserisinrole()
        CM = CType(BindingContext(BindingSource1), CurrencyManager)
        loadmenuinrole()
        ResumeLayout()
    End Sub

    Private Sub loadmenuinrole()
        TreeView1.Nodes.Clear()
        Dim drs = Dataset.Tables(2).Select("parentid=0", "myorder asc")
        For Each dr As DataRow In drs
            Dim root As TreeNode = New TreeNode(dr("description").ToString)
            root.Tag = dr("programid")
            root.SelectedImageIndex = 0
            TreeView1.Nodes.Add(root)
            PopulateTree(dr, root)
        Next

        TreeView1.SelectedNode = TreeView1.Nodes(0)        
        TreeView1.ExpandAll()

    End Sub
    Private Sub PopulateTree(ByVal dr As DataRow, ByVal pnode As TreeNode)
        Dim drs = Dataset.Tables(2).Select("parentid=" & dr("programid"), "myorder asc")

        For Each row As DataRow In drs
            Dim cChild As TreeNode = New TreeNode(row("description").ToString)
            cChild.Tag = row("programid")
            cChild.SelectedImageIndex = 0          
            pnode.Nodes.Add(cChild)
            PopulateTree(row, cChild)
        Next
    End Sub

    Private Sub loaduserisinrole()
        For i = 0 To Dataset.Tables(1).Rows.Count - 1
            Dim dr As DataRow = Dataset.Tables(1).Rows(i)
            Dim mycheckbox As New CheckBox
            mycheckbox.Text = dr.Item(1).ToString
            mycheck.Add(mycheckbox)

            FlowLayoutPanel1.Controls.Add(mycheck(i))
            AddHandler CType(mycheck(i), CheckBox).Click, AddressOf onUserCheckboxClicked
        Next
    End Sub

    Private Sub onUserCheckboxClicked(ByVal sender As Object, ByVal e As EventArgs)
        Dim chk As CheckBox = CType(sender, CheckBox)
        Dim datasetchanges As DataSet
        datasetchanges = Dataset.GetChanges()
        Try
            If datasetchanges.Tables(0).Rows.Count > 0 Then
                MessageBox.Show(String.Format("{0} unsaved data. Please save data first before assign user to role.", datasetchanges.Tables(0).Rows.Count))
                chk.Checked = Not chk.Checked
                Exit Sub
            End If
        Catch ex As Exception
        End Try

        Dim users() As String = {chk.Text}
        Dim roles() As String = {CurrentRole}
        Try
            If chk.Checked Then
                DJLib.AppConfig.RoleAttribute.AddUsersToRoles(users, roles)
            Else
                DJLib.AppConfig.RoleAttribute.RemoveUsersFromRoles(users, roles)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub



    Private Sub bindingText()
        'TextBox1.DataBindings.Add("Text", BindingSource1, "rolename", True)
    End Sub
    Private Sub loadData()
        Dataset = New DataSet
        DataGridView1.DataSource = BindingSource1
        DataGridView1.AutoGenerateColumns = True

        BindingNavigator1.BindingSource = BindingSource1
        sqlstr = "select * from roles where applicationname = '" & applicationname & "' order by rolename;" & _
                 "select * from users  where applicationname='" & applicationname & "' order by username;" & _
                 "select * from tbprogram  where applicationname='" & applicationname & "' order by parentid,myorder;"
        DbAdapter1.TbgetDataSet(sqlstr, Dataset)
        Dataset.Tables(0).TableName = "roles"
        BindingSource1.DataSource = Dataset.Tables("roles")
        Dataset.Tables(2).TableName = "TBProgram"
        DataGridView1.Columns(1).Visible = False
    End Sub

    Public Sub UpdateRecord()

        Dim ra As Integer
        Dim message As String = String.Empty
        Dim sb As New System.Text.StringBuilder
        Try
            BindingSource1.EndEdit()
            Dim ds2 = Dataset.GetChanges
            If DbAdapter1.TBRolesSaveChanges(ds2, message, ra) Then
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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        If DJLib.AppConfig.Principal.IsInRole("administrators") Then
            UpdateRecord()
        End If
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
        e.Row(1) = applicationname
    End Sub
    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message.ToString)
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        Dim Userinroles() As String
        Try
            Dim dg As DataGridView = CType(sender, DataGridView)
            CurrentRole = DataGridView1.Rows(CM.Position).Cells(0).Value
            Label2.Text = String.Format("Role ""{0}""", CurrentRole)
            Label5.Text = String.Format("Role ""{0}""", CurrentRole)
            'assign checkbox
            Userinroles = DJLib.AppConfig.RoleAttribute.GetUsersInRole(CurrentRole)
            Dim chk As CheckBox
            For Each mycontrol In FlowLayoutPanel1.Controls
                If TypeOf (mycontrol) Is CheckBox Then
                    chk = CType(mycontrol, CheckBox)
                    chk.Checked = Userinroles.Contains(chk.Text)
                End If
            Next
        Catch ex As Exception
        End Try
        updatetabs()

    End Sub
    Private Sub updatetabs()
        Try
            Dim n As TreeNode
            For Each n In TreeView1.Nodes
                Printnode(n, CurrentRole)
            Next
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Printnode(ByVal n As TreeNode, ByVal CurrentRole As String)
        Dim allmember() As String
        Dim dr = Dataset.Tables(2).Rows.Find(n.Tag)
        allmember = dr("members").ToString.Split(" ")
        n.Checked = allmember.Contains(CurrentRole)
        For Each anode As TreeNode In n.Nodes
            Printnode(anode, CurrentRole)
        Next
    End Sub
    Private Sub TreeView1_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        If DJLib.AppConfig.Principal.IsInRole("administrators") Then
            Dim n As TreeNode = e.Node
            Dim allmember() As String
            Dim mylist As New ArrayList
            Dim dr = Dataset.Tables(2).Rows.Find(n.Tag)
            allmember = dr("members").ToString.Split(" ")

            For i = 0 To allmember.Length - 1
                mylist.Add(allmember(i))
            Next
            If e.Node.Checked Then
                If Not mylist.Contains(CurrentRole) Then
                    mylist.Add(CurrentRole)
                Else
                    Exit Sub
                End If
            Else
                If mylist.Contains(CurrentRole) Then
                    mylist.Remove(CurrentRole)
                Else
                    Exit Sub
                End If
            End If
            mylist.Sort()
            Dim sb As New System.Text.StringBuilder
            For i = 0 To mylist.Count - 1
                If mylist(i) <> "" Then
                    If sb.Length > 0 Then
                        sb.Append(" ")
                    End If
                    sb.Append(mylist(i))
                End If
            Next

            Dim q = From a As DataRow In Dataset.Tables(2).Rows Where
                    a.Item(3) = e.Node.Text And a.Item(10) = DJLib.AppConfig.RoleAttribute.ApplicationName
                    Select a
            For Each a In q
                a.Item(8) = sb.ToString
            Next

            DbAdapter1.TBProgramSaveChanges(Dataset)
            'loadmenuinrole()
            updatetabs()
        Else
            e.Node.Checked = Not e.Node.Checked
        End If
    End Sub





End Class
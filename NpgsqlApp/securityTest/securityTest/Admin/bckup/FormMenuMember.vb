Imports securityTest.HelperClass
Public Class FormMenuMember
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
        TreeView1.ImageList = FormMenu.ImageList1
        loadData()
        bindingCheckbox()
        CM.Position = Dataset.Tables(0).DefaultView.Find(TreeView1.Nodes(0).Tag)
        loadCheckbox()
        AddHandler Dataset.Tables(0).TableNewRow, AddressOf onNewRow
        moverecordposition()
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

    Private Sub loadData()
        SuspendLayout()
        Dataset = New DataSet
        sqlstr = "select * from tbprogram  where applicationname='" & applicationname & "' order by programid;" & _
                 "select * from roles  where applicationname='" & applicationname & "' order by rolename;"
        DbAdapter1.TbgetDataSet(sqlstr, Dataset)
        Dataset.Tables(0).TableName = "tbprogram"
        Dataset.Tables(0).DefaultView.Sort = "programid"
        BindingSource1.DataSource = Dataset.Tables("tbprogram")
        CM = CType(BindingContext(BindingSource1), CurrencyManager)
        Dataset.Relations.Add("relParentChild", Dataset.Tables(0).Columns("programid"), Dataset.Tables(0).Columns("parentid"))
        loadmenuinrole()
        ResumeLayout()
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
            If ds2 IsNot Nothing Then
                MessageBox.Show(sb.ToString)
            Else
                MessageBox.Show("Nothing to save.")
            End If
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

    Private Sub BindingNavigatorAddNewItem_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        BindingSource1.AddNew()
    End Sub

    Private Sub loadCheckbox()
        Try
            Dim dr As DataRow = Dataset.Tables(0).Rows.Find(TreeView1.SelectedNode.Tag)
            Dim myvalue = dr("members")
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
            Dim dr As DataRow = Dataset.Tables(0).Rows.Find(TreeView1.SelectedNode.Tag)
            dr.BeginEdit()
            dr("members") = sb.ToString
            dr.EndEdit()
            BindingSource1.EndEdit()
        End If
    End Sub
    Private Sub loadmenuinrole()
        TreeView1.Nodes.Clear()
        'skip root node
        Dim drs = Dataset.Tables(0).Select("parentid=0", "myorder asc")
        For Each dr As DataRow In drs
            Dim root As TreeNode = New TreeNode(dr("description").ToString)
            root.Tag = dr("programid")
            root.SelectedImageIndex = 4
            root.ImageIndex = 0
            TreeView1.Nodes.Add(root)
            PopulateTree(dr, root)
        Next
        TreeView1.SelectedNode = TreeView1.Nodes(0)
        TreeView1.ExpandAll()
    End Sub
    Private Function PopulateTree(ByVal dr As DataRow, ByVal pnode As TreeNode) As Boolean
        Dim result As Boolean = False
        For Each row As DataRow In dr.GetChildRows("RelParentChild")
            Dim cChild As TreeNode = New TreeNode(row("description").ToString)
            result = PopulateTree(row, cChild)
            cChild.Tag = row("programid")
            If result Then
                cChild.SelectedImageIndex = 4
                cChild.ImageIndex = 0
            Else
                cChild.SelectedImageIndex = 4
                cChild.ImageIndex = 3
            End If
            pnode.Nodes.Add(cChild)
            result = True
        Next
        Return result
    End Function

    Private Sub moverecordposition()
        Dim tn As TreeNode = TreeView1.SelectedNode
        Label8.Text = String.Format("Member : ""{0}""", tn.Text)
        Dim row = Dataset.Tables(0).DefaultView.Find(tn.Tag)
        cm.Position = row
        loadCheckbox()
    End Sub

    Private Sub TreeView1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TreeView1.KeyUp
        Dim tn As TreeNode = TreeView1.SelectedNode
        TreeView1.SelectedNode = tn
        moverecordposition()
    End Sub

    Private Sub TreeView1_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        Dim tn As TreeNode = e.Node
        TreeView1.SelectedNode = tn
        moverecordposition()
    End Sub
End Class
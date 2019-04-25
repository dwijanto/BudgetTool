Imports HR.HelperClass
Public Class FormMenuEditor
    Dim Dataset As DataSet
    Dim sqlstr As String
    Dim applicationname = DJLib.AppConfig.RoleAttribute.ApplicationName
    Private WithEvents bindingsource1 As BindingSource
    Private count As Integer = 0
    Private tn As TreeNode
    Dim cm As CurrencyManager
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()


    End Sub


    Private Sub FormMenuEditor_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dataset = New DataSet
        loaddata()

        ' Add any initialization after the InitializeComponent() call.
        tn = New TreeNode
        ' Set the index of image from the 
        TreeView1.ImageList = FormMenu.ImageList1
        panelbuttonvisible(TreeView1.Nodes.Count > 0)
        updownenabled(TreeView1.Nodes.Count > 1)
        TreeView1.ContextMenuStrip = ContextMenuStrip1
        ' ImageList for selected and unselected tree nodes.
        'TreeView1.ImageIndex = 0
        'TreeView1.SelectedImageIndex = 0
    End Sub
    Private Sub bindingtext()
        TextBox1.DataBindings.Add("Text", bindingsource1, "description")
        TextBox2.DataBindings.Add("Text", bindingsource1, "programname")
        CheckBox1.DataBindings.Add("Checked", bindingsource1, "isactive")
        TextBox3.DataBindings.Add("Text", bindingsource1, "icon")
        TextBox4.DataBindings.Add("Text", bindingsource1, "iconindex")
        TextBox5.DataBindings.Add("Text", bindingsource1, "latestupdate")
        TextBox6.DataBindings.Add("Text", bindingsource1, "formname")
    End Sub

#Region "TreeView"
    Private Sub NewMenuToolStripButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NewMenuToolStripButton.Click, ToolStripMenuItem1.Click
        Dim root = TreeView1.SelectedNode
        TreeView1.Focus()
        Dim tn As New TreeNode
        tn.Text = "Untitled" & getcount()

        If IsNothing(root.Parent) Then
            TreeView1.Nodes.Add(tn)

        Else
            root.Parent.Nodes.Add(tn)
        End If
        'TreeView1.SelectedNode = tn
        panelbuttonvisible(TreeView1.Nodes.Count > 0)
        updownenabled(TreeView1.Nodes.Count > 1)

        'addnewrecord()
    End Sub
    Private Sub InsertMenuItemToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InsertMenuItemToolStripButton.Click, ToolStripMenuItem3.Click
        Dim tn = TreeView1.SelectedNode
        TreeView1.Focus()
        Dim tni As New TreeNode
        tni.Text = "Untitled" & getcount()

        If IsNothing(tn.Parent) Then
            TreeView1.Nodes.Insert(tn.Index, tni)
        Else
            tn.Parent.Nodes.Insert(tn.Index, tni)
        End If
        tni.EnsureVisible()
        'addnewrecord()
    End Sub
    Private Sub panelbuttonvisible(ByVal state As Boolean)
        Panel1.Visible = state
        InsertMenuItemToolStripButton.Enabled = state
        MenuItemsToolStripButton.Enabled = state
        DeleteToolStripButton.Enabled = state

    End Sub
    Private Sub updownenabled(ByVal state As Boolean)
        MoveDownToolStripButton.Enabled = state
        MoveUpToolStripButton.Enabled = state
    End Sub
    Private Sub MenuItemsToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemsToolStripButton.Click, ToolStripMenuItem2.Click
        Dim tn = TreeView1.SelectedNode
        TreeView1.Focus()
        Dim tnc As New TreeNode
        tnc.Text = "Untitled" & getcount()
        tn.Nodes.Add(tnc)
        tnc.EnsureVisible()
    End Sub

    Private Function getcount()
        count += 1
        Return count
    End Function


    Private Sub addnewrecord()
        Dim myparent As Integer = 0
        If TreeView1.SelectedNode.Parent IsNot Nothing Then
            myparent = TreeView1.SelectedNode.Parent.Tag
        End If
        Dim tn As TreeNode = TreeView1.SelectedNode
        'bindingsource1.AddNew() not using this anymore

        Dim dr = Dataset.Tables(0).NewRow
        dr("applicationname") = applicationname
        dr("members") = "superuser"
        dr("parentid") = myparent
        dr("myorder") = tn.Index
        dr("description") = tn.Text
        dr("isactive") = True
        tn.Tag = CInt(dr("programid"))

        Dataset.Tables(0).Rows.Add(dr)


    End Sub

    Private Sub DeleteToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripButton.Click, ToolStripMenuItem4.Click
        If MessageBox.Show("Delete selected record?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Dim tn = TreeView1.SelectedNode
            TreeView1.Focus()
            Dim row = Dataset.Tables(0).Rows.Find(tn.Tag)
            row.Delete()
            tn.Remove()
            panelbuttonvisible(TreeView1.Nodes.Count > 0)
            updownenabled(TreeView1.Nodes.Count > 1)
            moverecordposition()
        End If
    End Sub

    Private Sub TreeView1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TreeView1.KeyUp
        moverecordposition()
    End Sub

    Private Sub TreeView1_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        Dim tn As TreeNode = e.Node
        TreeView1.SelectedNode = tn
        moverecordposition()
    End Sub

    Private Sub moverecordposition()
        Dim tn As TreeNode = TreeView1.SelectedNode
        'check for new node
        If IsNothing(tn.Tag) Then
            addnewrecord()

            bindingsource1.EndEdit()
            DbAdapter1.TBProgramSaveChanges(Dataset)
            Dataset.AcceptChanges()
            tn.Tag = Dataset.Tables(0).Rows(Dataset.Tables(0).Rows.Count - 1).Item("programid")

        Else

        End If

        Dataset.Tables(0).DefaultView.Sort = "programid"
        Dim row = Dataset.Tables(0).DefaultView.Find(tn.Tag)
        cm.Position = row

    End Sub

    Private Sub MoveUpToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MoveUpToolStripButton.Click
        movenode(Direction.MoveUp)
    End Sub

    Private Sub MoveDownToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MoveDownToolStripButton.Click
        movenode(Direction.MoveDown)
    End Sub

    Private Sub movenode(ByVal dir As Integer)
        Dim tn As TreeNode = TreeView1.SelectedNode
        Dim Index As Integer = tn.Index
        If dir = Direction.MoveUp Then
            If Index = 0 Then Exit Sub
        Else
            If IsNothing(tn.Parent) Then
                If Index = TreeView1.Nodes.Count - 1 Then Exit Sub
            Else
                If Index = tn.Parent.Nodes.Count - 1 Then Exit Sub
            End If
        End If
        If IsNothing(tn.Parent) Then

            TreeView1.Nodes.RemoveAt(Index)
            TreeView1.Nodes.Insert(Index + (1 * dir), tn)
            TreeView1.SelectedNode = TreeView1.Nodes(Index + (1 * dir))
        Else
            Dim Parent As TreeNode = tn.Parent
            Parent.Nodes.RemoveAt(Index)
            Parent.Nodes.Insert(Index + (1 * dir), tn)
            TreeView1.SelectedNode = Parent.Nodes(Index + (1 * dir))
        End If
    End Sub
#End Region

#Region "Drag Drop"
    Private Sub TreeView1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TreeView1.DragDrop
        If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) = False Then Exit Sub
        Dim st As TreeView = CType(sender, TreeView)
        Dim dropnode As TreeNode = CType(e.Data.GetData("System.Windows.Forms.TreeNode"), TreeNode)
        Dim targetnode As TreeNode = st.SelectedNode

        'Recursive check
        If targetnode IsNot Nothing Then
            Dim arr() = targetnode.FullPath.Split("\")
            'If targetnode.FullPath.Contains(dropnode.Text) Then
            If arr.Contains(dropnode.Text) Then
                MsgBox("Recursive node not allowed.")
                st.SelectedNode = dropnode
                Exit Sub
            End If
        End If

        dropnode.Remove()

        If targetnode Is Nothing Then
            st.Nodes.Add(dropnode)
        Else
            targetnode.Nodes.Add(dropnode)
        End If
        dropnode.EnsureVisible()
        TreeView1.SelectedNode = dropnode
        moverecordposition()
    End Sub

    Private Sub TreeView1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TreeView1.DragEnter
        If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) Then
            e.Effect = DragDropEffects.Move
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub TreeView1_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TreeView1.DragOver
        If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) = False Then Exit Sub
        Dim st As TreeView = CType(sender, TreeView)
        Dim pt As Point = CType(sender, TreeView).PointToClient(New Point(e.X, e.Y))
        Dim targetnode As TreeNode = st.GetNodeAt(pt)

        If Not (st.SelectedNode Is targetnode) Then
            st.SelectedNode = targetnode

            Dim dropnode As TreeNode = CType(e.Data.GetData("System.Windows.Forms.Treenode"), TreeNode)
            Do Until targetnode Is Nothing
                If targetnode Is dropnode Then
                    e.Effect = DragDropEffects.None
                    Exit Sub
                End If
                targetnode = targetnode.Parent
            Loop
            e.Effect = DragDropEffects.Move
        End If
    End Sub


    Private Sub TreeView1_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles TreeView1.ItemDrag
        DoDragDrop(e.Item, DragDropEffects.Move)
    End Sub
#End Region

#Region "Data"
    Private Sub loaddata()
        bindingsource1 = New BindingSource
        sqlstr = "select * from tbprogram  where applicationname='" & applicationname & "' order by parentid,myorder;"

        DbAdapter1.TbgetDataSet(sqlstr, Dataset)
        bindingsource1.DataSource = Dataset.Tables(0)
        cm = CType(BindingContext(bindingsource1), CurrencyManager)
        Dataset.Tables(0).TableName = "TbProgram"
        Dataset.Tables(0).DefaultView.Sort = "programid"

        Dataset.Relations.Add("relParentChild", Dataset.Tables(0).Columns("programid"), Dataset.Tables(0).Columns("parentid"))
        loadmenuinrole()
        bindingtext()
        cm.Position = Dataset.Tables(0).DefaultView.Find(TreeView1.Nodes(0).Tag)
    End Sub

    Private Sub loadmenuinrole()
        TreeView1.SuspendLayout()
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
        TreeView1.ResumeLayout()
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


#End Region

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        'getchanges
        Me.Validate()
        tvgetchanges()
        UpdateRecord()
    End Sub

    Private Sub tvgetchanges()
        bindingsource1.EndEdit()
        For Each n As TreeNode In TreeView1.Nodes
            Dim row = Dataset.Tables(0).Rows.Find(n.Tag)
            UpdateRow(row, n)
            gotochild(n)
        Next
    End Sub
    Private Sub gotochild(ByVal n As TreeNode)
        For Each r As TreeNode In n.Nodes
            Dim row = Dataset.Tables(0).Rows.Find(r.Tag)
            UpdateRow(row, r)            
            gotochild(r)
        Next
    End Sub
    Private Sub UpdateRow(ByRef row As DataRow, ByVal n As TreeNode)
        Dim check As Boolean = False
        Dim myparent As Integer = 0
        If n.Parent IsNot Nothing Then
            myparent = n.Parent.Tag
        End If
        row.BeginEdit()
        If row("myorder") <> n.Index Then
            row("myorder") = n.Index
            check = True
        End If
        If row("parentid") <> myparent Then
            row("parentid") = myparent
            check = True
        End If
        If check Then row.EndEdit()
    End Sub
    Private Sub UpdateRecord()

        Dim ds = Dataset.GetChanges
        If Not IsNothing(ds) Then
            If DbAdapter1.TBProgramSaveChanges(Dataset) Then
                'Dataset.Merge(ds)
                Dataset.AcceptChanges()
            End If
            If Dataset.HasErrors Then
                MessageBox.Show(Dataset.Tables(0).Rows(0).RowError)
            Else
                MsgBox("Saved")
            End If
        Else
            MessageBox.Show("Nothing to save.")
        End If
        
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If TreeView1.SelectedNode IsNot Nothing Then
            TreeView1.SelectedNode.Text = TextBox1.Text
        End If
    End Sub

    Private Sub FormMenuEditor_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        tvgetchanges()
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
   

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim tb As ToolStripButton = DirectCast(sender, ToolStripButton)
        If tb.Text = "Expand All" Then
            TreeView1.ExpandAll()
            tb.Text = "Collapse All"
        Else
            TreeView1.CollapseAll()
            tb.Text = "Expand All"
        End If
    End Sub
End Class

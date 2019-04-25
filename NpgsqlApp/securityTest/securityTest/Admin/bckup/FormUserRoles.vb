Imports System.Data
Imports securityTest.HelperClass
Public Class FormUserRoles
    Dim Dataset As DataSet
    Dim AllRoles() As String
    Dim inRoles() As String
    Dim CurrentUser As String
    Dim cm As CurrencyManager
    Dim myCheck As New ArrayList

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        BindingSource1 = New BindingSource
        loadData()
        InitRoles()
    End Sub

    Private Sub loadData()
        Dataset = New DataSet
        BindingNavigator1.BindingSource = BindingSource1
        DbAdapter1.TbgetDataSet("select * from users order by username", Dataset)
        Dataset.Tables(0).TableName = "users"
        BindingSource1.DataSource = Dataset.Tables("users")
        bindingDataGridView()
        cm = CType(BindingContext(BindingSource1), CurrencyManager)

    End Sub

    Private Sub InitRoles()
        AllRoles = DJLib.AppConfig.RoleAttribute.GetAllRoles

        Array.Sort(AllRoles)

        For i = 0 To AllRoles.Length - 1
            Dim mycheckbox As New CheckBox
            mycheckbox.Text = AllRoles(i)
            myCheck.Add(mycheckbox)

            FlowLayoutPanel1.Controls.Add(myCheck(i))
            AddHandler CType(myCheck(i), CheckBox).Click, AddressOf onCheckboxClicked
        Next
    End Sub
    Private Sub bindingDataGridView()
        DataGridView1.DataSource = BindingSource1

        DataGridView1.AutoGenerateColumns = True
        For i = 0 To 21
            DataGridView1.Columns(i).Visible = False
        Next
        DataGridView1.Columns(1).Visible = True
        DataGridView1.Columns(3).Visible = True
        DataGridView1.Columns(4).Visible = True
        DataGridView1.Columns(8).Visible = True        
    End Sub
    Public Sub UpdateRecord()
        Dim ra As Integer
        Dim message As String = String.Empty
        BindingSource1.EndEdit()

        Dim sb As New System.Text.StringBuilder
        Dim ds2 = Dataset.GetChanges

        If DbAdapter1.TBUsersSaveChanges(ds2, message, ra) Then
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

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
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
                    BindingSource1.RemoveAt(cm.Position)
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


    Private Sub BindingNavigatorAddNewItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem.Click
        Dim result As System.Windows.Forms.DialogResult = DialogCreateUser.ShowDialog()
        loadData()
    End Sub

    Private Sub ResetPasswordToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetPasswordToolStripButton.Click
        If MessageBox.Show("Reset Passsword selected record?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Dim newpassword As String = String.Empty
            If DataGridView1.SelectedRows.Count = 0 Then
                newpassword = DJLib.AppConfig.MembershipService.ResetPassword(DataGridView1.Rows(cm.Position).Cells(1).Value, "password123")
            Else
                For Each a As DataGridViewRow In DataGridView1.SelectedRows
                    newpassword = DJLib.AppConfig.MembershipService.ResetPassword(DataGridView1.Rows(a.Index).Cells(1).Value, "password123")
                Next
            End If            
            MsgBox("New Password: " & newpassword)
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        Try
            Dim dg As DataGridView = CType(sender, DataGridView)
            CurrentUser = DataGridView1.Rows(cm.Position).Cells(1).Value
            Label1.Text = String.Format("Add ""{0}"" to Roles", CurrentUser)
            'assign checkbox
            inRoles = DJLib.AppConfig.RoleAttribute.GetRolesForUser(CurrentUser)
            Dim chk As CheckBox
            For Each mycontrol In FlowLayoutPanel1.Controls
                If TypeOf (mycontrol) Is CheckBox Then
                    chk = CType(mycontrol, CheckBox)
                    chk.Checked = inRoles.Contains(chk.Text)
                End If
            Next

        Catch ex As Exception
        End Try
    End Sub

    Private Sub onCheckboxClicked(ByVal sender As Object, ByVal e As EventArgs)
        Dim chk As CheckBox = CType(sender, CheckBox)
        Dim users() As String = {CurrentUser}
        Dim roles() As String = {chk.Text}
        If chk.Checked Then
            DJLib.AppConfig.RoleAttribute.AddUsersToRoles(users, roles)
        Else
            DJLib.AppConfig.RoleAttribute.RemoveUsersFromRoles(users, roles)
        End If
    End Sub



End Class
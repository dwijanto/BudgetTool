Imports System.Data
Imports HR.HelperClass
Public Class FormUser
    Dim Dataset As DataSet
    Dim WithEvents bindingsource1 As BindingSource
    Dim CM As CurrencyManager

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        bindingsource1 = New BindingSource
        loadData()

    End Sub

    Private Sub loadData()
        Dataset = New DataSet
        DataGridView1.DataSource = bindingsource1
        DataGridView1.AutoGenerateColumns = True
        BindingNavigator1.BindingSource = bindingsource1
        'CM = CType(BindingContext(bindingsource1), CurrencyManager)
        DbAdapter1.TbgetDataSet("select * from users order by username", Dataset)
        Dataset.Tables(0).TableName = "users"
        bindingsource1.DataSource = Dataset.Tables("users")
    End Sub

    Public Sub UpdateRecord()
        Dim ra As Integer
        Dim message As String = String.Empty
        bindingsource1.EndEdit()
        If DbAdapter1.TBProgramSaveChanges(Dataset, message, ra) Then
            MessageBox.Show(ra & " Record" & IIf(ra > 1, "s", "") & " Affected.")
            If Dataset.HasErrors Then
                MessageBox.Show("Some Record(s) has been modified/deleted by other user. Records will refresh shortly.")
            End If
        End If
        loadData()
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
            For Each a As DataGridViewRow In DataGridView1.SelectedRows
                bindingsource1.RemoveAt(a.Index)
            Next
            UpdateRecord()
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
        Dim mydialogCreateuser = New DialogCreateUser
        Dim result As System.Windows.Forms.DialogResult = mydialogCreateuser.ShowDialog()
        loadData()
    End Sub

    Private Sub ResetPasswordToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetPasswordToolStripButton.Click
        If MessageBox.Show("Reset Passsword selected record?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Dim newpassword As String = String.Empty
            For Each a As DataGridViewRow In DataGridView1.SelectedRows

                newpassword = DJLib.AppConfig.MembershipService.ResetPassword(DataGridView1.Rows(a.Index).Cells(1).Value, "password123")
            Next
            MsgBox("New Password: " & newpassword)
        End If
    End Sub
End Class
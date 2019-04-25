Imports securityTest.HelperClass
Public Class FormAnimal
    Dim Dataset As DataSet
    Dim sqlstr As String
    Dim myCheck As New ArrayList
    Dim CM As CurrencyManager
    Dim applicationname As String = DJLib.AppConfig.RoleAttribute.ApplicationName
    Private newRecord As Boolean = False
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        BindingSource1 = New BindingSource
        loadData()
        CM = CType(BindingContext(BindingSource1), CurrencyManager)
    End Sub
    Private Sub onNewRow(ByVal sender As Object, ByVal e As DataTableNewRowEventArgs)
        
    End Sub


    Private Sub loadData()
        Dataset = New DataSet
        DataGridView1.DataSource = BindingSource1
        DataGridView2.DataSource = BindingSource2
        DataGridView3.DataSource = BindingSource3
        DataGridView1.AutoGenerateColumns = True
        BindingNavigator1.BindingSource = BindingSource1

        If DbAdapter1.GetRefCursorAnimal(Dataset) Then
            Dataset.Tables(0).TableName = "TBExAnimal"
            Dataset.Tables(1).TableName = "TBExPet"
            Dataset.Tables(2).TableName = "TBExPetBelonging"
            BindingSource1.DataSource = Dataset.Tables("TBExAnimal")           
            Dataset.Relations.Add("AniToPet", Dataset.Tables("TBExAnimal").Columns("exanimalid"), Dataset.Tables("TBExPet").Columns("exanimalid"))
            Dataset.Relations.Add("PetToBelonging", Dataset.Tables("TBExPet").Columns("expetid"), Dataset.Tables("TBExPetBelonging").Columns("expetid"))
            BindingSource2.DataSource = BindingSource1
            BindingSource2.DataMember = "AniToPet"
            BindingSource3.DataSource = BindingSource2
            BindingSource3.DataMember = "PetToBelonging"        
        End If
        
    End Sub

    Public Sub UpdateRecord()
        Try
            Dim ra As Integer
            Dim message As String = String.Empty
            Dim sb As New System.Text.StringBuilder

            BindingSource1.EndEdit()
            If DbAdapter1.TBAnimalSaveChanges(Dataset, message, ra) Then
                sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected. (Table Animal)")
            Else

            End If
            BindingSource2.EndEdit()
            If DbAdapter1.TBPetSaveChanges(Dataset, message, ra) Then
                If sb.Length > 0 Then
                    sb.Append(vbCrLf)
                End If
                sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected. (Table Pet)")
            Else

            End If
            BindingSource3.EndEdit()
            If DbAdapter1.TBPetBelongingSaveChanges(Dataset, message, ra) Then
                If sb.Length > 0 Then
                    sb.Append(vbCrLf)
                End If
                sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected. (Table PetBelonging)")
            Else
            End If
            If Dataset.HasErrors Then
                If sb.Length > 0 Then
                    sb.Append(vbCrLf)
                End If
                sb.Append("Some Record(s) has been modified/deleted by other user. Records will refresh shortly.")
            End If
            MessageBox.Show(sb.ToString)
        Catch ex As Exception
        End Try
        'loadData()
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


End Class
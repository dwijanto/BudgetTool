'Imports DJLib.AppConfig
Imports System.Data
Imports securityTest.HelperClass
Public Class ConcurrencyHandling
    Dim Dataset As DataSet
    Dim CM As CurrencyManager
    Dim addNew As Boolean = False

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        loadData()

    End Sub

    Private Sub loadData()
        BindingSource1 = New BindingSource
        Dataset = New DataSet
        DataGridView1.DataSource = BindingSource1
        DataGridView1.AutoGenerateColumns = True
        BindingNavigator1.BindingSource = BindingSource1
        'CM = CType(BindingContext(bindingsource1), CurrencyManager)
        DbAdapter1.TbgetDataSet("select * from tbprogram order by programid", Dataset)
        Dataset.Tables(0).TableName = "tbprogram"
        BindingSource1.DataSource = Dataset.Tables("tbprogram")
    End Sub



    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        Dim ra As Integer
        Dim message As String = String.Empty
        BindingSource1.EndEdit()
        Try
            If DbAdapter1.TBProgramSaveChanges(Dataset, message, ra, False) Then
                MessageBox.Show(ra & " Record" & IIf(ra > 1, "s", "") & " Affected.")
                If Dataset.HasErrors Then
                    MessageBox.Show("Some Record(s) has been modified/deleted by other user.Records are going to reload shortly.")
                    loadData()
                End If
            Else
                MessageBox.Show("Some Records has been deleted by other user.Reload data in progress")
                loadData()
            End If
        Catch dbcx As Data.DBConcurrencyException
            Dim response As Windows.Forms.DialogResult
            

            ProcessDialogResult(response)


        Catch ex As Exception
            MsgBox("An error was thrown while attempting to update the database.")

        End Try
        
    End Sub
    Private Function CreateMessage(ByVal cr As DataSet) As String
        Return _
            "Database: " & GetRowData(GetCurrentRowInDB(cr), Data.DataRowVersion.Default) & vbCrLf & _
            "Original: " & GetRowData(cr, Data.DataRowVersion.Original) & vbCrLf & _
            "Proposed: " & GetRowData(cr, Data.DataRowVersion.Current) & vbCrLf & _
            "Do you still want to update the database with the proposed value?"
    End Function
    '--------------------------------------------------------------------------
    ' This method loads a temporary table with current records from the database
    ' and returns the current values from the row that caused the exception.
    '--------------------------------------------------------------------------
    Private TempCustomersDataTable As New DataSet

    Private Function GetCurrentRowInDB(ByVal RowWithError As DataSet) _
        As DataSet

        'Me.CustomersTableAdapter.Fill(TempCustomersDataTable)

        'Dim currentRowInDb As DataSet = _
        '    TempCustomersDataTable.FindByCustomerID(RowWithError.CustomerID)

        'Return currentRowInDb
        Return Dataset
    End Function


    '--------------------------------------------------------------------------
    ' This method takes a CustomersRow and RowVersion 
    ' and returns a string of column values to display to the user.
    '--------------------------------------------------------------------------
    Private Function GetRowData(ByVal custRow As DataSet, _
        ByVal RowVersion As Data.DataRowVersion) As String

        Dim rowData As String = ""

        'For i As Integer = 0 To custRow.ItemArray.Length - 1
        '    rowData += custRow.Item(i, RowVersion).ToString() & " "
        'Next

        Return rowData
    End Function


    Private Sub ProcessDialogResult(ByVal response As Windows.Forms.DialogResult)

        'Select Case response

        '    Case Windows.Forms.DialogResult.Yes
        '        NorthwindDataSet.Customers.Merge(TempCustomersDataTable, True)
        '        UpdateDatabase()

        '    Case Windows.Forms.DialogResult.No
        '        NorthwindDataSet.Customers.Merge(TempCustomersDataTable)
        '        MsgBox("Update cancelled")
        'End Select
    End Sub


    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message & "Some Records has been deleted by other user.Refreshing record in progress")
    End Sub

    Private Sub RefreshToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripButton.Click
        loadData()
    End Sub


End Class
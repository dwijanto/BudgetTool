Imports System.Threading
Delegate Sub UpdateTextHandler(ByVal mystring As String)
Delegate Sub ChgTextDelegate()
Delegate Sub WorkerEnd()

Public Class FormVirtualMode
    Private customers As New System.Collections.ArrayList()
    Private customerInEdit As Customer
    Private rowInEdit As Integer = -1
    Private rowScopeCommit As Boolean = True
    Private bindingsource1 As BindingSource
    Private dataset As DataSet
    Shared _Instance As FormVirtualMode
    Private test As String

    Public Shared ReadOnly Property Instance As FormVirtualMode
        Get
            If IsNothing(_Instance) Then
                _Instance = New FormVirtualMode
            End If
            Return _Instance
        End Get
    End Property

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        Me.Text = "DataGridView virtual-mode Demo"
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub FormVirtualMode_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = False
        _Instance = Nothing
    End Sub


    Private Sub FormVirtualMode_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        dataset = New DataSet
        Dim datatable As DataTable
        datatable = makenametable()
        dataset.Tables.Add(datatable)


        bindingsource1 = New BindingSource


        'DataGridView2.DataSource = bindingsource1
        DataGridView1.VirtualMode = True
        Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Me.customers.Add(New Customer("Bon app'", "Laurence Lebihan"))
        Me.customers.Add(New Customer("Bottom-Dollar Markets", _
            "Elizabeth Lincoln"))
        Me.customers.Add(New Customer("B's Beverages", "Victoria Ashworth"))
        'Dim row As DataRow
        'Dim stopwatch As New Stopwatch

        'For i = 0 To 1000000
        '    'Me.customers.Add(New Customer("Cust" & i, "Contact" & i))
        '    row = datatable.NewRow
        '    row("Company Name") = "Cust" & i
        '    row("ContactName") = "Contact" & i
        '    datatable.Rows.Add(row)
        '    Application.DoEvents()
        'Next
        ''MsgBox("hello")
        'stopwatch.Start()
        '' Set the row count, including the row for new records.
        ''Me.DataGridView1.RowCount = 100004
        'bindingsource1.DataSource = dataset.Tables(0)
        'DataGridView2.DataSource = bindingsource1
        'stopwatch.Stop()
        'MsgBox(stopwatch.Elapsed.ToString)
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim t As Thread = New Thread(AddressOf GetData)
        t.Start()
        txtoutput.Text &= "Updates complete"

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        WorkerThread()
    End Sub
    Public Sub UpdateTextMethod(ByVal strMessage As String)
        txtOutput.Text &= (vbCrLf & Now().ToShortTimeString() & vbTab & strMessage)
    End Sub

    Public Sub UpdateResult()
        txtoutput.Text = "hellolagi"
    End Sub
    Public Sub WorkerThread()
        Cursor.Current = Cursors.WaitCursor
        Dim abc As String = "hellolagi"
        Thread.Sleep(New TimeSpan(0, 0, 0, 5, 0))

        Me.Invoke(New ChgTextDelegate(AddressOf UpdateResult))
        Thread.Sleep(New TimeSpan(0, 0, 0, 5, 0))
        Cursor.Current = Cursors.Arrow
        Me.Invoke(New WorkerEnd(AddressOf EndWorker))
    End Sub
    Private Sub EndWorker()
        txtoutput.Text &= "EndWorker" & test

    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        test = "dwi"
        Dim t As Thread = New Thread(New ThreadStart(AddressOf WorkerThread))
        t.Start()
        MsgBox("hello")
    End Sub

    Private Function makenametable() As DataTable
        Dim nametable As DataTable = New DataTable("Customer")
        Dim CompanyName As DataColumn = New DataColumn
        With CompanyName
            .DataType = System.Type.GetType("System.String")
            .ColumnName = "Company Name"
        End With

        Dim ContactName As DataColumn = New DataColumn
        With ContactName
            .DataType = System.Type.GetType("System.String")
            .ColumnName = "Contact Name"
        End With
        nametable.Columns.Add(CompanyName)
        nametable.Columns.Add("ContactName", System.Type.GetType("System.String"))
        Return nametable
    End Function


    Private Sub DataGridView1_CancelRowEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.QuestionEventArgs) Handles DataGridView1.CancelRowEdit
        If Me.rowInEdit = Me.DataGridView1.Rows.Count - 2 AndAlso _
            Me.rowInEdit = Me.customers.Count Then

            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding Customer object with a new, empty one.
            Me.customerInEdit = New Customer()

        Else

            ' If the user has canceled the edit of an existing row, 
            ' release the corresponding Customer object.
            Me.customerInEdit = Nothing
            Me.rowInEdit = -1

        End If

    End Sub

    Private Sub DataGridView1_CellValueNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles DataGridView1.CellValueNeeded
        If e.RowIndex = Me.DataGridView1.RowCount - 1 Then
            Return
        End If

        Dim customerTmp As Customer = Nothing

        ' Store a reference to the Customer object for the row being painted.
        If e.RowIndex = rowInEdit Then
            customerTmp = Me.customerInEdit
        Else
            customerTmp = CType(Me.customers(e.RowIndex), Customer)
        End If

        ' Set the cell value to paint using the Customer object retrieved.
        Select Case Me.DataGridView1.Columns(e.ColumnIndex).HeaderText
            Case "Company Name"
                e.Value = customerTmp.CompanyName

            Case "Contact Name"
                e.Value = customerTmp.ContactName
        End Select


    End Sub

    Private Sub DataGridView1_CellValuePushed(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles DataGridView1.CellValuePushed
        Dim customerTmp As Customer = Nothing

        ' Store a reference to the Customer object for the row being edited.
        If e.RowIndex < Me.customers.Count Then

            ' If the user is editing a new row, create a new Customer object.
            If Me.customerInEdit Is Nothing Then
                Me.customerInEdit = New Customer( _
                    CType(Me.customers(e.RowIndex), Customer).CompanyName, _
                    CType(Me.customers(e.RowIndex), Customer).ContactName)
            End If
            customerTmp = Me.customerInEdit
            Me.rowInEdit = e.RowIndex

        Else
            customerTmp = Me.customerInEdit
        End If

        ' Set the appropriate Customer property to the cell value entered.
        Dim newValue As String = TryCast(e.Value, String)
        Select Case Me.DataGridView1.Columns(e.ColumnIndex).HeaderText
            Case "Company Name"
                customerTmp.CompanyName = newValue
            Case "Contact Name"
                customerTmp.ContactName = newValue
        End Select

    End Sub

    Private Sub DataGridView1_NewRowNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DataGridView1.NewRowNeeded
        Me.customerInEdit = New Customer()
        Me.rowInEdit = Me.DataGridView1.Rows.Count - 1

    End Sub

    Private Sub DataGridView1_RowDirtyStateNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.QuestionEventArgs) Handles DataGridView1.RowDirtyStateNeeded
        If Not rowScopeCommit Then

            ' In cell-level commit scope, indicate whether the value
            ' of the current cell has been modified.
            e.Response = Me.DataGridView1.IsCurrentCellDirty

        End If

    End Sub

    Private Sub DataGridView1_RowValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.RowValidated
        If e.RowIndex >= Me.customers.Count AndAlso _
           e.RowIndex <> Me.DataGridView1.Rows.Count - 1 Then

            ' Add the new Customer object to the data store.
            Me.customers.Add(Me.customerInEdit)
            Me.customerInEdit = Nothing
            Me.rowInEdit = -1

        ElseIf (Me.customerInEdit IsNot Nothing) AndAlso _
            e.RowIndex < Me.customers.Count Then

            ' Save the modified Customer object in the data store.
            Me.customers(e.RowIndex) = Me.customerInEdit
            Me.customerInEdit = Nothing
            Me.rowInEdit = -1

        ElseIf Me.DataGridView1.ContainsFocus Then

            Me.customerInEdit = Nothing
            Me.rowInEdit = -1

        End If

    End Sub

    Private Sub DataGridView1_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DataGridView1.UserDeletingRow
        If e.Row.Index < Me.customers.Count Then

            ' If the user has deleted an existing row, remove the 
            ' corresponding Customer object from the data store.
            Me.customers.RemoveAt(e.Row.Index)

        End If

        If e.Row.Index = Me.rowInEdit Then

            ' If the user has deleted a newly created row, release
            ' the corresponding Customer object. 
            Me.rowInEdit = -1
            Me.customerInEdit = Nothing

        End If

    End Sub




End Class
Public Class Customer

    Private companyNameValue As String
    Private contactNameValue As String

    Public Sub New()
        ' Leave fields empty.
    End Sub

    Public Sub New(ByVal companyName As String, ByVal contactName As String)
        companyNameValue = companyName
        contactNameValue = contactName
    End Sub

    Public Property CompanyName() As String
        Get
            Return companyNameValue
        End Get
        Set(ByVal value As String)
            companyNameValue = value
        End Set
    End Property

    Public Property ContactName() As String
        Get
            Return contactNameValue
        End Get
        Set(ByVal value As String)
            contactNameValue = value
        End Set
    End Property

End Class

Module BackgroundMethods
    Public Sub GetData()
        WaitForData("Message 1")
        WaitForData("Message 2")
    End Sub
    Public Sub WaitForData(ByVal strMessage As String)

        Thread.Sleep(2000)
        Dim f As FormVirtualMode = My.Application.OpenForms("formvirtualmode")
        f.Invoke(New UpdateTextHandler(AddressOf f.UpdateTextMethod), _
                 New Object() {strMessage})

    End Sub
End Module
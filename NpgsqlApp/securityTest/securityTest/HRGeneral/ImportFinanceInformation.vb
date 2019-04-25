Imports System.ComponentModel
Imports HR.HelperClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports DJLib.Dbtools

Public Class ImportFinanceInformation
    Private WithEvents BackgroundWorker1 As New BackgroundWorker
    Dim FileName As String = String.Empty
    Dim Status As Boolean = False
    Dim Dataset1 As DataSet
    Dim ConnectionString As String = dbtools1.getConnectionString
    Dim aprocesses() As Process = Nothing '= Process.GetProcesses
    Dim aprocess As Process = Nothing

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not (BackgroundWorker1.IsBusy) Then
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                FileName = OpenFileDialog1.FileName
                TextBox1.Text = FileName
                Try
                    BackgroundWorker1.WorkerReportsProgress = True
                    BackgroundWorker1.WorkerSupportsCancellation = True
                    BackgroundWorker1.RunWorkerAsync()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Select Case e.ProgressPercentage
            Case 2
                TextBox2.Text = e.UserState
            Case 3
                TextBox3.Text = e.UserState
        End Select
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        BackgroundWorker1.ReportProgress(3, TextBox3.Text & "Start")

        Dim errMsg As String = String.Empty
        Status = ImportData(FileName, errMsg)
        If Status Then
            BackgroundWorker1.ReportProgress(2, TextBox2.Text & " Done.")
        Else
            BackgroundWorker1.ReportProgress(3, "Error::" & errMsg)
        End If
    End Sub

    Private Function ImportData(ByVal FileName As String, Optional ByRef errMessage As String = "") As Boolean
        Dim myreturn As Boolean = False
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim oRange As Excel.Range = Nothing

        Dim iRows As Long = 0

        'Need these variable to kill excel
        Dim aprocesses() As Process = Nothing '= Process.GetProcesses
        Dim aprocess As Process = Nothing

        Try
            BackgroundWorker1.ReportProgress(2, "Preparing Data...")
            Dataset1 = New DataSet
            BackgroundWorker1.ReportProgress(3, "Opening Excel File....")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            Application.DoEvents()
            oXl.Visible = True
            'get process pid
            aprocesses = Process.GetProcesses
            For i = 0 To aprocesses.GetUpperBound(0)
                If aprocesses(i).MainWindowHandle.ToString = oXl.Hwnd.ToString Then
                    aprocess = aprocesses(i)
                    Exit For
                End If
                Application.DoEvents()
            Next
            oXl.Visible = False
            oXl.DisplayAlerts = False
            oWb = oXl.Workbooks.Open(FileName)

            Dim checksheet As Boolean = False
            For i = 1 To oWb.Worksheets.Count
                If oWb.Worksheets(i).name = "input sheet" Then
                    checksheet = True
                End If
            Next
            If Not checksheet Then
                Throw New System.Exception("Excel File is not valid!")
            End If

            Dim stopwatch As New Stopwatch
            stopwatch.Start()

            BackgroundWorker1.ReportProgress(2, "Select Worksheet ""input sheet""")
            oSheet = oWb.Worksheets("input sheet")


            iRows = oSheet.UsedRange.Rows.Count
            BackgroundWorker1.ReportProgress(2, "Connect to Db...")
            BackgroundWorker1.ReportProgress(3, "Preparing Tables...")

            Dim sqlstr As String = "select expensesdetailid,myyear,expensesdetailtxid  from expensesdetailtx where myyear = " & DateTimePicker1.Value.Year & ";"
            If Not dbtools1.getDataSet(sqlstr, Dataset1, errMessage) Then
                Return myreturn
            End If

            Dataset1.Tables(0).TableName = "expensesdetailtx"
            Dim keys0(0) As DataColumn
            keys0(0) = Dataset1.Tables(0).Columns(0)
            Dataset1.Tables(0).PrimaryKey = keys0

            Dim sapaccountid As Integer = 0
            Dim sapindexid As Integer = 0
            Dim sapccid As Integer = 0
            Dim indexcostcenterid As Integer = 0
            Dim deptid As Integer = 0
            Dim indexcostcenterdeptid As Integer = 0
            Dim expensesnatureid As Integer = 0
            Dim sapaccnameid As Integer = 0
            Dim accexpensesid As Integer = 0
            Dim expensesdetailid As Integer = 0
            Dim myDictionary As New Dictionary(Of String, String)
            Dim stringbuilder1 As New StringBuilder

            For i = 10 To iRows
                If IsNothing(oSheet.Cells(i, 1).value) Then
                    Exit For
                End If
                BackgroundWorker1.ReportProgress(2, "Row " & i & " of " & iRows)

                If oSheet.Cells(i, 1).Value.ToString <> "HEADCOUNT" Then


                    '1. Check table sapaccount if not avail then create 
                    sapaccountid = DbAdapter1.getSapAccountid(oSheet.Cells(i, 3).value.ToString)
                    '2. Check table sapindex if not avail then create
                    sapindexid = DbAdapter1.getSapIndexid(sapaccountid, oSheet.Cells(i, 5).value.ToString)
                    '3. Check table CostCenter if not avail then create
                    'Debug.WriteLine("{0} {1}", oSheet.Cells(i, 4).value.ToString, oSheet.Cells(i, 7).value.ToString)
                    sapccid = DbAdapter1.getsapccid(oSheet.Cells(i, 4).value.ToString, oSheet.Cells(i, 7).value.ToString)

                    '4. Check table IndexCostCenter if not avail then create
                    indexcostcenterid = DbAdapter1.getindexcostcenterid(sapindexid, sapccid)
                    '5. Check table dept if not avail then create
                    deptid = DbAdapter1.getdeptid(oSheet.Cells(i, 6).value.ToString)
                    '6. Check table indexcostcenterdept if not avail then create
                    indexcostcenterdeptid = DbAdapter1.getIndexCostCenterDeptId(indexcostcenterid, deptid)
                    '7. Check table expensesnature if not avail then create
                    expensesnatureid = DbAdapter1.getExpensesNatureId(oSheet.Cells(i, 2).value.ToString)
                    '8. Check table sapaccname if not avail then create
                    sapaccnameid = DbAdapter1.getSapAccNameId(oSheet.Cells(i, 1).value.ToString)
                    '9. Check table accexpenses if not avail then create
                    accexpensesid = DbAdapter1.getAccExpensesId(sapaccnameid, expensesnatureid)
                    '10. check table expensesdetail if not avail then create
                    expensesdetailid = DbAdapter1.getExpensesDetailid(accexpensesid, indexcostcenterdeptid)
                    'Find Table ExpensesDetail using accexpensesid and indexcostcenterdeptid

                    Try
                        Dim mykey = expensesdetailid
                        myDictionary.Add(mykey, mykey)
                        Dim pkey(0) As Object
                        pkey(0) = expensesdetailid

                        Dim DataRow1 As DataRow = Dataset1.Tables(0).Rows.Find(pkey)
                        If DataRow1 Is Nothing Then
                            stringbuilder1.Append(expensesdetailid & vbTab)
                            stringbuilder1.Append(DateTimePicker1.Value.Year & vbCrLf)
                        End If

                    Catch ex As Exception

                    End Try

                End If

                Application.DoEvents()

            Next

            'copy expensesdetail

            sqlstr = "copy expensesdetailtx(expensesdetailid,myyear) from stdin;"
            BackgroundWorker1.ReportProgress(2, "Copy To Db (ExpensesDetailtx)")
            If stringbuilder1.ToString <> "" Then
                errMessage = dbtools1.copy(sqlstr, stringbuilder1.ToString, myreturn)
                BackgroundWorker1.ReportProgress(2, "Copy To Db.")
            Else
                BackgroundWorker1.ReportProgress(2, "Nothing to Copy.")
                myreturn = True
            End If
            stopwatch.Stop()
            BackgroundWorker1.ReportProgress(2, "Elapsed Time: " & Format(stopwatch.Elapsed.Minutes, "00") & ":" & Format(stopwatch.Elapsed.Seconds, "00") & "." & stopwatch.Elapsed.Milliseconds.ToString)

            BackgroundWorker1.ReportProgress(3, "")

        Catch ex As Exception
            errMessage = ex.Message
        Finally
            oXl.Quit()
            'releaseComObject(oRange)
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)

            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                If Not aprocess Is Nothing Then
                    aprocess.Kill()
                End If
            Catch ex As Exception
            End Try

        End Try
        Return myreturn
    End Function

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If CheckBox1.Checked Then
            Me.Close()
        End If
    End Sub

    Private Sub ImportFinanceInformation_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If (BackgroundWorker1.IsBusy) Then
            MsgBox("Please wait until the current process is finished")
            e.Cancel = True
        End If
    End Sub



    Private Sub ImportFinanceInformation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = BudgetYear
    End Sub
End Class
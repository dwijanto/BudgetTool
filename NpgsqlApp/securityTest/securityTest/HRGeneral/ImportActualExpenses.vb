Imports System.ComponentModel
Imports HR.HelperClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports DJLib.Dbtools

Public Class ImportActualExpenses
    Private WithEvents BackgroundWorker1 As New BackgroundWorker
    Dim FileName As String = String.Empty
    Dim Status As Boolean = False
    Dim Dataset1 As DataSet
    Dim ConnectionString As String = dbtools1.getConnectionString
    Dim aprocesses() As Process = Nothing '= Process.GetProcesses
    Dim aprocess As Process = Nothing
    Dim myyear As Integer = 0

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not (BackgroundWorker1.IsBusy) Then
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                FileName = OpenFileDialog1.FileName
                TextBox1.Text = FileName
                Try
                    myyear = DateTimePicker1.Value.Date.Year
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
        Dim iCols As Long = 0

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
                If oWb.Worksheets(i).name = "PL" Then
                    checksheet = True
                End If
            Next
            If Not checksheet Then
                Throw New System.Exception("Excel File is not valid!")
            End If

            Dim stopwatch As New Stopwatch
            stopwatch.Start()

            BackgroundWorker1.ReportProgress(2, "Select Worksheet ""PL""")
            oSheet = oWb.Worksheets("PL")


            iRows = oSheet.UsedRange.Rows.Count
            iCols = oSheet.UsedRange.Columns.Count

            BackgroundWorker1.ReportProgress(2, "Connect to Db...")
            BackgroundWorker1.ReportProgress(3, "Preparing Tables...")

            Dim sqlstr As String = "select accountnameid,accountname from actaccountname;"
            If Not dbtools1.getDataSet(sqlstr, Dataset1, errMessage) Then
                Return myreturn
            End If

            Dataset1.Tables(0).TableName = "actaccountname"
            Dim keys0(0) As DataColumn
            keys0(0) = Dataset1.Tables(0).Columns(0)
            Dataset1.Tables(0).PrimaryKey = keys0





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

            Dim AccountNameDict As New Dictionary(Of String, Integer)
            Dim myMonth As String = "JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC"
            assignDictionary(Dataset1, AccountNameDict)
            Dim accountnameid As Long
            For i = 5 To iRows
                

                If IsNothing(oSheet.Cells(i, 1).value) Then
                    Exit For
                End If
                BackgroundWorker1.ReportProgress(2, "Row " & i & " of " & iRows)
                Dim regionshort As String = oSheet.Cells(i, 3).value.ToString
                Dim sapaccountid As String = oSheet.Cells(i, 5).value.ToString
                Dim account As String = oSheet.Cells(i, 7).value.ToString
                Dim sapcc As String = oSheet.Cells(i, 8).value.ToString
                Dim mydate As Date

                If oSheet.Cells(i, 1).value.ToString = "Y" And oSheet.Cells(i, 2).value = "Actual" Then
                    'get accountname
                    Dim accountname As String = oSheet.Cells(i, 6).value.ToString
                    Try
                        accountnameid = AccountNameDict(accountname)
                    Catch ex As Exception
                        accountnameid = DbAdapter1.createAccountName(accountname)
                        AccountNameDict.Add(accountname, accountnameid)
                    End Try
                    For col = 1 To iCols
                        If myMonth.Contains(oSheet.Cells(4, col).value.ToString.ToUpper) Then
                            Select Case oSheet.Cells(4, col).value.ToString.ToUpper
                                Case "JAN"
                                    mydate = CDate(myyear & "-1-1")
                                Case "FEB"
                                    mydate = CDate(myyear & "-2-1")
                                Case "MAR"
                                    mydate = CDate(myyear & "-3-1")
                                Case "APR"
                                    mydate = CDate(myyear & "-4-1")
                                Case "MAY"
                                    mydate = CDate(myyear & "-5-1")
                                Case "JUN"
                                    mydate = CDate(myyear & "-6-1")
                                Case "JUL"
                                    mydate = CDate(myyear & "-7-1")
                                Case "AUG"
                                    mydate = CDate(myyear & "-8-1")
                                Case "SEP"
                                    mydate = CDate(myyear & "-9-1")
                                Case "OCT"
                                    mydate = CDate(myyear & "-10-1")
                                Case "NOV"
                                    mydate = CDate(myyear & "-11-1")
                                Case "DEC"
                                    mydate = CDate(myyear & "-12-1")
                            End Select
                            Dim myvalue As String = oSheet.Cells(i, col).value.ToString
                            'create record
                            createrecord(stringbuilder1, myyear, regionshort, sapaccountid, accountnameid, account, sapcc, DateFormatyyyyMMdd(mydate), myvalue)
                        End If

                        Application.DoEvents()

                    Next
                End If
            Next
            'copy expensesdetail
            sqlstr = "delete from acttx where myyear = " & myyear & "; copy acttx(myyear,regionshort,sapaccountid,accountnameid,sapaccount,sapcc,mydate,amount) from stdin;"
            BackgroundWorker1.ReportProgress(2, "Copy To Db (acttx)")
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

    Private Sub assignDictionary(ByRef Dataset1 As DataSet, ByRef AccountNameDict As Dictionary(Of String, Integer))
        Dim q = From rec In Dataset1.Tables(0)
                Select rec
        For Each r In q
            AccountNameDict.Add(r.Item("accountname"), r.Item("accountnameid"))
        Next
    End Sub
    'Private Sub createrecord(ByRef stringBuilder1 As StringBuilder, ByVal personexpensesid As Integer, ByVal amount As Double, ByVal myverid As Integer, ByVal mydate As String)
    '    stringBuilder1.Append(personexpensesid & vbTab)
    '    stringBuilder1.Append(amount & vbTab)
    '    stringBuilder1.Append(myverid & vbTab)
    '    stringBuilder1.Append(mydate & vbCrLf)
    'End Sub

    Private Sub createrecord(ByVal stringbuilder1 As StringBuilder, ByVal myyear As Integer, ByVal regionshort As String, ByVal sapaccountid As String, ByVal accountnameid As Long, ByVal account As String, ByVal sapcc As String, ByVal txdate As String, ByVal myvalue As String)
        stringbuilder1.Append(myyear & vbTab)
        stringbuilder1.Append(regionshort & vbTab)
        stringbuilder1.Append(sapaccountid & vbTab)
        stringbuilder1.Append(accountnameid & vbTab)
        stringbuilder1.Append(account & vbTab)
        stringbuilder1.Append(sapcc & vbTab)
        stringbuilder1.Append(txdate & vbTab)
        stringbuilder1.Append(myvalue & vbCrLf)
    End Sub

End Class
Imports System.ComponentModel
Imports HR.HelperClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports DJLib.Dbtools
Imports DJLib.ExcelStuff

Public Class FGenerateReport001
    Private WithEvents BackgroundWorker1 As New BackgroundWorker
    Dim FileName As String = String.Empty
    Dim Status As Boolean = False
    Dim Dataset1 As DataSet
    Dim ConnectionString As String = dbtools1.getConnectionString
    Dim aprocesses() As Process = Nothing '= Process.GetProcesses
    Dim aprocess As Process = Nothing
    Dim VersionId As Integer = 0
    Dim Myyear As Integer = 0
    Dim MyRegionId As Integer
    Dim ShortRegion As String
    Dim RegionName As String



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Then
            MsgBox("Please select from list!")
            ComboBox1.Select()
            Exit Sub
        End If
        VersionId = ComboBox1.SelectedValue
        Myyear = DateTimePicker1.Value.Year
        'Button1.Enabled = False


        'MyRegionId = ComboBox2.SelectedValue
        'If MyRegionId > 0 Then
        '    ShortRegion = RegionDict(MyRegionId)
        'End If
        'RegionName = ComboBox2.Text.ToString

        If Not (BackgroundWorker1.IsBusy) Then
            'Dim FileName As String = String.Empty
            Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
            DirectoryBrowser.Description = "Which directory do you want to use?"

            If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                FileName = DirectoryBrowser.SelectedPath & "\" & "Budget-" & dbtools1.Region & "-" & ComboBox1.Text & "-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
                'Myform1 = New MyForm With {.combobox1 = ComboBox1.SelectedItem.ToString}
                'Label1.Text = ""
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

        Timer1.Start()
        Dim errMsg As String = String.Empty
        Status = GenerateExcel(FileName, errMsg)
        If Status Then
            BackgroundWorker1.ReportProgress(2, TextBox2.Text & " Done.")
        Else
            BackgroundWorker1.ReportProgress(3, errMsg)
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Timer1.Stop()
        FormMenu.setBubbleMessage("Export To Excel", "Done")
        If Status Then
            'If CheckBox1.Checked Then
            '    Me.Close()
            'End If
        End If
        If Status Then
            If MsgBox("File name: " & FileName & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                Process.Start(FileName)
            End If
        End If
        Button1.Enabled = True
    End Sub

    Private Sub MySub_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If (BackgroundWorker1.IsBusy) Then
            MsgBox("Please wait until the current process is finished")
            e.Cancel = True
        End If
    End Sub

    Private Sub MySub_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MyRegionId = dbtools1.RegionId
        'Dim sqlstr As String = "select 'Last Version',0 union all select verid,hrver from ver order by myorder;"
        Dim sqlstr As String = "(select 0 as verid,'Last Version' as hrver) union all (select verid,hrver from ver order by myorder)"
        dbtools1.FillComboboxDataSource(ComboBox1, sqlstr)
        ShortRegion = dbtools1.Region
        RegionName = dbtools1.RegionName

        If ShortRegion = "HK" Then         
            sqlstr = "select 0,'All Region' union all (select regionid,regionname from region order by regionname);"
            dbtools1.FillComboboxDataSource(ComboBox2, sqlstr)
            Label3.Visible = True
            ComboBox2.Visible = True
            MyRegionId = 0
            Dataset1 = New DataSet
            Dim errmessage As String = ""
            sqlstr = "Select * from region"
            If Not dbtools1.getDataSet(sqlstr, Dataset1, errmessage) Then
                MessageBox.Show(errmessage)
            End If
            Dataset1.Tables(0).TableName = "Region"
        Else

        End If



        DateTimePicker1.Value = BudgetYear
    End Sub
    Private Function GenerateExcel(ByRef FileName As String, ByRef errorMsg As String) As Boolean

        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False
        Dim dataset1 As New DataSet

        Dim StopWatch As New Stopwatch
        StopWatch.Start()

        'Open Excel
        Application.DoEvents()

        Cursor.Current = Cursors.WaitCursor
        Dim source As String = FileName
        Dim StringBuilder1 As New System.Text.StringBuilder

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim Sqlstr As String = String.Empty

        'Need these variable to kill excel
        Dim aprocesses() As Process = Nothing '= Process.GetProcesses
        Dim aprocess As Process = Nothing
        Try
            'Create Object Excel 
            BackgroundWorker1.ReportProgress(3, "Creating Excel..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            BackgroundWorker1.ReportProgress(3, TextBox3.Text & "Get PID..")
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
            BackgroundWorker1.ReportProgress(2, "Opening Template...")
            BackgroundWorker1.ReportProgress(3, "Generating Report..")
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ExcelTemplate.xltx")
            Dim iSheetDAta As Integer = 2
            'Loop for chart
            'Go to worksheetData
            oSheet = oWb.Worksheets(iSheetDAta)
            oWb.Worksheets(iSheetDAta).select()
            BackgroundWorker1.ReportProgress(2, "DB Query...")

            Call QueryData(oWb, iSheetDAta)


            oWb.Worksheets(iSheetDAta).select()
            oSheet = oWb.Worksheets(iSheetDAta)
            If IsNothing(oSheet.Cells(2, 1).value) Then
                Throw New System.Exception("Data not available!")
            End If
            'Check data

            oWb.Names.Add(Name:="DBRangeAll", RefersToR1C1:="=OFFSET(" & oSheet.Name & "!R1C1,0,0,COUNTA(" & oSheet.Name & "!C1),COUNTA(" & oSheet.Name & "!R1))")
            oSheet.Name = "RAW_DATA"

            'Generate Chart&Pivot start from worksheet 2
            iSheetDAta = 1
            BackgroundWorker1.ReportProgress(2, "Generating PivotTable...")

            Call GeneratePivotTable(oWb, iSheetDAta, errorMsg)
            If errorMsg <> "" Then
                Throw New System.Exception(errorMsg)
            End If
            oWb.Worksheets(1).select()
            oSheet = oWb.Worksheets(1)
            oSheet.Range("B9").Select()
            oXl.ActiveWindow.FreezePanes = True

            StopWatch.Stop()

            BackgroundWorker1.ReportProgress(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            FileName = ValidateFileName(System.IO.Path.GetDirectoryName(source), source)
            BackgroundWorker1.ReportProgress(3, "Saving File...")
            oXl.DisplayAlerts = False
            'oWb.Worksheets("DBAll").delete()
            oWb.SaveAs(FileName)
            oXl.DisplayAlerts = True
            result = True
        Catch ex As Exception
            errormsg = ex.Message
        Finally
            BackgroundWorker1.ReportProgress(3, "Releasing Memory...")
            'clear excel from memory
            oXl.Quit()
            'oXl.Visible = True
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
            Cursor.Current = Cursors.Default
            BackgroundWorker1.ReportProgress(3, "")
        End Try


        'If result Then
        '    If MsgBox("File name: " & FileName & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
        '        Process.Start(FileName)
        '    End If
        'End If
        'Button1.Enabled = True
        Return result

    End Function

    Public Sub QueryData(ByRef owb As Excel.Workbook, ByVal isheet As Integer)
        Dim sqlstr As String = String.Empty
        Dim stringbuilder1 As New System.Text.StringBuilder

        'Check Worksheet
        For i = owb.Worksheets.Count To isheet - 1
            owb.Worksheets.Add(After:=owb.Worksheets(i))
        Next
        Dim firstdate As String = "'" & DateTimePicker1.Value.Year & "-1-1'"
        Dim lastdate As String = "'" & DateTimePicker1.Value.Year & "-12-31'"



        'GET DATA
        If MyRegionId <> 0 Then
            If VersionId = 0 Then
                VersionId = DbAdapter1.getlastversion(DateTimePicker1.Value.Year, "CV" & dbtools1.Region, dbtools1.Region)
            End If
            sqlstr = "select * from fgetbudgettx(" & VersionId & "," & firstdate & "::date," & lastdate & "::date," & MyRegionId & ") as b(sapaccname character varying,expensesnature character varying, sapaccount character varying, sapaccid character varying, sapcc character varying, dept character varying, category character varying, joindate date,enddate date, personname character varying, othername character varying, title character varying, expat boolean,mydate date,amount numeric,headcount numeric,regionname character varying,crcytype character varying,version character varying)"
            stringbuilder1.Append(sqlstr)
        Else
            If VersionId = 0 Then
                'Loop for each country
                Dim tmp As String = String.Empty
                For Each rec In Dataset1.Tables(0).Rows
                    Try
                        VersionId = DbAdapter1.getlastversion(DateTimePicker1.Value.Year, "CV" & rec.item("regionshortname"), rec.item("regionshortname"))
                        sqlstr = "(select * from fgetbudgettx(" & VersionId & "," & firstdate & "::date," & lastdate & "::date," & rec.item("regionid") & ") as b(sapaccname character varying,expensesnature character varying, sapaccount character varying, sapaccid character varying, sapcc character varying, dept character varying, category character varying, joindate date,enddate date, personname character varying, othername character varying, title character varying, expat boolean,mydate date,amount numeric,headcount numeric,regionname character varying,crcytype character varying,version character varying))"
                        'tmp = tmp + IIf(tmp = "", "", " union all ") + sqlstr
                        stringbuilder1.Append(IIf(stringbuilder1.ToString = "", "", " union all ") & sqlstr)
                    Catch ex As Exception

                    End Try
                Next
            Else
                sqlstr = "select * from fgetbudgettx(" & VersionId & "," & firstdate & "::date," & lastdate & "::date) as b(sapaccname character varying, expensesnature character varying, sapaccount character varying, sapaccid character varying, sapcc character varying, dept character varying, category character varying, joindate date,enddate date, personname character varying, othername character varying, title character varying, expat boolean,mydate date,amount numeric,headcount numeric,regionname character varying,crcytype character varying,version character varying)"
                stringbuilder1.Append(sqlstr)
            End If

        End If
        DJLib.ExcelStuff.FillDataSource(owb, isheet, stringbuilder1.ToString, dbtools1)
    End Sub

    Private Sub GeneratePivotTable(ByVal oWb As Excel.Workbook, ByVal iSheet As Integer, ByRef errMsg As String)
        Dim osheet As Excel.Worksheet
        Try

            osheet = oWb.Worksheets(iSheet)
            oWb.Worksheets(iSheet).select()

            'oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeAll").CreatePivotTable(osheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
            oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeAll").CreatePivotTable(osheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion12)
            'osheet.Cells(1, 1) = "as of " & Today.ToString("MMM dd yyyy")
            'osheet.PivotTables("PivotTable1").columngrand = False
            oWb.ShowPivotTableFieldList = False
            'osheet.PivotTables("PivotTable1").rowgrand = False
            'osheet.PivotTables("PivotTable1").ColumnGrand = False
            osheet.PivotTables("PivotTable1").AllowMultipleFilters = True
            osheet.PivotTables("PivotTable1").ingriddropzones = True
            osheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)

            'Compact mode
            'osheet.PivotTables("PivotTable1").RowAxisLayout(Excel.XlLayoutRowType.xlCompactRow)
            osheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight8"

            'Calculated Field if any
            'osheet.PivotTables("PivotTable1").CalculatedFields.Add("Flagbottleneck", "=IF(filter>0,1,0)", True)


            'add PageField
            osheet.PivotTables("PivotTable1").PivotFields("regionname").orientation = Excel.XlPivotFieldOrientation.xlPageField
            osheet.PivotTables("PivotTable1").PivotFields("regionname").currentpage = "All"
            osheet.PivotTables("PivotTable1").PivotFields("regionname").caption = "Region Name"


            'add Rowfields
            'osheet.PivotTables("PivotTable1").PivotFields("categorybudget").orientation = Excel.XlPivotFieldOrientation.xlRowField
            osheet.PivotTables("PivotTable1").PivotFields("sapaccname").orientation = Excel.XlPivotFieldOrientation.xlRowField
            osheet.PivotTables("PivotTable1").PivotFields("expensesnature").orientation = Excel.XlPivotFieldOrientation.xlRowField
            osheet.PivotTables("PivotTable1").PivotFields("sapaccount").orientation = Excel.XlPivotFieldOrientation.xlRowField

            osheet.PivotTables("PivotTable1").PivotFields("sapcc").orientation = Excel.XlPivotFieldOrientation.xlRowField
            osheet.PivotTables("PivotTable1").PivotFields("sapaccid").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'osheet.PivotTables("PivotTable1").PivotFields("sapcc").layoutform = Excel.XlLayoutFormType.xlTabular
            osheet.PivotTables("PivotTable1").PivotFields("dept").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'osheet.PivotTables("PivotTable1").PivotFields("dept").layoutform = Excel.XlLayoutFormType.xlTabular
            osheet.PivotTables("PivotTable1").PivotFields("crcytype").orientation = Excel.XlPivotFieldOrientation.xlRowField
            osheet.PivotTables("PivotTable1").PivotFields("category").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'osheet.PivotTables("PivotTable1").PivotFields("category").layoutform = Excel.XlLayoutFormType.xlTabular
            osheet.PivotTables("PivotTable1").PivotFields("joindate").orientation = Excel.XlPivotFieldOrientation.xlRowField
            osheet.PivotTables("PivotTable1").PivotFields("enddate").orientation = Excel.XlPivotFieldOrientation.xlRowField
            osheet.PivotTables("PivotTable1").PivotFields("personname").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'osheet.PivotTables("PivotTable1").PivotFields("personname").layoutform = Excel.XlLayoutFormType.xlTabular
            osheet.PivotTables("PivotTable1").PivotFields("othername").orientation = Excel.XlPivotFieldOrientation.xlRowField
            osheet.PivotTables("PivotTable1").PivotFields("title").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'osheet.PivotTables("PivotTable1").PivotFields("othername").layoutform = Excel.XlLayoutFormType.xlTabular
            osheet.PivotTables("PivotTable1").PivotFields("mydate").orientation = Excel.XlPivotFieldOrientation.xlRowField
            'osheet.PivotTables("PivotTable1").PivotFields("mydate").layoutform = Excel.XlLayoutFormType.xlTabular
            osheet.Range("N7").Group(Start:=True, End:=True, Periods:={False, False, False, False, True, True, True})
            osheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlHidden
            osheet.PivotTables("PivotTable1").PivotFields("Quarters").Orientation = Excel.XlPivotFieldOrientation.xlHidden
            'add columnfield
            osheet.PivotTables("PivotTable1").PivotFields("mydate").orientation = Excel.XlPivotFieldOrientation.xlColumnField
            osheet.PivotTables("PivotTable1").PivotFields("mydate").caption = "Months"
            'osheet.PivotTables("PivotTable1").PivotFields("monthly").orientation = Excel.XlPivotFieldOrientation.xlColumnField
            'osheet.PivotTables("PivotTable1").PivotFields("monthly").numberformat = "MMM-yy"
            'osheet.PivotTables("PivotTable1").PivotFields("monthly").Caption = "Month"

            'remove subtotal
            'osheet.PivotTables("PivotTable1").PivotFields("sapaccname").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").PivotFields("expensesnature").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").pivotfields("sapaccount").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").pivotfields("sapcc").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").pivotfields("sapaccid").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").PivotFields("dept").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").pivotfields("crcytype").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").PivotFields("joindate").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").PivotFields("enddate").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").PivotFields("category").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").PivotFields("personname").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").PivotFields("title").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            osheet.PivotTables("PivotTable1").PivotFields("othername").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}


            'add datafield
            osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("headcount"), "HC", Excel.XlConsolidationFunction.xlSum)
            osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amount"), "BUD", Excel.XlConsolidationFunction.xlSum)

            'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("orderconfirmed"), "Order Confirmed", Excel.XlConsolidationFunction.xlSum)
            'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("orderunconfirmed"), "Order Unconfirmed", Excel.XlConsolidationFunction.xlSum)
            'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("forecast"), " Forecast", Excel.XlConsolidationFunction.xlSum)
            'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("bottleneck"), " Bottleneck", Excel.XlConsolidationFunction.xlSum)
            'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("Flagbottleneck"), "Flag for demand exceeds bottleneck", Excel.XlConsolidationFunction.xlSum)

            osheet.PivotTables("PivotTable1").PivotFields("BUD").NumberFormat = "#,##0"


            'osheet.PivotTables("PivotTable1").PivotFields("BottleNeck < TTL Demand").NumberFormat = "#,##0_);[Red](#,##0)"
            'osheet.PivotTables("PivotTable1").PivotFields(" Bottleneck").NumberFormat = "#,##0"
            'osheet.PivotTables("PivotTable1").PivotFields(" Forecast").NumberFormat = "#,##0"
            'osheet.PivotTables("PivotTable1").PivotFields("Order Confirmed").NumberFormat = "#,##0"
            'osheet.PivotTables("PivotTable1").PivotFields("Order Unconfirmed").NumberFormat = "#,##0"


            'Change DataPivot Orientation
            'osheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlRowField

            'Filter Datafield
            'osheet.PivotTables("PivotTable1").PivotFields("sopdescription").PivotFilters.Add(Excel.XlPivotFilterType.xlValueIsGreaterThan, osheet.PivotTables("PivotTable1").PivotFields("Flag for demand exceeds bottleneck"), 0)
            'sort column period 

            'oSheet.PivotTables("PivotTable1").pivotfields("Period").autosort(Excel.XlSortOrder.xlDescending, "period")
            osheet.Name = "PivotTable"
            osheet.Cells.EntireColumn.AutoFit()



            iSheet += 1

        Catch ex As Exception
            errMsg = ex.Message
        End Try

    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        BackgroundWorker1.ReportProgress(3, TextBox3.Text & ".")
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            MyRegionId = ComboBox2.SelectedValue
            ShortRegion = RegionDict(MyRegionId)
            RegionName = ComboBox2.Text.ToString
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub FGenerateReport001_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged

    End Sub
End Class
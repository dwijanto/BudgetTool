Imports System.ComponentModel
Imports HR.HelperClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports DJLib.Dbtools
Public Class ImportHRInformation

    Private WithEvents BackgroundWorker1 As New BackgroundWorker
    Dim FileName As String = String.Empty
    Dim Status As Boolean = False
    Dim Dataset1 As DataSet
    Dim ConnectionString As String = dbtools1.getConnectionString
    Dim aprocesses() As Process = Nothing '= Process.GetProcesses
    Dim aprocess As Process = Nothing
    Dim EndOfYear As Date
    Dim myYear As Integer = 0
    Dim myVerid As Integer = 0
    Dim regionImport As Integer
    Enum RegionImportEnum
        HK = 1
        SZ = 2
        TW = 3
        PH = 4
    End Enum
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not (BackgroundWorker1.IsBusy) Then
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                FileName = OpenFileDialog1.FileName
                TextBox1.Text = FileName
                EndOfYear = CDate(DateTimePicker1.Value.Year & "-12-31")
                myYear = DateTimePicker1.Value.Year
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
        Dim iCols As Long = 0
        'Need these variable to kill excel
        Dim aprocesses() As Process = Nothing '= Process.GetProcesses
        Dim aprocess As Process = Nothing
        Dim stringbuilder1 As New StringBuilder
        Dim myyear As Integer = 0

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

            oSheet = oWb.Worksheets(1)
            If oSheet.Name <> "HR-INFO" Then
                Throw New System.Exception("File not valid")
            End If
            If IsNothing(oSheet.Cells(1, 1).value) Then
                Throw New System.Exception("File not valid")
            End If
            If IsNothing(oSheet.Cells(1, 2).value) Then
                Throw New System.Exception("Year control is blank")
            Else
                myyear = oSheet.Cells(1, 2).value.ToString
                If oSheet.Cells(1, 2).value.ToString <> DateTimePicker1.Value.Year Then
                    Throw New System.Exception("Year selection = " & DateTimePicker1.Value.Year & ", File year = " & oSheet.Cells(1, 2).value.ToString)
                End If
            End If

            'Check version
            If IsNothing(oSheet.Cells(2, 2).value) Then
                Throw New System.Exception("Version is blank.")
            End If
            Dim versionReason As String = String.Empty
            If Not IsNothing(oSheet.Cells(3, 2).value) Then
                versionReason = oSheet.Cells(3, 2).value.ToString
            End If
            'Check region
            'Dim regionimport As Integer = 0
            If Not IsNothing(oSheet.Cells(4, 2).value) Then
                Dim regionname = oSheet.Cells(4, 2).value.ToString
                regionimport = DbAdapter1.getRegionIDFromRegionName(regionname)
                If IsNothing(regionimport) Then
                    Throw New System.Exception("Region " & regionname & " is not registered!")
                End If
                'check region
                If regionimport <> dbtools1.RegionId Then
                    Throw New System.Exception("user :" & dbtools1.RegionName & ", File : " & regionname)
                End If
            End If

            Dim stopwatch As New Stopwatch
            stopwatch.Start()

            myVerid = DbAdapter1.getVerId(oSheet.Cells(2, 2).value.ToString, versionReason, myyear)

            If Not DbAdapter1.validversion(myVerid, DateTimePicker1.Value.Year, "CV" & dbtools1.Region, dbtools1.Region) Then
                Throw New System.Exception("This version is closed")
            End If


            iRows = oSheet.UsedRange.Rows.Count
            iCols = oSheet.UsedRange.Columns.Count

            BackgroundWorker1.ReportProgress(2, "Connect to Db...")

            'clean up data for same version,year and region
            cleanCurrentData(myyear, myVerid, regionImport)


            BackgroundWorker1.ReportProgress(3, "Preparing Tables...")


            If CheckBox2.Checked Then
                'Check COA
                Dim myCheck As Boolean = False
                Dim sbchk As New StringBuilder

                myCheck = CheckCOA(oSheet, sbchk)
                If Not myCheck Then
                    Dim errorFilename As String = Path.GetDirectoryName(FileName) & "\" & dbtools1.Region & "PersonalError.txt"
                    Using sw As StreamWriter = File.CreateText(errorFilename)
                        sw.WriteLine(sbchk.ToString)
                        sw.Close()
                    End Using
                    Process.Start(errorFilename)
                    Throw New System.Exception("Incorrect COA")
                End If
            End If
            



            'Dim sqlstr As String = "select expensesdetailid,myyear,expensesdetailtxid  from expensesdetailtx ;"
            Dim sqlstr As String = "select expensesdetailid,myyear,expensesdetailtxid  from expensesdetailtx where myyear = " & DateTimePicker1.Value.Year & ";"
            If Not dbtools1.getDataSet(sqlstr, Dataset1, errMessage) Then
                Return myreturn
            End If

            Dataset1.Tables(0).TableName = "expensesdetailtx"
            Dim keys0(0) As DataColumn
            keys0(0) = Dataset1.Tables(0).Columns(0)
            Dataset1.Tables(0).PrimaryKey = keys0



            Dim categorytypeid As Integer = 0
            Dim expensesnatureid As Integer = 0
            Dim sapaccnameid As Integer = 0
            Dim accexpensesid As Integer = 0
            Dim sapaccountid As Integer = 0
            Dim sapindexid As Integer = 0
            Dim FullYear As String = "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"
            Dim firstTime As Boolean = True
            Dim categorydtlid As Integer = 0
            'Check For ExpensesNature.
            Dim updatelocation As Boolean = True
            For i = 12 To iCols

                BackgroundWorker1.ReportProgress(2, "Processing Column " & i & " of " & iCols & "...")
                If i = 18 Then
                    Debug.WriteLine("Debug mode")
                End If



                If Not IsNothing(oSheet.Cells(6, i).value) Then
                    'If Not IsNothing(oSheet.Cells(6, i).value) Or i < 12 Then
                    'Check table expensesnature if not avail then create
                    If i >= 14 Then
                        If i = 28 Then
                            'Debug.WriteLine("Debug Mode")
                        End If
                        expensesnatureid = DbAdapter1.getExpensesNatureId(Replace(oSheet.Cells(6, i).value.ToString, "'", "''"))
                        '8. Check table sapaccname if not avail then create
                        sapaccnameid = DbAdapter1.getSapAccNameId(oSheet.Cells(5, i).value.ToString)
                        '9. Check table accexpenses if not avail then create
                        accexpensesid = DbAdapter1.getAccExpensesId(sapaccnameid, expensesnatureid)
                        '10. check table expensesdetail if not avail then create
                        sapaccountid = DbAdapter1.getSapAccountid(oSheet.Cells(7, i).value.ToString)
                    End If
                    '::update expensesFullyear

                    If Not IsNothing(oSheet.Cells(13, i).value) Or i < 14 Then

                        If oSheet.Cells(13, i).value = "Category" Then
                            'create categorytype
                            categorytypeid = DbAdapter1.getcategorytypeId(Replace(oSheet.Cells(6, i).value.ToString, "'", "''"))
                        End If
                        ''End If
                    End If

                    Dim categoryid As Integer = 0
                    Dim costcenterid As Integer = 0
                    Dim titleid As Integer = 0
                    Dim personid As Integer = 0
                    Dim persontitleid As Integer = 0
                    Dim personjoindateid As Integer = 0
                    Dim personjoindatecategoryid As Integer = 0
                    Dim expensesdetailtxid As Integer = 0
                    Dim salarytxid As Integer = 0
                    Dim deptid As Integer = 0
                    Dim indexcostcenterid As Integer = 0
                    Dim indexcostcenterdeptid As Integer = 0
                    Dim expensesdetailid As Integer = 0
                    Dim familymemberid As Integer = 0
                    Dim personexpensesid As Integer = 0
                    Dim personexpensesdtlid As Integer = 0
                    Dim icdpjcid As Integer = 0

                    'Dim enddateSTR As String = String.Empty
                    For J = 14 To iRows
                        If J = 88 Then
                            Debug.WriteLine("Debug Mode")
                            'If Not IsNothing(oSheet.Cells(J, 1).value) Then
                            '    If oSheet.Cells(J, 8).value = "ZHOU Cheng Jin" Then
                            '        'MessageBox.Show("ZHOU Cheng Jin is here.")
                            '    End If
                            'End If


                        End If

                        If IsNothing(oSheet.Cells(J, 1).value) Then
                            Exit For
                        End If

                        BackgroundWorker1.ReportProgress(2, "Processing Column " & i & " of " & iCols & ".Row " & J)

                        'Check has value or not J=ROW,i = Col
                        If Not IsNothing(oSheet.Cells(J, i).value) Then
                            'If J = 21 Then
                            '    Debug.WriteLine("Debug Mode")
                            '    If Not IsNothing(oSheet.Cells(J, 1).value) Then
                            '        If oSheet.Cells(J, 8).value = "ZHOU Cheng Jin" Then
                            '            'MessageBox.Show(oSheet.Cells(J, i).value)
                            '        End If
                            '    End If


                            'End If


                            categoryid = DbAdapter1.getcategoryId(oSheet.Cells(J, 1).value.ToString)
                            If updatelocation Then
                                If Not IsNothing(oSheet.Cells(J, 2).value) Then
                                    DbAdapter1.UpdateLocation(oSheet.Cells(J, 1).value.ToString, oSheet.Cells(J, 2).value.ToString)

                                End If
                            End If

                            'check dept
                            Dim mydept As String = oSheet.Cells(J, 3).value.ToString
                            deptid = DbAdapter1.getdeptid(mydept)

                            'check costcentre

                            Dim mysapcc As String = String.Empty

                            'get costcenter  based on sapcc
                            mysapcc = mydept
                            Dim last2digit = mydept.Substring(4, 2)
                            If Not IsNothing(oSheet.Cells(8, i).value) Then
                                Dim sapcc As String = oSheet.Cells(8, i).value.ToString
                                Dim sapccobj() As String = sapcc.Split(",")
                                If sapccobj.Length > 1 Then
                                    If last2digit <> sapccobj(1) Then
                                        mysapcc = sapccobj(0) & last2digit
                                    End If
                                Else
                                    mysapcc = sapcc & last2digit

                                End If
                            End If
                            costcenterid = DbAdapter1.getsapccid(mysapcc)

                            'check title
                            titleid = DbAdapter1.gettitleid(oSheet.Cells(J, 10).value.ToString)
                            'Check person

                            If Not IsNothing(oSheet.Cells(7, i).value) Then
                                '2. Check table sapindex if not avail then create
                                sapindexid = DbAdapter1.getSapIndexid(sapaccountid, oSheet.Cells(7, i).value.ToString & "-" & mysapcc)
                                '3. Check table CostCenter if not avail then create
                                indexcostcenterid = DbAdapter1.getIndexcostcenterId(sapindexid, costcenterid)
                                '5. Check table dept if not avail then create
                                indexcostcenterdeptid = DbAdapter1.getIndexCostCenterDeptId(indexcostcenterid, deptid)
                                '7. Check table expensesnature if not avail then create
                                expensesdetailid = DbAdapter1.getExpensesDetailid(accexpensesid, indexcostcenterdeptid)
                                'Find Table ExpensesDetail using accexpensesid and indexcostcenterdeptid
                            End If


                            Dim expat As Boolean = False
                            '******************* column changed 
                            If Not IsNothing(oSheet.Cells(J, 11).value) Then
                                If oSheet.Cells(J, 11).value.ToString.ToUpper.Trim = "TRUE" Then
                                    expat = True
                                End If
                            End If

                            'personid = DbAdapter1.getpersonid(oSheet.Cells(J, 8).value.ToString, expat, oSheet.Cells(J, 1).value.ToString.Substring(0, 2))
                            personid = DbAdapter1.getpersonid(oSheet.Cells(J, 8).value.ToString)

                            Dim JoinDateD As Date = CDate(oSheet.Cells(J, 5).value.ToString)
                            Dim enddate As Date
                            Dim enddateSTR As String = String.Empty
                            If Not IsNothing(oSheet.Cells(J, 6).value) Then
                                enddate = CDate(oSheet.Cells(J, 6).value.ToString)
                                enddateSTR = "Not Null"
                            End If


                            Dim othername As String = String.Empty
                            If Not IsNothing(oSheet.Cells(J, 9).value) Then
                                othername = oSheet.Cells(J, 9).value.ToString
                            Else
                                othername = "Null"
                            End If

                            If enddateSTR = "Not Null" Then
                                personjoindateid = DbAdapter1.getpersonjoindateid(personid, JoinDateD, othername, True, enddate)
                            Else
                                personjoindateid = DbAdapter1.getpersonjoindateid(personid, JoinDateD, othername, True)
                            End If

                            'If Hong Kong, add effective date start and effective date end
                            If regionImport = RegionImportEnum.HK Then
                                'check effectivedate start cannot be blank for personjoindatecategoryid
                                Dim EffectiveDateStart As Date = Nothing
                                Dim EffectiveDateEnd As Date? = Nothing
                                Dim bonusfactor As Integer = Nothing
                                If IsNothing(oSheet.Cells(J, 12).value) Then
                                    Throw New System.Exception("Effective Date Start cannot be blank")
                                Else
                                    EffectiveDateStart = CDate(oSheet.Cells(J, 12).value)
                                End If
                                If Not IsNothing(oSheet.Cells(J, 13).value) Then
                                    EffectiveDateEnd = CDate(oSheet.Cells(J, 13).value)
                                End If
                                If Not IsNothing(oSheet.Cells(J, 14).value) Then
                                    bonusfactor = oSheet.Cells(J, 14).value
                                End If

                                personjoindatecategoryid = DbAdapter1.getpersonjoindatecategoryid(categoryid, personjoindateid, deptid, CDbl(oSheet.Cells(J, 7).value.ToString), expat, EffectiveDateStart, EffectiveDateEnd, bonusfactor)

                            Else
                                personjoindatecategoryid = DbAdapter1.getpersonjoindatecategoryid(categoryid, personjoindateid, deptid, CDbl(oSheet.Cells(J, 7).value.ToString), expat)
                            End If


                            persontitleid = DbAdapter1.insertpersontitle(titleid, JoinDateD, personjoindatecategoryid)

                            If Not IsNothing(oSheet.Cells(7, i).value) Then
                                'Get expensesdetailtxid
                                expensesdetailtxid = DbAdapter1.getExpensesDetailtxid(myyear, expensesdetailid)
                                'expensesdetailtxid = DbAdapter1.getExpensesDetailtxid(DateTimePicker1.Value.Year, expensesdetailid)
                                'insert personexpenses

                                icdpjcid = DbAdapter1.geticdpjcid(personjoindatecategoryid, indexcostcenterdeptid, accexpensesid, myyear, myVerid, regionImport)
                                'personexpensesid = DbAdapter1.insertpersonexpensesid(personjoindatecategoryid, expensesdetailtxid)
                                personexpensesid = DbAdapter1.insertpersonexpensesid(icdpjcid, expensesdetailtxid)

                                'check i (expenses) then assign the value

                                If IsNothing(oSheet.Cells(13, i).value) Then
                                    'insert personexpensesdtl
                                    DbAdapter1.insertpersonexpensesdtl(personexpensesid, CDbl(oSheet.Cells(J, i).value.ToString), CDate(oSheet.Cells(J, 5).value.ToString))

                                    'If expensesFullYear "/12" then true
                                    'Else create ExpensesNatureMonths
                                    Dim monthToInsert As String = String.Empty
                                    If Not IsNothing(oSheet.Cells(J, i + 1).value) Then
                                        If oSheet.Cells(J, i + 1).value.ToString.Contains("/") Then
                                            'update expensesnaturefullyear = true
                                            DbAdapter1.setexpensesnaturefullyear(expensesnatureid, True)
                                            'insert table expnesesnaturemonths 12 records
                                            monthToInsert = FullYear
                                        Else
                                            'update expensesnaturefullyear = false
                                            DbAdapter1.setexpensesnaturefullyear(expensesnatureid, False)
                                            'insert table expensesnaturemonths based on "everymonth, each month  or list of months"
                                            If oSheet.Cells(J, i + 1).value.ToString.ToLower = "every month" Then
                                                monthToInsert = FullYear
                                            Else
                                                monthToInsert = oSheet.Cells(J, i + 1).value.ToString
                                            End If

                                        End If
                                        'DbAdapter1.insertexpensesnaturemonths(expensesnatureid, monthToInsert, myYear)
                                        'DbAdapter1.insertCategorytxMonths(categorytxid, monthToInsert, myYear)

                                        'this insertcategorytxmonth will create categorytx without amount for link to categorytxmonths. disable this if you use for category import 
                                        'DbAdapter1.insertCategorytxMonths(categoryid, Replace(oSheet.Cells(6, i).value.ToString, "'", "''"), monthToInsert, myyear)
                                        DbAdapter1.insertPersonalTxMonths(personjoindatecategoryid, expensesnatureid, monthToInsert, myyear, myVerid, regionImport)
                                    End If
                                Else
                                    If oSheet.Cells(13, i).value.ToString = "Category" Then
                                        If Not IsNothing(oSheet.Cells(J, i).value) Then
                                            'Dim myvalue As Double = CDbl(oSheet.Cells(J, i).value.ToString)
                                            Dim myvalue = oSheet.Cells(J, i).value.ToString

                                            If Not IsNumeric(myvalue) Then
                                                myvalue = DbAdapter1.getplanid(oSheet.Cells(J, i).value.ToString)
                                            End If


                                            categorydtlid = DbAdapter1.insertcategorydtl(categoryid, categorytypeid, CDbl(myvalue), myyear, myVerid)

                                        End If

                                    End If
                                End If

                            End If

                        End If

                        'Debug.WriteLine("Next")
                    Next
                    updatelocation = False
                Else
                    'Check For Hong Kong-> Hardcoded column 12 = effective date start and column 13 = effective date end
                    If regionImport = RegionImportEnum.HK And (i >= 12 And i <= 13) Then
                        If i = 12 Then
                            If Not oSheet.Cells(13, i).value.ToString.ToLower.Contains("effective date start") Then
                                Throw New System.Exception("Column J should be effective date start")
                                Exit For
                            End If
                        ElseIf i = 13 Then
                            If Not oSheet.Cells(13, 13).value.ToString.ToLower.Contains("effective date end") Then
                                Throw New System.Exception("Column K should be effective date end")
                                Exit For
                            End If
                        End If
                    End If


                    'check for catch up increment
                    If Not IsNothing(oSheet.Cells(13, i).value) Then
                        Dim txtype As String = String.Empty
                        Dim medicalplan As String = String.Empty
                        If Not IsNothing(oSheet.Cells(9, i).value) Then
                            If oSheet.Cells(9, i).value.ToString.ToLower.Contains("catch up") Then
                                txtype = "catch up"
                            ElseIf oSheet.Cells(9, i).value.ToString.ToLower.Contains("general increment") Then
                                txtype = "incr"
                            End If
                        ElseIf Not IsNothing(oSheet.Cells(13, i).value) Then
                            If oSheet.Cells(13, i).value.ToString.ToLower.Contains("1.5") Then
                                medicalplan = "1.5"
                            ElseIf oSheet.Cells(13, i).value.ToString.ToLower.Contains("2.2") Then
                                medicalplan = "2.2"
                            ElseIf oSheet.Cells(13, i).value.ToString.ToLower.Contains("3.2") Then
                                medicalplan = "3.2"
                            ElseIf oSheet.Cells(13, i).value.ToString.ToLower.Contains("7.2") Then
                                medicalplan = "7.2"
                            End If
                        End If

                        If txtype <> "" Then
                            For J = 14 To iRows
                                If Not IsNothing(oSheet.Cells(J, i + 1).value) Then
                                    Dim startingdate As Date = getdate(oSheet.Cells(J, i + 1).value.ToString, myyear)
                                    Dim amount As Double = oSheet.Cells(J, i).value.ToString
                                    If txtype = "catch up" Then
                                        'DbAdapter1.insertsalarytx(oSheet.Cells(J, 1).value.ToString, oSheet.Cells(J, 3).value.ToString, oSheet.Cells(J, 5).value.ToString, oSheet.Cells(J, 8).value.ToString, CDbl(oSheet.Cells(J, i).value.ToString), startingdate, txtype, myVerid)
                                        DbAdapter1.insertsalarytx(oSheet.Cells(J, 1).value.ToString, oSheet.Cells(J, 3).value.ToString, oSheet.Cells(J, 5).value.ToString, oSheet.Cells(J, 8).value.ToString, CDbl(oSheet.Cells(J, i).value.ToString), startingdate, txtype, myVerid, regionImport, myyear)
                                    ElseIf txtype = "general increment" Then
                                        'get categoryid
                                        Dim categoryid As Integer = DbAdapter1.getcategoryId(oSheet.Cells(J, 1).value.ToString)
                                        'DbAdapter1.insertsalarytx(oSheet.Cells(J, 1).value.ToString, oSheet.Cells(J, 3).value.ToString, oSheet.Cells(J, 5).value.ToString, oSheet.Cells(J, 8).value.ToString, CDbl(oSheet.Cells(J, i).value.ToString), startingdate, txtype, myVerid)

                                        DbAdapter1.insertsalarytx(categoryid, amount, CDate(myyear & "-1-1"), txtype, myVerid)
                                    End If

                                End If
                            Next
                        End If
                        If medicalplan <> "" Then
                            Dim familymemberid As Integer = 0
                            For J = 14 To iRows
                                If Not IsNothing(oSheet.Cells(J, i).value) Then

                                    Dim expat As Boolean = False
                                    If Not IsNothing(oSheet.Cells(J, 11).value) Then
                                        expat = True
                                    End If
                                    'Dim personid = DbAdapter1.getpersonid(oSheet.Cells(J, 8).value.ToString, expat, oSheet.Cells(J, 1).value.ToString.Substring(0, 2))
                                    Dim personid = DbAdapter1.getpersonid(oSheet.Cells(J, 8).value.ToString)
                                    'Dim joindate As String = DateFormatyyyyMMdd(CDate(oSheet.Cells(j, 5).value.ToString))
                                    Dim JoinDateD As Date = CDate(oSheet.Cells(J, 5).value.ToString)
                                    Dim othername As String = String.Empty
                                    If Not IsNothing(oSheet.Cells(J, 9).value) Then
                                        othername = oSheet.Cells(J, 9).value.ToString
                                    Else
                                        othername = "Null"
                                    End If
                                    Dim personjoindateid = DbAdapter1.getpersonjoindateid(personid, JoinDateD, othername, True)
                                    If medicalplan = "1.5" Then
                                        familymemberid = DbAdapter1.insertfamilymemberid(personjoindateid, "Plan 1.5", oSheet.Cells(J, i).value.ToString, myyear, myVerid, regionImport)
                                    ElseIf medicalplan = "2.2" Then
                                        familymemberid = DbAdapter1.insertfamilymemberid(personjoindateid, "Plan 2.2", oSheet.Cells(J, i).value.ToString, myyear, myVerid, regionImport)
                                    ElseIf medicalplan = "3.2" Then
                                        familymemberid = DbAdapter1.insertfamilymemberid(personjoindateid, "Plan 3.2", oSheet.Cells(J, i).value.ToString, myyear, myVerid, regionImport)
                                    ElseIf medicalplan = "7.2" Then
                                        familymemberid = DbAdapter1.insertfamilymemberid(personjoindateid, "Plan 7.2", oSheet.Cells(J, i).value.ToString, myyear, myVerid, regionImport)
                                    End If
                                End If

                            Next
                        End If

                    End If
                End If
                firstTime = False
            Next
            stopwatch.Stop()
            BackgroundWorker1.ReportProgress(2, "Elapsed Time: " & Format(stopwatch.Elapsed.Minutes, "00") & ":" & Format(stopwatch.Elapsed.Seconds, "00") & "." & stopwatch.Elapsed.Milliseconds.ToString)

            BackgroundWorker1.ReportProgress(3, "")
            myreturn = True
        Catch ex As Exception
            errMessage = ex.Message
            'cleanCurrentData(myyear, myVerid, regionimport)
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

    'Private Function ImportDataOri(ByVal FileName As String, Optional ByRef errMessage As String = "") As Boolean
    '    Dim myreturn As Boolean = False
    '    Dim oXl As Excel.Application = Nothing
    '    Dim oWb As Excel.Workbook = Nothing
    '    Dim oSheet As Excel.Worksheet = Nothing
    '    Dim oRange As Excel.Range = Nothing

    '    Dim iRows As Long = 0
    '    Dim iCols As Long = 0
    '    'Need these variable to kill excel
    '    Dim aprocesses() As Process = Nothing '= Process.GetProcesses
    '    Dim aprocess As Process = Nothing
    '    Dim stringbuilder1 As New StringBuilder
    '    Dim myyear As Integer = 0

    '    Try

    '        BackgroundWorker1.ReportProgress(2, "Preparing Data...")
    '        Dataset1 = New DataSet
    '        BackgroundWorker1.ReportProgress(3, "Opening Excel File....")
    '        oXl = CType(CreateObject("Excel.Application"), Excel.Application)
    '        Application.DoEvents()
    '        oXl.Visible = True
    '        'get process pid
    '        aprocesses = Process.GetProcesses
    '        For i = 0 To aprocesses.GetUpperBound(0)
    '            If aprocesses(i).MainWindowHandle.ToString = oXl.Hwnd.ToString Then
    '                aprocess = aprocesses(i)
    '                Exit For
    '            End If
    '            Application.DoEvents()
    '        Next
    '        oXl.Visible = False
    '        oXl.DisplayAlerts = False
    '        oWb = oXl.Workbooks.Open(FileName)

    '        oSheet = oWb.Worksheets(1)
    '        If oSheet.Name <> "HR-INFO" Then
    '            Throw New System.Exception("File not valid")
    '        End If
    '        If IsNothing(oSheet.Cells(1, 1).value) Then
    '            Throw New System.Exception("File not valid")
    '        End If
    '        If IsNothing(oSheet.Cells(1, 2).value) Then
    '            Throw New System.Exception("Year control is blank")
    '        Else
    '            myyear = oSheet.Cells(1, 2).value.ToString
    '            If oSheet.Cells(1, 2).value.ToString <> DateTimePicker1.Value.Year Then
    '                Throw New System.Exception("Year selection = " & DateTimePicker1.Value.Year & ", File year = " & oSheet.Cells(1, 2).value.ToString)
    '            End If
    '        End If

    '        'Check version
    '        If IsNothing(oSheet.Cells(2, 2).value) Then
    '            Throw New System.Exception("Version is blank.")
    '        End If
    '        Dim versionReason As String = String.Empty
    '        If Not IsNothing(oSheet.Cells(3, 2).value) Then
    '            versionReason = oSheet.Cells(3, 2).value.ToString
    '        End If
    '        'Check region
    '        'Dim regionimport As Integer = 0
    '        If Not IsNothing(oSheet.Cells(4, 2).value) Then
    '            Dim regionname = oSheet.Cells(4, 2).value.ToString
    '            regionImport = DbAdapter1.getRegionIDFromRegionName(regionname)
    '            If IsNothing(regionImport) Then
    '                Throw New System.Exception("Region " & regionname & " is not registered!")
    '            End If
    '            'check region
    '            If regionImport <> dbtools1.RegionId Then
    '                Throw New System.Exception("user :" & dbtools1.RegionName & ", File : " & regionname)
    '            End If
    '        End If

    '        Dim stopwatch As New Stopwatch
    '        stopwatch.Start()

    '        myVerid = DbAdapter1.getVerId(oSheet.Cells(2, 2).value.ToString, versionReason, myyear)

    '        If Not DbAdapter1.validversion(myVerid, DateTimePicker1.Value.Year, "CV" & dbtools1.Region, dbtools1.Region) Then
    '            Throw New System.Exception("This version is closed")
    '        End If


    '        iRows = oSheet.UsedRange.Rows.Count
    '        iCols = oSheet.UsedRange.Columns.Count

    '        BackgroundWorker1.ReportProgress(2, "Connect to Db...")

    '        'clean up data for same version,year and region
    '        cleanCurrentData(myyear, myVerid, regionImport)


    '        BackgroundWorker1.ReportProgress(3, "Preparing Tables...")


    '        If CheckBox2.Checked Then
    '            'Check COA
    '            Dim myCheck As Boolean = False
    '            Dim sbchk As New StringBuilder

    '            myCheck = CheckCOA(oSheet, sbchk)
    '            If Not myCheck Then
    '                Dim errorFilename As String = Path.GetDirectoryName(FileName) & "\" & dbtools1.Region & "PersonalError.txt"
    '                Using sw As StreamWriter = File.CreateText(errorFilename)
    '                    sw.WriteLine(sbchk.ToString)
    '                    sw.Close()
    '                End Using
    '                Process.Start(errorFilename)
    '                Throw New System.Exception("Incorrect COA")
    '            End If
    '        End If




    '        'Dim sqlstr As String = "select expensesdetailid,myyear,expensesdetailtxid  from expensesdetailtx ;"
    '        Dim sqlstr As String = "select expensesdetailid,myyear,expensesdetailtxid  from expensesdetailtx where myyear = " & DateTimePicker1.Value.Year & ";"
    '        If Not dbtools1.getDataSet(sqlstr, Dataset1, errMessage) Then
    '            Return myreturn
    '        End If

    '        Dataset1.Tables(0).TableName = "expensesdetailtx"
    '        Dim keys0(0) As DataColumn
    '        keys0(0) = Dataset1.Tables(0).Columns(0)
    '        Dataset1.Tables(0).PrimaryKey = keys0



    '        Dim categorytypeid As Integer = 0
    '        Dim expensesnatureid As Integer = 0
    '        Dim sapaccnameid As Integer = 0
    '        Dim accexpensesid As Integer = 0
    '        Dim sapaccountid As Integer = 0
    '        Dim sapindexid As Integer = 0
    '        Dim FullYear As String = "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"
    '        Dim firstTime As Boolean = True
    '        Dim categorydtlid As Integer = 0
    '        'Check For ExpensesNature.
    '        Dim updatelocation As Boolean = True
    '        For i = 12 To iCols

    '            BackgroundWorker1.ReportProgress(2, "Processing Column " & i & " of " & iCols & "...")
    '            If i = 18 Then
    '                Debug.WriteLine("Debug mode")
    '            End If
    '            If Not IsNothing(oSheet.Cells(6, i).value) Then
    '                'If Not IsNothing(oSheet.Cells(6, i).value) Or i < 12 Then
    '                'Check table expensesnature if not avail then create
    '                If i >= 14 Then
    '                    If i = 28 Then
    '                        'Debug.WriteLine("Debug Mode")
    '                    End If
    '                    expensesnatureid = DbAdapter1.getExpensesNatureId(Replace(oSheet.Cells(6, i).value.ToString, "'", "''"))
    '                    '8. Check table sapaccname if not avail then create
    '                    sapaccnameid = DbAdapter1.getSapAccNameId(oSheet.Cells(5, i).value.ToString)
    '                    '9. Check table accexpenses if not avail then create
    '                    accexpensesid = DbAdapter1.getAccExpensesId(sapaccnameid, expensesnatureid)
    '                    '10. check table expensesdetail if not avail then create
    '                    sapaccountid = DbAdapter1.getSapAccountid(oSheet.Cells(7, i).value.ToString)
    '                End If
    '                '::update expensesFullyear

    '                If Not IsNothing(oSheet.Cells(13, i).value) Or i < 14 Then

    '                    If oSheet.Cells(13, i).value = "Category" Then
    '                        'create categorytype
    '                        categorytypeid = DbAdapter1.getcategorytypeId(Replace(oSheet.Cells(6, i).value.ToString, "'", "''"))
    '                    End If
    '                    ''End If
    '                End If

    '                Dim categoryid As Integer = 0
    '                Dim costcenterid As Integer = 0
    '                Dim titleid As Integer = 0
    '                Dim personid As Integer = 0
    '                Dim persontitleid As Integer = 0
    '                Dim personjoindateid As Integer = 0
    '                Dim personjoindatecategoryid As Integer = 0
    '                Dim expensesdetailtxid As Integer = 0
    '                Dim salarytxid As Integer = 0
    '                Dim deptid As Integer = 0
    '                Dim indexcostcenterid As Integer = 0
    '                Dim indexcostcenterdeptid As Integer = 0
    '                Dim expensesdetailid As Integer = 0
    '                Dim familymemberid As Integer = 0
    '                Dim personexpensesid As Integer = 0
    '                Dim personexpensesdtlid As Integer = 0
    '                Dim icdpjcid As Integer = 0

    '                'Dim enddateSTR As String = String.Empty
    '                For J = 14 To iRows
    '                    If J = 88 Then
    '                        Debug.WriteLine("Debug Mode")
    '                        'If Not IsNothing(oSheet.Cells(J, 1).value) Then
    '                        '    If oSheet.Cells(J, 8).value = "ZHOU Cheng Jin" Then
    '                        '        'MessageBox.Show("ZHOU Cheng Jin is here.")
    '                        '    End If
    '                        'End If


    '                    End If

    '                    If IsNothing(oSheet.Cells(J, 1).value) Then
    '                        Exit For
    '                    End If

    '                    BackgroundWorker1.ReportProgress(2, "Processing Column " & i & " of " & iCols & ".Row " & J)

    '                    'Check has value or not J=ROW,i = Col
    '                    If Not IsNothing(oSheet.Cells(J, i).value) Then
    '                        'If J = 21 Then
    '                        '    Debug.WriteLine("Debug Mode")
    '                        '    If Not IsNothing(oSheet.Cells(J, 1).value) Then
    '                        '        If oSheet.Cells(J, 8).value = "ZHOU Cheng Jin" Then
    '                        '            'MessageBox.Show(oSheet.Cells(J, i).value)
    '                        '        End If
    '                        '    End If


    '                        'End If





    '                        categoryid = DbAdapter1.getcategoryId(oSheet.Cells(J, 1).value.ToString)
    '                        If updatelocation Then
    '                            If Not IsNothing(oSheet.Cells(J, 2).value) Then
    '                                DbAdapter1.UpdateLocation(oSheet.Cells(J, 1).value.ToString, oSheet.Cells(J, 2).value.ToString)

    '                            End If
    '                        End If

    '                        'check dept
    '                        Dim mydept As String = oSheet.Cells(J, 3).value.ToString
    '                        deptid = DbAdapter1.getdeptid(mydept)

    '                        'check costcentre

    '                        Dim mysapcc As String = String.Empty

    '                        'get costcenter  based on sapcc
    '                        mysapcc = mydept
    '                        Dim last2digit = mydept.Substring(4, 2)
    '                        If Not IsNothing(oSheet.Cells(8, i).value) Then
    '                            Dim sapcc As String = oSheet.Cells(8, i).value.ToString
    '                            Dim sapccobj() As String = sapcc.Split(",")
    '                            If sapccobj.Length > 1 Then
    '                                If last2digit <> sapccobj(1) Then
    '                                    mysapcc = sapccobj(0) & last2digit
    '                                End If
    '                            Else
    '                                mysapcc = sapcc & last2digit

    '                            End If
    '                        End If
    '                        costcenterid = DbAdapter1.getsapccid(mysapcc)

    '                        'check title
    '                        titleid = DbAdapter1.gettitleid(oSheet.Cells(J, 10).value.ToString)
    '                        'Check person

    '                        If Not IsNothing(oSheet.Cells(7, i).value) Then
    '                            '2. Check table sapindex if not avail then create
    '                            sapindexid = DbAdapter1.getSapIndexid(sapaccountid, oSheet.Cells(7, i).value.ToString & "-" & mysapcc)
    '                            '3. Check table CostCenter if not avail then create
    '                            indexcostcenterid = DbAdapter1.getIndexcostcenterId(sapindexid, costcenterid)
    '                            '5. Check table dept if not avail then create
    '                            indexcostcenterdeptid = DbAdapter1.getIndexCostCenterDeptId(indexcostcenterid, deptid)
    '                            '7. Check table expensesnature if not avail then create
    '                            expensesdetailid = DbAdapter1.getExpensesDetailid(accexpensesid, indexcostcenterdeptid)
    '                            'Find Table ExpensesDetail using accexpensesid and indexcostcenterdeptid
    '                        End If


    '                        Dim expat As Boolean = False
    '                        '******************* column changed 
    '                        If Not IsNothing(oSheet.Cells(J, 11).value) Then
    '                            If oSheet.Cells(J, 11).value.ToString.ToUpper.Trim = "TRUE" Then
    '                                expat = True
    '                            End If
    '                        End If

    '                        'personid = DbAdapter1.getpersonid(oSheet.Cells(J, 8).value.ToString, expat, oSheet.Cells(J, 1).value.ToString.Substring(0, 2))
    '                        personid = DbAdapter1.getpersonid(oSheet.Cells(J, 8).value.ToString)

    '                        Dim JoinDateD As Date = CDate(oSheet.Cells(J, 5).value.ToString)
    '                        Dim enddate As Date
    '                        Dim enddateSTR As String = String.Empty
    '                        If Not IsNothing(oSheet.Cells(J, 6).value) Then
    '                            enddate = CDate(oSheet.Cells(J, 6).value.ToString)
    '                            enddateSTR = "Not Null"
    '                        End If


    '                        Dim othername As String = String.Empty
    '                        If Not IsNothing(oSheet.Cells(J, 9).value) Then
    '                            othername = oSheet.Cells(J, 9).value.ToString
    '                        Else
    '                            othername = "Null"
    '                        End If

    '                        If enddateSTR = "Not Null" Then
    '                            personjoindateid = DbAdapter1.getpersonjoindateid(personid, JoinDateD, othername, True, enddate)
    '                        Else
    '                            personjoindateid = DbAdapter1.getpersonjoindateid(personid, JoinDateD, othername, True)
    '                        End If


    '                        personjoindatecategoryid = DbAdapter1.getpersonjoindatecategoryid(categoryid, personjoindateid, deptid, CDbl(oSheet.Cells(J, 7).value.ToString), expat)

    '                        persontitleid = DbAdapter1.insertpersontitle(titleid, JoinDateD, personjoindatecategoryid)

    '                        If Not IsNothing(oSheet.Cells(7, i).value) Then
    '                            'Get expensesdetailtxid
    '                            expensesdetailtxid = DbAdapter1.getExpensesDetailtxid(myyear, expensesdetailid)
    '                            'expensesdetailtxid = DbAdapter1.getExpensesDetailtxid(DateTimePicker1.Value.Year, expensesdetailid)
    '                            'insert personexpenses

    '                            icdpjcid = DbAdapter1.geticdpjcid(personjoindatecategoryid, indexcostcenterdeptid, accexpensesid, myyear, myVerid, regionImport)
    '                            'personexpensesid = DbAdapter1.insertpersonexpensesid(personjoindatecategoryid, expensesdetailtxid)
    '                            personexpensesid = DbAdapter1.insertpersonexpensesid(icdpjcid, expensesdetailtxid)

    '                            'check i (expenses) then assign the value

    '                            If IsNothing(oSheet.Cells(13, i).value) Then
    '                                'insert personexpensesdtl
    '                                DbAdapter1.insertpersonexpensesdtl(personexpensesid, CDbl(oSheet.Cells(J, i).value.ToString), CDate(oSheet.Cells(J, 5).value.ToString))

    '                                'If expensesFullYear "/12" then true
    '                                'Else create ExpensesNatureMonths
    '                                Dim monthToInsert As String = String.Empty
    '                                If Not IsNothing(oSheet.Cells(J, i + 1).value) Then
    '                                    If oSheet.Cells(J, i + 1).value.ToString.Contains("/") Then
    '                                        'update expensesnaturefullyear = true
    '                                        DbAdapter1.setexpensesnaturefullyear(expensesnatureid, True)
    '                                        'insert table expnesesnaturemonths 12 records
    '                                        monthToInsert = FullYear
    '                                    Else
    '                                        'update expensesnaturefullyear = false
    '                                        DbAdapter1.setexpensesnaturefullyear(expensesnatureid, False)
    '                                        'insert table expensesnaturemonths based on "everymonth, each month  or list of months"
    '                                        If oSheet.Cells(J, i + 1).value.ToString.ToLower = "every month" Then
    '                                            monthToInsert = FullYear
    '                                        Else
    '                                            monthToInsert = oSheet.Cells(J, i + 1).value.ToString
    '                                        End If

    '                                    End If
    '                                    'DbAdapter1.insertexpensesnaturemonths(expensesnatureid, monthToInsert, myYear)
    '                                    'DbAdapter1.insertCategorytxMonths(categorytxid, monthToInsert, myYear)

    '                                    'this insertcategorytxmonth will create categorytx without amount for link to categorytxmonths. disable this if you use for category import 
    '                                    'DbAdapter1.insertCategorytxMonths(categoryid, Replace(oSheet.Cells(6, i).value.ToString, "'", "''"), monthToInsert, myyear)
    '                                    DbAdapter1.insertPersonalTxMonths(personjoindatecategoryid, expensesnatureid, monthToInsert, myyear, myVerid, regionImport)
    '                                End If
    '                            Else
    '                                If oSheet.Cells(13, i).value.ToString = "Category" Then
    '                                    If Not IsNothing(oSheet.Cells(J, i).value) Then
    '                                        'Dim myvalue As Double = CDbl(oSheet.Cells(J, i).value.ToString)
    '                                        Dim myvalue = oSheet.Cells(J, i).value.ToString

    '                                        If Not IsNumeric(myvalue) Then
    '                                            myvalue = DbAdapter1.getplanid(oSheet.Cells(J, i).value.ToString)
    '                                        End If


    '                                        categorydtlid = DbAdapter1.insertcategorydtl(categoryid, categorytypeid, CDbl(myvalue), myyear, myVerid)

    '                                    End If

    '                                End If
    '                            End If

    '                        End If

    '                    End If

    '                    'Debug.WriteLine("Next")
    '                Next
    '                updatelocation = False
    '            Else
    '                'check for catch up increment
    '                If Not IsNothing(oSheet.Cells(13, i).value) Then
    '                    Dim txtype As String = String.Empty
    '                    Dim medicalplan As String = String.Empty
    '                    If Not IsNothing(oSheet.Cells(9, i).value) Then
    '                        If oSheet.Cells(9, i).value.ToString.ToLower.Contains("catch up") Then
    '                            txtype = "catch up"
    '                        ElseIf oSheet.Cells(9, i).value.ToString.ToLower.Contains("general increment") Then
    '                            txtype = "incr"
    '                        End If
    '                    ElseIf Not IsNothing(oSheet.Cells(13, i).value) Then
    '                        If oSheet.Cells(13, i).value.ToString.ToLower.Contains("1.5") Then
    '                            medicalplan = "1.5"
    '                        ElseIf oSheet.Cells(13, i).value.ToString.ToLower.Contains("2.2") Then
    '                            medicalplan = "2.2"
    '                        ElseIf oSheet.Cells(13, i).value.ToString.ToLower.Contains("3.2") Then
    '                            medicalplan = "3.2"
    '                        ElseIf oSheet.Cells(13, i).value.ToString.ToLower.Contains("7.2") Then
    '                            medicalplan = "7.2"
    '                        End If
    '                    End If

    '                    If txtype <> "" Then
    '                        For J = 14 To iRows
    '                            If Not IsNothing(oSheet.Cells(J, i + 1).value) Then
    '                                Dim startingdate As Date = getdate(oSheet.Cells(J, i + 1).value.ToString, myyear)
    '                                Dim amount As Double = oSheet.Cells(J, i).value.ToString
    '                                If txtype = "catch up" Then
    '                                    'DbAdapter1.insertsalarytx(oSheet.Cells(J, 1).value.ToString, oSheet.Cells(J, 3).value.ToString, oSheet.Cells(J, 5).value.ToString, oSheet.Cells(J, 8).value.ToString, CDbl(oSheet.Cells(J, i).value.ToString), startingdate, txtype, myVerid)
    '                                    DbAdapter1.insertsalarytx(oSheet.Cells(J, 1).value.ToString, oSheet.Cells(J, 3).value.ToString, oSheet.Cells(J, 5).value.ToString, oSheet.Cells(J, 8).value.ToString, CDbl(oSheet.Cells(J, i).value.ToString), startingdate, txtype, myVerid, regionImport, myyear)
    '                                ElseIf txtype = "general increment" Then
    '                                    'get categoryid
    '                                    Dim categoryid As Integer = DbAdapter1.getcategoryId(oSheet.Cells(J, 1).value.ToString)
    '                                    'DbAdapter1.insertsalarytx(oSheet.Cells(J, 1).value.ToString, oSheet.Cells(J, 3).value.ToString, oSheet.Cells(J, 5).value.ToString, oSheet.Cells(J, 8).value.ToString, CDbl(oSheet.Cells(J, i).value.ToString), startingdate, txtype, myVerid)

    '                                    DbAdapter1.insertsalarytx(categoryid, amount, CDate(myyear & "-1-1"), txtype, myVerid)
    '                                End If

    '                            End If
    '                        Next
    '                    End If
    '                    If medicalplan <> "" Then
    '                        Dim familymemberid As Integer = 0
    '                        For J = 14 To iRows
    '                            If Not IsNothing(oSheet.Cells(J, i).value) Then

    '                                Dim expat As Boolean = False
    '                                If Not IsNothing(oSheet.Cells(J, 11).value) Then
    '                                    expat = True
    '                                End If
    '                                'Dim personid = DbAdapter1.getpersonid(oSheet.Cells(J, 8).value.ToString, expat, oSheet.Cells(J, 1).value.ToString.Substring(0, 2))
    '                                Dim personid = DbAdapter1.getpersonid(oSheet.Cells(J, 8).value.ToString)
    '                                'Dim joindate As String = DateFormatyyyyMMdd(CDate(oSheet.Cells(j, 5).value.ToString))
    '                                Dim JoinDateD As Date = CDate(oSheet.Cells(J, 5).value.ToString)
    '                                Dim othername As String = String.Empty
    '                                If Not IsNothing(oSheet.Cells(J, 9).value) Then
    '                                    othername = oSheet.Cells(J, 9).value.ToString
    '                                Else
    '                                    othername = "Null"
    '                                End If
    '                                Dim personjoindateid = DbAdapter1.getpersonjoindateid(personid, JoinDateD, othername, True)
    '                                If medicalplan = "1.5" Then
    '                                    familymemberid = DbAdapter1.insertfamilymemberid(personjoindateid, "Plan 1.5", oSheet.Cells(J, i).value.ToString, myyear, myVerid, regionImport)
    '                                ElseIf medicalplan = "2.2" Then
    '                                    familymemberid = DbAdapter1.insertfamilymemberid(personjoindateid, "Plan 2.2", oSheet.Cells(J, i).value.ToString, myyear, myVerid, regionImport)
    '                                ElseIf medicalplan = "3.2" Then
    '                                    familymemberid = DbAdapter1.insertfamilymemberid(personjoindateid, "Plan 3.2", oSheet.Cells(J, i).value.ToString, myyear, myVerid, regionImport)
    '                                ElseIf medicalplan = "7.2" Then
    '                                    familymemberid = DbAdapter1.insertfamilymemberid(personjoindateid, "Plan 7.2", oSheet.Cells(J, i).value.ToString, myyear, myVerid, regionImport)
    '                                End If
    '                            End If

    '                        Next
    '                    End If

    '                End If
    '            End If
    '            firstTime = False
    '        Next
    '        stopwatch.Stop()
    '        BackgroundWorker1.ReportProgress(2, "Elapsed Time: " & Format(stopwatch.Elapsed.Minutes, "00") & ":" & Format(stopwatch.Elapsed.Seconds, "00") & "." & stopwatch.Elapsed.Milliseconds.ToString)

    '        BackgroundWorker1.ReportProgress(3, "")
    '        myreturn = True
    '    Catch ex As Exception
    '        errMessage = ex.Message
    '        'cleanCurrentData(myyear, myVerid, regionimport)
    '    Finally
    '        oXl.Quit()
    '        'releaseComObject(oRange)
    '        releaseComObject(oSheet)
    '        releaseComObject(oWb)
    '        releaseComObject(oXl)

    '        GC.Collect()
    '        GC.WaitForPendingFinalizers()
    '        Try
    '            If Not aprocess Is Nothing Then
    '                aprocess.Kill()
    '            End If
    '        Catch ex As Exception
    '        End Try

    '    End Try
    '    Return myreturn
    'End Function


    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If CheckBox1.Checked Then
            Me.Close()
        End If
    End Sub

    Private Sub ImportHRInformation_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If (BackgroundWorker1.IsBusy) Then
            MsgBox("Please wait until the current process is finished")
            e.Cancel = True
        End If
    End Sub

    Private Sub ImportHRInformation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = BudgetYear
    End Sub

    Private Function CheckCOA(ByVal oSheet As Excel.Worksheet, ByVal sbchk As StringBuilder) As Boolean
        Dim myreturn As Boolean = True
        Dim i As Integer
        Dim j As Integer

        'Check 
        'Sapaccname
        'sapaccount
        'sapcc
        Dim dataset As New DataSet

        Dim iRows = oSheet.UsedRange.Rows.Count
        Dim iCols = oSheet.UsedRange.Columns.Count
        Dim sapaccname As String
        Dim sapaccnameid As Integer
        Dim sapaccount As String
        Dim sapaccountid As String
        Dim expensescheck As Integer = 0
        Dim sapccfid As Integer
        Dim sapindexfid As Integer
        Dim sapindexaccnamefid As Integer

        Dim sqlstr = "Select * from sapaccnamef; select * from sapaccountf;select * from sapccf; select * from sapindexf;select * from sapindexaccnamef"
        Dim errMessage As String = String.Empty
        If Not dbtools1.getDataSet(sqlstr, Dataset1, errMessage) Then
            Return myreturn
        End If

        Dataset1.Tables(0).TableName = "sapaccnamef"
        Dim keys0(0) As DataColumn
        keys0(0) = Dataset1.Tables(0).Columns(0)
        Dataset1.Tables(0).PrimaryKey = keys0

        Dataset1.Tables(1).TableName = "sapaccountf"
        Dim keys1(0) As DataColumn
        keys1(0) = Dataset1.Tables(1).Columns(0)
        Dataset1.Tables(1).PrimaryKey = keys1

        Dataset1.Tables(2).TableName = "sapccf"
        Dim keys2(0) As DataColumn
        keys2(0) = Dataset1.Tables(2).Columns(0)
        Dataset1.Tables(2).PrimaryKey = keys2

        Dataset1.Tables(3).TableName = "sapindexf"
        Dim keys3(0) As DataColumn
        keys3(0) = Dataset1.Tables(3).Columns(0)
        Dataset1.Tables(3).PrimaryKey = keys3

        Dataset1.Tables(4).TableName = "sapindexaccnamef"
        Dim keys4(0) As DataColumn
        keys4(0) = Dataset1.Tables(4).Columns(0)
        Dataset1.Tables(4).PrimaryKey = keys4

        'Read from coloumn and row
        For i = 14 To iCols
            BackgroundWorker1.ReportProgress(2, "Col: " & i & " of " & iCols)
            expensescheck = 0

            'If iCols = 31 Then
            '    Debug.WriteLine("debug mode")
            'End If
            If Not IsNothing(oSheet.Cells(5, i).value) Then
                sapaccname = oSheet.Cells(5, i).value.ToString
                sapaccount = oSheet.Cells(7, i).value.ToString
                sapaccnameid = 0
                sapaccountid = 0
                sapindexfid = 0
                sapindexaccnamefid = 0
                Dim myqry1 = From record In Dataset1.Tables(0)
                             Where record.Item("sapaccnamef") = sapaccname
                             Select record

                For Each rec In myqry1
                    sapaccnameid = rec.Item("sapaccnamefid")
                Next

                If sapaccnameid = 0 Then

                    'sbchk.Append(String.Format("Col: {0}, Row: {1} :: SAP ACCName Not Avail: {2}.", i, 5, sapaccname) & vbCrLf)
                    sbchk.Append(getErrorMessage(i, 5, "SapAccName", sapaccname))
                    'sbchk.Append("Col: " & i & " Row: " & 5 & " SAP ACCName Not Avail: " & sapaccname & vbCrLf)
                    expensescheck += 1
                    myreturn = False
                End If

                Dim myqry2 = From record In Dataset1.Tables("sapaccountf")
                             Where record.Item("sapaccountf") = sapaccount
                             Select record

                For Each rec In myqry2
                    sapaccountid = rec.Item("sapaccountfid")
                Next

                If sapaccountid = 0 Then
                    'sbchk.Append("Col: " & i & " Row: " & 7 & " SAP Account Not Avail: " & sapaccount & vbCrLf)
                    'sbchk.Append(String.Format("Col: {0}, Row: {1} :: SAP Account Not Avail: {2}.", i, 7, sapaccount) & vbCrLf)
                    sbchk.Append(getErrorMessage(i, 7, "SapAccount", sapaccount))
                    expensescheck += 1
                    myreturn = False
                End If



                If expensescheck = 0 Then 'No errors sapaccname and sapaccount
                    For j = 14 To iRows
                        'find 
                        'check for sapcc
                        'BackgroundWorker1.ReportProgress(2, "Col: " & i & " of " & iCols & ", Row: " & j & " of " & iRows)
                        BackgroundWorker1.ReportProgress(2, String.Format("Col: {0} of {1}, Row: {2} of {3}", i, iCols, j, iRows))
                        Dim mysapcc As String = String.Empty

                        'get costcenter  based on sapcc
                        If Not IsNothing(oSheet.Cells(j, i).value) Then
                            Dim mydept = oSheet.Cells(j, 3).value.ToString
                            mysapcc = mydept
                            Dim last2digit = mydept.Substring(4, 2)
                            If Not IsNothing(oSheet.Cells(8, i).value) Then
                                Dim sapcc As String = oSheet.Cells(8, i).value.ToString
                                Dim sapccobj() As String = sapcc.Split(",")
                                If sapccobj.Length > 1 Then
                                    If last2digit <> sapccobj(1) Then
                                        mysapcc = sapccobj(0) & last2digit
                                    End If
                                Else
                                    mysapcc = sapcc & last2digit

                                End If
                            End If
                            'get sapcc

                            Dim myqry3 = From record In Dataset1.Tables("sapccf")
                                         Where record.Item("sapccf") = mysapcc
                                         Select record

                            For Each rec In myqry3
                                sapccfid = rec.Item("sapccfid")
                            Next

                            If sapccfid = 0 Then
                                'sbchk.Append("Col: " & i & " Row: " & j & " SAPCC Not Avail: " & mysapcc & " Dept: " & mydept & " SAPACCName :" & sapaccname & " SAPAccount: " & sapaccount & vbCrLf)
                                'sbchk.Append(String.Format("Col: {0}, Row: {1} ::  SAPCC Not Avail: {2}, Dept: {3}, SAPAccount: {4}, SAPACCName : {5}.", i, j, mysapcc, mydept, sapaccount, sapaccname) & vbCrLf)
                                sbchk.Append(getErrorMessage(i, j, mysapcc, mydept, sapaccount, sapaccname, "SAP CC"))
                                expensescheck += 1
                                myreturn = False
                            Else
                                'check sap index

                                Dim myqry4 = From record In Dataset1.Tables("sapindexf")
                                         Where record.Item("sapaccountfid") = sapaccountid And record.Item("sapccfid") = sapccfid
                                         Select record

                                For Each rec In myqry4
                                    sapindexfid = rec.Item("sapindexfid")
                                Next
                                'check sap index accname
                                If sapindexfid = 0 Then
                                    'sbchk.Append("Col: " & i & " Row: " & j & " SAPCC Not Avail: " & mysapcc & " Dept: " & mydept & " SAPACCName :" & sapaccname & " SAPAccount: " & sapaccount & vbCrLf)
                                    'sbchk.Append(String.Format("Col: {0}, Row: {1} ::  SAPCC Not Avail: {2}, Dept: {3}, SAPAccount: {4}, SAPACCName : {5}.", i, j, mysapcc, mydept, sapaccount, sapaccname) & vbCrLf)
                                    sbchk.Append(getErrorMessage(i, j, mysapcc, mydept, sapaccount, sapaccname, "SAP-INDEX"))
                                    myreturn = False
                                Else
                                    Dim myqry5 = From record In Dataset1.Tables("sapindexaccnamef")
                                         Where record.Item("sapindexfid") = sapindexfid And record.Item("sapaccnamefid") = sapaccnameid
                                         Select record

                                    For Each rec In myqry5
                                        sapindexaccnamefid = rec.Item("sapindexaccnamefid")
                                    Next
                                    If sapindexaccnamefid = 0 Then
                                        'sbchk.Append("Col: " & i & " Row: " & j & " SAPCC Not Avail: " & mysapcc & " Dept: " & mydept & " SAPACCName :" & sapaccname & " SAPAccount: " & sapaccount & vbCrLf)
                                        'sbchk.Append(String.Format("Col: {0}, Row: {1} ::  SAPCC Not Avail: {2}, Dept: {3}, SAPAccount: {4}, SAPACCName : {5}.", i, j, mysapcc, mydept, sapaccount, sapaccname) & vbCrLf)
                                        sbchk.Append(getErrorMessage(i, j, mysapcc, mydept, sapaccount, sapaccname, "SAP-INDEX & SAPACCNAME"))
                                        myreturn = False
                                    End If
                                End If

                            End If
                        End If
                        

                    Next
                End If

            End If
        Next

        
        Return myreturn
    End Function

    'Private Function getErrorMessage(ByVal i As Integer, ByVal j As Integer, ByVal mysapcc As String, ByVal mydept As String, ByVal sapaccount As String, ByVal sapaccname As String, ByVal CheckDesc As String) As String
    '    Dim myReturn = String.Format("Col: {0}, Row: {1} :: {6} - SAPCC Not Avail: {2}, Dept: {3}, SAPAccount: {4}, SAPACCName : {5}.", i, j, mysapcc, mydept, sapaccount, sapaccname) & vbCrLf
    '    Return myReturn
    'End Function

    'Private Function getErrorMessage(ByVal i As Integer, ByVal row As Integer, ByVal errDesc As String, ByVal desc As String) As String
    '    Dim myreturn As String = String.Format("Col: {0}, Row: {1} :: {2} Not Avail: {3}.", i, row, errDesc, desc) & vbCrLf
    '    Return myreturn
    'End Function

End Class
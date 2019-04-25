Imports System.ComponentModel
Imports HR.HelperClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports DJLib.Dbtools

Public Class CalculateExpenses01

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
            Try
                TextBox2.Text = ""
                TextBox3.Text = ""
                BackgroundWorker1.WorkerReportsProgress = True
                BackgroundWorker1.WorkerSupportsCancellation = True

                BackgroundWorker1.RunWorkerAsync()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

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
        Status = Calculate(errMsg)
        If Status Then
            BackgroundWorker1.ReportProgress(2, TextBox2.Text & " Done.")
        Else
            BackgroundWorker1.ReportProgress(3, "Error::" & errMsg)
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If CheckBox1.Checked Then
            Me.Close()
        End If
    End Sub

    Private Sub CalculateExpenses_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If (BackgroundWorker1.IsBusy) Then
            MsgBox("Please wait until the current process is finished")
            e.Cancel = True
        End If
    End Sub

    Private Sub CalculateExpenses_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = BudgetYear
    End Sub

    Private Function Calculate(ByRef message) As Boolean
        Dim myReturn As Boolean = False
        'Threading.Thread.Sleep(5000)
        Dim Dataset1 As New DataSet
        Dim sqlstr As String = String.Empty
        Dim sb As New StringBuilder
        Dim stringBuilder1 As New StringBuilder

        myyear = DateTimePicker1.Value.Year
        Dim BeginingYear As Date = CDate(DateTimePicker1.Value.Year & "-1-1")
        Dim EndOFYear As Date = CDate(DateTimePicker1.Value.Year & "-12-31")
        'get expensesdetailtx

        'Table 0 expenses nature
        sb.Append("select expensesdetailtxid,sapaccname,expensesnatureid,expensesnature,sapaccount,sapaccid,sapcc,dept,currency,fullyear from fgetexpensesdetailtx(")
        sb.Append(myyear)
        sb.Append(") as mytable(expensesdetailtxid integer,sapaccname character varying,expensesnatureid integer,expensesnature character varying,sapaccount character varying,sapaccid character varying,sapcc character varying,dept character varying,currency character varying,fullyear boolean);")

        'Table 1 expensesnaturemonths
        sb.Append("select expensesnatureid,mymonth,mymonthint from expensesnaturemonths where myyear =")
        sb.Append(myyear)
        sb.Append(";")

        'Table 2 personjoindatecategory
        'sb.Append("select * from personcategoryview;")
        sb.Append("select * from  fgetpersoncategory(")
        sb.Append(DateFormatyyyyMMdd(BeginingYear))
        sb.Append(") as a(personjoindatecategoryid integer,personname character varying, othername character varying,joindate date,title character varying, category character varying,dept character varying,expat boolean,enddate date,headcount numeric);")

        'Table 3 category
        sb.Append("select * from fgetcategory(")
        sb.Append(myyear)
        sb.Append(") as c(categoryid integer,category character varying,categorytypeid integer,categorytype character varying, amount numeric,myyear integer);")

        'Table 4 paramdt
        sb.Append("select * from paramdt;")

        'table 5 personexpensesdtl
        sb.Append("select personjoindatecategoryid , ped.validdate,ped.amount,en.expensesnature,ped.personexpensesid " & _
                  " from  personexpensesdtl ped " & _
                  " left join personexpenses pe on pe.personexpensesid = ped.personexpensesid" & _
                  " left join expensesdetailtx edtx on edtx.expensesdetailtxid = pe.expensesdetailtxid" & _
                  " left join expensesdetail ed on ed.expensesdetailid = edtx.expensesdetailid" & _
                  " left join accexpenses ac on ac.accexpensesid = ed.accexpensesid" & _
                  " left join sapaccname san on san.sapaccnameid = ac.sapaccnameid" & _
                  " left join expensesnature en on en.expensesnatureid = ac.expensesnatureid" & _
                  " where edtx.myyear = ")
        sb.Append(myyear)
        sb.Append(";")

        'table 6 catch up Salary
        sb.Append("select personjoindatecategoryid,amount,validfrom from salarytx where validfrom >= ")
        sb.Append(DateFormatyyyyMMdd(BeginingYear) & " and validfrom <= ")
        sb.Append(DateFormatyyyyMMdd(EndOFYear) & ";")

        'table 7 version
        sb.Append("select verid from ver order by verid limit 1;")

        'table 8 personexpenses
        sb.Append("select * from personexpenses;")

        'table 9 personexpensesdtl for new comer
        sb.Append("select pjc.personjoindatecategoryid , ped.validdate,ped.amount,en.expensesnature,ped.personexpensesid " & _
                  " from personjoindate pj" & _
                  " left join person p on p.personid = pj.personid" & _
                  " left join personjoindatecategory pjc on pjc.personjoindateid = pj.personjoindateid" & _
                  " left join personexpenses pe on pe.personjoindatecategoryid = pjc.personjoindatecategoryid" & _
                  " left join personexpensesdtl ped on ped.personexpensesid = pe.personexpensesid" & _
                  " left join expensesdetailtx edtx on edtx.expensesdetailtxid = pe.expensesdetailtxid" & _
                  " left join expensesdetail ed on ed.expensesdetailid = edtx.expensesdetailid" & _
                  " left join accexpenses ac on ac.accexpensesid = ed.accexpensesid" & _
                  " left join sapaccname san on san.sapaccnameid = ac.sapaccnameid" & _
                  " left join expensesnature en on en.expensesnatureid = ac.expensesnatureid" & _
                  " where not ped.validdate isnull and pj.joindate >= ")
        sb.Append(DateFormatyyyyMMdd(BeginingYear) & " and pj.joindate <= ")
        sb.Append(DateFormatyyyyMMdd(EndOFYear) & ";")

        'table 10 13month
        sb.Append("select pe.personexpensesid,pe.personjoindatecategoryid,e.expensesnatureid,e.expensesnature from personexpenses pe" & _
                   " left join expensesdetailtx edtx on edtx.expensesdetailtxid = pe.expensesdetailtxid" & _
                   " left join expensesdetail ed on ed.expensesdetailid = edtx.expensesdetailid" & _
                   " left join accexpenses ae on ae.accexpensesid = ed.accexpensesid" & _
                   " left join expensesnature e on e.expensesnatureid = ae.expensesnatureid" & _
                   " where(edtx.myyear = " & myyear & ")" & _
                   " order by e.expensesnature")


        sqlstr = sb.ToString
        Try
            '::Rule::
            '1. Read from table expensesdetailtx
            '2. Find eligible person for each expenses
            '2. Check NatureExpenses for special calculation
            '3. Apply common calculation
            '4. Create Budget with version 1

            '1.Read from table expensesdetailtx
            If dbtools1.getDataSet(sqlstr, Dataset1, message) Then

                Dim tbexpensesdt = Dataset1.Tables(1)
                Dim tbperson = Dataset1.Tables(2)

                Dim tbcategory = Dataset1.Tables(3)
                Dim key3(2) As DataColumn
                key3(0) = tbcategory.Columns("category")
                key3(1) = tbcategory.Columns("categorytype")
                key3(2) = tbcategory.Columns("myyear")
                tbcategory.PrimaryKey = key3

                Dim tbparamdt = Dataset1.Tables(4)
                Dim keyTbparamdt(0) As DataColumn
                keyTbparamdt(0) = tbparamdt.Columns("paramname")
                tbparamdt.PrimaryKey = keyTbparamdt

                Dim tbpersonexpensesdtl = Dataset1.Tables(5)
                Dim KeyTbPersonExpensesDtl(2) As DataColumn
                KeyTbPersonExpensesDtl(0) = tbpersonexpensesdtl.Columns("personjoindatecategoryid")
                KeyTbPersonExpensesDtl(1) = tbpersonexpensesdtl.Columns("expensesnature")
                KeyTbPersonExpensesDtl(2) = tbpersonexpensesdtl.Columns("validdate")
                tbpersonexpensesdtl.PrimaryKey = KeyTbPersonExpensesDtl


                Dim tbCatchUpSalary = Dataset1.Tables(6)
                Dim key6(0) As DataColumn
                key6(0) = tbCatchUpSalary.Columns("personjoindatecategoryid")
                tbCatchUpSalary.PrimaryKey = key6

                Dim tbVer = Dataset1.Tables(7)
                Dim verid As Integer = tbVer.Rows(0).Item("verid")

                Dim tbpersonexpenses = Dataset1.Tables(8)
                Dim key8(1) As DataColumn
                key8(0) = tbpersonexpenses.Columns("personjoindatecategoryid")
                key8(1) = tbpersonexpenses.Columns("expensesdetailtxid")
                tbpersonexpenses.PrimaryKey = key8

                Dim tbpersonexpensesdtlnewcomer = Dataset1.Tables(9)
                Dim Key9(2) As DataColumn
                Key9(0) = tbpersonexpensesdtlnewcomer.Columns("personjoindatecategoryid")
                Key9(1) = tbpersonexpensesdtlnewcomer.Columns("expensesnature")
                Key9(2) = tbpersonexpensesdtlnewcomer.Columns("validdate")
                tbpersonexpensesdtlnewcomer.PrimaryKey = Key9

                Dim tbPersonJDExpenses = Dataset1.Tables(10)
                Dim Key10(2) As DataColumn
                Key10(0) = tbPersonJDExpenses.Columns(0)
                Key10(1) = tbPersonJDExpenses.Columns(1)
                Key10(2) = tbPersonJDExpenses.Columns(2)
                tbPersonJDExpenses.PrimaryKey = Key10


                'Get Parameter from tbparam
                Dim GeneralRate As Double
                Dim GeneralIncrMonth As Integer
                Dim ExpatRate As Double
                Dim ExpatIncrMonth As Integer
                Dim serviceyear10 As Double
                Dim serviceyear15 As Double
                Dim incrMonth As Integer
                Dim MonthDesc As New Dictionary(Of Integer, String)
                fillMonthDesc(MonthDesc)


                For i = 0 To tbparamdt.Rows.Count - 1
                    Dim dr = tbparamdt.Rows(i)
                    Select Case dr.Item(2).ToString
                        Case "General Rate"
                            GeneralRate = CDbl(dr.Item("nvalue").ToString)
                            GeneralIncrMonth = CInt(dr.Item("ivalue").ToString)
                        Case "Expat Rate"
                            ExpatRate = CDbl(dr.Item("nvalue").ToString)
                            ExpatIncrMonth = CInt(dr.Item("ivalue").ToString)
                        Case "10"
                            serviceyear10 = CDbl(dr.Item("nvalue").ToString)
                        Case "15"
                            serviceyear15 = CDbl(dr.Item("nvalue").ToString)
                    End Select
                Next

                Dim mymessage As String = String.Empty
                For i = 0 To Dataset1.Tables(0).Rows.Count - 1
                    Dim dr As DataRow = Dataset1.Tables(0).Rows(i)
                    Dim expensesdetailtxid = dr.Item("expensesdetailtxid")
                    Dim expensesnature As String = dr.Item("expensesnature").ToString
                    Dim sapaccount As String = dr.Item("sapaccount").ToString
                    Dim sapcc As String = dr.Item("sapcc").ToString
                    Dim dept As String = dr.Item("dept").ToString

                    If Not (expensesnature.Contains("13th Month")) Then


                        '2 select person for each expenses from table person categoryview using linq
                        Dim myquery = From persons In Dataset1.Tables(2)
                                      Where persons.Item("dept").ToString = dept
                                      Select persons Order By persons.Item("personname")


                        For Each p In myquery

                            'BackgroundWorker1.ReportProgress(2, String.Format("{0} {1} {2} {3} ", dr.Item(1).ToString, dr.Item(3).ToString, dr.Item(5).ToString, dr.Item(6).ToString))
                            Dim serviceyear As Double = 0.0
                            Dim incr As Double = 0
                            Dim lastmonth As Integer
                            serviceyear = Math.Round(((CDate(DateTimePicker1.Value.Year & "-12-31") - CDate(p.Item(3).ToString)).Days / 365), 2)

                            'get personexpensesid for budgetRecord
                            Dim personjoindatecategoryid = p.Item(0)

                            Dim personexpensesid As Integer = 0
                            Dim headcount As Double = p.Item(9)
                            Dim mycategory As String = p.Item(5)
                            Dim joindate As Date = p.Item(3)

                            Dim pq = From personexpenses In Dataset1.Tables(8)
                                     Where personexpenses.Item("personjoindatecategoryid") = personjoindatecategoryid And personexpenses.Item("expensesdetailtxid") = expensesdetailtxid
                                     Select personexpenses
                            For Each pe In pq
                                personexpensesid = pe.Item("personexpensesid")
                            Next



                            If personexpensesid = 0 Then
                                'Create personexpensesid
                                personexpensesid = DbAdapter1.getpersonexpensesid(personjoindatecategoryid, expensesdetailtxid)
                            End If



                            Dim SalaryDict As New Dictionary(Of Integer, Double)
                            If dr.Item(3).ToString = "Basic Salary" Then

                                Dim baseSalary As Double = 0
                                Dim validdate As Date
                                BackgroundWorker1.ReportProgress(2, String.Format("Processing {0}", "*** Basic Salary, 13th Month ***"))


                                Dim q = From catchup In Dataset1.Tables(6)
                                        Where catchup.Item("personjoindatecategoryid") = personjoindatecategoryid
                                        Select catchup

                                Dim catchupValue As Double
                                Dim catchupDate As Date

                                Dim CatchupValueDict As New Dictionary(Of Integer, Double)

                                For L = 1 To 12
                                    CatchupValueDict.Add(L, 0)
                                    SalaryDict.Add(L, 0)
                                Next

                                For Each myresult In q
                                    catchupValue = myresult.Item("amount")
                                    catchupDate = myresult.Item("validfrom")
                                    CatchupValueDict(catchupDate.Month) = catchupValue
                                Next
                                '

                                'getBaseSalary from personexpensesdtl (the first salary in year budget)
                                Dim qry = From expenses In Dataset1.Tables(5)
                                          Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = "Basic Salary" And expenses.Item("validdate") < CDate(DateTimePicker1.Value.Year & "/1/1")
                                          Select expenses Order By expenses.Item("validdate") Descending

                                For Each myresult In qry
                                    baseSalary = myresult.Item("amount")
                                    validdate = myresult.Item("validdate")

                                    If baseSalary = 0 Then

                                    End If
                                    Exit For
                                Next


                                Dim qry2 = From expenses In Dataset1.Tables(9)
                                        Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = "Basic Salary" And expenses.Item("validdate") >= CDate(DateTimePicker1.Value.Year & "/01/01") And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                                        Select expenses Order By expenses.Item("validdate") Ascending
                                For Each myresult2 In qry2
                                    baseSalary = myresult2.Item("amount")
                                    validdate = myresult2.Item("validdate")
                                    Exit For
                                Next


                                'If baseSalary = 0 Then
                                '    Debug.WriteLine("debugMode")
                                'End If
                                If p.Item(7) Then 'Expat Calculation
                                    incr = IIf(serviceyear > 1, ExpatRate, 0)
                                    incrMonth = ExpatIncrMonth
                                Else 'General Calculation
                                    incr = IIf(serviceyear > 1, GeneralRate, 0)
                                    incrMonth = GeneralIncrMonth
                                End If

                                If p.Item(8).ToString = "" Then
                                    lastmonth = 12
                                Else
                                    lastmonth = CDate(p.Item("enddate").ToString).Month
                                End If
                                Dim totalsalary As Double = 0
                                For K = 1 To lastmonth
                                    Dim kDate As Date = CDate(myyear & "-" & K & "-1")
                                    If K = incrMonth Then
                                        If serviceyear > 1 Then
                                            baseSalary *= (incr + 1)
                                            'No!! don't create new record in personexpensesdetailtx for basic salary change
                                            'Try to see the result
                                            DbAdapter1.insertpersonexpensesdtl(personexpensesid, baseSalary, kDate)
                                        End If
                                    End If
                                    'Check Catchup
                                    If CatchupValueDict(K) <> 0 Then
                                        baseSalary *= (1 + CatchupValueDict(K))
                                        'No !! don't create new record in personexpensesdetailtx for basic salary change
                                        'Try to see the result
                                        DbAdapter1.insertpersonexpensesdtl(personexpensesid, baseSalary, kDate)
                                    End If
                                    'create record budget
                                    If validdate <= kDate Then
                                        'totalsalary += baseSalary
                                        SalaryDict(K) = baseSalary
                                        'If baseSalary = 0 Then
                                        '    Debug.WriteLine("debugMode")
                                        'End If
                                        'Dim mydate As String = "'" & DateTimePicker1.Value.Year & "-" & K & "-1'"
                                        Dim mydate As String = DateFormatyyyyMMdd(kDate)
                                        Debug.WriteLine("Person Name {0} Dept {1} personexpensesid {2} ExpensesNature {3} Personjoindatecategory {4} myDate {5}", p.Item("personname"), p.Item("dept"), personexpensesid, dr.Item("expensesnature"), p.Item("personjoindatecategoryid"), mydate)
                                        createrecord(stringBuilder1, personexpensesid, baseSalary, verid, mydate)
                                    End If
                                Next



                                '*************** 13th Month
                                personexpensesid = 0
                                Dim qry3 = From category In Dataset1.Tables(3)
                                      Where category.Item("category") = mycategory And category.Item("categorytype") = "13th Month" And category.Item("myyear") = myyear
                                      Select category
                                Dim checkcategory As Boolean = False
                                For Each dt In qry3
                                    'if listed in category then calculate
                                    '
                                    Debug.WriteLine("CheckTrue")
                                    checkcategory = True

                                    If checkcategory Then
                                        Dim enttitlement = IIf(serviceyear > 1, 1, serviceyear)
                                        Dim lastsalary As Double = 0
                                        Try
                                            lastsalary = SalaryDict(12)
                                        Catch ex As Exception

                                        End Try
                                        If joindate > CDate(myyear & "-10-01") Then lastsalary = 0
                                        If lastsalary > 0 Then

                                            Dim qry4 = From e In Dataset1.Tables(10)
                                                       Where e.Item("expensesnature") = "13th Month" And e.Item("personjoindatecategoryid") = personjoindatecategoryid
                                                       Select e.Item("personexpensesid")

                                            For Each result In qry4
                                                personexpensesid = result
                                            Next

                                            If personexpensesid = 0 Then
                                                'personexpensesid = DbAdapter1.getExpensesDetailtxid(myyear, "13th Month", dept, personjoindatecategoryid)
                                            End If
                                            createrecord(stringBuilder1, personexpensesid, lastsalary * enttitlement, verid, DateFormatyyyyMMdd(CDate(myyear & "-12-1")))
                                            Debug.WriteLine("PersonExpensesid {0} {1} ", personexpensesid, personjoindatecategoryid)
                                        End If
                                    End If
                                Next

                            ElseIf dr.Item(3).ToString = "Red Pocket" Then
                                '    Debug.WriteLine("Apply Red Pocket")
                            ElseIf Not IsDBNull(dr.Item(9)) Then
                                '    If dr.Item(9) = True Then
                                '        Debug.WriteLine("Apply common Fullyear")
                                '    Else
                                '        Debug.WriteLine("Apply common Monthly")
                                '    End If
                            Else
                                '    Debug.WriteLine("Expenses is not applicable")
                            End If
                            'Insert Head Count



                            '3. Check NatureExpenses for special calculation



                        Next
                    End If

                Next

                'copy expensesdetail
                myReturn = True
                sqlstr = "delete from budgettx where mydate >= " & DateFormatyyyyMMdd(BeginingYear) & " and mydate <= " & DateFormatyyyyMMdd(EndOFYear) & " and ver = " & verid & ";copy budgettx(personexpensesid,amount,ver,mydate) from stdin;"

                If stringBuilder1.ToString <> "" Then
                    message = dbtools1.copy(sqlstr, stringBuilder1.ToString, myReturn)
                    If message <> "" Then
                        myReturn = False
                    End If
                    BackgroundWorker1.ReportProgress(2, "Copy To Db (BudgetTx)")

                Else
                    BackgroundWorker1.ReportProgress(2, "Nothing to Copy.")
                    myReturn = True
                End If
                BackgroundWorker1.ReportProgress(3, message)


            End If
        Catch ex As Exception
            myReturn = False
            message = ex.Message
        Finally

        End Try
        Return myReturn
    End Function

    Private Sub fillMonthDesc(ByRef MonthDesc As Dictionary(Of Integer, String))
        MonthDesc.Add(1, "01-Jan")
        MonthDesc.Add(2, "02-Feb")
        MonthDesc.Add(3, "03-Mar")
        MonthDesc.Add(4, "04-Apr")
        MonthDesc.Add(5, "05-May")
        MonthDesc.Add(6, "06-Jun")
        MonthDesc.Add(7, "07-Jul")
        MonthDesc.Add(8, "08-Aug")
        MonthDesc.Add(9, "09-Sep")
        MonthDesc.Add(10, "10-Oct")
        MonthDesc.Add(11, "11-Nov")
        MonthDesc.Add(12, "12-Dec")

    End Sub

    Private Sub createrecord(ByRef stringBuilder1 As StringBuilder, ByVal personexpensesid As Integer, ByVal amount As Double, ByVal verid As Integer, ByVal mydate As String)
        stringBuilder1.Append(personexpensesid & vbTab)
        stringBuilder1.Append(amount & vbTab)
        stringBuilder1.Append(verid & vbTab)
        stringBuilder1.Append(mydate & vbCrLf)
    End Sub






End Class
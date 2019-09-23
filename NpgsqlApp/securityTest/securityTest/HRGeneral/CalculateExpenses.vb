Imports System.ComponentModel
Imports HR.HelperClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports DJLib.Dbtools

Public Class CalculateExpenses
    Dim bypasserror As Boolean = False
    Private WithEvents BackgroundWorker1 As New BackgroundWorker
    Dim FileName As String = String.Empty
    Dim Status As Boolean = False
    Dim Dataset1 As DataSet
    Dim ConnectionString As String = dbtools1.getConnectionString
    Dim aprocesses() As Process = Nothing '= Process.GetProcesses
    Dim aprocess As Process = Nothing
    Dim myyear As Integer
    Dim BeginingYear As Date
    Dim EndOFYear As Date

    'Tbparamdtl fields
    Dim GeneralRate As Double
    Dim GeneralIncrMonth As Integer
    Dim ExpatRate As Double
    Dim ExpatIncrMonth As Integer
    Dim serviceyear10 As Double
    Dim serviceyear15 As Double
    Dim incrMonth As Integer
    Dim AmountA As Double
    Dim AmountB As Double
    Dim AmountC As Double
    Dim MPFValue As Double
    Dim MPFFloorValue As Double
    Dim myRegionId As Integer

    Dim stringBuilder1 As New StringBuilder
    Dim verid As Integer
    Dim myverid As Integer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not (BackgroundWorker1.IsBusy) Then
            Try
                TextBox2.Text = ""
                TextBox3.Text = ""
                BackgroundWorker1.WorkerReportsProgress = True
                BackgroundWorker1.WorkerSupportsCancellation = True
                stringBuilder1.Clear()

                'check lastversion
                If DbAdapter1.validversion(ComboBox1.SelectedValue, DateTimePicker1.Value.Year, "CV" & dbtools1.Region, dbtools1.Region) Then
                    'update currentversion
                    If DbAdapter1.setcurrentversion(ComboBox1.SelectedValue, DateTimePicker1.Value.Year, "CV" & dbtools1.Region, dbtools1.Region, ComboBox1.Text) Then
                        BackgroundWorker1.RunWorkerAsync()
                    End If

                Else
                    Throw New System.Exception("This version is closed")
                End If

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
        Dim stopwatch As New Stopwatch

        stopwatch.Start()
        Dim errMsg As String = String.Empty
        Status = Calculate(errMsg)
        stopwatch.Stop()
        If Status Then
            BackgroundWorker1.ReportProgress(2, TextBox2.Text & " Done.")
            BackgroundWorker1.ReportProgress(3, TextBox3.Text & "Elapsed Time::" & Format(stopwatch.Elapsed.Minutes, "00") & ":" & Format(stopwatch.Elapsed.Seconds, "00") & "." & stopwatch.Elapsed.Milliseconds.ToString)
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
        myRegionId = dbtools1.RegionId
        'dbtools1.Region = "HK" Then

        DateTimePicker1.Value = BudgetYear
        loadcombobox()
    End Sub
    Private Sub loadcombobox()
        'Clear combobox first
        Dim sqlstr As String = "select verid,hrver from ver where ver.myyear = " & DateTimePicker1.Value.Year & "order by myorder;"

        dbtools1.FillComboboxDataSource(ComboBox1, sqlstr)
        'Call calculateExpensesLoad()
    End Sub
    Private Sub buildQuery(ByRef sb As StringBuilder)
        'Table 0 expenses nature
        'Dim sqlstr As String = String.Empty
        'sb.Append("select expensesdetailtxid,sapaccname,expensesnatureid,expensesnature,sapaccount,sapaccid,sapcc,dept,currency,fullyear,indexcostcenterdeptid from fgetexpensesdetailtx(")
        'sb.Append(myyear)
        'sb.Append(") as mytable(expensesdetailtxid integer,sapaccname character varying,expensesnatureid integer,expensesnature character varying,sapaccount character varying,sapaccid character varying,sapcc character varying,dept character varying,currency character varying,fullyear boolean,indexcostcenterdeptid integer)")
        'sb.Append(" where expensesnature = 'Basic Salary' or expensesnature = 'Salary & Wage';")
        'sqlstr = sb.ToString
        Dim sqlstr As String = String.Empty
        sb.Append("select expensesdetailtxid,sapaccname,expensesnatureid,expensesnature,sapaccount,sapaccid,sapcc,dept,currency,fullyear,indexcostcenterdeptid from fgetexpensesdetailtx(")
        sb.Append(myyear)
        sb.Append(") as mytable(expensesdetailtxid integer,sapaccname character varying,expensesnatureid integer,expensesnature character varying,sapaccount character varying,sapaccid character varying,sapcc character varying,dept character varying,currency character varying,fullyear boolean,indexcostcenterdeptid integer)")
        sb.Append(" where expensesnature = 'Basic Salary';")
        sqlstr = sb.ToString

        'Table 1 expensesnaturemonths
        sb.Append("select expensesnatureid,mymonth,mymonthint from expensesnaturemonths where myyear =")
        sb.Append(myyear)
        sb.Append(";")
        sqlstr = sb.ToString
        'Table 2 personjoindatecategory
        'sb.Append("select * from personcategoryview;")
        'sqlstr = "select * from  fgetpersoncategory(" & DateFormatyyyyMMdd(BeginingYear) & "," & dbtools1.RegionId & "," & myverid & ") as a(personjoindatecategoryid integer,personname character varying, othername character varying,joindate date,title character varying, category character varying,dept character varying,expat boolean,enddate date,headcount numeric,personjoindateid integer);"
        sqlstr = "select * from  fgetpersoncategory(" & DateFormatyyyyMMdd(BeginingYear) & "," & dbtools1.RegionId & "," & myverid & ") as a(personjoindatecategoryid integer,personname character varying, othername character varying,joindate date,title character varying, category character varying,dept character varying,expat boolean,enddate date,headcount numeric,personjoindateid integer,effectivedatestart date,effectivedateend date,bonusfactor integer);"
        sb.Append(sqlstr)

        'Table 3 category
        sqlstr = "select * from fgetcategory(" & myyear & "," & myverid & ") as c(categoryid integer,category character varying,categorytypeid integer,categorytype character varying, amount numeric,myyear integer,mpfcategory character varying);"
        sb.Append(sqlstr)

        'Table 4 paramdt
        sqlstr = "select * from paramdt pd " & _
                  " left join paramhd ph on ph.paramhdid = pd.paramhdid" & _
                  " where pd.dvalue <= " & DateFormatyyyyMMdd(EndOFYear) & " and ph.cvalue = '" & dbtools1.Region & "';"
        sb.Append(sqlstr)
        'table 5 personexpensesdtl
        'sqlstr = "select i.personjoindatecategoryid , ped.validdate,ped.amount,en.expensesnature,ped.personexpensesid,i.indexcostcenterdeptid,i.icdpjcid " & _
        '          " from  personexpensesdtl ped " & _
        '          " left join personexpenses pe on pe.personexpensesid = ped.personexpensesid" & _
        '          " left join icdpjc i on i.icdpjcid = pe.icdpjcid" & _
        '          " left join personjoindatecategory pjc on pjc.personjoindatecategoryid = i.personjoindatecategoryid" & _
        '          " left join personjoindate pd on pd.personjoindateid = pjc.personjoindateid" & _
        '          " left join person p on p.personid = pd.personid" & _
        '          " left join expensesdetailtx edtx on edtx.expensesdetailtxid = pe.expensesdetailtxid" & _
        '          " left join expensesdetail ed on ed.expensesdetailid = edtx.expensesdetailid" & _
        '          " left join accexpenses ac on ac.accexpensesid = ed.accexpensesid" & _
        '          " left join sapaccname san on san.sapaccnameid = ac.sapaccnameid" & _
        '          " left join expensesnature en on en.expensesnatureid = ac.expensesnatureid" & _
        '          " where edtx.myyear = " & myyear & " and p.regionid = " & dbtools1.RegionId & " and i.verid = " & myverid & ";"

        sqlstr = "select i.personjoindatecategoryid , ped.validdate,ped.amount,en.expensesnature,ped.personexpensesid,i.indexcostcenterdeptid,i.icdpjcid,cc.sapcc,c.category " & _
                  " from  personexpensesdtl ped " & _
                  " left join personexpenses pe on pe.personexpensesid = ped.personexpensesid" & _
                  " left join icdpjc i on i.icdpjcid = pe.icdpjcid" & _
                  " left join indexcostcenterdept icd on icd.indexcostcenterdeptid = i.indexcostcenterdeptid" & _
                  " left join indexcostcenter ic on ic.indexcostcenterid = icd.indexcostcenterid" & _
                  " left join sapcc cc on cc.sapccid = ic.sapccid" & _
                  " left join personjoindatecategory pjc on pjc.personjoindatecategoryid = i.personjoindatecategoryid" & _
                  " left join category c on c.categoryid = pjc.categoryid" & _
                  " left join personjoindate pd on pd.personjoindateid = pjc.personjoindateid" & _
                  " left join person p on p.personid = pd.personid" & _
                  " left join expensesdetailtx edtx on edtx.expensesdetailtxid = pe.expensesdetailtxid" & _
                  " left join expensesdetail ed on ed.expensesdetailid = edtx.expensesdetailid" & _
                  " left join accexpenses ac on ac.accexpensesid = ed.accexpensesid" & _
                  " left join sapaccname san on san.sapaccnameid = ac.sapaccnameid" & _
                  " left join expensesnature en on en.expensesnatureid = ac.expensesnatureid" & _
                  " where edtx.myyear = " & myyear & " and c.regionid = " & dbtools1.RegionId & " and i.verid = " & myverid & ";"

        sb.Append(sqlstr)
        'table 6 catch up Salary
        sb.Append("select stx.personjoindatecategoryid,amount,validfrom,txtype from salarytx stx" & _
                  " left join personjoindatecategory pjc on pjc.personjoindatecategoryid = stx.personjoindatecategoryid" & _
                  " left join category c on c.categoryid = pjc.categoryid" & _
                  " left join personjoindate pj on pj.personjoindateid = pjc.personjoindateid" & _
                  " left join person p on p.personid = pj.personid" & _
                  " where validfrom >= " & DateFormatyyyyMMdd(BeginingYear) & " and validfrom <= " & DateFormatyyyyMMdd(EndOFYear) & "and stx.verid = " & myverid & " and c.regionid = " & dbtools1.RegionId & ";")
        sqlstr = sb.ToString
        'table 7 version
        sb.Append("select verid from ver order by verid limit 1;")
        sqlstr = sb.ToString
        'table 8 personexpenses

        '**************check personexpenses 

        sqlstr = "select * from personexpenses pe" & _
                  " left join icdpjc i on i.icdpjcid = pe.icdpjcid" & _
                  " left join personjoindatecategory pjc on pjc.personjoindatecategoryid = i.personjoindatecategoryid" & _
                  " left join category c on c.categoryid = pjc.categoryid " & _
                  " left join personjoindate pj on pj.personjoindateid = pjc.personjoindateid" & _
                  " left join accexpenses ae on ae.accexpensesid = i.accexpensesid" & _
                  " left join expensesnature en on en.expensesnatureid = ae.expensesnatureid" & _
                  " left join person p on p.personid = pj.personid" & _
                  " where c.regionid = " & dbtools1.RegionId & " and i.verid = " & myverid & ";"
        sb.Append(sqlstr)



        'table 9 personexpensesdtl for new comer
        sqlstr = "select pjc.personjoindatecategoryid , ped.validdate,ped.amount,en.expensesnature,ped.personexpensesid,i.icdpjcid " & _
                  " from personjoindate pj" & _
                  " left join person p on p.personid = pj.personid" & _
                  " left join personjoindatecategory pjc on pjc.personjoindateid = pj.personjoindateid" & _
                  " left join category c on c.categoryid = pjc.categoryid" & _
                  " left join icdpjc i on i.personjoindatecategoryid = pjc.personjoindatecategoryid" & _
                  " left join personexpenses pe on pe.icdpjcid = i.icdpjcid" & _
                  " left join personexpensesdtl ped on ped.personexpensesid = pe.personexpensesid" & _
                  " left join expensesdetailtx edtx on edtx.expensesdetailtxid = pe.expensesdetailtxid" & _
                  " left join expensesdetail ed on ed.expensesdetailid = edtx.expensesdetailid" & _
                  " left join accexpenses ac on ac.accexpensesid = ed.accexpensesid" & _
                  " left join sapaccname san on san.sapaccnameid = ac.sapaccnameid" & _
                  " left join expensesnature en on en.expensesnatureid = ac.expensesnatureid" & _
                  " where not ped.validdate isnull and pj.joindate >= " & DateFormatyyyyMMdd(BeginingYear) & " and pj.joindate <= " & DateFormatyyyyMMdd(EndOFYear) & " and c.regionid = " & dbtools1.RegionId & " and i.verid = " & myverid & " ;"
        sb.Append(sqlstr)
        'sb.Append("select pjc.personjoindatecategoryid , ped.validdate,ped.amount,en.expensesnature,ped.personexpensesid,i.icdpjcid " & _
        '         " from personjoindate pj" & _
        '         " left join person p on p.personid = pj.personid" & _
        '         " left join personjoindatecategory pjc on pjc.personjoindateid = pj.personjoindateid" & _
        '         " left join icdpjc i on i.personjoindatecategoryid = pjc.personjoindatecategoryid" & _
        '         " left join personexpenses pe on pe.icdpjcid = i.icdpjcid" & _
        '         " left join personexpensesdtl ped on ped.personexpensesid = pe.personexpensesid" & _
        '         " left join expensesdetailtx edtx on edtx.expensesdetailtxid = pe.expensesdetailtxid" & _
        '         " left join expensesdetail ed on ed.expensesdetailid = edtx.expensesdetailid" & _
        '         " left join accexpenses ac on ac.accexpensesid = ed.accexpensesid" & _
        '         " left join sapaccname san on san.sapaccnameid = ac.sapaccnameid" & _
        '         " left join expensesnature en on en.expensesnatureid = ac.expensesnatureid" & _
        '         " where not ped.validdate isnull and pj.joindate >= ")
        'sb.Append(DateFormatyyyyMMdd(BeginingYear) & " and pj.joindate <= ")
        'sb.Append(DateFormatyyyyMMdd(EndOFYear) & " and p.regionid = " & dbtools1.RegionId & " ;")
        sqlstr = sb.ToString

        'Table 10 expenses nature
        sqlstr = "select expensesdetailtxid,sapaccname,expensesnatureid,expensesnature,sapaccount,sapaccid,sapcc,dept,currency,fullyear,indexcostcenterdeptid from fgetexpensesdetailtx(" & _
                  myyear & ") as mytable(expensesdetailtxid integer,sapaccname character varying,expensesnatureid integer,expensesnature character varying,sapaccount character varying,sapaccid character varying,sapcc character varying,dept character varying,currency character varying,fullyear boolean,indexcostcenterdeptid integer)" & _
                  " where expensesnature <> 'Basic Salary' and expensesnature <> 'Salary & Wage' and expensesnature <> 'Emp Comp. insur.' and expensesnature <> 'EMPLOYEE MPF';"
        sb.Append(sqlstr)
        'sb.Append("select expensesdetailtxid,sapaccname,expensesnatureid,expensesnature,sapaccount,sapaccid,sapcc,dept,currency,fullyear,indexcostcenterdeptid from fgetexpensesdetailtx(")
        'sb.Append(myyear)
        'sb.Append(") as mytable(expensesdetailtxid integer,sapaccname character varying,expensesnatureid integer,expensesnature character varying,sapaccount character varying,sapaccid character varying,sapcc character varying,dept character varying,currency character varying,fullyear boolean,indexcostcenterdeptid integer)")
        'sb.Append(" where expensesnature <> 'Basic Salary' and expensesnature <> 'Salary & Wage' and expensesnature <> 'Emp Comp. insur.' and expensesnature <> 'EMPLOYEE MPF';")
        'sqlstr = sb.ToString

        ''Table 10 expenses nature
        'sb.Append("select expensesdetailtxid,sapaccname,expensesnatureid,expensesnature,sapaccount,sapaccid,sapcc,dept,currency,fullyear,indexcostcenterdeptid from fgetexpensesdetailtx(")
        'sb.Append(myyear)
        'sb.Append(") as mytable(expensesdetailtxid integer,sapaccname character varying,expensesnatureid integer,expensesnature character varying,sapaccount character varying,sapaccid character varying,sapcc character varying,dept character varying,currency character varying,fullyear boolean,indexcostcenterdeptid integer)")
        'sb.Append(" where expensesnature <> 'Basic Salary'  and expensesnature <> 'Emp Comp. insur.' and expensesnature <> 'EMPLOYEE MPF';")
        'sqlstr = sb.ToString

        'Table 11 plan
        'sqlstr = "SELECT c.categoryid, c.category,  ptx.amount,ptx.validfrom ,ctx.myyear" & _
        '          " FROM categorydtl ctd" & _
        '          " LEFT JOIN category c ON c.categoryid = ctd.categoryid" & _
        '          " LEFT JOIN categorytype ct ON ct.categorytypeid = ctd.categorytypeid" & _
        '          " LEFT JOIN categorytx ctx ON ctx.categorydtlid = ctd.categorydtlid" & _
        '          " left join ( SELECT p.planid, p.planname, sum(px.nominal) as amount, px.validfrom" & _
        '          " FROM plantx px" & _
        '          " LEFT JOIN plan p ON p.planid = px.planid" & _
        '          " LEFT JOIN plantype pt ON pt.plantypeid = px.plantypeid" & _
        '          "    where px.validfrom <= " & DateFormatyyyyMMdd(EndOFYear) & _
        '          " group by p.planid, p.planname,px.validfrom )as ptx on ptx.planid = ctx.amount" & _
        '          " where categorytype = 'Local medical expenses' and ctx.myyear = " & myyear & " and ctx.verid = " & myverid & " ;"
        'Added Percentage
        sqlstr = "SELECT c.categoryid, c.category,ptx.planid, ptx.amount,ptx.validfrom ,ctx.myyear,inspct.pct" & _
                 " FROM categorydtl ctd" & _
                 " LEFT JOIN category c ON c.categoryid = ctd.categoryid" & _
                 " LEFT JOIN categorytype ct ON ct.categorytypeid = ctd.categorytypeid" & _
                 " LEFT JOIN categorytx ctx ON ctx.categorydtlid = ctd.categorydtlid" & _
                 " left join ( SELECT p.planid, p.planname, sum(px.nominal) as amount, px.validfrom" & _
                 " FROM plantx px" & _
                 " LEFT JOIN plan p ON p.planid = px.planid" & _
                 " LEFT JOIN plantype pt ON pt.plantypeid = px.plantypeid" & _
                 "    where px.nominal > 1 and px.validfrom <= " & DateFormatyyyyMMdd(EndOFYear) & _
                 " group by p.planid, p.planname,px.validfrom )as ptx on ptx.planid = ctx.amount" & _
                 " left join  ( SELECT p.planid, p.planname, sum(px.nominal) as pct, px.validfrom" & _
                 " FROM plantx px" & _
                 " LEFT JOIN plan p ON p.planid = px.planid" & _
                 " LEFT JOIN plantype pt ON pt.plantypeid = px.plantypeid" & _
                 "    where px.nominal < 1 and px.plantypeid = 2 and px.validfrom <= " & DateFormatyyyyMMdd(EndOFYear) & _
                 " group by p.planid, p.planname,px.validfrom )as inspct on inspct.planid = ctx.amount" & _
                 " where categorytype = 'Local medical expenses' and ctx.myyear = " & myyear & " and ctx.verid = " & myverid & " ;"
        sb.Append(sqlstr)
        'Table 12 family member plan
        'sb.Append("select personjoindateid,planname,f.planid,count,amount,validfrom  from" & _
        '          " familymemberplan f" & _
        '          " left join (SELECT p.planid, p.planname, sum(px.nominal) as amount, px.validfrom " & _
        '       " FROM plantx px " & _
        '       " LEFT JOIN plan p ON p.planid = px.planid" & _
        '       " LEFT JOIN plantype pt ON pt.plantypeid = px.plantypeid" & _
        '       " where px.validfrom <= " & DateFormatyyyyMMdd(EndOFYear) & "group by p.planid, p.planname,px.validfrom )as ptx on ptx.planid = f.planid;")
        sb.Append("select personjoindateid,planname,f.planid,count,amount,validfrom  from" & _
                  " familymemberplan f" & _
                  " left join (SELECT p.planid, p.planname, sum(px.nominal) as amount, px.validfrom " & _
               " FROM plantx px " & _
               " LEFT JOIN plan p ON p.planid = px.planid" & _
               " LEFT JOIN plantype pt ON pt.plantypeid = px.plantypeid" & _
               " where px.validfrom <= " & DateFormatyyyyMMdd(EndOFYear) & "group by p.planid, p.planname,px.validfrom )as ptx on ptx.planid = f.planid where verid = " & myverid & " and myyear = " & myyear & ";")
        sqlstr = sb.ToString
        'Table 13 expenses nature
        'sb.Append("select expensesdetailtxid,sapaccname,expensesnatureid,expensesnature,sapaccount,sapaccid,sapcc,dept,currency,fullyear,indexcostcenterdeptid,icdpjcid from fgetexpensesdetailtx(")
        sb.Append("select * from fgetexpensesdetailtx(")
        sb.Append(myyear)
        sb.Append(") as mytable(expensesdetailtxid integer,sapaccname character varying,expensesnatureid integer,expensesnature character varying,sapaccount character varying,sapaccid character varying,sapcc character varying,dept character varying,currency character varying,fullyear boolean,indexcostcenterdeptid integer)")
        sb.Append(" where  expensesnature <> 'Emp Comp. insur.' or expensesnature <> 'EMPLOYEE MPF';")

        sqlstr = sb.ToString
        'Table 14 CategoryTxMonths
        sb.Append("select cm.*,ct.categorytypeid,ct.categorytype,c.categoryid,c.category  from categorytxmonths cm" & _
                  " left join categorytx ctx on ctx.categorytxid = cm.categorytxid" & _
                  " left join categorydtl cd on cd.categorydtlid = ctx.categorydtlid" & _
                  " left join category c on c.categoryid = cd.categoryid" & _
                  " left join categorytype ct on ct.categorytypeid = cd.categorytypeid" & _
                  " where cm.myyear = " & myyear & " and cm.verid = " & myverid & ";")
        sqlstr = sb.ToString

        'Table 15
        sb.Append("select expensesdetailtxid,sapaccname,expensesnatureid,expensesnature,sapaccount,sapaccid,sapcc,dept,currency,fullyear,indexcostcenterdeptid from fgetexpensesdetailtx(")
        sb.Append(myyear)
        sb.Append(") as mytable(expensesdetailtxid integer,sapaccname character varying,expensesnatureid integer,expensesnature character varying,sapaccount character varying,sapaccid character varying,sapcc character varying,dept character varying,currency character varying,fullyear boolean,indexcostcenterdeptid integer)")
        sb.Append(" where expensesnature = 'Salary & Wage' ;")
        sqlstr = sb.ToString

        'Table 16
        sqlstr = "select ptx.*,en.expensesnature from persontxmonth ptx left join expensesnature en on en.expensesnatureid = ptx.expensesnatureid where myyear = " & myyear & " and ptx.verid = " & myverid & ";"
        sb.Append(sqlstr)
        'Table 17 Plan
        sqlstr = "select p.planid,p.planname,sum(px.nominal) as amount,px.validfrom from plan p left join plantx px on px.planid = p.planid group by p.planid,p.planname,px.validfrom order by planid,validfrom asc;"
        sb.Append(sqlstr)
    End Sub

    Private Sub AddPrimaryKey(ByRef Dataset1 As DataSet)
        Dataset1.Tables(0).TableName = "PersonExpensesNature1"
        Dim Key0(3) As DataColumn
        Key0(0) = Dataset1.Tables(0).Columns("sapccname")
        Key0(1) = Dataset1.Tables(0).Columns("expensesnature")
        Key0(2) = Dataset1.Tables(0).Columns("sapaccount")
        Key0(3) = Dataset1.Tables(0).Columns("dept")
        Dataset1.Tables(0).PrimaryKey = Key0

        Dataset1.Tables(1).TableName = "tbexpensesmonth"
        Dataset1.Tables(2).TableName = "tbperson"

        Dataset1.Tables(3).TableName = "tbcategory"
        Dim key3(2) As DataColumn
        key3(0) = Dataset1.Tables(3).Columns("category")
        key3(1) = Dataset1.Tables(3).Columns("categorytype")
        key3(2) = Dataset1.Tables(3).Columns("myyear")
        Dataset1.Tables(3).PrimaryKey = key3

        Dataset1.Tables(4).TableName = "tbparamdt"
        Dim keyTbparamdt(1) As DataColumn

        keyTbparamdt(0) = Dataset1.Tables(4).Columns("paramname")
        keyTbparamdt(1) = Dataset1.Tables(4).Columns("dvalue")
        Dataset1.Tables(4).PrimaryKey = keyTbparamdt

        Dataset1.Tables(5).TableName = "tbpersonexpensesdtl"
        Dim KeyTbPersonExpensesDtl(3) As DataColumn
        KeyTbPersonExpensesDtl(0) = Dataset1.Tables(5).Columns("personjoindatecategoryid")
        KeyTbPersonExpensesDtl(1) = Dataset1.Tables(5).Columns("expensesnature")
        KeyTbPersonExpensesDtl(2) = Dataset1.Tables(5).Columns("validdate")
        KeyTbPersonExpensesDtl(3) = Dataset1.Tables(5).Columns("indexcostcenterdeptid")
        Dataset1.Tables(5).PrimaryKey = KeyTbPersonExpensesDtl


        Dataset1.Tables(6).TableName = "tbcatchupsalary"
        Dim key6(2) As DataColumn
        key6(0) = Dataset1.Tables(6).Columns("personjoindatecategoryid")
        key6(1) = Dataset1.Tables(6).Columns("txtype")
        key6(2) = Dataset1.Tables(6).Columns("validfrom")
        Dataset1.Tables(6).PrimaryKey = key6

        Dataset1.Tables(7).TableName = "tbVer"
        verid = Dataset1.Tables(7).Rows(0).Item("verid")

        Dataset1.Tables(8).TableName = "tbpersonexpenses"
        Dim key8(1) As DataColumn
        key8(0) = Dataset1.Tables(8).Columns("personjoindatecategoryid")
        key8(1) = Dataset1.Tables(8).Columns("expensesdetailtxid")
        Dataset1.Tables(8).PrimaryKey = key8

        Dataset1.Tables(9).TableName = "tbpersonexpensesdtlnewcomer"
        Dim Key9(2) As DataColumn
        Key9(0) = Dataset1.Tables(9).Columns("personjoindatecategoryid")
        Key9(1) = Dataset1.Tables(9).Columns("expensesnature")
        Key9(2) = Dataset1.Tables(9).Columns("validdate")
        Dataset1.Tables(9).PrimaryKey = Key9

        Dataset1.Tables(10).TableName = "PersonExpenses"
        Dim Key10(3) As DataColumn
        Key10(0) = Dataset1.Tables(10).Columns("sapccname")
        Key10(1) = Dataset1.Tables(10).Columns("expensesnature")
        Key10(2) = Dataset1.Tables(10).Columns("sapaccount")
        Key10(3) = Dataset1.Tables(10).Columns("dept")
        Dataset1.Tables(10).PrimaryKey = Key10

        Dataset1.Tables(11).TableName = "CategoryPlan"
        Dim Key11(1) As DataColumn
        Key11(0) = Dataset1.Tables(11).Columns("category")
        Key11(1) = Dataset1.Tables(11).Columns("validfrom")
        Dataset1.Tables(11).PrimaryKey = Key11


        Dataset1.Tables(12).TableName = "FamilyMemberPlan"
        Dim Key12(1) As DataColumn
        Key12(0) = Dataset1.Tables(12).Columns("personjoindatecategoryid")
        Key12(1) = Dataset1.Tables(12).Columns("validfrom")
        Dataset1.Tables(12).PrimaryKey = Key12

        Dataset1.Tables(14).TableName = "Categorytxmonths"
        Dim Key14(2) As DataColumn
        Key14(0) = Dataset1.Tables(14).Columns("category")
        Key14(1) = Dataset1.Tables(14).Columns("categorytype")
        Key14(2) = Dataset1.Tables(14).Columns("months")
        Dataset1.Tables(14).PrimaryKey = Key14
    End Sub

    Private Function Calculate(ByRef message) As Boolean
        Dim myReturn As Boolean = False
        Dim Dataset1 As New DataSet
        Dim sqlstr As String = String.Empty
        Dim sb As New StringBuilder

        'get expensesdetailtx
        myyear = DateTimePicker1.Value.Year
        BeginingYear = CDate(DateTimePicker1.Value.Year & "-1-1")
        EndOFYear = CDate(DateTimePicker1.Value.Year & "-12-31")
        Call buildQuery(sb)

        sqlstr = sb.ToString
        Try

            If dbtools1.getDataSet(sqlstr, Dataset1, message) Then
                Call AddPrimaryKey(Dataset1)
                'GetParamVar(Dataset1, GeneralRate, GeneralIncrMonth, ExpatRate, ExpatIncrMonth, serviceyear10, serviceyear15, AmountA, AmountB, AmountC, MPFValue)
                GetParamVar(Dataset1, GeneralRate, GeneralIncrMonth, ExpatRate, ExpatIncrMonth, serviceyear10, serviceyear15, AmountA, AmountB, AmountC, MPFValue, MPFFloorValue)
                Dim mymessage As String = String.Empty

                Dim PersonSalaryDict As New Dictionary(Of Integer, persondata)
                Dim PersonWageDict As New Dictionary(Of Integer, persondata)
                'calculate Salary Base Salary

                For i = 0 To Dataset1.Tables(0).Rows.Count - 1

                    Dim dr As DataRow = Dataset1.Tables(0).Rows(i)
                    'expensesdetailtxid,sapaccname,expensesnatureid,expensesnature,sapaccount,sapaccid,sapcc,dept,currency,fullyear,indexcostcenterdeptid
                    Debug.WriteLine(String.Format("i:{0} sapccname:{1} expensesnature:{2} sapccid:{3} sapcc:{4}", i, dr.Item(1).ToString, dr.Item(3).ToString, dr.Item(5).ToString, dr.Item(6).ToString))
                    CalculateSalary(dr, PersonSalaryDict, Dataset1, PersonWageDict)
                Next

                For i = 0 To Dataset1.Tables(15).Rows.Count - 1

                    Dim dr As DataRow = Dataset1.Tables(15).Rows(i)
                    Debug.WriteLine(String.Format("{0} {1} {2} {3} {4}", i, dr.Item(1).ToString, dr.Item(3).ToString, dr.Item(5).ToString, dr.Item(6).ToString))
                    CalculateSalary(dr, PersonSalaryDict, Dataset1, PersonWageDict)
                Next

                calculateExpenses2(Dataset1, PersonSalaryDict)

                'Calculate Emp Comp
                calculateExpenses3(Dataset1, PersonSalaryDict)

                'copy expensesdetail
                BackgroundWorker1.ReportProgress(2, "Copy To Db (BudgetTx)")
                myReturn = True
                sqlstr = "delete from budgettx where extract('Year' from mydate)  =" & myyear & " and ver = " & myverid & " and personexpensesid in (select personexpensesid from personexpenses pe " & _
                         " left join icdpjc i on i.icdpjcid = pe.icdpjcid" & _
                         " left join personjoindatecategory pjc on pjc.personjoindatecategoryid = i.personjoindatecategoryid" & _
                         " left join category c on c.categoryid = pjc.categoryid" & _
                         " left join personjoindate pj on pj.personjoindateid= pjc.personjoindateid " & _
                         " left join person p on p.personid = pj.personid where c.regionid =  " & dbtools1.RegionId & " order by personexpensesid) ;copy budgettx(personexpensesid,amount,ver,mydate,headcount) from stdin;"

                If stringBuilder1.ToString <> "" Then

                    '****
                    'Dim errorFilename As String = "c:\junk\PersonalError.txt"
                    'Using sw As StreamWriter = File.CreateText(errorFilename)
                    '    sw.WriteLine(stringBuilder1.ToString)
                    '    sw.Close()
                    'End Using
                    'Process.Start(errorFilename)

                    '*****

                    message = dbtools1.copy(sqlstr, stringBuilder1.ToString, myReturn)
                    If message <> "" Then
                        myReturn = False
                    End If
                    BackgroundWorker1.ReportProgress(2, "Copy To Db (BudgetTx)")
                    BackgroundWorker1.ReportProgress(3, message)
                Else
                    BackgroundWorker1.ReportProgress(3, "Nothing to Copy.")
                    myReturn = True
                End If
                'BackgroundWorker1.ReportProgress(3, message)


            End If
        Catch ex As Exception
            myReturn = False
            message = ex.Message
        Finally

        End Try
        Return myReturn
    End Function

    Private Sub CalculateSalary(ByVal dr As DataRow, ByRef PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal dataset1 As DataSet, ByRef PersonWageDict As Dictionary(Of Integer, persondata))
        Dim expensesdetailtxid = dr.Item("expensesdetailtxid")
        Dim expensesnature As String = dr.Item("expensesnature").ToString
        Dim sapaccount As String = dr.Item("sapaccount").ToString
        Dim sapcc As String = dr.Item("sapcc").ToString
        Dim dept As String = dr.Item("dept").ToString
        Dim indexcostcenterdeptid = dr.Item("indexcostcenterdeptid").ToString

        Dim mysb As New StringBuilder

        '2 select person for each expenses from table person categoryview using linq
        Dim myquery = From persons In dataset1.Tables(2)
                       Where persons.Item("dept").ToString = dept
                      Select persons Order By persons.Item("personname")
        Dim i = 0

        For Each p In myquery
            Dim persondata1 As New persondata
            i += 1
            Debug.WriteLine(String.Format("{0} StaffName:{1} OtherName:{2} PersonJoindateCategoryId:{3}", i, p.Item(1).ToString, p.Item(2).ToString, p.Item("personjoindatecategoryid")))
            If p.Item(1) = "FU Kang" Then 'Or p.Item(1) = "03405769" Then
                Debug.Print("debug")
            End If
            Dim serviceyear As Double = 0.0
            Dim incr As Double = 0
            Dim lastmonth As Integer
            serviceyear = Math.Round(((CDate(DateTimePicker1.Value.Year & "-12-31") - CDate(p.Item(3).ToString)).Days / 365), 2)

            'get personexpensesid for budgetRecord
            Dim personjoindatecategoryid = p.Item(0)
            'Dim icdpjcid = dr.Item("icdpjcid")

            Dim personexpensesid As Integer = 0
            Dim headcount As Double = p.Item(9)
            Dim mycategory As String = p.Item(5)
            Dim joindate As Date = p.Item(3)

            If personjoindatecategoryid = 375 Then
                Debug.WriteLine("debug mode")
            End If

            'get icdpjcid 
            'find icdpjcid kalok ga ada jangan di create
            'Dim icdpjcid = DbAdapter1.geticdpjcid(personjoindatecategoryid, indexcostcenterdeptid)
            'personexpenses.Item("expensesdetailtxid") = expensesdetailtxid
            'Dim pq = From personexpenses In dataset1.Tables(8)
            'Where(personexpenses.Item("personjoindatecategoryid") = personjoindatecategoryid And personexpenses.Item("expensesnature") = "Basic Salary" Or personexpenses.Item("expensesnature") = "Salary & Wage")
            Dim pq = From personexpenses In dataset1.Tables(8)
                     Where personexpenses.Item("personjoindatecategoryid") = personjoindatecategoryid And personexpenses.Item("expensesdetailtxid") = expensesdetailtxid


            '*********************check relasi expenses dan person sarch pake icdpjc***********

            'Dim pq = From personexpenses In dataset1.Tables(8)
            '         Where personexpenses.Item("personjoindatecategoryid") = personjoindatecategoryid And personexpenses.Item("icdpjcid") = icdpjcid
            '         Select personexpenses
            For Each pe In pq
                personexpensesid = pe.Item("personexpensesid")
            Next



            If personexpensesid = 0 Then
                'Create personexpensesid
                'personexpensesid = DbAdapter1.getpersonexpensesid(icdpjcid, expensesdetailtxid)
            End If
            If personexpensesid > 0 Then
                Dim SalaryDict As New Dictionary(Of Integer, Double)
                If personjoindatecategoryid = 375 Then
                    Debug.WriteLine("debug mode")
                End If
                If dr.Item("expensesnature").ToString = "Basic Salary" Or dr.Item("expensesnature").ToString = "Salary & Wage" Then

                    Dim baseSalary As Double = 0
                    Dim validdate As Date
                    BackgroundWorker1.ReportProgress(2, String.Format("Processing ******{0}********", expensesnature))

                    Dim catchupValue As Double = 0
                    Dim catchupDate As Date
                    Dim CatchupValueDict As New Dictionary(Of Integer, Double)

                    Dim incrVal As Double = 0
                    'Dim incrDate As Date
                    Dim incrValDict As New Dictionary(Of Integer, Double)

                    For L = 1 To 12
                        CatchupValueDict.Add(L, 0)
                        incrValDict.Add(L, 0)
                        SalaryDict.Add(L, 0)
                        persondata1.salaryDict.Add(L, 0)

                    Next

                    Dim q = From catchup In dataset1.Tables(6)
                            Where catchup.Item("personjoindatecategoryid") = personjoindatecategoryid And catchup.Item("txtype") = "catch up"
                            Select catchup

                    For Each myresult In q
                        catchupValue = myresult.Item("amount")
                        catchupDate = myresult.Item("validfrom")
                        CatchupValueDict(catchupDate.Month) = catchupValue
                    Next
                    '

                    'getBaseSalary from personexpensesdtl (the first salary in year budget) this part for person who joined BEFORE year budget
                    Dim qry = From expenses In dataset1.Tables(5)
                              Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And (expenses.Item("expensesnature") = "Basic Salary" Or expenses.Item("expensesnature") = "Salary & Wage") And expenses.Item("validdate") < CDate(DateTimePicker1.Value.Year & "/1/1")
                              Select expenses Order By expenses.Item("validdate") Descending

                    For Each myresult In qry
                        'baseSalary = myresult.Item("amount")
                        baseSalary = myresult.Item("amount") * p.Item("headcount")
                        validdate = myresult.Item("validdate")
                        Exit For
                    Next

                    'getBaseSalary from personexpensesdtl (the first salary in year budget) this part for person who joined IN year budget
                    Dim qry2 = From expenses In dataset1.Tables(9)
                            Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And (expenses.Item("expensesnature") = "Basic Salary" Or expenses.Item("expensesnature") = "Salary & Wage") And expenses.Item("validdate") >= CDate(DateTimePicker1.Value.Year & "/01/01") And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                            Select expenses Order By expenses.Item("validdate") Ascending
                    For Each myresult2 In qry2
                        'baseSalary = myresult2.Item("amount")
                        baseSalary = myresult2.Item("amount") * p.Item("headcount")

                        validdate = myresult2.Item("validdate")
                        Exit For
                    Next

                    'If baseSalary = 0 Then
                    '    Debug.WriteLine("debugMode")
                    'End If
                    'If p.Item("personname") = "ZHOU Cheng Jin II" Then
                    '    Debug.WriteLine("debug mode")
                    '    'MessageBox.Show("ZHOU Cheng Jin is here.")
                    'End If

                    Dim enddate As Date? = Nothing
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                    End If
                    If Not IsDBNull(p.Item("effectivedateend")) Then
                        enddate = p.Item("effectivedateend")
                    End If

                    'Last month will be override below
                    If p.Item("enddate").ToString = "" Then                      
                        lastmonth = 12
                    Else
                        lastmonth = CDate(p.Item("enddate").ToString).Month                        
                    End If

                    Dim totalsalary As Double = 0

                    'If dbtools1.Region = "HK" Or dbtools1.Region = "TW" Or dbtools1.Region = "SZ" Then
                    If p.Item("expat") Then 'Expat Calculation
                        incr = IIf(serviceyear > 1, ExpatRate, 0)
                        incrMonth = ExpatIncrMonth
                    Else 'General Calculation
                        incr = IIf(serviceyear > 1, GeneralRate, 0)
                        'Shenzen SZD2 SZD3 no need increment
                        incrMonth = 0
                        If Not "SZD2,SZD3".Contains(mycategory) Then
                            incrMonth = GeneralIncrMonth
                        Else
                            Debug.Print("incrMonth debug.")
                        End If

                    End If

                    'End If

                    'If personjoindatecategoryid = 51 Then
                    '    Debug.WriteLine("debug mode")
                    'End If
                    'If p.Item(2) = "Charlene" Then
                    '    Debug.WriteLine("debug mode")
                    'End If

                    Dim mytempsalary = 0
                    Dim myorisalary = baseSalary
                    For K = 1 To 12
                        'For K = 1 To 12
                        Dim kDate As Date = CDate(myyear & "-" & K & "-1")

                        If K = incrMonth Then
                            If serviceyear > 1 Then
                                If Not "SZD2,SZD3".Contains(mycategory) Then
                                    baseSalary *= (incr + 1)
                                End If

                                DbAdapter1.insertpersonexpensesdtl(personexpensesid, baseSalary, kDate)
                            End If
                        End If
                        'Check Catchup
                        If CatchupValueDict(K) <> 0 Then
                            If K = incrMonth Then
                                baseSalary = myorisalary * (1 + CatchupValueDict(K) + incr)
                            Else
                                baseSalary *= (1 + CatchupValueDict(K))
                            End If

                            DbAdapter1.insertpersonexpensesdtl(personexpensesid, baseSalary, kDate)
                        End If
                        'create record budget

                        'Modified on 2013-08-08
                        'If validdate <= kDate And K <= lastmonth Then
                        'If Month(validdate) <= K And K <= lastmonth Then

                        'Check Lastmonth with effectivedateend, check effectivedatestart
                        If Not IsDBNull(p.Item("effectivedateend")) Then
                            lastmonth = Month(p.Item("effectivedateend"))
                        End If
                        If Not IsDBNull(p.Item("effectivedatestart")) Then
                            validdate = p.Item("effectivedatestart")
                        End If
                        If ValidJoinDate(validdate, K, myyear) And K <= lastmonth Then

                            'totalsalary += baseSalary
                            SalaryDict(K) = baseSalary
                            persondata1.salaryDict(K) = baseSalary
                            mytempsalary = baseSalary
                            'If baseSalary = 0 Then
                            '    Debug.WriteLine("debugMode")
                            'End If
                            'Dim mydate As String = "'" & DateTimePicker1.Value.Year & "-" & K & "-1'"

                            'Debug.WriteLine("Person Name {0} Dept {1} personexpensesid {2} ExpensesNature {3} Personjoindatecategory {4} myDate {5}", p.Item("personname"), p.Item("dept"), personexpensesid, dr.Item("expensesnature"), p.Item("personjoindatecategoryid"), mydate)
                            'Debug.WriteLine("Create Record {0}", baseSalary)
                        Else

                            mytempsalary = 0
                        End If
                        Dim mydate As String = DateFormatyyyyMMdd(kDate)

                        'add validation starting date and enddate
                        If ValidJoinDate(validdate, K, myyear) And K <= lastmonth Then
                            'check effectivedate
                            'createrecord(stringBuilder1, personexpensesid, mytempsalary, myverid, mydate, p.Item("headcount"), p.Item("enddate"))
                            createrecord(stringBuilder1, personexpensesid, mytempsalary, myverid, mydate, p.Item("headcount"), enddate)
                        End If

                        If p.Item("personname") = "FU Kang" Then
                            'MessageBox.Show("ZHOU Cheng Jin - Create Record Salary. Personexpensesid = " & personexpensesid & ", Month = " & K)
                        End If

                        'Added catchup in feb - Mar
                        myorisalary = mytempsalary
                    Next

                    '*******Debug.WriteLine("Person Name {0}  personexpensesid {1} ExpensesNature {2} Personjoindatecategory {3} sapaccname  {4} sapaccount {5} costcenter{6} indexcostcenterdeptid {7}", p.Item("personname"), personexpensesid, dr.Item("expensesnature"), p.Item("personjoindatecategoryid"), dr.Item("sapaccname"), dr.Item("sapaccount"), dr.Item("sapcc"), dr.Item("sapaccid"), dr.Item("indexcostcenterdeptid"))
                    'If dr.Item("expensesnature") = "Basic Salary" Then
                    'If personjoindatecategoryid = 372 Then
                    '    Debug.WriteLine("debug mode")
                    'End If
                    'MessageBox.Show(personjoindatecategoryid & " " & p.Item("personname"))
                    Try
                        If Not PersonSalaryDict.ContainsKey(personjoindatecategoryid) Then


                            PersonSalaryDict.Add(personjoindatecategoryid, persondata1)
                        End If
                    Catch ex As Exception
                        Debug.Print("Error")
                    End Try




                    'PersonSalaryDict.Add(icdpjcid, persondata1)
                    'ElseIf dr.Item("expensesnature") = "Salary & Wage" Then
                    '    PersonWageDict.Add(personjoindatecategoryid, persondata1)
                    'PersonSalaryDict.Add(icdpjcid, persondata1)
                    'End If
                End If

                End If
                'Debug.WriteLine("Person Name: {0}, ExpensesNature: {1}, Personjoindatecategory: {2} ,sapaccname: {3},sapaccount: {4}, costcenter: {5},  sapaccid:{6} , indexcostcenterdeptid {7}", p.Item("personname"), dr.Item("expensesnature"), p.Item("personjoindatecategoryid"), dr.Item("sapaccname"), dr.Item("sapaccount"), dr.Item("sapcc"), dr.Item("sapaccid"), dr.Item("indexcostcenterdeptid"))
        Next
    End Sub

    Private Sub calculateExpenses2(ByVal Dataset1 As DataSet, ByRef PersonSalaryDict As Dictionary(Of Integer, persondata))
        For i = 0 To Dataset1.Tables(10).Rows.Count - 1
            Dim dr As DataRow = Dataset1.Tables(10).Rows(i)
            Dim expensesdetailtxid = dr.Item("expensesdetailtxid")
            Dim expensesnature As String = dr.Item("expensesnature").ToString
            Dim expensesnatuerid As Integer = dr.Item("expensesnatureid")

            Dim sapaccount As String = dr.Item("sapaccount").ToString
            Dim sapcc As String = dr.Item("sapcc").ToString

            Dim dept As String = dr.Item("dept").ToString

            '2 select person for each expenses from table person categoryview using linq
            Dim myquery = From persons In Dataset1.Tables(2)
                          Where persons.Item("dept").ToString = dept
                          Select persons Order By persons.Item("personname")


            For Each p In myquery

                Dim persondata1 As New persondata

                BackgroundWorker1.ReportProgress(3, String.Format("{0} {1} {2} {3} ", dr.Item("sapaccname").ToString, dr.Item("expensesnature").ToString, dr.Item("sapaccid").ToString, dr.Item("dept").ToString))
                Dim serviceyear As Double = 0.0
                Dim incr As Double = 0
                'Dim lastmonth As Integer
                serviceyear = Math.Round(((CDate(DateTimePicker1.Value.Year & "-12-31") - CDate(p.Item(3).ToString)).Days / 365), 2)

                'get personexpensesid for budgetRecord
                Dim personjoindatecategoryid = p.Item("personjoindatecategoryid")

                Dim personexpensesid As Integer = 0
                Dim personjoindateid = p.Item("personjoindateid")
                Dim headcount As Double = p.Item("headcount")
                Dim mycategory As String = p.Item("category")
                Dim joindate As Date = p.Item("joindate")

                Dim pq = From personexpenses In Dataset1.Tables(8)
                         Where personexpenses.Item("personjoindatecategoryid") = personjoindatecategoryid And personexpenses.Item("expensesdetailtxid") = expensesdetailtxid
                         Select personexpenses
                For Each pe In pq
                    personexpensesid = pe.Item("personexpensesid")
                Next



                If personexpensesid = 0 Then
                    'personexpensesid = DbAdapter1.getpersonexpensesid(personjoindatecategoryid, expensesdetailtxid)

                End If
                If personexpensesid > 0 Then

                    'Calculate Expenses
                    BackgroundWorker1.ReportProgress(2, String.Format("Processing  *** {0} ***", dr.Item("expensesnature")))
                    Debug.WriteLine("Expenses nature: {0}", dr.Item("expensesnature"))

                    If dr.Item("expensesnature") = "Salary tax by ER" Or
                        dr.Item("expensesnature") = "Insurance - Housing" Or
                        dr.Item("expensesnature") = "Trip credit" Or
                         dr.Item("expensesnature") = "Bonus" Or
                        dr.Item("expensesnature") = "Travel Insur AIG (PUR)" Then
                        Debug.Print("Test")
                    End If

                    If dr.Item("expensesnature") = "13th Month" Then
                        Call DoublePay(stringBuilder1, Dataset1, personexpensesid, myverid, myyear, joindate, serviceyear, mycategory, personjoindatecategoryid, PersonSalaryDict, persondata1, p)
                    ElseIf dr.Item("expensesnature").ToString = "Group Bonus" Then 'Each Person
                        Call GroupBonus(stringBuilder1, Dataset1, personexpensesid, myverid, myyear, joindate, serviceyear, mycategory, personjoindatecategoryid, PersonSalaryDict, persondata1, p)
                    ElseIf dr.Item("expensesnature").ToString = "Local Bonus" Then
                        Call LocalBonus(stringBuilder1, Dataset1, personexpensesid, myverid, myyear, joindate, serviceyear, mycategory, personjoindatecategoryid, PersonSalaryDict, persondata1, p)
                    ElseIf dr.Item("expensesnature").ToString = "Other Bonus" Then
                        Call OtherBonus(stringBuilder1, Dataset1, personexpensesid, myverid, myyear, joindate, serviceyear, mycategory, personjoindatecategoryid, PersonSalaryDict, persondata1, p)
                    ElseIf dr.Item("expensesnature") = "Local medical expenses" Then
                        Call LocalMedicalExpenses(stringBuilder1, Dataset1, personexpensesid, myverid, myyear, joindate, serviceyear, mycategory, personjoindatecategoryid, PersonSalaryDict, persondata1, personjoindateid, p)
                    ElseIf dr.Item("expensesnature").ToString = "Red Pocket" Then
                        'Call RedPocket(stringBuilder1, Dataset1, dr,p,serviceyer,verid, personexpensesid, verid, myyear, joindate, serviceyear, mycategory, personjoindatecategoryid, PersonSalaryDict, persondata1, personjoindateid)
                        Call RedPocket(stringBuilder1, Dataset1, dr, p, personexpensesid, serviceyear, myverid, PersonSalaryDict, persondata1)
                    ElseIf dr.Item("expensesnature").ToString = "staff award" Then
                        Call StaffAward(stringBuilder1, Dataset1, dr, p, personexpensesid, serviceyear, myverid, PersonSalaryDict, persondata1, personjoindatecategoryid)
                    ElseIf dr.Item("expensesnature").ToString = "Training (Dept)" Then
                        Call TrainingDept(stringBuilder1, Dataset1, dr, p, personexpensesid, serviceyear, myverid, PersonSalaryDict, persondata1)
                    ElseIf dr.Item("expensesnature").ToString = "Recruitment Expenses" Then
                        Call RecruitmentExpenses(stringBuilder1, Dataset1, dr, p, personexpensesid, serviceyear, myverid, PersonSalaryDict, persondata1)
                    ElseIf Not IsDBNull(dr.Item("fullyear")) Then

                        Call OtherExpenses(stringBuilder1, Dataset1, dr, personexpensesid, myverid, myyear, joindate, serviceyear, mycategory, personjoindatecategoryid, PersonSalaryDict, persondata1, p)

                    End If
                End If

            Next
        Next
    End Sub

    Private Sub calculateExpenses3(ByVal Dataset1 As DataSet, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata))


        Dim myQuery = From myExpenses In Dataset1.Tables(13)
                      Where myExpenses.Item("expensesnature") = "Emp Comp. insur." Or myExpenses.Item("expensesnature") = "EMPLOYEE MPF"
                      Select myExpenses



        For Each dr In myQuery

            Dim dept = dr.Item("dept")
            Dim expensesdetailtxid = dr.Item("expensesdetailtxid")
            Dim expensesnature = dr.Item("expensesnature")

            BackgroundWorker1.ReportProgress(2, String.Format("Processing ******{0}********", expensesnature))

            If expensesnature = "EMPLOYEE MPF" Then
                Debug.WriteLine("Debug mode")
            End If

            Dim myperson = From persons In Dataset1.Tables(2)
              Where persons.Item("dept").ToString = dept
              Select persons Order By persons.Item("personname")

            Dim personexpensesid As Integer = 0
            For Each person In myperson
                If person.Item("personname") = "03425996" Then
                    Debug.WriteLine("Debug mode for new commer")
                End If

                Dim joindate As Date = person.Item("joindate")

                Dim personjoindatecategoryid = person.Item("personjoindatecategoryid")
                Dim mycategory = person.Item("category")
                Dim pq = From personexpenses In Dataset1.Tables(8)
                         Where personexpenses.Item("personjoindatecategoryid") = personjoindatecategoryid And personexpenses.Item("expensesdetailtxid") = expensesdetailtxid
                         Select personexpenses
                For Each pe In pq
                    personexpensesid = pe.Item("personexpensesid")
                Next
                If personexpensesid = 0 Then
                    'personexpensesid = DbAdapter1.getpersonexpensesid(personjoindatecategoryid, expensesdetailtxid)
                End If

                
                If personexpensesid > 0 Then


                    'if listed in category then apply it
                    Debug.WriteLine("Apply Calculate3 Personjoindatecategoryid: {0}, expensesnatureid : {1} - {2}, personexpensesid: {3} ", person.Item("personjoindatecategoryid"), dr.Item("expensesnatureid"), dr.Item("expensesnature"), personexpensesid)
                    If personexpensesid = 742270 Then
                        Debug.Print("debug")
                    End If
                    Dim qry3 = From category In Dataset1.Tables(3)
                                   Where category.Item("category") = mycategory And category.Item("categorytype") = expensesnature And category.Item("myyear") = myyear
                                   Select category
                    Dim totalsalary As Double = 0

                    '****** check join date



                    'Dim ValidMonth As Integer = getvalidmonth(joindate)
                    Dim validmonth As Integer
                    If person.Item("enddate").ToString <> "" Then
                        validmonth = getvalidmonthnewcommer(joindate, person.Item("enddate"))
                    Else
                        validmonth = getvalidmonthnewcommer(joindate)
                    End If

                    'Dim joindate As Date = p.Item("joindate")
                    Dim enddate As Date? = Nothing
                    If Not IsDBNull(person.Item("enddate")) Then
                        'If Not p.Item("enddate").ToString() <> "" Then
                        enddate = person.Item("enddate")
                    End If

                    If dbtools1.Region = "HK" Then
                        If Not IsDBNull(person.Item("effectivedatestart")) Then
                            joindate = person.Item("effectivedatestart")
                        End If
                        If Not IsDBNull(person.Item("enddate")) Then
                            enddate = person.Item("enddate")
                        End If
                        If Not IsDBNull(person.Item("effectivedateend")) Then
                            enddate = person.Item("effectivedateend")
                        End If
                    Else
                        If Not IsDBNull(person.Item("enddate")) Then
                            enddate = person.Item("enddate")
                        End If
                    End If

                    For Each myresult In qry3
                        If expensesnature = "Emp Comp. insur." Then
                            If person.Item("personname") = "03425996" Then
                                Debug.WriteLine("Debug mode for new commer")
                            End If
                            Dim myamount = myresult.Item("amount")
                            For i = 1 To 12
                                totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                                totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).commision(i)
                                totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).housing(i)
                            Next

                            totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).bonus
                            totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).doublepay
                            'Dim empcompinsur = totalsalary * 0.006
                            Dim empcompinsur = totalsalary * myamount
                            For i = 1 To 12
                                'If joindate <= CDate(myyear & "-" & i & "-1") Then
                                '    createrecord(stringBuilder1, personexpensesid, empcompinsur / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & i & "-1")))
                                'End If
                                Dim currentdate As Date = CDate(myyear & "-" & i & "-1")
                                Dim check As Boolean = True
                                'If person.Item("enddate").ToString <> "" Then
                                '    check = currentdate < person.Item("enddate")
                                'End If
                                If Not IsNothing(enddate) Then
                                    check = currentdate < enddate
                                End If
                                'If currentdate >= person.Item("joindate") And check Then
                                'If ValidJoinDate(person.Item("joindate"), i, myyear) And check Then
                                If ValidJoinDate(joindate, i, myyear) And check Then
                                    createrecord(stringBuilder1, personexpensesid, empcompinsur / validmonth, myverid, currentdate, person.Item("headcount"), person.Item("enddate"))
                                End If
                            Next
                        ElseIf expensesnature = "EMPLOYEE MPF" Then
                            If personjoindatecategoryid = 488 Then
                                Debug.Print("debug mpf")
                            End If
                            Dim myamount = myresult.Item("amount")
                            Dim mpfcategory = myresult.Item("mpfcategory")
                            Try

                                Dim serviceyear = Math.Round(((CDate(DateTimePicker1.Value.Year & "-12-31") - CDate(person.Item(3).ToString)).Days / 365), 2)
                                For i = 1 To 12
                                    totalsalary = 0
                                    totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                                    totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).commision(i)
                                    'totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).housing(i)
                                    'Dim mpf As Double = totalsalary * 0.05
                                    Dim mpf As Double = totalsalary * myamount
                                    'totalsalary = IIf(mpf > 1000, 1000, mpf)
                                    'Check MPFFloor
                                    If mpf >= (MPFFloorValue * person.Item("headcount")) Then
                                        'totalsalary = IIf(mpf > (MPFValue * person.Item("headcount")), (MPFValue * person.Item("headcount")), mpf)
                                        totalsalary = mpf
                                        Dim additional As Double = 0
                                        If Not IsDBNull(mpfcategory) Then
                                            Select Case mpfcategory
                                                Case "A"
                                                    If serviceyear > 5 Then
                                                        additional = 0.025 * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                                                    End If
                                                Case "B"
                                                    If serviceyear > 10 Then
                                                        additional = 0.05 * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                                                    ElseIf serviceyear > 5 Then
                                                        additional = 0.025 * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                                                    End If
                                                Case "C"
                                                    additional = 0.05 * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                                                Case "D"
                                                    additional = 0.1 * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                                                Case "E1"
                                                    additional = 0.1 * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                                                Case "E2"
                                                    additional = 0.15 * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                                                Case "E3"
                                                    additional = 0.2 * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                                            End Select
                                        Else
                                            Debug.Print("dbNull Category")
                                        End If

                                        totalsalary += additional


                                        'If joindate.AddDays(62) < CDate(myyear & "-" & i & "-1") Then
                                        'If joindate <= CDate(myyear & "-" & i & "-1") Then  Request By Tracy 2018-12-21
                                        createrecord(stringBuilder1, personexpensesid, totalsalary, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & i & "-1")), person.Item("headcount"), person.Item("enddate"))
                                        'Else
                                        'Debug.WriteLine("just checking {0} {1}", joindate, CDate(myyear & "-" & i & "-1"))
                                        'End If
                                    End If
                                Next
                            Catch ex As Exception
                                Debug.Print(ex.Message)
                            End Try


                        End If

                    Next

                End If
            Next
        Next
    End Sub


    Private Sub calculateExpenses3Ori(ByVal Dataset1 As DataSet, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata))


        Dim myQuery = From myExpenses In Dataset1.Tables(13)
                      Where myExpenses.Item("expensesnature") = "Emp Comp. insur." Or myExpenses.Item("expensesnature") = "EMPLOYEE MPF"
                      Select myExpenses



        For Each dr In myQuery

            Dim dept = dr.Item("dept")
            Dim expensesdetailtxid = dr.Item("expensesdetailtxid")
            Dim expensesnature = dr.Item("expensesnature")

            BackgroundWorker1.ReportProgress(2, String.Format("Processing ******{0}********", expensesnature))

            If expensesnature = "EMPLOYEE MPF" Then
                'Debug.WriteLine("Debug mode")
            End If

            Dim myperson = From persons In Dataset1.Tables(2)
              Where persons.Item("dept").ToString = dept
              Select persons Order By persons.Item("personname")

            Dim personexpensesid As Integer = 0
            For Each person In myperson
                'If person.Item("personname") = "New Staff 7" Then
                '    Debug.WriteLine("Debug mode for new commer")
                'End If

                Dim joindate As Date = person.Item("joindate")

                Dim personjoindatecategoryid = person.Item("personjoindatecategoryid")
                Dim mycategory = person.Item("category")
                Dim pq = From personexpenses In Dataset1.Tables(8)
                         Where personexpenses.Item("personjoindatecategoryid") = personjoindatecategoryid And personexpenses.Item("expensesdetailtxid") = expensesdetailtxid
                         Select personexpenses
                For Each pe In pq
                    personexpensesid = pe.Item("personexpensesid")
                Next
                If personexpensesid = 0 Then
                    'personexpensesid = DbAdapter1.getpersonexpensesid(personjoindatecategoryid, expensesdetailtxid)
                End If
                If personexpensesid > 0 Then


                    'if listed in category then apply it
                    Debug.WriteLine("Apply Calculate3 Personjoindatecategoryid: {0}, expensesnatureid : {1} - {2}, personexpensesid: {3} ", person.Item("personjoindatecategoryid"), dr.Item("expensesnatureid"), dr.Item("expensesnature"), personexpensesid)

                    Dim qry3 = From category In Dataset1.Tables(3)
                                   Where category.Item("category") = mycategory And category.Item("categorytype") = expensesnature And category.Item("myyear") = myyear
                                   Select category
                    Dim totalsalary As Double = 0

                    '****** check join date

                    'Dim ValidMonth As Integer = getvalidmonth(joindate)
                    Dim validmonth As Integer
                    If person.Item("enddate").ToString <> "" Then
                        validmonth = getvalidmonthnewcommer(joindate, person.Item("enddate"))
                    Else
                        validmonth = getvalidmonthnewcommer(joindate)
                    End If

                    For Each myresult In qry3
                        If expensesnature = "Emp Comp. insur." Then
                            Dim myamount = myresult.Item("amount")
                            For i = 1 To 12
                                totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                                totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).commision(i)
                                totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).housing(i)
                            Next

                            totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).bonus
                            totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).doublepay
                            'Dim empcompinsur = totalsalary * 0.006
                            Dim empcompinsur = totalsalary * myamount
                            For i = 1 To 12
                                'If joindate <= CDate(myyear & "-" & i & "-1") Then
                                '    createrecord(stringBuilder1, personexpensesid, empcompinsur / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & i & "-1")))
                                'End If
                                Dim currentdate As Date = CDate(myyear & "-" & i & "-1")
                                Dim check As Boolean = True
                                If person.Item("enddate").ToString <> "" Then
                                    check = currentdate < person.Item("enddate")
                                End If
                                'If currentdate >= person.Item("joindate") And check Then
                                If ValidJoinDate(person.Item("joindate"), i, myyear) And check Then
                                    createrecord(stringBuilder1, personexpensesid, empcompinsur / validmonth, myverid, currentdate, person.Item("headcount"), person.Item("enddate"))
                                End If
                            Next
                        ElseIf expensesnature = "EMPLOYEE MPF" Then
                            Dim myamount = myresult.Item("amount")
                            For i = 1 To 12
                                totalsalary = 0
                                totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                                totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).commision(i)
                                totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).housing(i)
                                'Dim mpf As Double = totalsalary * 0.05
                                Dim mpf As Double = totalsalary * myamount
                                'totalsalary = IIf(mpf > 1000, 1000, mpf)
                                totalsalary = IIf(mpf > (MPFValue * person.Item("headcount")), (MPFValue * person.Item("headcount")), mpf)
                                'If joindate.AddDays(62) < CDate(myyear & "-" & i & "-1") Then
                                If joindate < CDate(myyear & "-" & i & "-1") Then
                                    createrecord(stringBuilder1, personexpensesid, totalsalary, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & i & "-1")), person.Item("headcount"), person.Item("enddate"))
                                Else
                                    'Debug.WriteLine("just checking {0} {1}", joindate, CDate(myyear & "-" & i & "-1"))
                                End If

                            Next



                        End If

                    Next

                End If
            Next
        Next
    End Sub
    Private Sub DoublePay(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Integer, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)

        'For HK, double pay is based on Personal Data column N (No. of Months (Bonus))
        '13 -> has 13th Month
        'other than that doesn't have 13th Month

        'If dbtools1.Region = "HK" Then
        '    'Check bonusfactor. if 13 then continue with the calculation
        '    If Not IsDBNull(p.Item("bonusfactor")) Then
        '        If p.Item("bonusfactor") = 13 Then
        '            Dim validmonth As Integer = 12
        '            Dim enttitlement = IIf(serviceyear > 1, 1, serviceyear)
        '            Dim LastSalary As Decimal
        '            'Calculate 13th month 
        '            Dim enddate As Date? = Nothing
        '            If Not IsDBNull(p.Item("enddate")) Then
        '                enddate = p.Item("enddate")
        '            End If
        '            Dim bonusfactor As Integer = 13
        '            Dim StartMonthSalary As Integer = 1
        '            Dim LastMonthSalary As Integer = 12

        '            If Not IsDBNull(p.Item("effectivedatestart")) Then
        '                joindate = p.Item("effectivedatestart")
        '                StartMonthSalary = Month(joindate)
        '            End If
        '            If Not IsDBNull(p.Item("enddate")) Then
        '                enddate = p.Item("enddate")
        '                LastMonthSalary = 12 'Month(enddate)
        '            ElseIf Not IsDBNull(p.Item("effectivedateend")) Then
        '                enddate = p.Item("effectivedateend")
        '                LastMonthSalary = Month(enddate)
        '            End If
        '            bonusfactor = p.Item("bonusfactor")
        '            LastSalary = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(LastMonthSalary)
        '            validmonth = LastMonthSalary - StartMonthSalary + 1
        '            Dim doublepay As Decimal = (LastSalary * enttitlement) / 12
        '            'Dim doublepay As Decimal = (LastSalary * enttitlement) / validmonth
        '            For k = StartMonthSalary To LastMonthSalary
        '                createrecord(stringBuilder1, personexpensesid, doublepay, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & k & "-1")), p.Item("headcount"), p.Item("enddate"))
        '            Next

        '            'Create record
        '        End If
        '    End If
        'Else
        Dim qry3 = From category In Dataset1.Tables(3)
                              Where category.Item("category") = mycategory And category.Item("categorytype") = "13th Month" And category.Item("myyear") = myyear
                              Select category
        'Category listed

        For Each dt In qry3

            'Add Checking BonusFactor for HK
            If dbtools1.Region = "HK" Then
                If Not IsDBNull(p.Item("bonusfactor")) Then
                    If p.Item("bonusfactor") <> 13 Then
                        Exit For
                    End If
                End If

                Dim validmonth As Integer = 12
                Dim enttitlement = IIf(serviceyear > 1, 1, serviceyear)
                Dim LastSalary As Decimal
                'Calculate 13th month 
                Dim enddate As Date? = Nothing
                If Not IsDBNull(p.Item("enddate")) Then
                    enddate = p.Item("enddate")
                    Exit For
                End If
                ' Dim bonusfactor As Integer = 13
                Dim StartMonthSalary As Integer = 1
                Dim LastMonthSalary As Integer = 12

                If Not IsDBNull(p.Item("effectivedatestart")) Then
                    joindate = p.Item("effectivedatestart")
                    StartMonthSalary = Month(joindate)
                End If
                If Not IsDBNull(p.Item("enddate")) Then
                    enddate = p.Item("enddate")
                    LastMonthSalary = 12 'Month(enddate)
                ElseIf Not IsDBNull(p.Item("effectivedateend")) Then
                    enddate = p.Item("effectivedateend")
                    LastMonthSalary = Month(enddate)
                End If

                LastSalary = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(LastMonthSalary)
                validmonth = LastMonthSalary - StartMonthSalary + 1
                Dim doublepay As Decimal = (LastSalary * enttitlement) / 12
                'Dim doublepay As Decimal = (LastSalary * enttitlement) / validmonth
                For k = StartMonthSalary To LastMonthSalary
                    createrecord(stringBuilder1, personexpensesid, doublepay, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & k & "-1")), p.Item("headcount"), p.Item("enddate"))
                Next


            Else
                Dim enttitlement = IIf(serviceyear > 1, 1, serviceyear)
                Dim lastsalary As Double = 0

                Try
                    lastsalary = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12)
                Catch ex As Exception

                End Try

                'Probation
                If joindate > CDate(myyear & "-10-01") Then lastsalary = 0
                If lastsalary > 0 Then
                    'Dim doublepay = lastsalary * enttitlement * p.Item("headcount")
                    Dim doublepay = lastsalary * enttitlement '* p.Item("headcount")
                    PersonSalaryDict(personjoindatecategoryid).doublepay = doublepay
                    If dbtools1.Region = "TW" Or dbtools1.Region = "HK" Or dbtools1.Region = "PH" Then
                        'doublepay = doublepay / 12
                        doublepay = ValidateJoinDate(joindate, p.Item("enddate"), doublepay) 'monthly
                    End If


                    'get categorytxmonths
                    Dim mymonths = From myrecord In Dataset1.Tables(14)
                                   Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = "13th Month"
                                   Select myrecord

                    For Each record In mymonths
                        If ValidJoinDate(joindate, record.Item("months"), myyear) Then
                            createrecord(stringBuilder1, personexpensesid, doublepay, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                        End If

                    Next
                    'persondata1.doublepay = doublepay
                    'PersonSalaryDict(personjoindatecategoryid).doublepay = doublepay 'move doublepay to top
                    'Debug.WriteLine("PersonExpensesid {0} {1} ", personexpensesid, personjoindatecategoryid)
                End If
            End If

            

        Next
        'End If


    End Sub
    Private Sub DoublePay02(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Integer, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)

        'For HK, double pay is based on Personal Data column N (No. of Months (Bonus))
        '13 -> has 13th Month
        'other than that doesn't have 13th Month

        'If dbtools1.Region = "HK" Then
        '    'Check bonusfactor. if 13 then continue with the calculation
        '    If Not IsDBNull(p.Item("bonusfactor")) Then
        '        If p.Item("bonusfactor") = 13 Then
        '            Dim validmonth As Integer = 12
        '            Dim enttitlement = IIf(serviceyear > 1, 1, serviceyear)
        '            Dim LastSalary As Decimal
        '            'Calculate 13th month 
        '            Dim enddate As Date? = Nothing
        '            If Not IsDBNull(p.Item("enddate")) Then
        '                enddate = p.Item("enddate")
        '            End If
        '            Dim bonusfactor As Integer = 13
        '            Dim StartMonthSalary As Integer = 1
        '            Dim LastMonthSalary As Integer = 12

        '            If Not IsDBNull(p.Item("effectivedatestart")) Then
        '                joindate = p.Item("effectivedatestart")
        '                StartMonthSalary = Month(joindate)
        '            End If
        '            If Not IsDBNull(p.Item("enddate")) Then
        '                enddate = p.Item("enddate")
        '                LastMonthSalary = 12 'Month(enddate)
        '            ElseIf Not IsDBNull(p.Item("effectivedateend")) Then
        '                enddate = p.Item("effectivedateend")
        '                LastMonthSalary = Month(enddate)
        '            End If
        '            bonusfactor = p.Item("bonusfactor")
        '            LastSalary = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(LastMonthSalary)
        '            validmonth = LastMonthSalary - StartMonthSalary + 1
        '            Dim doublepay As Decimal = (LastSalary * enttitlement) / 12
        '            'Dim doublepay As Decimal = (LastSalary * enttitlement) / validmonth
        '            For k = StartMonthSalary To LastMonthSalary
        '                createrecord(stringBuilder1, personexpensesid, doublepay, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & k & "-1")), p.Item("headcount"), p.Item("enddate"))
        '            Next

        '            'Create record
        '        End If
        '    End If
        'Else
        Dim qry3 = From category In Dataset1.Tables(3)
                              Where category.Item("category") = mycategory And category.Item("categorytype") = "13th Month" And category.Item("myyear") = myyear
                              Select category
        'Category listed

        For Each dt In qry3

            'Add Checking BonusFactor for HK
            If dbtools1.Region = "HK" Then
                If Not IsDBNull(p.Item("bonusfactor")) Then
                    If p.Item("bonusfactor") <> 13 Then
                        Exit For
                    End If
                End If
            Else

            End If

            Dim enttitlement = IIf(serviceyear > 1, 1, serviceyear)
            Dim lastsalary As Double = 0

            Try
                lastsalary = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12)
            Catch ex As Exception

            End Try

            'Probation
            If joindate > CDate(myyear & "-10-01") Then lastsalary = 0
            If lastsalary > 0 Then
                'Dim doublepay = lastsalary * enttitlement * p.Item("headcount")
                Dim doublepay = lastsalary * enttitlement '* p.Item("headcount")
                PersonSalaryDict(personjoindatecategoryid).doublepay = doublepay
                If dbtools1.Region = "TW" Or dbtools1.Region = "HK" Or dbtools1.Region = "PH" Then
                    'doublepay = doublepay / 12
                    doublepay = ValidateJoinDate(joindate, p.Item("enddate"), doublepay) 'monthly
                End If


                'get categorytxmonths
                Dim mymonths = From myrecord In Dataset1.Tables(14)
                               Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = "13th Month"
                               Select myrecord

                For Each record In mymonths
                    If ValidJoinDate(joindate, record.Item("months"), myyear) Then
                        createrecord(stringBuilder1, personexpensesid, doublepay, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                    End If

                Next
                'persondata1.doublepay = doublepay
                'PersonSalaryDict(personjoindatecategoryid).doublepay = doublepay 'move doublepay to top
                'Debug.WriteLine("PersonExpensesid {0} {1} ", personexpensesid, personjoindatecategoryid)
            End If

        Next
        'End If


    End Sub
    Private Sub DoublePayOld(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Integer, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)

        'For HK, double pay is based on Personal Data column N (No. of Months (Bonus))
        '13 -> has 13th Month
        'other than that doesn't have 13th Month

        'If dbtools1.Region = "HK" Then
        '    'Check bonusfactor. if 13 then continue with the calculation
        '    If Not IsDBNull(p.Item("bonusfactor")) Then
        '        If p.Item("bonusfactor") = 13 Then
        '            Dim validmonth As Integer = 12
        '            Dim enttitlement = IIf(serviceyear > 1, 1, serviceyear)
        '            Dim LastSalary As Decimal
        '            'Calculate 13th month 
        '            Dim enddate As Date? = Nothing
        '            If Not IsDBNull(p.Item("enddate")) Then
        '                enddate = p.Item("enddate")
        '            End If
        '            Dim bonusfactor As Integer = 13
        '            Dim StartMonthSalary As Integer = 1
        '            Dim LastMonthSalary As Integer = 12

        '            If Not IsDBNull(p.Item("effectivedatestart")) Then
        '                joindate = p.Item("effectivedatestart")
        '                StartMonthSalary = Month(joindate)
        '            End If
        '            If Not IsDBNull(p.Item("enddate")) Then
        '                enddate = p.Item("enddate")
        '                LastMonthSalary = 12 'Month(enddate)
        '            ElseIf Not IsDBNull(p.Item("effectivedateend")) Then
        '                enddate = p.Item("effectivedateend")
        '                LastMonthSalary = Month(enddate)
        '            End If
        '            bonusfactor = p.Item("bonusfactor")
        '            LastSalary = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(LastMonthSalary)
        '            validmonth = LastMonthSalary - StartMonthSalary + 1
        '            Dim doublepay As Decimal = (LastSalary * enttitlement) / 12
        '            'Dim doublepay As Decimal = (LastSalary * enttitlement) / validmonth
        '            For k = StartMonthSalary To LastMonthSalary
        '                createrecord(stringBuilder1, personexpensesid, doublepay, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & k & "-1")), p.Item("headcount"), p.Item("enddate"))
        '            Next

        '            'Create record
        '        End If
        '    End If
        'Else
        Dim qry3 = From category In Dataset1.Tables(3)
                              Where category.Item("category") = mycategory And category.Item("categorytype") = "13th Month" And category.Item("myyear") = myyear
                              Select category
        'Category listed

        For Each dt In qry3

            'Add Checking BonusFactor for HK
            If dbtools1.Region = "HK" Then
                If Not IsDBNull(p.Item("bonusfactor")) Then
                    If p.Item("bonusfactor") <> 13 Then
                        Exit For
                    End If
                End If
            End If

            Dim enttitlement = IIf(serviceyear > 1, 1, serviceyear)
            Dim lastsalary As Double = 0

            Try
                lastsalary = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12)
            Catch ex As Exception

            End Try

            'Probation
            If joindate > CDate(myyear & "-10-01") Then lastsalary = 0
            If lastsalary > 0 Then
                'Dim doublepay = lastsalary * enttitlement * p.Item("headcount")
                Dim doublepay = lastsalary * enttitlement '* p.Item("headcount")
                PersonSalaryDict(personjoindatecategoryid).doublepay = doublepay
                If dbtools1.Region = "TW" Or dbtools1.Region = "HK" Or dbtools1.Region = "PH" Then
                    'doublepay = doublepay / 12
                    doublepay = ValidateJoinDate(joindate, p.Item("enddate"), doublepay) 'monthly
                End If


                'get categorytxmonths
                Dim mymonths = From myrecord In Dataset1.Tables(14)
                               Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = "13th Month"
                               Select myrecord

                For Each record In mymonths
                    If ValidJoinDate(joindate, record.Item("months"), myyear) Then
                        createrecord(stringBuilder1, personexpensesid, doublepay, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                    End If

                Next
                'persondata1.doublepay = doublepay
                'PersonSalaryDict(personjoindatecategoryid).doublepay = doublepay 'move doublepay to top
                'Debug.WriteLine("PersonExpensesid {0} {1} ", personexpensesid, personjoindatecategoryid)
            End If

        Next
        'End If


    End Sub
    Private Sub GroupBonus1(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)
        'Debug.WriteLine("apply: {2} For Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} ", personjoindatecategoryid, dr.Item("expensesnatureid"), dr.Item("expensesnature"), personexpensesid)

        'get the rate
        'find amount for personexpensesdetail
        Dim qry = From expenses In Dataset1.Tables(5)
                    Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = "Group Bonus" And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                    Select expenses Order By expenses.Item("validdate") Descending
        Dim myamount As Double = 0
        Dim validdate As Date
        For Each myresult In qry
            myamount = myresult.Item("amount")
            validdate = myresult.Item("validdate")
            Exit For
        Next
        If myamount > 0 Then
            'get grosssalary jan-Dec + 13th Month
            Try
                'Dim grosssalary As Double = 0
                Dim grosssalary As Double = getGrossSalary(PersonSalaryDict, personjoindatecategoryid)

                'For gross = 1 To 12
                '    grosssalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(gross)
                'Next
                'grosssalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).doublepay


                persondata1.bonus += myamount

                'Dim ValidMonth As Integer = getvalidmonth(joindate)
                Dim validmonth As Integer
                If p.Item("enddate").ToString <> "" Then
                    validmonth = getvalidmonth(joindate, p.Item("enddate"))
                Else
                    validmonth = getvalidmonth(joindate)
                End If



                For M = 1 To 12
                    'If joindate <= CDate(myyear & "-" & M & "-1") Then
                    '    createrecord(stringBuilder1, personexpensesid, (grosssalary * myamount / ValidMonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                    'End If
                    Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                    Dim check As Boolean = True
                    If p.Item("enddate").ToString <> "" Then
                        check = currentdate < p.Item("enddate")
                    End If
                    'If currentdate >= p.Item("joindate") And check Then
                    If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then
                        createrecord(stringBuilder1, personexpensesid, (grosssalary * myamount / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")), p.Item("headcount"), p.Item("enddate"))
                    End If
                Next
            Catch ex As Exception
                Debug.WriteLine("Group Bonus Error: personexpensesid {0} personjoindatecategoryid {1}", personexpensesid, personjoindatecategoryid)
            End Try

        End If
    End Sub
    'Private Sub GroupBonus2(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)
    '    'Debug.WriteLine("apply: {2} For Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} ", personjoindatecategoryid, dr.Item("expensesnatureid"), dr.Item("expensesnature"), personexpensesid)

    '    'get the rate
    '    'find amount for personexpensesdetail
    '    Dim qry = From expenses In Dataset1.Tables(5)
    '                Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = "Group Bonus" And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
    '                Select expenses Order By expenses.Item("validdate") Descending
    '    Dim myamount As Double = 0
    '    Dim validdate As Date
    '    For Each myresult In qry
    '        myamount = myresult.Item("amount")
    '        validdate = myresult.Item("validdate")
    '        Exit For
    '    Next
    '    If myamount > 0 Then
    '        'get grosssalary jan-Dec + 13th Month
    '        Try
    '            'Dim grosssalary As Double = 0
    '            Dim grosssalary As Double = getGrossSalary(PersonSalaryDict, personjoindatecategoryid)

    '            'For gross = 1 To 12
    '            '    grosssalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(gross)
    '            'Next
    '            'grosssalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).doublepay


    '            persondata1.bonus += myamount
    '            'Dim ValidMonth As Integer = getvalidmonth(joindate)
    '            Dim validmonth As Integer
    '            If p.Item("enddate").ToString <> "" Then
    '                validmonth = getvalidmonth(joindate, p.Item("enddate"))
    '            Else
    '                validmonth = getvalidmonth(joindate)
    '            End If

    '            Dim mymonths = From myrecord In Dataset1.Tables(16)
    '                           Where myrecord.Item("personjoindatecategoryid") = personjoindatecategoryid And myrecord.Item("expensesnature") = "Group Bonus"
    '                           Select myrecord


    '            For Each record In mymonths
    '                'If joindate <= CDate(myyear & "-" & M & "-1") Then
    '                '    createrecord(stringBuilder1, personexpensesid, (grosssalary * myamount / ValidMonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
    '                'End If
    '                Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
    '                Dim check As Boolean = True
    '                If p.Item("enddate").ToString <> "" Then
    '                    check = currentdate < p.Item("enddate")
    '                End If
    '                'If currentdate >= p.Item("joindate") And check Then
    '                If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
    '                    createrecord(stringBuilder1, personexpensesid, (grosssalary * myamount / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
    '                End If
    '            Next
    '        Catch ex As Exception
    '            Debug.WriteLine("Group Bonus Error: personexpensesid {0} personjoindatecategoryid {1}", personexpensesid, personjoindatecategoryid)
    '        End Try

    '    End If
    'End Sub
    'Private Sub GroupBonus3(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)
    '    'Debug.WriteLine("apply: {2} For Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} ", personjoindatecategoryid, dr.Item("expensesnatureid"), dr.Item("expensesnature"), personexpensesid)

    '    'get the rate
    '    'find amount for personexpensesdetail
    '    If personjoindatecategoryid = 2439 Then
    '        Debug.WriteLine("Mettew")
    '    End If
    '    Dim category As String = String.Empty
    '    Dim qry = From expenses In Dataset1.Tables(5)
    '                Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = "Group Bonus" And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
    '                Select expenses Order By expenses.Item("validdate") Descending
    '    Dim myamount As Double = 0
    '    Dim validdate As Date
    '    For Each myresult In qry
    '        myamount = myresult.Item("amount")
    '        validdate = myresult.Item("validdate")
    '        category = myresult.Item("Category")
    '        Exit For
    '    Next
    '    If myamount > 0 Then
    '        'get grosssalary jan-Dec + 13th Month
    '        Try
    '            'Dim grosssalary As Double = 0
    '            'Dim grosssalary As Double = getGrossSalary(PersonSalaryDict, personjoindatecategoryid)

    '            'Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 12 
    '            Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 12 * DbAdapter1.gettargetrate(personjoindatecategoryid, EndOFYear)
    '            'For gross = 1 To 12
    '            '    grosssalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(gross)
    '            'Next
    '            'grosssalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).doublepay
    '            If dbtools1.Region = "HK" Then
    '                'myamount = myamount * p.Item("headcount") 'no need because the salary already splited by headcount
    '            End If
    '            If p.Item("expat") Then
    '                'persondata1.bonus += myamount
    '                PersonSalaryDict(personjoindatecategoryid).bonus += myamount

    '            Else
    '                'persondata1.bonus += (grosssalary * myamount)
    '                PersonSalaryDict(personjoindatecategoryid).bonus += (grosssalary * myamount)
    '            End If

    '            'Dim ValidMonth As Integer = getvalidmonth(joindate)
    '            Dim validmonth As Integer
    '            If p.Item("enddate").ToString <> "" Then
    '                validmonth = getvalidmonth(joindate, p.Item("enddate"))
    '            Else
    '                validmonth = getvalidmonth(joindate)
    '            End If


    '            'For ShenZhen
    '            '
    '            'Shenzhen only 1 month  Not Anymore
    '            ' on 27 Oct 2011 Connie Ask to split into 12 months
    '            'If dbtools1.Region = "SZ" And category = "SZM1" Then
    '            '    validmonth = 1
    '            'End If

    '            Dim mymonths = From myrecord In Dataset1.Tables(16)
    '                           Where myrecord.Item("personjoindatecategoryid") = personjoindatecategoryid And myrecord.Item("expensesnature") = "Group Bonus"
    '                           Select myrecord


    '            For Each record In mymonths
    '                'If joindate <= CDate(myyear & "-" & M & "-1") Then
    '                '    createrecord(stringBuilder1, personexpensesid, (grosssalary * myamount / ValidMonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
    '                'End If
    '                Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
    '                Dim check As Boolean = True
    '                If p.Item("enddate").ToString <> "" Then
    '                    check = currentdate < p.Item("enddate")
    '                End If
    '                'If currentdate >= p.Item("joindate") And check Then
    '                Dim myValidDate As Date = p.Item("joindate")

    '                If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
    '                    If p.Item("expat") Then
    '                        createrecord(stringBuilder1, personexpensesid, (myamount / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
    '                    Else
    '                        createrecord(stringBuilder1, personexpensesid, (grosssalary * myamount / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
    '                    End If

    '                End If
    '            Next
    '        Catch ex As Exception
    '            Debug.WriteLine("Group Bonus Error: personexpensesid {0} personjoindatecategoryid {1}", personexpensesid, personjoindatecategoryid)
    '        End Try

    '    End If
    'End Sub

    Private Sub GroupBonus(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)
        'Debug.WriteLine("apply: {2} For Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} ", personjoindatecategoryid, dr.Item("expensesnatureid"), dr.Item("expensesnature"), personexpensesid)

        'get the rate
        'find amount for personexpensesdetail
        If personjoindatecategoryid = 2439 Then
            Debug.WriteLine("Mettew")
        End If
        Dim category As String = String.Empty
        Dim qry = From expenses In Dataset1.Tables(5)
                    Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = "Group Bonus" And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                    Select expenses Order By expenses.Item("validdate") Descending
        Dim myamount As Double = 0
        Dim validdate As Date
        For Each myresult In qry
            myamount = myresult.Item("amount")
            validdate = myresult.Item("validdate")
            category = myresult.Item("Category")
            Exit For
        Next
        If myamount > 0 Then
            'get grosssalary jan-Dec + 13th Month
            Try

                Dim validmonth As Double 'valid month always 12 
                If p.Item("enddate").ToString <> "" Then
                    validmonth = 12 'getvalidmonth(joindate, p.Item("enddate"))
                Else
                    validmonth = getvalidmonth(joindate)
                End If



                'Override For HK EffectiveDate, 

                Dim enddate As Date? = Nothing
                If Not IsDBNull(p.Item("enddate")) Then
                    enddate = p.Item("enddate")
                End If
                Dim bonusfactor As Integer = 12
                Dim LastMonthSalary As Integer = 12
                Dim StartMonthSalary As Integer = 1

                If dbtools1.Region = "HK" Then
                    If Not IsDBNull(p.Item("effectivedatestart")) Then
                        joindate = p.Item("effectivedatestart")
                        StartMonthSalary = Month(joindate)
                    End If
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                        LastMonthSalary = 12 'Month(enddate)
                    ElseIf Not IsDBNull(p.Item("effectivedateend")) Then
                        enddate = p.Item("effectivedateend")
                        LastMonthSalary = Month(enddate)
                    End If

                    'If p.Item("enddate").ToString <> "" Then
                    If Not IsNothing(enddate) Then
                        'validmonth = getvalidmonth(joindate, enddate) / 12
                    Else
                        'validmonth = getvalidmonthEF(joindate) / 12 'Get ValidMonth Effective Date
                    End If
                    bonusfactor = p.Item("bonusfactor")
                Else
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                    End If
                End If

                'validmonth = LastMonthSalary - StartMonthSalary + 1

                'Dim grosssalary As Double = 0
                'Dim grosssalary As Double = getGrossSalary(PersonSalaryDict, personjoindatecategoryid)

                'Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 12 

                'Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 12 * DbAdapter1.gettargetrate(personjoindatecategoryid, EndOFYear)
                Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(LastMonthSalary) * bonusfactor * DbAdapter1.gettargetrate(personjoindatecategoryid, EndOFYear)


                'For gross = 1 To 12
                '    grosssalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(gross)
                'Next
                'grosssalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).doublepay
                If dbtools1.Region = "HK" Then
                    'myamount = myamount * p.Item("headcount") 'no need because the salary already splited by headcount
                End If
                If p.Item("expat") Then
                    'persondata1.bonus += myamount
                    PersonSalaryDict(personjoindatecategoryid).bonus += myamount

                Else
                    'persondata1.bonus += (grosssalary * myamount)
                    PersonSalaryDict(personjoindatecategoryid).bonus += (grosssalary * myamount)
                End If


                'Dim validmonth As Integer
                'Dim enddate As Date? = Nothing
                'If dbtools1.Region = "HK" Then


                '    If Not IsDBNull(p.Item("effectivedatestart")) Then
                '        joindate = p.Item("effectivedatestart")
                '    End If
                '    If Not IsDBNull(p.Item("enddate")) Then
                '        enddate = p.Item("enddate")
                '    End If
                '    If Not IsDBNull(p.Item("effectivedateend")) Then
                '        enddate = p.Item("effectivedateend")
                '    End If

                '    'If p.Item("enddate").ToString <> "" Then
                '    If Not IsNothing(enddate) Then
                '        validmonth = getvalidmonth(joindate, enddate)
                '    Else
                '        validmonth = getvalidmonth(joindate)
                '    End If
                'Else

                '    If p.Item("enddate").ToString <> "" Then
                '        validmonth = getvalidmonth(joindate, p.Item("enddate"))
                '    Else
                '        validmonth = getvalidmonth(joindate)
                '    End If
                'End If
                'Dim ValidMonth As Integer = getvalidmonth(joindate)




                'For ShenZhen
                '
                'Shenzhen only 1 month  Not Anymore
                ' on 27 Oct 2011 Connie Ask to split into 12 months
                'If dbtools1.Region = "SZ" And category = "SZM1" Then
                '    validmonth = 1
                'End If

                Dim mymonths = From myrecord In Dataset1.Tables(16)
                               Where myrecord.Item("personjoindatecategoryid") = personjoindatecategoryid And myrecord.Item("expensesnature") = "Group Bonus"
                               Select myrecord


                For Each record In mymonths
                    'If joindate <= CDate(myyear & "-" & M & "-1") Then
                    '    createrecord(stringBuilder1, personexpensesid, (grosssalary * myamount / ValidMonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                    'End If
                    Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
                    Dim check As Boolean = True
                    'If p.Item("enddate").ToString <> "" Then
                    '    check = currentdate < p.Item("enddate")
                    'End If
                    If Not IsNothing(enddate) Then
                        check = currentdate < enddate
                    End If
                    'If currentdate >= p.Item("joindate") And check Then
                    'Dim myValidDate As Date = p.Item("joindate")

                    'If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
                    If ValidJoinDate(joindate, record.Item("months"), myyear) And check Then
                        If p.Item("expat") Then
                            'createrecord(stringBuilder1, personexpensesid, (myamount / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                            createrecord(stringBuilder1, personexpensesid, (myamount / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), enddate)
                        Else
                            'createrecord(stringBuilder1, personexpensesid, (grosssalary * myamount / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                            'createrecord(stringBuilder1, personexpensesid, (grosssalary * myamount / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), enddate)
                            createrecord(stringBuilder1, personexpensesid, (grosssalary * myamount / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), enddate)
                        End If

                    End If
                Next
            Catch ex As Exception
                Debug.WriteLine("Group Bonus Error: personexpensesid {0} personjoindatecategoryid {1}", personexpensesid, personjoindatecategoryid)
            End Try

        End If
    End Sub

    Private Sub LocalBonus(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)
        Dim expensesnature = "Local Bonus"
        Debug.WriteLine("apply: {2} For Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} ,person: {4} ", personjoindatecategoryid, 0, expensesnature, personexpensesid, p.Item("othername"))
        If personjoindatecategoryid = 860 Then
            Debug.WriteLine("Jasmine")
        End If
        Dim qry = From expenses In Dataset1.Tables(5)
                   Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = expensesnature And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                   Select expenses Order By expenses.Item("validdate") Descending

        Dim myamount As Double = 0
        Dim validdate As Date
        For Each myresult In qry
            myamount = myresult.Item("amount")
            validdate = myresult.Item("validdate")
            Exit For
        Next
        If myamount > 0 Then
            Try
                'If p.Item("enddate").ToString <> "" Then
                Dim validmonth As Integer
                If p.Item("enddate").ToString <> "" Then
                    validmonth = getvalidmonth(joindate, p.Item("enddate"))
                Else
                    validmonth = getvalidmonth(joindate)
                End If
                'Dim validmonth As Integer
                'If Not IsNothing(enddate) Then
                '    validmonth = getvalidmonth(joindate, enddate)
                'Else
                '    validmonth = getvalidmonth(joindate)
                'End If

                Dim enddate As Date? = Nothing
                If Not IsDBNull(p.Item("enddate")) Then
                    enddate = p.Item("enddate")
                End If
                Dim bonusfactor As Integer = 13
                Dim LastMonthSalary As Integer = 12
                If dbtools1.Region = "HK" Then
                    If Not IsDBNull(p.Item("effectivedatestart")) Then
                        joindate = p.Item("effectivedatestart")
                    End If
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                        LastMonthSalary = Month(enddate)
                    End If
                    If Not IsDBNull(p.Item("effectivedateend")) Then
                        enddate = p.Item("effectivedateend")
                        LastMonthSalary = Month(enddate)
                    End If

                    'If p.Item("enddate").ToString <> "" Then
                    If Not IsNothing(enddate) Then
                        validmonth = getvalidmonth(joindate, enddate)
                    Else
                        validmonth = getvalidmonthEF(joindate)  'Get ValidMonth Effective Date
                    End If
                    bonusfactor = p.Item("bonusfactor")                  
                End If

                'Debug.WriteLine("personexpensesid {1} amount {2} expensesnature {3} personjoindatecategoryid {4}", personexpensesid, myamount, expensesnature, personjoindatecategoryid)
                'Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 13 * myamount * DbAdapter1.gettargetrate(personjoindatecategoryid, EndOFYear)
                Dim grosssalary As Double
                grosssalary = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(LastMonthSalary) * bonusfactor * myamount * DbAdapter1.gettargetrate(personjoindatecategoryid, EndOFYear)
               
                'persondata1.bonus += grosssalary
                PersonSalaryDict(personjoindatecategoryid).bonus += grosssalary

                'Dim ValidMonth As Integer = getvalidmonth(joindate)
                'Dim validmonth As Integer
                'If p.Item("enddate").ToString <> "" Then
                '    validmonth = getvalidmonth(joindate, p.Item("enddate"))
                'Else
                '    validmonth = getvalidmonth(joindate)
                'End If

                'Dim enddate As Date? = Nothing
                'If Not IsDBNull(p.Item("effectivedatestart")) Then
                '    joindate = p.Item("effectivedatestart")
                'End If

                'If Not IsDBNull(p.Item("enddate")) Then
                '    enddate = p.Item("enddate")
                'End If
                'If Not IsDBNull(p.Item("effectivedateend")) Then
                '    enddate = p.Item("effectivedateend")
                'End If

                ''If p.Item("enddate").ToString <> "" Then
                'Dim validmonth As Integer
                'If Not IsNothing(enddate) Then
                '    validmonth = getvalidmonth(joindate, enddate)
                'Else
                '    validmonth = getvalidmonth(joindate)
                'End If
                'Shenzhen only 1 month 
                'On 27 Oct Connie ask to change into 12 months
                '
                'If dbtools1.Region = "SZ" And mycategory <> "SZM5" Then
                ' validmonth = 1
                ' End If

                Dim mymonths = From myrecord In Dataset1.Tables(16)
                            Where myrecord.Item("personjoindatecategoryid") = personjoindatecategoryid And myrecord.Item("expensesnature") = expensesnature
                            Select myrecord


                For Each record In mymonths
                    Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
                    Dim check As Boolean = True
                    If p.Item("enddate").ToString <> "" Then
                        check = currentdate < p.Item("enddate")
                    End If

                    'If currentdate >= p.Item("joindate") And check Then

                    'If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
                    If ValidJoinDate(joindate, record.Item("months"), myyear) And check Then
                        'createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                        createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), enddate)
                    End If
                Next
                'For M = 1 To 12
                '    'If joindate <= CDate(myyear & "-" & M & "-1") Then
                '    '    createrecord(stringBuilder1, personexpensesid, grosssalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                '    'End If
                '    Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                '    Dim check As Boolean = True
                '    If p.Item("enddate").ToString <> "" Then
                '        check = currentdate < p.Item("enddate")
                '    End If
                '    If currentdate > p.Item("joindate") And check Then
                '        createrecord(stringBuilder1, personexpensesid, grosssalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                '    End If
                'Next
            Catch ex As Exception
                Debug.WriteLine("Local Bonus Error: personexpensesid {0} personjoindatecategoryid {1}", personexpensesid, personjoindatecategoryid)
            End Try
        ElseIf myamount <= 0 Then
            'Get based on Category
            Dim qry3 = From category In Dataset1.Tables(3)
                                 Where category.Item("category") = mycategory And category.Item("categorytype") = "Local Bonus" And category.Item("myyear") = myyear
                                 Select category
            'Category listed

            For Each dt In qry3
                myamount = dt.Item("amount")
                'Dim ValidMonth As Integer = getvalidmonth(joindate)

                Dim validmonth As Integer
                If p.Item("enddate").ToString <> "" Then
                    validmonth = getvalidmonth(joindate, p.Item("enddate"))
                Else
                    validmonth = getvalidmonth(joindate)
                End If

                Dim bonusfactor As Integer = 13
                Dim LastMonthSalary As Integer = 12
                Dim enddate As Date? = Nothing
                If Not IsDBNull(p.Item("effectivedatestart")) Then
                    joindate = p.Item("effectivedatestart")
                End If

                If Not IsDBNull(p.Item("enddate")) Then
                    enddate = p.Item("enddate")
                    LastMonthSalary = Month(enddate)
                End If
                If Not IsDBNull(p.Item("effectivedateend")) Then
                    enddate = p.Item("effectivedateend")
                    LastMonthSalary = Month(enddate)
                End If

                'If p.Item("enddate").ToString <> "" Then
                'Dim validmonth As Integer
                If Not IsNothing(enddate) Then
                    'validmonth = getvalidmonth(joindate, enddate) '/ 12

                Else
                    'validmonth = getvalidmonth(joindate) '/ 12
                End If
                'Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 13 * myamount * DbAdapter1.gettargetrate(personjoindatecategoryid, EndOFYear)
                Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(LastMonthSalary) * bonusfactor * myamount * DbAdapter1.gettargetrate(personjoindatecategoryid, EndOFYear)
                'persondata1.bonus += (grosssalary * p.Item("headcount"))
                'PersonSalaryDict(personjoindatecategoryid).bonus += (grosssalary * p.Item("headcount"))
                PersonSalaryDict(personjoindatecategoryid).bonus += (grosssalary)

                

                'Shenzhen only 1 month
                'If dbtools1.Region = "SZ" And mycategory <> "SZM5" Then
                'validmonth = 1
                'End If
                Dim mymonths = From myrecord In Dataset1.Tables(14)
                               Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = expensesnature
                               Select myrecord



                For Each record In mymonths
                    If joindate <= CDate(myyear & "-" & record.Item("months") & "-1") Then
                        'createrecord(stringBuilder1, personexpensesid, grosssalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                        Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
                        Dim check As Boolean = True
                        'If p.Item("enddate").ToString <> "" Then
                        '    check = currentdate < p.Item("enddate")
                        'End If
                        If Not IsNothing(enddate) Then
                            check = currentdate < enddate
                        End If
                        'If currentdate >= p.Item("joindate") And check Then
                        'If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
                        If ValidJoinDate(joindate, record.Item("months"), myyear) And check Then
                            'createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth) * p.Item("headcount"), myverid, currentdate)
                            'createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth), myverid, currentdate, p.Item("headcount"), p.Item("enddate"))
                            createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth), myverid, currentdate, p.Item("headcount"), enddate)
                        End If
                    End If
                Next
                'For M = 1 To 12
                '    If joindate <= CDate(myyear & "-" & M & "-1") Then
                '        'createrecord(stringBuilder1, personexpensesid, grosssalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                '        Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                '        Dim check As Boolean = True
                '        If p.Item("enddate").ToString <> "" Then
                '            check = currentdate < p.Item("enddate")
                '        End If
                '        If currentdate > p.Item("joindate") And check Then
                '            createrecord(stringBuilder1, personexpensesid, grosssalary / validmonth, myverid, currentdate)
                '        End If
                '    End If

                'Next
            Next
        End If

    End Sub
    'Private Sub LocalBonus1(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)
    '    Dim expensesnature = "Local Bonus"
    '    Debug.WriteLine("apply: {2} For Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} ", personjoindatecategoryid, 0, expensesnature, personexpensesid)

    '    Dim qry = From expenses In Dataset1.Tables(5)
    '               Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = expensesnature And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
    '               Select expenses Order By expenses.Item("validdate") Descending

    '    Dim myamount As Double = 0
    '    Dim validdate As Date
    '    For Each myresult In qry
    '        myamount = myresult.Item("amount")
    '        validdate = myresult.Item("validdate")
    '        Exit For
    '    Next
    '    If myamount > 0 Then
    '        Try
    '            'Debug.WriteLine("personexpensesid {1} amount {2} expensesnature {3} personjoindatecategoryid {4}", personexpensesid, myamount, expensesnature, personjoindatecategoryid)
    '            Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 13 * myamount
    '            persondata1.bonus += grosssalary

    '            'Dim ValidMonth As Integer = getvalidmonth(joindate)
    '            Dim validmonth As Integer
    '            If p.Item("enddate").ToString <> "" Then
    '                validmonth = getvalidmonth(joindate, p.Item("enddate"))
    '            Else
    '                validmonth = getvalidmonth(joindate)
    '            End If
    '            For M = 1 To 12
    '                'If joindate <= CDate(myyear & "-" & M & "-1") Then
    '                '    createrecord(stringBuilder1, personexpensesid, grosssalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
    '                'End If
    '                Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
    '                Dim check As Boolean = True
    '                If p.Item("enddate").ToString <> "" Then
    '                    check = currentdate < p.Item("enddate")
    '                End If
    '                'If currentdate >= p.Item("joindate") And check Then
    '                If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then
    '                    createrecord(stringBuilder1, personexpensesid, grosssalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")), p.Item("headcount"), p.Item("enddate"))
    '                End If
    '            Next
    '        Catch ex As Exception
    '            Debug.WriteLine("Local Bonus Error: personexpensesid {0} personjoindatecategoryid {1}", personexpensesid, personjoindatecategoryid)
    '        End Try
    '    ElseIf myamount <= 0 Then
    '        'Get based on Category
    '        Dim qry3 = From category In Dataset1.Tables(3)
    '                             Where category.Item("category") = mycategory And category.Item("categorytype") = "Local Bonus" And category.Item("myyear") = myyear
    '                             Select category
    '        'Category listed

    '        For Each dt In qry3
    '            myamount = dt.Item("amount")
    '            Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 13 * myamount
    '            persondata1.bonus += grosssalary

    '            'Dim ValidMonth As Integer = getvalidmonth(joindate)
    '            Dim validmonth As Integer
    '            If p.Item("enddate").ToString <> "" Then
    '                validmonth = getvalidmonth(joindate, p.Item("enddate"))
    '            Else
    '                validmonth = getvalidmonth(joindate)
    '            End If
    '            For M = 1 To 12
    '                If joindate <= CDate(myyear & "-" & M & "-1") Then
    '                    'createrecord(stringBuilder1, personexpensesid, grosssalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
    '                    Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
    '                    Dim check As Boolean = True
    '                    If p.Item("enddate").ToString <> "" Then
    '                        check = currentdate < p.Item("enddate")
    '                    End If
    '                    'If currentdate >= p.Item("joindate") And check Then
    '                    If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then
    '                        createrecord(stringBuilder1, personexpensesid, grosssalary / validmonth, myverid, currentdate, p.Item("headcount"), p.Item("enddate"))
    '                    End If
    '                End If

    '            Next
    '        Next
    '    End If

    'End Sub


    

    'Private Sub LocalBonus2(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)
    '    Dim expensesnature = "Local Bonus"
    '    Debug.WriteLine("apply: {2} For Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} ,person: {4} ", personjoindatecategoryid, 0, expensesnature, personexpensesid, p.Item("othername"))
    '    If personjoindatecategoryid = 2439 Then
    '        Debug.WriteLine("Mettew")
    '    End If
    '    Dim qry = From expenses In Dataset1.Tables(5)
    '               Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = expensesnature And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
    '               Select expenses Order By expenses.Item("validdate") Descending

    '    Dim myamount As Double = 0
    '    Dim validdate As Date
    '    For Each myresult In qry
    '        myamount = myresult.Item("amount")
    '        validdate = myresult.Item("validdate")
    '        Exit For
    '    Next
    '    If myamount > 0 Then
    '        Try
    '            'Debug.WriteLine("personexpensesid {1} amount {2} expensesnature {3} personjoindatecategoryid {4}", personexpensesid, myamount, expensesnature, personjoindatecategoryid)
    '            Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 13 * myamount * DbAdapter1.gettargetrate(personjoindatecategoryid, EndOFYear)
    '            'persondata1.bonus += grosssalary
    '            PersonSalaryDict(personjoindatecategoryid).bonus += grosssalary

    '            'Dim ValidMonth As Integer = getvalidmonth(joindate)
    '            Dim validmonth As Integer
    '            If p.Item("enddate").ToString <> "" Then
    '                validmonth = getvalidmonth(joindate, p.Item("enddate"))
    '            Else
    '                validmonth = getvalidmonth(joindate)
    '            End If
    '            'Shenzhen only 1 month 
    '            'On 27 Oct Connie ask to change into 12 months
    '            '
    '            'If dbtools1.Region = "SZ" And mycategory <> "SZM5" Then
    '            ' validmonth = 1
    '            ' End If

    '            Dim mymonths = From myrecord In Dataset1.Tables(16)
    '                        Where myrecord.Item("personjoindatecategoryid") = personjoindatecategoryid And myrecord.Item("expensesnature") = expensesnature
    '                        Select myrecord


    '            For Each record In mymonths
    '                Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
    '                Dim check As Boolean = True
    '                If p.Item("enddate").ToString <> "" Then
    '                    check = currentdate < p.Item("enddate")
    '                End If

    '                'If currentdate >= p.Item("joindate") And check Then
    '                If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
    '                    createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
    '                End If
    '            Next
    '            'For M = 1 To 12
    '            '    'If joindate <= CDate(myyear & "-" & M & "-1") Then
    '            '    '    createrecord(stringBuilder1, personexpensesid, grosssalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
    '            '    'End If
    '            '    Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
    '            '    Dim check As Boolean = True
    '            '    If p.Item("enddate").ToString <> "" Then
    '            '        check = currentdate < p.Item("enddate")
    '            '    End If
    '            '    If currentdate > p.Item("joindate") And check Then
    '            '        createrecord(stringBuilder1, personexpensesid, grosssalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
    '            '    End If
    '            'Next
    '        Catch ex As Exception
    '            Debug.WriteLine("Local Bonus Error: personexpensesid {0} personjoindatecategoryid {1}", personexpensesid, personjoindatecategoryid)
    '        End Try
    '    ElseIf myamount <= 0 Then
    '        'Get based on Category
    '        Dim qry3 = From category In Dataset1.Tables(3)
    '                             Where category.Item("category") = mycategory And category.Item("categorytype") = "Local Bonus" And category.Item("myyear") = myyear
    '                             Select category
    '        'Category listed

    '        For Each dt In qry3
    '            myamount = dt.Item("amount")
    '            Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 13 * myamount * DbAdapter1.gettargetrate(personjoindatecategoryid, EndOFYear)
    '            'persondata1.bonus += (grosssalary * p.Item("headcount"))
    '            'PersonSalaryDict(personjoindatecategoryid).bonus += (grosssalary * p.Item("headcount"))
    '            PersonSalaryDict(personjoindatecategoryid).bonus += (grosssalary)

    '            'Dim ValidMonth As Integer = getvalidmonth(joindate)
    '            Dim validmonth As Integer
    '            If p.Item("enddate").ToString <> "" Then
    '                validmonth = getvalidmonth(joindate, p.Item("enddate"))
    '            Else
    '                validmonth = getvalidmonth(joindate)
    '            End If

    '            'Shenzhen only 1 month
    '            'If dbtools1.Region = "SZ" And mycategory <> "SZM5" Then
    '            'validmonth = 1
    '            'End If
    '            Dim mymonths = From myrecord In Dataset1.Tables(14)
    '                           Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = expensesnature
    '                           Select myrecord



    '            For Each record In mymonths
    '                If joindate <= CDate(myyear & "-" & record.Item("months") & "-1") Then
    '                    'createrecord(stringBuilder1, personexpensesid, grosssalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
    '                    Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
    '                    Dim check As Boolean = True
    '                    If p.Item("enddate").ToString <> "" Then
    '                        check = currentdate < p.Item("enddate")
    '                    End If
    '                    'If currentdate >= p.Item("joindate") And check Then
    '                    If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
    '                        'createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth) * p.Item("headcount"), myverid, currentdate)
    '                        createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth), myverid, currentdate, p.Item("headcount"), p.Item("enddate"))
    '                    End If
    '                End If
    '            Next
    '            'For M = 1 To 12
    '            '    If joindate <= CDate(myyear & "-" & M & "-1") Then
    '            '        'createrecord(stringBuilder1, personexpensesid, grosssalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
    '            '        Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
    '            '        Dim check As Boolean = True
    '            '        If p.Item("enddate").ToString <> "" Then
    '            '            check = currentdate < p.Item("enddate")
    '            '        End If
    '            '        If currentdate > p.Item("joindate") And check Then
    '            '            createrecord(stringBuilder1, personexpensesid, grosssalary / validmonth, myverid, currentdate)
    '            '        End If
    '            '    End If

    '            'Next
    '        Next
    '    End If

    'End Sub

    Private Sub OtherBonus1(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)
        Dim expensesnature = "Other Bonus"
        Debug.WriteLine("apply: {2} For  Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} ", personjoindatecategoryid, 0, expensesnature, personexpensesid)

        Dim qry = From expenses In Dataset1.Tables(5)
                   Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = expensesnature And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                   Select expenses Order By expenses.Item("validdate") Descending

        Dim myamount As Double = 0
        Dim validdate As Date
        For Each myresult In qry
            myamount = myresult.Item("amount")
            validdate = myresult.Item("validdate")
            Exit For
        Next
        If myamount > 0 Then
            'Personal
            Try
                Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 12 * myamount
                persondata1.bonus += grosssalary
                'Dim ValidMonth As Integer = getvalidmonth(joindate)
                Dim validmonth As Integer
                If p.Item("enddate").ToString <> "" Then
                    validmonth = getvalidmonth(joindate, p.Item("enddate"))
                Else
                    validmonth = getvalidmonth(joindate)
                End If
                For M = 1 To 12
                    'If joindate <= CDate(myyear & "-" & M & "-1") Then
                    '    createrecord(stringBuilder1, personexpensesid, grosssalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                    'End If
                    Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                    Dim check As Boolean = True
                    If p.Item("enddate").ToString <> "" Then
                        check = currentdate < p.Item("enddate")
                    End If
                    'If currentdate >= p.Item("joindate") And check Then
                    If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then
                        createrecord(stringBuilder1, personexpensesid, grosssalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")), p.Item("headcount"), p.Item("enddate"))
                    End If
                Next
            Catch ex As Exception
                Debug.WriteLine("Other Bonus Error: personexpensesid {0} personjoindatecategoryid {1}", personexpensesid, personjoindatecategoryid)
            End Try
        ElseIf myamount <= 0 Then
            'Get based on Category
            Dim qry3 = From category In Dataset1.Tables(3)
                                 Where category.Item("category") = mycategory And category.Item("categorytype") = expensesnature And category.Item("myyear") = myyear
                                 Select category
            'Category listed

            For Each dt In qry3
                myamount = dt.Item("amount")
                Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 13 * myamount
                persondata1.bonus += grosssalary
                If personjoindatecategoryid = 28 Then
                    Debug.WriteLine("debug mode")
                End If
                'Dim ValidMonth As Integer = getvalidmonth(joindate)
                Dim validmonth As Integer
                If p.Item("enddate").ToString <> "" Then
                    validmonth = getvalidmonth(joindate, p.Item("enddate"))
                Else
                    validmonth = getvalidmonth(joindate)
                End If
                For M = 1 To 12
                    If joindate <= CDate(myyear & "-" & M & "-1") Then
                        createrecord(stringBuilder1, personexpensesid, grosssalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")), p.Item("headcount"), p.Item("enddate"))
                    End If

                Next
            Next
        End If
    End Sub
    Private Sub OtherBonus(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)
        Dim expensesnature = "Other Bonus"
        Debug.WriteLine("apply: {2} For  Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} ", personjoindatecategoryid, 0, expensesnature, personexpensesid)
        If personjoindatecategoryid = 541 Then
            Debug.Print("other bonus")
        End If

        Dim qry = From expenses In Dataset1.Tables(5)
                   Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = expensesnature And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                   Select expenses Order By expenses.Item("validdate") Descending

        Dim myamount As Double = 0
        Dim validdate As Date
        For Each myresult In qry
            myamount = myresult.Item("amount")
            validdate = myresult.Item("validdate")
            Exit For
        Next
        If myamount > 0 Then
            'Personal
            Try
                Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 13 * myamount * DbAdapter1.gettargetrate(personjoindatecategoryid, EndOFYear)
                'persondata1.bonus += grosssalary
                PersonSalaryDict(personjoindatecategoryid).bonus += grosssalary

                'Dim ValidMonth As Integer = getvalidmonth(joindate)
                Dim validmonth As Integer
                If p.Item("enddate").ToString <> "" Then
                    validmonth = getvalidmonth(joindate, p.Item("enddate"))
                Else
                    validmonth = getvalidmonth(joindate)
                End If

                'Shenzhen only 1 month Not Anymore
                'If dbtools1.Region = "SZ" Then
                '    validmonth = 1
                'End If

                Dim enddate As Date? = Nothing
                If dbtools1.Region = "HK" Then
                    If Not IsDBNull(p.Item("effectivedatestart")) Then
                        joindate = p.Item("effectivedatestart")
                    End If
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                    End If
                    If Not IsDBNull(p.Item("effectivedateend")) Then
                        enddate = p.Item("effectivedateend")
                    End If
                Else
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                    End If
                End If

                Dim mymonths = From myrecord In Dataset1.Tables(16)
                              Where myrecord.Item("personjoindatecategoryid") = personjoindatecategoryid And myrecord.Item("expensesnature") = expensesnature
                              Select myrecord


                For Each record In mymonths
                    Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
                    Dim check As Boolean = True
                    'If p.Item("enddate").ToString <> "" Then
                    '    check = currentdate < p.Item("enddate")
                    'End If
                    If Not IsNothing(enddate) Then
                        check = currentdate < enddate
                    End If
                    'If currentdate >= p.Item("joindate") And check Then
                    'If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
                    If ValidJoinDate(joindate, record.Item("months"), myyear) And check Then
                        createrecord(stringBuilder1, personexpensesid, grosssalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                    End If
                Next

            Catch ex As Exception
                Debug.WriteLine("Other Bonus Error: personexpensesid {0} personjoindatecategoryid {1}", personexpensesid, personjoindatecategoryid)
            End Try
        ElseIf myamount <= 0 Then
            'Get based on Category
            Dim qry3 = From category In Dataset1.Tables(3)
                                 Where category.Item("category") = mycategory And category.Item("categorytype") = expensesnature And category.Item("myyear") = myyear
                                 Select category
            'Category listed

            For Each dt In qry3
                myamount = dt.Item("amount")
                Dim grosssalary As Double = DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(12) * 13 * myamount * DbAdapter1.gettargetrate(personjoindatecategoryid, EndOFYear)
                'persondata1.bonus += grosssalary
                PersonSalaryDict(personjoindatecategoryid).bonus += grosssalary
                If personjoindatecategoryid = 28 Then
                    Debug.WriteLine("debug mode")
                End If
                'Dim ValidMonth As Integer = getvalidmonth(joindate)
                Dim validmonth As Integer
                If p.Item("enddate").ToString <> "" Then
                    validmonth = getvalidmonth(joindate, p.Item("enddate"))
                Else
                    validmonth = getvalidmonth(joindate)
                End If
                'Shenzhen only 1 month Not Anymore
                'If dbtools1.Region = "SZ" Then
                '    validmonth = 1
                'End If

                Dim enddate As Date? = Nothing
                If Not IsDBNull(p.Item("enddate")) Then
                    enddate = p.Item("enddate")
                End If
                If dbtools1.Region = "HK" Then
                    If Not IsDBNull(p.Item("effectivedatestart")) Then
                        joindate = p.Item("effectivedatestart")
                    End If
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                    End If
                    If Not IsDBNull(p.Item("effectivedateend")) Then
                        enddate = p.Item("effectivedateend")
                    End If
                Else
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                    End If
                End If
                Dim mymonths = From myrecord In Dataset1.Tables(14)
                               Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = expensesnature
                               Select myrecord


                For Each record In mymonths
                    Dim check As Boolean = True
                    Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
                    'If p.Item("enddate").ToString <> "" Then
                    '    check = currentdate < p.Item("enddate")
                    'End If

                    If Not IsNothing(enddate) Then
                        check = currentdate < enddate
                    End If
                    'If joindate <= CDate(myyear & "-" & record.Item("months") & "-1") And check Then
                    'If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then

                    If ValidJoinDate(joindate, record.Item("months"), myyear) And check Then
                        'createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth) * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")))
                        createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                    End If
                Next

                'For M = 1 To 12
                '    If joindate <= CDate(myyear & "-" & M & "-1") Then
                '        createrecord(stringBuilder1, personexpensesid, grosssalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                '    End If

                'Next
            Next
        End If
    End Sub

    Private Sub LocalMedicalExpenses(ByVal stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal personjoindateid As Integer, ByVal p As DataRow)
        Dim myamount As Double = 0
        Dim myamountarr(11) As Double
        Dim myamountarr2(11) As Double
        'For i = 0 To 11
        '    myamountarr(i) = 0
        'Next

        'Get familymember plan
        Dim mypct As Double = 0
        Dim qry = From expenses In Dataset1.Tables(12)
                        Where expenses.Item("personjoindateid") = personjoindateid
                        Select expenses Order By expenses.Item("validfrom") Descending
        For Each myresult In qry
            Dim planid = myresult.Item("planid")
            Dim mycount = myresult.Item("count")
            myamount = myresult.Item("amount") * myresult.Item("count")
            'Find table plan
            'Populate array myamountarr
            Dim qry2 = From expenses In Dataset1.Tables(17)
                        Where expenses.Item("planid") = planid
                        Select expenses Order By expenses.Item("validfrom") Ascending

            For Each myresult2 In qry2
                For i = 0 To 11
                    If myresult2.Item("validfrom") <= CDate(String.Format("{0}-{1}-1", myyear, i + 1)) Then
                        myamountarr(i) = myresult2.Item("amount") * mycount
                    End If
                Next
            Next
            Exit For 'Get only the first from the latest
        Next


        Dim qry3 = From category In Dataset1.Tables(11)
                              Where category.Item("category") = mycategory
                              Select category
                              Order By category.Item("validfrom") Descending

        For Each myresult In qry3
            Dim planid = myresult.Item("planid")

            myamount += myresult.Item("amount")
            If Not IsDBNull(myresult.Item("pct")) Then
                mypct += myresult.Item("pct")
            End If

            'Find table plan
            'Populate array myamountarr
            Dim qry2 = From expenses In Dataset1.Tables(17)
                        Where expenses.Item("planid") = planid
                        Select expenses Order By expenses.Item("validfrom") Ascending

            For Each myresult2 In qry2
                For i = 0 To 11
                    If myresult2.Item("validfrom") <= CDate(String.Format("{0}-{1}-1", myyear, i + 1)) Then
                        'myamountarr(i) += myresult2.Item("amount")
                        myamountarr2(i) = myresult2.Item("amount")
                    End If
                Next
            Next

            'add myamountarr with myamountarr2
            For i = 0 To 11
                myamountarr(i) += myamountarr2(i)
            Next

            Exit For 'Get only the first from the latest
        Next
        If myamount > 0 Then
            'Dim ValidMonth As Integer = getvalidmonth(joindate)
            Dim validmonth As Integer
            If p.Item("enddate").ToString <> "" Then
                validmonth = getvalidmonth(joindate, p.Item("enddate"))
            Else
                validmonth = getvalidmonth(joindate)
            End If


            Dim enddate As Date? = Nothing
            If Not IsDBNull(p.Item("enddate")) Then
                enddate = p.Item("enddate")
            End If
            If dbtools1.Region = "HK" Then
                If Not IsDBNull(p.Item("effectivedatestart")) Then
                    joindate = p.Item("effectivedatestart")
                End If
                If Not IsDBNull(p.Item("enddate")) Then
                    enddate = p.Item("enddate")                    
                End If
                If Not IsDBNull(p.Item("effectivedateend")) Then
                    enddate = p.Item("effectivedateend")
                End If
            Else
                If Not IsDBNull(p.Item("enddate")) Then
                    enddate = p.Item("enddate")
                End If
            End If

            For mymonth = 1 To 12
                'If joindate <= CDate(myyear & "-" & mymonth & "-1") Then
                '    createrecord(stringBuilder1, personexpensesid, myamount / 12, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & mymonth & "-1")))
                'End If
                Dim currentdate As Date = CDate(myyear & "-" & mymonth & "-1")
                Dim check As Boolean = True
                'If p.Item("enddate").ToString <> "" Then
                '    check = currentdate < p.Item("enddate")
                'End If
                If Not IsNothing(enddate) Then
                    check = currentdate < enddate
                End If
                'If currentdate >= p.Item("joindate") And check Then

                'If ValidJoinDate(p.Item("joindate"), mymonth, myyear) And check Then
                If ValidJoinDate(joindate, mymonth, myyear) And check Then
                    'createrecord(stringBuilder1, personexpensesid, (myamount / 12) * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & mymonth & "-1")), p.Item("headcount"), p.Item("enddate"))
                    'createrecord(stringBuilder1, personexpensesid, (myamount / 12) * p.Item("headcount") + (mypct * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(mymonth) * p.Item("headcount")), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & mymonth & "-1")), p.Item("headcount"), p.Item("enddate"))
                    createrecord(stringBuilder1, personexpensesid, (myamountarr(mymonth - 1) / 12) * p.Item("headcount") + (mypct * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(mymonth) * p.Item("headcount")), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & mymonth & "-1")), p.Item("headcount"), p.Item("enddate"))

                End If

            Next

        End If
    End Sub
    Private Sub LocalMedicalExpenses1(ByVal stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal personjoindateid As Integer, ByVal p As DataRow)
        Dim myamount As Double = 0
        'Get familymember plan
        Dim mypct As Double = 0
        Dim qry = From expenses In Dataset1.Tables(12)
                        Where expenses.Item("personjoindateid") = personjoindateid
                        Select expenses Order By expenses.Item("validfrom") Descending
        For Each myresult In qry
            myamount = myresult.Item("amount") * myresult.Item("count")
            Exit For
        Next


        Dim qry3 = From category In Dataset1.Tables(11)
                              Where category.Item("category") = mycategory
                              Select category
                              Order By category.Item("validfrom") Descending

        For Each myresult In qry3
            myamount += myresult.Item("amount")
            If Not IsDBNull(myresult.Item("pct")) Then
                mypct += myresult.Item("pct")
            End If
            Exit For
        Next
        If myamount > 0 Then
            'Dim ValidMonth As Integer = getvalidmonth(joindate)
            Dim validmonth As Integer
            If p.Item("enddate").ToString <> "" Then
                validmonth = getvalidmonth(joindate, p.Item("enddate"))
            Else
                validmonth = getvalidmonth(joindate)
            End If
            For mymonth = 1 To 12
                'If joindate <= CDate(myyear & "-" & mymonth & "-1") Then
                '    createrecord(stringBuilder1, personexpensesid, myamount / 12, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & mymonth & "-1")))
                'End If
                Dim currentdate As Date = CDate(myyear & "-" & mymonth & "-1")
                Dim check As Boolean = True
                If p.Item("enddate").ToString <> "" Then
                    check = currentdate < p.Item("enddate")
                End If
                'If currentdate >= p.Item("joindate") And check Then

                If ValidJoinDate(p.Item("joindate"), mymonth, myyear) And check Then
                    'createrecord(stringBuilder1, personexpensesid, (myamount / 12) * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & mymonth & "-1")), p.Item("headcount"), p.Item("enddate"))
                    createrecord(stringBuilder1, personexpensesid, (myamount / 12) * p.Item("headcount") + (mypct * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(mymonth) * p.Item("headcount")), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & mymonth & "-1")), p.Item("headcount"), p.Item("enddate"))

                End If

            Next

        End If
    End Sub
    Private Sub LocalMedicalExpenses2(ByVal stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal personjoindateid As Integer, ByVal p As DataRow)
        Dim myamount As Double = 0
        Dim myamountarr(11) As Double
        Dim myamountarr2(11) As Double
        'For i = 0 To 11
        '    myamountarr(i) = 0
        'Next

        'Get familymember plan
        Dim mypct As Double = 0
        Dim qry = From expenses In Dataset1.Tables(12)
                        Where expenses.Item("personjoindateid") = personjoindateid
                        Select expenses Order By expenses.Item("validfrom") Descending
        For Each myresult In qry
            Dim planid = myresult.Item("planid")
            Dim mycount = myresult.Item("count")
            myamount = myresult.Item("amount") * myresult.Item("count")
            'Find table plan
            'Populate array myamountarr
            Dim qry2 = From expenses In Dataset1.Tables(17)
                        Where expenses.Item("planid") = planid
                        Select expenses Order By expenses.Item("validfrom") Ascending

            For Each myresult2 In qry2
                For i = 0 To 11
                    If myresult2.Item("validfrom") <= CDate(String.Format("{0}-{1}-1", myyear, i + 1)) Then
                        myamountarr(i) = myresult2.Item("amount") * mycount
                    End If
                Next
            Next
            Exit For 'Get only the first from the latest
        Next


        Dim qry3 = From category In Dataset1.Tables(11)
                              Where category.Item("category") = mycategory
                              Select category
                              Order By category.Item("validfrom") Descending

        For Each myresult In qry3
            Dim planid = myresult.Item("planid")

            myamount += myresult.Item("amount")
            If Not IsDBNull(myresult.Item("pct")) Then
                mypct += myresult.Item("pct")
            End If

            'Find table plan
            'Populate array myamountarr
            Dim qry2 = From expenses In Dataset1.Tables(17)
                        Where expenses.Item("planid") = planid
                        Select expenses Order By expenses.Item("validfrom") Ascending

            For Each myresult2 In qry2
                For i = 0 To 11
                    If myresult2.Item("validfrom") <= CDate(String.Format("{0}-{1}-1", myyear, i + 1)) Then
                        'myamountarr(i) += myresult2.Item("amount")
                        myamountarr2(i) = myresult2.Item("amount")
                    End If
                Next
            Next

            'add myamountarr with myamountarr2
            For i = 0 To 11
                myamountarr(i) += myamountarr2(i)
            Next

            Exit For 'Get only the first from the latest
        Next
        If myamount > 0 Then
            'Dim ValidMonth As Integer = getvalidmonth(joindate)
            Dim validmonth As Integer
            If p.Item("enddate").ToString <> "" Then
                validmonth = getvalidmonth(joindate, p.Item("enddate"))
            Else
                validmonth = getvalidmonth(joindate)
            End If
            For mymonth = 1 To 12
                'If joindate <= CDate(myyear & "-" & mymonth & "-1") Then
                '    createrecord(stringBuilder1, personexpensesid, myamount / 12, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & mymonth & "-1")))
                'End If
                Dim currentdate As Date = CDate(myyear & "-" & mymonth & "-1")
                Dim check As Boolean = True
                If p.Item("enddate").ToString <> "" Then
                    check = currentdate < p.Item("enddate")
                End If
                'If currentdate >= p.Item("joindate") And check Then

                If ValidJoinDate(p.Item("joindate"), mymonth, myyear) And check Then
                    'createrecord(stringBuilder1, personexpensesid, (myamount / 12) * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & mymonth & "-1")), p.Item("headcount"), p.Item("enddate"))
                    'createrecord(stringBuilder1, personexpensesid, (myamount / 12) * p.Item("headcount") + (mypct * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(mymonth) * p.Item("headcount")), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & mymonth & "-1")), p.Item("headcount"), p.Item("enddate"))
                    createrecord(stringBuilder1, personexpensesid, (myamountarr(mymonth - 1) / 12) * p.Item("headcount") + (mypct * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(mymonth) * p.Item("headcount")), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & mymonth & "-1")), p.Item("headcount"), p.Item("enddate"))

                End If

            Next

        End If
    End Sub
    Private Sub RedPocket(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal dr As DataRow, ByVal p As DataRow, ByVal personexpensesid As Integer, ByVal serviceyear As Double, ByVal myverid As Integer, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata)
        'find month
        Dim amount As Double
        Dim expensesnature As String = "Red Pocket"
        'find category
        Dim qry3 = From category In Dataset1.Tables(3)
                   Where category.Item("category") = p.Item("category") And category.Item("categorytype") = "Red Pocket" And category.Item("myyear") = myyear
                   Select category

        For Each ct In qry3
            If serviceyear >= 0.5 Then
                amount = AmountA
            ElseIf serviceyear < 0.2 Then
                amount = AmountC
            Else
                amount = AmountB
            End If
            'Dim myqr = From months In Dataset1.Tables(1)
            '                               Where months.Item("expensesnatureid") = dr.Item("expensesnatureid")
            '                               Select months

            'For Each K In myqr
            '    Dim currentdate As Date = CDate(myyear & "-" & K.Item("mymonthint") & "-1")
            '    If currentdate > p.Item("joindate") Then
            '        createrecord(stringBuilder1, personexpensesid, amount, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & K.Item("mymonthint") & "-1")))
            '    End If
            'Next

            'get categorytxmonths
            Dim mycategory = p.Item("category")
            Dim mymonths = From myrecord In Dataset1.Tables(14)
                           Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = expensesnature
                           Select myrecord

            Dim joindate As Date = p.Item("joindate")
            Dim enddate As Date? = Nothing
            If Not IsDBNull(p.Item("enddate")) Then
                'If Not p.Item("enddate").ToString() <> "" Then
                enddate = p.Item("enddate")
            End If

            If dbtools1.Region = "HK" Then
                If Not IsDBNull(p.Item("effectivedatestart")) Then
                    joindate = p.Item("effectivedatestart")
                End If
                If Not IsDBNull(p.Item("enddate")) Then
                    enddate = p.Item("enddate")
                End If
                If Not IsDBNull(p.Item("effectivedateend")) Then
                    enddate = p.Item("effectivedateend")
                End If
            Else
                If Not IsDBNull(p.Item("enddate")) Then
                    enddate = p.Item("enddate")
                End If
            End If

            For Each record In mymonths
                Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
                Dim check As Boolean = True
                'If p.Item("enddate").ToString <> "" Then
                '    check = currentdate < p.Item("enddate")
                'End If
                If Not IsNothing(enddate) Then
                    check = currentdate < enddate
                End If
                'If currentdate >= p.Item("joindate") And check Then
                'If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
                If ValidJoinDate(joindate, record.Item("months"), myyear) And check Then
                    createrecord(stringBuilder1, personexpensesid, amount * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                    'createrecord(stringBuilder1, personexpensesid, amount, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")))
                End If

            Next

        Next

    End Sub

    Private Sub RedPocket1(ByRef stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal dr As DataRow, ByVal p As DataRow, ByVal personexpensesid As Integer, ByVal serviceyear As Double, ByVal myverid As Integer, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata)
        'find month
        Dim amount As Double
        Dim expensesnature As String = "Red Pocket"
        'find category
        Dim qry3 = From category In Dataset1.Tables(3)
                   Where category.Item("category") = p.Item("category") And category.Item("categorytype") = "Red Pocket" And category.Item("myyear") = myyear
                   Select category

        For Each ct In qry3
            If serviceyear >= 0.5 Then
                amount = AmountA
            ElseIf serviceyear < 0.2 Then
                amount = AmountC
            Else
                amount = AmountB
            End If
            'Dim myqr = From months In Dataset1.Tables(1)
            '                               Where months.Item("expensesnatureid") = dr.Item("expensesnatureid")
            '                               Select months

            'For Each K In myqr
            '    Dim currentdate As Date = CDate(myyear & "-" & K.Item("mymonthint") & "-1")
            '    If currentdate > p.Item("joindate") Then
            '        createrecord(stringBuilder1, personexpensesid, amount, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & K.Item("mymonthint") & "-1")))
            '    End If
            'Next

            'get categorytxmonths
            Dim mycategory = p.Item("category")
            Dim mymonths = From myrecord In Dataset1.Tables(14)
                           Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = expensesnature
                           Select myrecord




            For Each record In mymonths
                Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
                Dim check As Boolean = True
                If p.Item("enddate").ToString <> "" Then
                    check = currentdate < p.Item("enddate")
                End If
                'If currentdate >= p.Item("joindate") And check Then
                If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
                    createrecord(stringBuilder1, personexpensesid, amount * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                    'createrecord(stringBuilder1, personexpensesid, amount, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")))
                End If

            Next

        Next

    End Sub

    Private Sub StaffAward1(ByVal stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal dr As DataRow, ByVal p As DataRow, ByVal personexpensesid As Integer, ByVal serviceyear As Double, ByVal myverid As Integer, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata)
        'find month
        Dim amount As Double = 0
        Dim expensesnature As String = "staff award"
        'find category
        Dim qry3 = From category In Dataset1.Tables(3)
                   Where category.Item("category") = p.Item("category") And category.Item("categorytype") = expensesnature And category.Item("myyear") = myyear
                   Select category

        For Each ct In qry3
            If serviceyear >= 10 And serviceyear < 11 Then
                amount = serviceyear10
            ElseIf serviceyear >= 15 And serviceyear < 15 Then
                amount = serviceyear15

            End If
            If amount > 0 Then
                'Dim myqr = From months In Dataset1.Tables(1)
                '                           Where months.Item("expensesnatureid") = dr.Item("expensesnatureid")
                '                           Select months

                'For Each K In myqr
                '    createrecord(stringBuilder1, personexpensesid, amount, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & K.Item("mymonthint") & "-1")))
                'Next
                'get categorytxmonths
                Dim mycategory = p.Item("category")
                Dim mymonths = From myrecord In Dataset1.Tables(14)
                               Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = expensesnature
                               Select myrecord

                For Each record In mymonths
                    Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
                    Dim check As Boolean = True
                    If p.Item("enddate").ToString <> "" Then
                        check = currentdate < p.Item("enddate")
                    End If
                    'If currentdate >= p.Item("joindate") And check Then
                    If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then

                        createrecord(stringBuilder1, personexpensesid, amount, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                    End If

                Next

            End If

        Next
    End Sub
    Private Sub StaffAward(ByVal stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal dr As DataRow, ByVal p As DataRow, ByVal personexpensesid As Integer, ByVal serviceyear As Double, ByVal myverid As Integer, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal personjoindatecategoryid As Integer)
        'find month

        Dim expensesnature As String = "staff award"
        'find category
        Dim qry = From expenses In Dataset1.Tables(5)
                    Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = expensesnature And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                    Select expenses Order By expenses.Item("validdate") Descending
        Dim myamount As Double = 0
        Dim validdate As Date
        For Each myresult In qry
            myamount = myresult.Item("amount")
            validdate = myresult.Item("validdate")
            Exit For
        Next

        If myamount > 0 Then
            'get grosssalary jan-Dec + 13th Month
            Try

                Dim joindate As Date = p.Item("joindate")
                Dim enddate As Date? = Nothing
                If Not IsDBNull(p.Item("enddate")) Then                    
                    enddate = p.Item("enddate")
                End If

                If dbtools1.Region = "HK" Then
                    If Not IsDBNull(p.Item("effectivedatestart")) Then
                        joindate = p.Item("effectivedatestart")
                    End If
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                    End If
                    If Not IsDBNull(p.Item("effectivedateend")) Then
                        enddate = p.Item("effectivedateend")
                    End If
                Else
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                    End If
                End If


                Dim mymonths = From myrecord In Dataset1.Tables(16)
                               Where myrecord.Item("personjoindatecategoryid") = personjoindatecategoryid And myrecord.Item("expensesnature") = expensesnature
                               Select myrecord


                For Each record In mymonths
                    Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
                    Dim check As Boolean = True
                    'If p.Item("enddate").ToString <> "" Then
                    '    check = currentdate < p.Item("enddate")
                    'End If
                    If Not IsNothing(enddate) Then
                        check = currentdate < enddate
                    End If
                    'If currentdate >= p.Item("joindate") And check Then
                    'If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
                    If ValidJoinDate(joindate, record.Item("months"), myyear) And check Then
                        createrecord(stringBuilder1, personexpensesid, myamount, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                    End If
                Next
            Catch ex As Exception
                Debug.WriteLine("Staff Award Error: personexpensesid {0} personjoindatecategoryid {1}", personexpensesid, personjoindatecategoryid)
            End Try

        End If

    End Sub

    Private Sub TrainingDept(ByVal stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal dr As DataRow, ByVal p As DataRow, ByVal personexpensesid As Integer, ByVal serviceyear As Double, ByVal myverid As Integer, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata)

        Dim amount As Double = 0
        Dim expensesnature As String = "Training (Dept)"
        Dim personjoindatecategoryid As Integer = p.Item("personjoindatecategoryid")
        If personjoindatecategoryid = 541 Then
            Debug.WriteLine("debug mode")
        End If
        Dim joindate = p.Item("joindate")
        'find category
        Dim qry3 = From category In Dataset1.Tables(3)
                   Where category.Item("category") = p.Item("category") And category.Item("categorytype") = expensesnature And category.Item("myyear") = myyear
                   Select category

        For Each ct In qry3

            'get amount in tablepersonexpensesdtl
            Debug.WriteLine("apply: {2} For  Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} ", personjoindatecategoryid, 0, expensesnature, personexpensesid)
            Dim qry = From expenses In Dataset1.Tables(5)
                  Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = expensesnature And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                  Select expenses Order By expenses.Item("validdate") Descending

            Dim myamount As Double = 0
            Dim validdate As Date
            For Each myresult In qry
                myamount = myresult.Item("amount")
                validdate = myresult.Item("validdate")
                Exit For
            Next
            If myamount > 0 Then
                'Debug.WriteLine("Personname {0} personexpensesid {1} amount {2} personjoindatecategoryid  {3} expensesnatuer {4}", p.Item("personname"), personexpensesid, myamount, personjoindatecategoryid, expensesnature)
                Dim totalsalary As Double = 0
                Try
                    'Dim ValidMonth As Integer = getvalidmonth(joindate)
                    Dim validmonth As Integer
                    If p.Item("enddate").ToString <> "" Then
                        validmonth = getvalidmonth(joindate, p.Item("enddate"))
                    Else
                        validmonth = getvalidmonth(joindate)
                    End If


                    Dim enddate As Date? = Nothing
                    If Not IsDBNull(p.Item("enddate")) Then                        
                        enddate = p.Item("enddate")
                    End If

                    If dbtools1.Region = "HK" Then
                        If Not IsDBNull(p.Item("effectivedatestart")) Then
                            joindate = p.Item("effectivedatestart")
                        End If
                        If Not IsDBNull(p.Item("enddate")) Then
                            enddate = p.Item("enddate")
                        End If
                        If Not IsDBNull(p.Item("effectivedateend")) Then
                            enddate = p.Item("effectivedateend")
                        End If
                    Else
                        If Not IsDBNull(p.Item("enddate")) Then
                            enddate = p.Item("enddate")
                        End If
                    End If


                    For i = 1 To 12
                        If joindate <= CDate(myyear & "-" & i & "-1") Then
                            totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                        End If
                    Next
                    totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).doublepay
                    totalsalary *= myamount
                    For M = 1 To 12
                        Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                        Dim check As Boolean = True
                        'If p.Item("enddate").ToString <> "" Then
                        '    check = currentdate < p.Item("enddate")
                        'End If
                        If Not IsNothing(enddate) Then
                            check = currentdate < enddate
                        End If
                        'If currentdate >= p.Item("joindate") And check Then
                        'If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then
                        If ValidJoinDate(joindate, M, myyear) And check Then
                            createrecord(stringBuilder1, personexpensesid, totalsalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")), p.Item("headcount"), p.Item("enddate"))
                        End If

                    Next
                Catch ex As Exception
                    Debug.WriteLine(ex.Message)

                End Try
            ElseIf myamount <= 0 Then
                'Personjoindatecategoryid: 285
                'If personjoindatecategoryid = 285 Then
                '    Debug.WriteLine("DebugMode")
                'End If
                myamount = ct.Item("amount")
                Try
                    Dim grosssalary As Double = getGrossSalary(PersonSalaryDict, personjoindatecategoryid) * myamount

                    'Dim ValidMonth As Integer = getvalidmonth(joindate)
                    Dim validmonth As Integer
                    If p.Item("enddate").ToString <> "" Then
                        validmonth = getvalidmonth(joindate, p.Item("enddate"))
                    Else
                        validmonth = getvalidmonth(joindate)
                    End If

                    Dim enddate As Date? = Nothing
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                    End If

                    If dbtools1.Region = "HK" Then
                        If Not IsDBNull(p.Item("effectivedatestart")) Then
                            joindate = p.Item("effectivedatestart")
                        End If
                        If Not IsDBNull(p.Item("enddate")) Then
                            enddate = p.Item("enddate")
                        End If
                        If Not IsDBNull(p.Item("effectivedateend")) Then
                            enddate = p.Item("effectivedateend")
                        End If
                    End If


                    For M = 1 To 12
                        'If joindate <= CDate(myyear & "-" & M & "-1") Then
                        'If ValidJoinDate(joindate, M, myyear) Then
                        'createrecord(stringBuilder1, personexpensesid, grosssalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                        Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                        Dim check As Boolean = True
                        'If p.Item("enddate").ToString <> "" Then
                        '    check = currentdate < p.Item("enddate")
                        'End If

                        If Not IsNothing(enddate) Then
                            check = currentdate < enddate
                        End If

                        'If currentdate >= p.Item("joindate") And check Then
                        'If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then
                        If ValidJoinDate(joindate, M, myyear) And check Then
                            'createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth) * p.Item("headcount"), myverid, currentdate)
                            createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth), myverid, currentdate, p.Item("headcount"), p.Item("enddate"))
                        End If

                        'End If

                    Next
                Catch ex As Exception
                    Debug.WriteLine("Debug mode")
                End Try

            End If

        Next
    End Sub
    Private Sub TrainingDept1(ByVal stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal dr As DataRow, ByVal p As DataRow, ByVal personexpensesid As Integer, ByVal serviceyear As Double, ByVal myverid As Integer, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata)

        Dim amount As Double = 0
        Dim expensesnature As String = "Training (Dept)"
        Dim personjoindatecategoryid As Integer = p.Item("personjoindatecategoryid")
        If personjoindatecategoryid = 541 Then
            Debug.WriteLine("debug mode")
        End If
        Dim joindate = p.Item("joindate")
        'find category
        Dim qry3 = From category In Dataset1.Tables(3)
                   Where category.Item("category") = p.Item("category") And category.Item("categorytype") = expensesnature And category.Item("myyear") = myyear
                   Select category

        For Each ct In qry3

            'get amount in tablepersonexpensesdtl
            Debug.WriteLine("apply: {2} For  Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} ", personjoindatecategoryid, 0, expensesnature, personexpensesid)
            Dim qry = From expenses In Dataset1.Tables(5)
                  Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = expensesnature And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                  Select expenses Order By expenses.Item("validdate") Descending

            Dim myamount As Double = 0
            Dim validdate As Date
            For Each myresult In qry
                myamount = myresult.Item("amount")
                validdate = myresult.Item("validdate")
                Exit For
            Next
            If myamount > 0 Then
                'Debug.WriteLine("Personname {0} personexpensesid {1} amount {2} personjoindatecategoryid  {3} expensesnatuer {4}", p.Item("personname"), personexpensesid, myamount, personjoindatecategoryid, expensesnature)
                Dim totalsalary As Double = 0
                Try
                    'Dim ValidMonth As Integer = getvalidmonth(joindate)
                    Dim validmonth As Integer
                    If p.Item("enddate").ToString <> "" Then
                        validmonth = getvalidmonth(joindate, p.Item("enddate"))
                    Else
                        validmonth = getvalidmonth(joindate)
                    End If

                    For i = 1 To 12
                        If joindate <= CDate(myyear & "-" & i & "-1") Then
                            totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                        End If
                    Next
                    totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).doublepay
                    totalsalary *= myamount
                    For M = 1 To 12
                        Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                        Dim check As Boolean = True
                        If p.Item("enddate").ToString <> "" Then
                            check = currentdate < p.Item("enddate")
                        End If
                        'If currentdate >= p.Item("joindate") And check Then
                        If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then
                            createrecord(stringBuilder1, personexpensesid, totalsalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")), p.Item("headcount"), p.Item("enddate"))
                        End If

                    Next
                Catch ex As Exception
                    Debug.WriteLine(ex.Message)

                End Try
            ElseIf myamount <= 0 Then
                'Personjoindatecategoryid: 285
                'If personjoindatecategoryid = 285 Then
                '    Debug.WriteLine("DebugMode")
                'End If
                myamount = ct.Item("amount")
                Try
                    Dim grosssalary As Double = getGrossSalary(PersonSalaryDict, personjoindatecategoryid) * myamount

                    'Dim ValidMonth As Integer = getvalidmonth(joindate)
                    Dim validmonth As Integer
                    If p.Item("enddate").ToString <> "" Then
                        validmonth = getvalidmonth(joindate, p.Item("enddate"))
                    Else
                        validmonth = getvalidmonth(joindate)
                    End If
                    For M = 1 To 12
                        'If joindate <= CDate(myyear & "-" & M & "-1") Then
                        'If ValidJoinDate(joindate, M, myyear) Then
                        'createrecord(stringBuilder1, personexpensesid, grosssalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                        Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                        Dim check As Boolean = True
                        If p.Item("enddate").ToString <> "" Then
                            check = currentdate < p.Item("enddate")
                        End If
                        'If currentdate >= p.Item("joindate") And check Then
                        If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then
                            'createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth) * p.Item("headcount"), myverid, currentdate)
                            createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth), myverid, currentdate, p.Item("headcount"), p.Item("enddate"))
                        End If

                        'End If

                    Next
                Catch ex As Exception
                    Debug.WriteLine("Debug mode")
                End Try

            End If

        Next
    End Sub
    Private Sub OtherExpenses(ByVal stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal dr As DataRow, ByVal personexpensesid As Integer, ByVal myverid As Integer, ByVal myyear As Integer, ByVal joindate As Date, ByVal serviceyear As Double, ByVal mycategory As String, ByVal personjoindatecategoryid As Object, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata, ByVal p As DataRow)
        If dr.Item("fullyear") = True Then

            Debug.WriteLine("Apply common fullyear Personjoindatecategoryid: {0}, expensesnatureid : {1} - {2}, personexpensesid: {3} ", personjoindatecategoryid, dr.Item("expensesnatureid"), dr.Item("expensesnature"), personexpensesid)

            'If dr.Item("expensesnature") = "Company Activites" Then
            '    Debug.WriteLine("Company Activites")
            'End If
            'If dr.Item("expensesnature") = "Company Activites" And personjoindatecategoryid = 387 Then
            '     Debug.WriteLine("Company Activites")
            'End If

            

            Dim expensesnature = dr.Item("expensesnature")

            If expensesnature = "LSP provision" Then
                Debug.WriteLine("Debug mode")
            End If

            If expensesnature = "Training (Indv)" And personjoindatecategoryid = 156 Then
                Debug.WriteLine("debug Training (Indv)")
            End If
            If expensesnature = "Car & Parking Allowance" And personjoindatecategoryid = 156 Then
                Debug.WriteLine("debug Training (Indv)")
            End If
            'find amount for personexpensesdetail
            Dim qry = From expenses In Dataset1.Tables(5)
                        Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = dr.Item("expensesnature") And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                        Select expenses Order By expenses.Item("validdate") Descending
            Dim myamount As Double = 0
            Dim validdate As Date
            For Each myresult In qry
                myamount = myresult.Item("amount")
                validdate = myresult.Item("validdate")
                Exit For
            Next

            If dbtools1.Region = "SZ" Then
                If dr.Item("expensesnature") = "Recruitment Expenses" Then
                    'myamount = (myamount * getGrossSalary(PersonSalaryDict, personjoindatecategoryid))
                    'Nikita
                    'myamount = (myamount * getGrossSalary(PersonSalaryDict, personjoindatecategoryid)) / p.Item("headcount")
                    myamount = (myamount * getGrossSalary(PersonSalaryDict, personjoindatecategoryid))
                ElseIf dr.Item("expensesnature") = "LSP provision" Then
                    'LSP Provision same as HK
                    'myamount = myamount * p.Item("headcount")
                    myamount = (myamount * getGrossSalary(PersonSalaryDict, personjoindatecategoryid))

                End If
            ElseIf dbtools1.Region = "PH" Then
                If dr.Item("expensesnature") = "Bonus" Then
                    'myamount = myamount * getGrossSalary(PersonSalaryDict, personjoindatecategoryid)  / p.Item("headcount")
                    myamount = myamount * getGrossSalary(PersonSalaryDict, personjoindatecategoryid)
                End If
            ElseIf dbtools1.Region = "HK" Then
                'myamount = myamount * p.Item("headcount")
                'no need multiply by headcount again
                myamount = myamount
            End If

            Dim enddate As Date? = Nothing
            If Not IsDBNull(p.Item("enddate")) Then
                enddate = p.Item("enddate")
            End If
            If dbtools1.Region = "HK" Then
                If Not IsDBNull(p.Item("effectivedatestart")) Then
                    joindate = p.Item("effectivedatestart")
                End If
                If Not IsDBNull(p.Item("enddate")) Then
                    enddate = p.Item("enddate")
                End If
                If Not IsDBNull(p.Item("effectivedateend")) Then
                    enddate = p.Item("effectivedateend")
                End If
            Else
                If Not IsDBNull(p.Item("enddate")) Then
                    enddate = p.Item("enddate")
                End If
            End If

            If myamount > 0 Then 'Fix Amount
                'Debug.WriteLine("Personname {0} personexpensesid {1} amount {2} expensesnatuer {3} personjoindatecategoryid {4}", p.Item("personname"), personexpensesid, myamount, expensesnature, personjoindatecategoryid)
                For M = 1 To 12
                    'createrecord(stringBuilder1, personexpensesid, myamount / 12, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                    Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                    Dim check As Boolean = True
                    'If p.Item("enddate").ToString <> "" Then
                    '    check = currentdate < p.Item("enddate")
                    'End If
                    If Not IsNothing(enddate) Then
                        check = currentdate < enddate
                    End If
                    'If currentdate >= p.Item("joindate") And check Then
                    'If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then
                    If ValidJoinDate(joindate, M, myyear) And check Then
                        'createrecord(stringBuilder1, personexpensesid, myamount / 12, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")), p.Item("headcount"), p.Item("enddate"))
                        createrecord(stringBuilder1, personexpensesid, myamount / 12 * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")), p.Item("headcount"), p.Item("enddate"))
                    End If
                Next
            ElseIf myamount <= 0 Then
                'Get based on Category
                'If dbtools1.Region = "HK" Then
                Dim qry3 = From category In Dataset1.Tables(3)
                                 Where category.Item("category") = mycategory And category.Item("categorytype") = expensesnature And category.Item("myyear") = myyear
                                 Select category
                'Category listed

                For Each dt In qry3
                    Dim grosssalary As Double
                    myamount = dt.Item("amount")
                    If expensesnature = "LSP provision" Then
                        'grosssalary = getGrossSalary(PersonSalaryDict, personjoindatecategoryid) * myamount
                        If p.Item("headcount") = 0 Then
                            grosssalary = 0
                        Else
                            If p.Item("headcount") < 0 Then
                                MessageBox.Show(p.Item("headcount"))
                            End If
                            'Nikita
                            'grosssalary = (getGrossSalary(PersonSalaryDict, personjoindatecategoryid) * myamount) / p.Item("headcount")
                            grosssalary = (getGrossSalary(PersonSalaryDict, personjoindatecategoryid) * myamount) / p.Item("headcount")
                        End If

                    Else
                        grosssalary = myamount
                    End If
                    Dim validmonth As Double
                    If p.Item("personname") = "TAM King Fai" Then
                        Debug.Print("debug mode")
                    End If
                    'If p.Item("enddate").ToString <> "" Then
                    If Not IsNothing(enddate) Then
                        If dr.Item("expensesnature") = "Company Activites" Then
                            validmonth = getvalidmonthCA(joindate)
                        Else
                            'validmonth = getvalidmonth(joindate, p.Item("enddate"))
                            validmonth = getvalidmonth(joindate, enddate)
                        End If
                    Else
                        If dr.Item("expensesnature") = "Company Activites" Then
                            validmonth = getvalidmonthCA(joindate)
                        Else
                            validmonth = getvalidmonth(joindate)

                        End If

                    End If



                    For M = 1 To 12
                        'If joindate <= CDate(myyear & "-" & M & "-1") Then
                        'modified on 2013-08-08
                        If ValidJoinDate(joindate, M, myyear) Then
                            'createrecord(stringBuilder1, personexpensesid, grosssalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                            Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                            Dim check As Boolean = True
                            'If p.Item("enddate").ToString <> "" Then
                            If Not IsNothing(enddate) Then

                                'Company Activities modified on 2013-08-08
                                If dr.Item("expensesnature") = "Company Activites" Then
                                    'check = p.Item("enddate") = CDate(myyear & "-12-31")
                                    check = enddate = CDate(myyear & "-12-31")
                                Else
                                    'check = currentdate < p.Item("enddate")
                                    check = currentdate < enddate
                                End If

                            End If
                            'If currentdate >= p.Item("joindate") And check Then
                            'modified on 2013-08-08
                            'If Year(p.Item("joindate")) = myyear Then 'join in the same year budget
                            '    If M >= Month(p.Item("joindate")) And check Then
                            '        createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth) * p.Item("headcount"), myverid, currentdate)
                            '    End If
                            'Else 'old employee
                            '    If currentdate >= p.Item("joindate") And check Then
                            '        createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth) * p.Item("headcount"), myverid, currentdate)
                            '        'createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth), myverid, currentdate)
                            '    End If
                            'End If
                            'If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then
                            If ValidJoinDate(joindate, M, myyear) And check Then

                                createrecord(stringBuilder1, personexpensesid, (grosssalary / validmonth) * p.Item("headcount"), myverid, currentdate, p.Item("headcount"), p.Item("enddate"))



                            End If
                        End If

                    Next



                Next
                'ElseIf dbtools1.Region = "TW" Then
                '    Dim qry3 = From category In Dataset1.Tables(3)
                '                     Where category.Item("category") = mycategory And category.Item("categorytype") = expensesnature And category.Item("myyear") = myyear
                '                     Select category
                '    'Category listed

                '    For Each dt In qry3
                '        myamount = dt.Item("amount")
                '        Dim ValidMonth As Integer = getvalidmonth(joindate)
                '        For M = 1 To 12
                '            If joindate <= CDate(myyear & "-" & M & "-1") Then
                '                createrecord(stringBuilder1, personexpensesid, myamount / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                '            End If

                '        Next
                '    Next
                'End If


            End If


        Else
            'find expensesnature months

            Debug.WriteLine("Apply common monthly  Personjoindatecategoryid: {0}, expensesnatureid: {1} - {2}, personexpensesid: {3} ", personjoindatecategoryid, dr.Item("expensesnatureid"), dr.Item("expensesnature"), personexpensesid)
            Dim expensesnature = dr.Item("expensesnature")

        

            If expensesnature = "Special Allowance" And personjoindatecategoryid = 174 Then
                Debug.WriteLine("debug Special allowance")
            End If
            If expensesnature = "IIT Chief Rep" And personjoindatecategoryid = 240 Then
                Debug.WriteLine("IIT Chief Rep debug.")

            End If
            If expensesnature = "Car & Parking Allowance" And personjoindatecategoryid = 156 Then
                Debug.WriteLine("debug Special allowance")
            End If
            If personjoindatecategoryid = 372 And expensesnature = "HEALTH INSURANCE" Then
                Debug.WriteLine("Debug mode")
            End If

            If dbtools1.Region = "HK" Or dbtools1.Region = "SZ" Or dbtools1.Region = "TW" Or dbtools1.Region = "PH" Then
                If dr.Item("expensesnature") = "Recruitment Expenses" Then
                    Debug.WriteLine("debug mode")
                End If
                'Personal
                Dim qry = From expenses In Dataset1.Tables(5)
                        Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = dr.Item("expensesnature") And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                        Select expenses Order By expenses.Item("validdate") Descending

                Dim myamount As Double = 0
                Dim validdate As Date
                Dim sapcc As String = String.Empty
                For Each myresult In qry
                    myamount = myresult.Item("amount")
                    validdate = myresult.Item("validdate")
                    'sapcc = myresult.Item("sapcc")
                    Exit For
                Next
                If dr.Item("expensesnature").ToString = "Housing fund" Then

                    Debug.WriteLine("debug mode")

                End If
                If dr.Item("expensesnature").ToString = "Accident" Then

                    Debug.WriteLine("debug mode")

                End If
                'If myamount <= 0 Then
                '    If dr.Item("expensesnature").ToString.ToLower = "housing fund" Then

                '        Debug.WriteLine("debug mode")

                '    End If
                '    'Find category
                '    Dim qry3 = From category In Dataset1.Tables(3)
                '               Where category.Item("category") = mycategory And category.Item("categorytype") = dr.Item("expensesnature") And category.Item("myyear") = myyear
                '               Select category

                '    For Each myresult In qry3
                '        myamount = myresult.Item("amount")
                '    Next
                'End If
                If myamount > 0 Then

                    If p.Item("personname") = "FU Kang" Then
                        Debug.Print("debug mode")
                    End If

                    'Dim mymonths = From myrecord In Dataset1.Tables(14)
                    '           Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = dr.Item("expensesnature")
                    '           Select myrecord
                    Dim mymonths = From myrecord In Dataset1.Tables(16)
                               Where myrecord.Item("personjoindatecategoryid") = personjoindatecategoryid And myrecord.Item("expensesnature") = dr.Item("expensesnature")
                               Select myrecord
                    Dim mypct As Double = myamount

                    'change join date
                    Dim enddate As Date? = Nothing
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                    End If
                    If dbtools1.Region = "HK" Then
                        If Not IsDBNull(p.Item("effectivedatestart")) Then
                            joindate = p.Item("effectivedatestart")
                        End If
                        If Not IsDBNull(p.Item("enddate")) Then
                            enddate = p.Item("enddate")
                        End If
                        If Not IsDBNull(p.Item("effectivedateend")) Then
                            enddate = p.Item("effectivedateend")
                        End If
                    Else
                        If Not IsDBNull(p.Item("enddate")) Then
                            enddate = p.Item("enddate")
                        End If
                    End If

                    For Each record In mymonths
                        'If personjoindatecategoryid = 51 Then
                        '    Debug.WriteLine("debug mode")
                        'End If
                        If dr.Item("expensesnature").ToString = "Commission" Then
                            'modified 2013-0-08
                            'If joindate <= CDate(myyear & "-" & record.Item("months") & "-1") Then
                            If ValidJoinDate(joindate, record.Item("months"), myyear) Then
                                'persondata1.commision(record.Item("months")) = myamount
                                PersonSalaryDict(personjoindatecategoryid).commision(record.Item("months")) = myamount
                            End If
                        ElseIf dr.Item("expensesnature").ToString = "Housing Allowance" Then
                            'modified 2013-0-08
                            'If joindate <= CDate(myyear & "-" & record.Item("months") & "-1") Then
                            If ValidJoinDate(joindate, record.Item("months"), myyear) Then
                                'persondata1.housing(record.Item("months")) = myamount
                                PersonSalaryDict(personjoindatecategoryid).housing(record.Item("months")) = myamount
                            End If

                        ElseIf (dr.Item("expensesnature").ToString = "Accident" Or
                                dr.Item("expensesnature").ToString.ToLower = "housing fund" Or
                                dr.Item("expensesnature").ToString = "Maternity" Or
                                dr.Item("expensesnature").ToString = "Medical" Or
                                dr.Item("expensesnature").ToString = "Pension" Or
                                dr.Item("expensesnature").ToString = "Social charges" Or
                                dr.Item("expensesnature").ToString = "Unemployment") And dbtools1.Region = "SZ" Then

                            myamount = mypct * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(record.Item("months"))

                        ElseIf (dr.Item("expensesnature").ToString = "PENSION" Or
                                dr.Item("expensesnature").ToString.ToLower = "HEALTH INSURANCE") And dbtools1.Region = "TW" Then
                            myamount = mypct * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(record.Item("months"))
                            'ElseIf (dr.Item("expensesnature").ToString = "Special Allowance" Or dr.Item("expensesnature") = "Car & Parking Allowance") Then
                            '    If dbtools1.Region = "HK" Then
                            '        myamount = mypct * p.Item("headcount")
                            '    End If
                        End If

                        'If joindate <= CDate(myyear & "-" & record.Item("months") & "-1") Then
                        '    createrecord(stringBuilder1, personexpensesid, myamount, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")))
                        'End If
                        Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
                        Dim check As Boolean = True

                        'If p.Item("enddate").ToString <> "" Then
                        '    check = currentdate < p.Item("enddate")
                        'End If
                        If Not IsNothing(enddate) Then
                            check = currentdate < enddate
                        End If

                        'modified 2013-08-08
                        'If currentdate >= p.Item("joindate") And check Then
                        'If currentdate >= p.Item("joindate") And check Then


                        ' If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
                        If ValidJoinDate(joindate, record.Item("months"), myyear) And check Then
                            If dbtools1.Region = "HK" Then
                                createrecord(stringBuilder1, personexpensesid, myamount * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                            Else
                                Dim myheadcount As Double = 0
                                If expensesnature = "Management Fee" Or expensesnature = "IIT Chief Rep" Then
                                    myheadcount = IIf(p.Item("headcount") = 0, 1, p.Item("headcount"))
                                Else
                                    myheadcount = p.Item("headcount")
                                End If
                                'createrecord(stringBuilder1, personexpensesid, myamount * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                                createrecord(stringBuilder1, personexpensesid, myamount * myheadcount, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & record.Item("months") & "-1")), p.Item("headcount"), p.Item("enddate"))
                            End If

                        End If
                    Next

                ElseIf myamount <= 0 Then
                    If p.Item("personname") = "FU Kang" Then
                        Debug.Print("debug mode")
                    End If
                    Dim qry3 = From category In Dataset1.Tables(3)
                               Where category.Item("category") = mycategory And category.Item("categorytype") = dr.Item("expensesnature") And category.Item("myyear") = myyear
                               Select category

                    For Each myresult In qry3
                        myamount = myresult.Item("amount")
                    Next

                    Dim mypct = myamount


                    'get categorytxmonths
                    Dim mymonths = From myrecord In Dataset1.Tables(14)
                                   Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = expensesnature
                                   Select myrecord


                    Dim enddate As Date? = Nothing
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                    End If
                    If dbtools1.Region = "HK" Then
                        If Not IsDBNull(p.Item("effectivedatestart")) Then
                            joindate = p.Item("effectivedatestart")
                        End If
                        If Not IsDBNull(p.Item("enddate")) Then
                            enddate = p.Item("enddate")
                        End If
                        If Not IsDBNull(p.Item("effectivedateend")) Then
                            enddate = p.Item("effectivedateend")
                        End If
                    Else
                        If Not IsDBNull(p.Item("enddate")) Then
                            enddate = p.Item("enddate")
                        End If
                    End If

                    'Try
                    For Each record In mymonths

                        If (dr.Item("expensesnature").ToString = "Accident" Or
                                dr.Item("expensesnature").ToString.ToLower = "housing fund" Or
                                dr.Item("expensesnature").ToString = "Maternity" Or
                                dr.Item("expensesnature").ToString = "Medical" Or
                                dr.Item("expensesnature").ToString = "Pension" Or
                                dr.Item("expensesnature").ToString = "Social charges" Or
                                dr.Item("expensesnature").ToString = "Unemployment") And dbtools1.Region = "SZ" Then
                            'Nikita
                            'myamount = (mypct * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(record.Item("months"))) / p.Item("headcount")
                            myamount = (mypct * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(record.Item("months"))) / p.Item("headcount")

                        ElseIf (dr.Item("expensesnature").ToString = "PENSION" Or
                            dr.Item("expensesnature").ToString = "HEALTH INSURANCE") And dbtools1.Region = "TW" Then
                            'Nikita
                            'myamount = mypct * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(record.Item("months")) / p.Item("headcount")
                            myamount = (mypct * DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(record.Item("months"))) / p.Item("headcount")



                        End If

                        Dim currentdate As Date = CDate(myyear & "-" & record.Item("months") & "-1")
                        'If currentdate > p.Item("joindate") Then
                        '    createrecord(stringBuilder1, personexpensesid, myamount, myverid, DateFormatyyyyMMdd(currentdate))
                        'End If

                        Dim check As Boolean = True
                        'If p.Item("enddate").ToString <> "" Then
                        '    check = currentdate < p.Item("enddate")
                        'End If
                        If Not IsNothing(enddate) Then
                            check = currentdate < enddate
                        End If
                        'If currentdate >= p.Item("joindate") And check Then
                        'If ValidJoinDate(p.Item("joindate"), record.Item("months"), myyear) And check Then
                        If ValidJoinDate(joindate, record.Item("months"), myyear) And check Then
                            Dim myheadcount As Double = 0
                            If expensesnature = "Management Fee" Or expensesnature = "IIT Chief Rep" Then
                                myheadcount = IIf(p.Item("headcount") = 0, 1, p.Item("headcount"))
                            Else
                                myheadcount = p.Item("headcount")
                            End If

                            'createrecord(stringBuilder1, personexpensesid, myamount * p.Item("headcount"), myverid, DateFormatyyyyMMdd(currentdate), p.Item("headcount"), p.Item("enddate"))
                            createrecord(stringBuilder1, personexpensesid, myamount * myheadcount, myverid, DateFormatyyyyMMdd(currentdate), p.Item("headcount"), p.Item("enddate"))
                            'createrecord(stringBuilder1, personexpensesid, myamount, myverid, DateFormatyyyyMMdd(currentdate))
                        End If

                    Next

                    'Catch ex As Exception
                    '    Debug.WriteLine("debug mode")
                    'End Try
                    'Debug.WriteLine("PersonExpensesid {0} {1} ", personexpensesid, personjoindatecategoryid)
                End If



            End If

        End If
    End Sub

    Private Sub GetParamVar(ByVal Dataset1 As DataSet, ByRef GeneralRate As Double, ByRef GeneralIncrMonth As Integer, ByRef ExpatRate As Double, ByRef ExpatIncrMonth As Integer, ByRef serviceyear10 As Double, ByRef serviceyear15 As Double, ByRef AmountA As Double, ByRef AmountB As Double, ByRef AmountC As Double, ByRef MPF As Double)
        Dim GeneralRateDate As Date
        Dim ExpatRateDate As Date
        Dim ServiceYear10Date As Date
        Dim ServiceYear15Date As Date
        Dim AmountADate As Date
        Dim AmountBDate As Date
        Dim AmountCDate As Date
        Dim MPFDate As Date

        For i = 0 To Dataset1.Tables(4).Rows.Count - 1
            Dim dr = Dataset1.Tables(4).Rows(i)
            Select Case dr.Item(2).ToString
                Case "General Rate"
                    If GeneralRateDate < dr.Item("dvalue") Then
                        GeneralRateDate = dr.Item("dvalue")
                        GeneralRate = CDbl(dr.Item("nvalue").ToString)
                        GeneralIncrMonth = CInt(dr.Item("ivalue").ToString)
                    End If
                Case "Expat Rate"
                    If ExpatRateDate < dr.Item("dvalue") Then
                        ExpatRateDate = dr.Item("dvalue")
                        ExpatRate = CDbl(dr.Item("nvalue").ToString)
                        ExpatIncrMonth = CInt(dr.Item("ivalue").ToString)
                    End If

                Case "10"
                    If ServiceYear10Date < dr.Item("dvalue") Then
                        ServiceYear10Date = dr.Item("dvalue")
                        serviceyear10 = CDbl(dr.Item("nvalue").ToString)
                    End If

                Case "15"
                    If ServiceYear15Date < dr.Item("dvalue") Then
                        ServiceYear15Date = dr.Item("dvalue")
                        serviceyear15 = CDbl(dr.Item("nvalue").ToString)
                    End If
                Case "Amount A"
                    If AmountADate < dr.Item("dvalue") Then
                        AmountADate = dr.Item("dvalue")
                        AmountA = CDbl(dr.Item("nvalue").ToString)
                    End If
                Case "Amount B"
                    If AmountBDate < dr.Item("dvalue") Then
                        AmountBDate = dr.Item("dvalue")
                        AmountB = CDbl(dr.Item("nvalue").ToString)
                    End If
                Case "Amount C"
                    If AmountCDate < dr.Item("dvalue") Then
                        AmountCDate = dr.Item("dvalue")
                        AmountC = CDbl(dr.Item("nvalue").ToString)
                    End If
                Case "MPF"
                    If MPFDate < dr.Item("dvalue") Then
                        MPFDate = dr.Item("dvalue")
                        MPF = CDbl(dr.Item("nvalue").ToString)
                    End If
            End Select
        Next
    End Sub
    Private Sub GetParamVar(ByVal Dataset1 As DataSet, ByRef GeneralRate As Double, ByRef GeneralIncrMonth As Integer, ByRef ExpatRate As Double, ByRef ExpatIncrMonth As Integer, ByRef serviceyear10 As Double, ByRef serviceyear15 As Double, ByRef AmountA As Double, ByRef AmountB As Double, ByRef AmountC As Double, ByRef MPF As Double, ByRef MPFFloor As Double)
        Dim GeneralRateDate As Date
        Dim ExpatRateDate As Date
        Dim ServiceYear10Date As Date
        Dim ServiceYear15Date As Date
        Dim AmountADate As Date
        Dim AmountBDate As Date
        Dim AmountCDate As Date
        Dim MPFDate As Date
        Dim MPFFloorDate As Date

        For i = 0 To Dataset1.Tables(4).Rows.Count - 1
            Dim dr = Dataset1.Tables(4).Rows(i)
            Select Case dr.Item(2).ToString
                Case "General Rate"
                    If GeneralRateDate < dr.Item("dvalue") Then
                        GeneralRateDate = dr.Item("dvalue")
                        GeneralRate = CDbl(dr.Item("nvalue").ToString)
                        GeneralIncrMonth = CInt(dr.Item("ivalue").ToString)
                    End If
                Case "Expat Rate"
                    If ExpatRateDate < dr.Item("dvalue") Then
                        ExpatRateDate = dr.Item("dvalue")
                        ExpatRate = CDbl(dr.Item("nvalue").ToString)
                        ExpatIncrMonth = CInt(dr.Item("ivalue").ToString)
                    End If

                Case "10"
                    If ServiceYear10Date < dr.Item("dvalue") Then
                        ServiceYear10Date = dr.Item("dvalue")
                        serviceyear10 = CDbl(dr.Item("nvalue").ToString)
                    End If

                Case "15"
                    If ServiceYear15Date < dr.Item("dvalue") Then
                        ServiceYear15Date = dr.Item("dvalue")
                        serviceyear15 = CDbl(dr.Item("nvalue").ToString)
                    End If
                Case "Amount A"
                    If AmountADate < dr.Item("dvalue") Then
                        AmountADate = dr.Item("dvalue")
                        AmountA = CDbl(dr.Item("nvalue").ToString)
                    End If
                Case "Amount B"
                    If AmountBDate < dr.Item("dvalue") Then
                        AmountBDate = dr.Item("dvalue")
                        AmountB = CDbl(dr.Item("nvalue").ToString)
                    End If
                Case "Amount C"
                    If AmountCDate < dr.Item("dvalue") Then
                        AmountCDate = dr.Item("dvalue")
                        AmountC = CDbl(dr.Item("nvalue").ToString)
                    End If
                Case "MPF"
                    If MPFDate < dr.Item("dvalue") Then
                        MPFDate = dr.Item("dvalue")
                        MPF = CDbl(dr.Item("nvalue").ToString)
                    End If
                Case "MPF Floor"
                    If MPFFloorDate < dr.Item("dvalue") Then
                        MPFFloorDate = dr.Item("dvalue")
                        MPFFloor = CDbl(dr.Item("nvalue").ToString)
                    End If
            End Select
        Next
    End Sub
    Private Function getGrossSalary(ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal personjoindatecategoryid As Integer) As Double
        Dim grosssalary As Double = 0
        For gross = 1 To 12
            grosssalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(gross)
        Next
        grosssalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).doublepay
        Return grosssalary
    End Function

    Private Function getvalidmonthCA(ByVal joindate As Date) As Double
        Dim validmonth As Double = 12
        If joindate >= CDate(myyear & "-01-01") And joindate <= CDate(myyear & "-12-31") Then
            'new commer
            'validmonth = 12 - (joindate.Month - 1)
            Dim serviceday = CDate(myyear & "-12-31").Subtract(joindate).Days + 1
            Dim servicemonth = 13 - Month(joindate)
            validmonth = 1 / (serviceday / 365 / servicemonth)

        End If
        Return validmonth
    End Function
    'GetValidMonthEffectiveDate
    Private Function getvalidmonthEF(ByVal joindate As Date) As Integer
        Dim validmonth As Integer = 12
        If joindate >= CDate(myyear & "-01-01") And joindate <= CDate(myyear & "-12-31") Then
            validmonth = 12 - (joindate.Month - 1)
        End If
        Return validmonth
    End Function
    Private Function getvalidmonth(ByVal joindate As Date) As Integer
        Dim validmonth As Integer = 12
        If joindate >= CDate(myyear & "-01-01") And joindate <= CDate(myyear & "-12-31") Then
            'new commer
            'validmonth = 12 - (joindate.Month - 1)
            'Dim serviceday = CDate(myyear & "-12-31").Subtract(joindate).Days + 1
            'Dim servicemonth = 13 - Month(joindate)
            'validmonth = serviceday / 365 / servicemonth
        End If
        Return validmonth
    End Function

    Private Function getvalidmonth(ByVal joindate As Date, ByVal enddate As Date) As Integer
        Dim validmonth As Integer = 12
        If (joindate >= CDate(myyear & "-01-01")) Then
            validmonth = enddate.Month - (joindate.Month - 1)
        Else
            'new commer
            If enddate <= CDate(myyear & "-12-31") Then

                validmonth = enddate.Month

            End If
        End If
        Return validmonth
    End Function
    'Private Sub createrecord(ByRef stringBuilder1 As StringBuilder, ByVal personexpensesid As Integer, ByVal amount As Double, ByVal myverid As Integer, ByVal mydate As String)
    '    stringBuilder1.Append(personexpensesid & vbTab)
    '    stringBuilder1.Append(amount & vbTab)
    '    stringBuilder1.Append(myverid & vbTab)
    '    stringBuilder1.Append(mydate & vbCrLf)
    'End Sub

    Private Sub createrecord(ByRef stringBuilder1 As StringBuilder, ByVal personexpensesid As Integer, ByVal amount As Double, ByVal myverid As Integer, ByVal mydate As String, ByVal HeadCount As Double, ByVal enddate As Object)

        stringBuilder1.Append(personexpensesid & vbTab)
        stringBuilder1.Append(amount & vbTab)
        stringBuilder1.Append(myverid & vbTab)
        stringBuilder1.Append(mydate & vbTab)

        'If Not IsDBNull(enddate) Then
        If Not (IsNothing(enddate) Or IsDBNull(enddate)) Then
            If CDate(mydate.Replace("'", "")).Month = enddate.month Then
                If enddate.AddDays(1).month = CDate(mydate.Replace("'", "")).Month Then 'Headcount will reset to zero if not at the end of month
                    HeadCount = 0
                End If
            End If
        End If
        stringBuilder1.Append(HeadCount & vbCrLf)
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            myverid = ComboBox1.SelectedValue
            verid = myverid
        Catch ex As Exception

        End Try


    End Sub

    Private Sub RecruitmentExpenses2(ByVal stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal dr As DataRow, ByVal p As DataRow, ByVal personexpensesid As Integer, ByVal serviceyear As Double, ByVal myverid As Integer, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata)
        Dim amount As Double = 0
        Dim expensesnature As String = "Recruitment Expenses"
        Dim personjoindatecategoryid As Integer = p.Item("personjoindatecategoryid")
        Dim joindate = p.Item("joindate")
        Dim mycategory = p.Item("category")
        'find category
        Dim qry3 = From category In Dataset1.Tables(3)
                   Where category.Item("category") = mycategory And category.Item("categorytype") = expensesnature And category.Item("myyear") = myyear
                   Select category
        'If p.Item(2) = "Nell" Then
        '    Debug.Print("nell")
        'End If
        For Each ct In qry3

            'get amount in tablepersonexpensesdtl
            Debug.WriteLine("apply: {2} For  Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} {4} ", personjoindatecategoryid, 0, expensesnature, personexpensesid, p.Item("othername"))
            Dim qry = From expenses In Dataset1.Tables(5)
                  Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = expensesnature And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                  Select expenses Order By expenses.Item("validdate") Descending

            Dim myamount As Double = 0
            Dim validdate As Date
            Dim sapcc As String = String.Empty
            For Each myresult In qry
                myamount = myresult.Item("amount")
                validdate = myresult.Item("validdate")
                sapcc = myresult.Item("sapcc")
                Exit For
            Next
            If myamount > 0 Then
                'Debug.WriteLine("Personname {0} personexpensesid {1} amount {2} personjoindatecategoryid  {3} expensesnatuer {4}", p.Item("personname"), personexpensesid, myamount, personjoindatecategoryid, expensesnature)
                Dim totalsalary As Double = 0
                Try
                    'Dim ValidMonth As Integer = getvalidmonth(joindate)
                    Dim validmonth As Integer
                    If p.Item("enddate").ToString <> "" Then
                        validmonth = getvalidmonth(joindate, p.Item("enddate"))
                    Else
                        validmonth = getvalidmonth(joindate)
                    End If
                    For i = 1 To 12
                        If joindate <= CDate(myyear & "-" & i & "-1") Then
                            totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                        End If
                    Next
                    totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).doublepay
                    totalsalary *= myamount
                    For M = 1 To 12
                        'If joindate <= CDate(myyear & "-" & M & "-1") Then
                        '    createrecord(stringBuilder1, personexpensesid, totalsalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                        'End If
                        Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                        Dim check As Boolean = True
                        If p.Item("enddate").ToString <> "" Then
                            check = currentdate < p.Item("enddate")
                        End If
                        'If currentdate >= p.Item("joindate") And check Then
                        If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then

                            'createrecord(stringBuilder1, personexpensesid, totalsalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                            createrecord(stringBuilder1, personexpensesid, totalsalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")), p.Item("headcount"), p.Item("enddate"))
                        End If
                    Next
                Catch ex As Exception
                    Debug.WriteLine(ex.Message)

                End Try
            ElseIf myamount <= 0 Then


                'Category listed
                sapcc = DbAdapter1.getsapcc(personjoindatecategoryid)

                'find table param to get sapcc and amount based on personcategoryid
                Dim qrytb4 = From param In Dataset1.Tables(4)
                             Where param.Item("paramname") = sapcc
                             Select param

                For Each result In qrytb4
                    Dim mypct = result.Item("nvalue")
                    myamount = result.Item("ivalue")
                    Dim grosssalary = getGrossSalary(PersonSalaryDict, personjoindatecategoryid)
                    Dim mylist() = result.Item("cvalue").ToString.Split(",")
                    Dim myresult = (grosssalary * mypct) + myamount
                    'Dim mc = From myrecord In Dataset1.Tables(14)
                    '               Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = expensesnature
                    '               Select myrecord

                    For Each mr In mylist
                        Dim mymonth = MonthToIntDict(mr.Trim)
                        Dim currentdate As Date = CDate(myyear & "-" & mymonth & "-1")
                        Dim check As Boolean = True
                        If p.Item("enddate").ToString <> "" Then
                            check = currentdate < p.Item("enddate")
                        End If
                        'If currentdate >= p.Item("joindate") And check Then
                        If ValidJoinDate(p.Item("joindate"), mymonth, myyear) And check Then
                            createrecord(stringBuilder1, personexpensesid, myresult * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(currentdate)), p.Item("headcount"), p.Item("enddate"))
                            'createrecord(stringBuilder1, personexpensesid, myresult, myverid, DateFormatyyyyMMdd(CDate(currentdate)))
                        End If

                    Next
                Next
            End If
        Next
    End Sub

    Private Sub RecruitmentExpenses(ByVal stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal dr As DataRow, ByVal p As DataRow, ByVal personexpensesid As Integer, ByVal serviceyear As Double, ByVal myverid As Integer, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata)
        Dim amount As Double = 0
        Dim expensesnature As String = "Recruitment Expenses"
        Dim personjoindatecategoryid As Integer = p.Item("personjoindatecategoryid")
        Dim joindate = p.Item("joindate")
        Dim mycategory = p.Item("category")
        'find category
        If personjoindatecategoryid = 535 Then
            Debug.Print("Recruitment Expenses Luis")
        End If
        Dim qry3 = From category In Dataset1.Tables(3)
                   Where category.Item("category") = mycategory And category.Item("categorytype") = expensesnature And category.Item("myyear") = myyear
                   Select category

        For Each ct In qry3

            'get amount in tablepersonexpensesdtl
            Debug.WriteLine("apply: {2} For  Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} {4} ", personjoindatecategoryid, 0, expensesnature, personexpensesid, p.Item("othername"))
            Dim qry = From expenses In Dataset1.Tables(5)
                  Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = expensesnature And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                  Select expenses Order By expenses.Item("validdate") Descending

            Dim myamount As Double = 0
            Dim validdate As Date
            Dim sapcc As String = String.Empty
            For Each myresult In qry
                myamount = myresult.Item("amount")
                validdate = myresult.Item("validdate")
                sapcc = myresult.Item("sapcc")
                Exit For
            Next
            If myamount > 0 Then
                'Debug.WriteLine("Personname {0} personexpensesid {1} amount {2} personjoindatecategoryid  {3} expensesnatuer {4}", p.Item("personname"), personexpensesid, myamount, personjoindatecategoryid, expensesnature)
                Dim totalsalary As Double = 0
                Try
                    'Dim ValidMonth As Integer = getvalidmonth(joindate)
                    Dim validmonth As Integer
                    If p.Item("enddate").ToString <> "" Then
                        validmonth = getvalidmonth(joindate, p.Item("enddate"))
                    Else
                        validmonth = getvalidmonth(joindate)
                    End If
                    For i = 1 To 12
                        If joindate <= CDate(myyear & "-" & i & "-1") Then
                            totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                        End If
                    Next
                    totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).doublepay
                    totalsalary *= myamount

                    Dim enddate As Date? = Nothing
                    If Not IsDBNull(p.Item("enddate")) Then
                        enddate = p.Item("enddate")
                    End If
                    If dbtools1.Region = "HK" Then
                        If Not IsDBNull(p.Item("effectivedatestart")) Then
                            joindate = p.Item("effectivedatestart")
                        End If
                        If Not IsDBNull(p.Item("enddate")) Then
                            enddate = p.Item("enddate")
                        End If
                        If Not IsDBNull(p.Item("effectivedateend")) Then
                            enddate = p.Item("effectivedateend")
                        End If
                    Else
                        If Not IsDBNull(p.Item("enddate")) Then
                            enddate = p.Item("enddate")
                        End If
                    End If

                    For M = 1 To 12
                        'If joindate <= CDate(myyear & "-" & M & "-1") Then
                        '    createrecord(stringBuilder1, personexpensesid, totalsalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                        'End If
                        Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                        Dim check As Boolean = True
                        'If p.Item("enddate").ToString <> "" Then
                        '    check = currentdate < p.Item("enddate")
                        'End If
                        If Not IsNothing(enddate) Then
                            check = currentdate < enddate
                        End If
                        'If currentdate >= p.Item("joindate") And check Then
                        'If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then
                        If ValidJoinDate(joindate, M, myyear) And check Then
                            createrecord(stringBuilder1, personexpensesid, totalsalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")), p.Item("headcount"), p.Item("enddate"))

                        End If
                    Next
                Catch ex As Exception
                    Debug.WriteLine(ex.Message)

                End Try
            ElseIf myamount <= 0 Then
                Try
                    'Category listed
                    sapcc = DbAdapter1.getsapcc(personjoindatecategoryid)

                    'find table param to get sapcc and amount based on personcategoryid
                    Dim qrytb4 = From param In Dataset1.Tables(4)
                                 Where param.Item("paramname") = sapcc
                                 Select param
                    If personjoindatecategoryid = 537 Then
                        Debug.WriteLine("debug")
                    End If
                    For Each result In qrytb4
                        Dim mypct = result.Item("nvalue")
                        myamount = result.Item("ivalue")
                        Dim grosssalary = getGrossSalary(PersonSalaryDict, personjoindatecategoryid)
                        Dim mylist() = result.Item("cvalue").ToString.Split(",")
                        Dim myresult = (grosssalary * mypct) + myamount
                        'Dim mc = From myrecord In Dataset1.Tables(14)
                        '               Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = expensesnature
                        '               Select myrecord

                        Dim enddate As Date? = Nothing
                        If Not IsDBNull(p.Item("enddate")) Then
                            enddate = p.Item("enddate")
                        End If
                        If dbtools1.Region = "HK" Then
                            If Not IsDBNull(p.Item("effectivedatestart")) Then
                                joindate = p.Item("effectivedatestart")
                            End If
                            If Not IsDBNull(p.Item("enddate")) Then
                                enddate = p.Item("enddate")
                            End If
                            If Not IsDBNull(p.Item("effectivedateend")) Then
                                enddate = p.Item("effectivedateend")
                            End If
                        Else
                            If Not IsDBNull(p.Item("enddate")) Then
                                enddate = p.Item("enddate")
                            End If
                        End If

                        For Each mr In mylist
                            Dim mymonth = MonthToIntDict(mr.Trim)
                            Dim currentdate As Date = CDate(myyear & "-" & mymonth & "-1")
                            Dim check As Boolean = True
                            'If p.Item("enddate").ToString <> "" Then
                            '    check = currentdate < p.Item("enddate")
                            'End If
                            If Not IsNothing(enddate) Then
                                check = currentdate < enddate
                            End If
                            'If currentdate >= p.Item("joindate") And check Then
                            'If ValidJoinDate(p.Item("joindate"), mymonth, myyear) And check Then
                            If ValidJoinDate(joindate, mymonth, myyear) And check Then
                                'createrecord(stringBuilder1, personexpensesid, myresult * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(currentdate)))
                                'createrecord(stringBuilder1, personexpensesid, PersonSalaryDict(personjoindatecategoryid).salaryDict(mymonth) * mypct * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(currentdate)))
                                createrecord(stringBuilder1, personexpensesid, PersonSalaryDict(personjoindatecategoryid).salaryDict(mymonth) * mypct, myverid, DateFormatyyyyMMdd(CDate(currentdate)), p.Item("headcount"), p.Item("enddate"))
                            End If

                        Next
                    Next
                Catch ex As Exception
                    'personname character varying, othername character varying
                    Throw New ApplicationException("Person Name : " & p.Item("personname") & ", Other Name: " & p.Item("othername") & " has no salary")
                End Try

            End If
        Next
    End Sub


    Private Sub RecruitmentExpenses1(ByVal stringBuilder1 As StringBuilder, ByVal Dataset1 As DataSet, ByVal dr As DataRow, ByVal p As DataRow, ByVal personexpensesid As Integer, ByVal serviceyear As Double, ByVal myverid As Integer, ByVal PersonSalaryDict As Dictionary(Of Integer, persondata), ByVal persondata1 As persondata)
        Dim amount As Double = 0
        Dim expensesnature As String = "Recruitment Expenses"
        Dim personjoindatecategoryid As Integer = p.Item("personjoindatecategoryid")
        Dim joindate = p.Item("joindate")
        Dim mycategory = p.Item("category")
        'find category
        Dim qry3 = From category In Dataset1.Tables(3)
                   Where category.Item("category") = mycategory And category.Item("categorytype") = expensesnature And category.Item("myyear") = myyear
                   Select category

        For Each ct In qry3

            'get amount in tablepersonexpensesdtl
            Debug.WriteLine("apply: {2} For  Personjoindatecategoryid: {0}, expensesnatureid: {1}  - {2},personexpensesid: {3} ", personjoindatecategoryid, 0, expensesnature, personexpensesid)
            Dim qry = From expenses In Dataset1.Tables(5)
                  Where expenses.Item("personjoindatecategoryid") = personjoindatecategoryid And expenses.Item("expensesnature") = expensesnature And expenses.Item("validdate") <= CDate(DateTimePicker1.Value.Year & "/12/31")
                  Select expenses Order By expenses.Item("validdate") Descending

            Dim myamount As Double = 0
            Dim validdate As Date
            Dim sapcc As String = String.Empty
            For Each myresult In qry
                myamount = myresult.Item("amount")
                validdate = myresult.Item("validdate")
                sapcc = myresult.Item("sapcc")
                Exit For
            Next
            If myamount > 0 Then
                'Debug.WriteLine("Personname {0} personexpensesid {1} amount {2} personjoindatecategoryid  {3} expensesnatuer {4}", p.Item("personname"), personexpensesid, myamount, personjoindatecategoryid, expensesnature)
                Dim totalsalary As Double = 0
                Try
                    'Dim ValidMonth As Integer = getvalidmonth(joindate)
                    Dim validmonth As Integer
                    If p.Item("enddate").ToString <> "" Then
                        validmonth = getvalidmonth(joindate, p.Item("enddate"))
                    Else
                        validmonth = getvalidmonth(joindate)
                    End If
                    For i = 1 To 12
                        If joindate <= CDate(myyear & "-" & i & "-1") Then
                            totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).salaryDict(i)
                        End If
                    Next
                    totalsalary += DirectCast(PersonSalaryDict(personjoindatecategoryid), persondata).doublepay
                    totalsalary *= myamount
                    For M = 1 To 12
                        'If joindate <= CDate(myyear & "-" & M & "-1") Then
                        '    createrecord(stringBuilder1, personexpensesid, totalsalary / ValidMonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")))
                        'End If
                        Dim currentdate As Date = CDate(myyear & "-" & M & "-1")
                        Dim check As Boolean = True
                        If p.Item("enddate").ToString <> "" Then
                            check = currentdate < p.Item("enddate")
                        End If
                        'If currentdate >= p.Item("joindate") And check Then
                        If ValidJoinDate(p.Item("joindate"), M, myyear) And check Then
                            createrecord(stringBuilder1, personexpensesid, totalsalary / validmonth, myverid, DateFormatyyyyMMdd(CDate(myyear & "-" & M & "-1")), p.Item("headcount"), p.Item("enddate"))
                        End If
                    Next
                Catch ex As Exception
                    Debug.WriteLine(ex.Message)

                End Try
            ElseIf myamount <= 0 Then


                'Category listed
                sapcc = DbAdapter1.getsapcc(personjoindatecategoryid)

                'find table param to get sapcc and amount based on personcategoryid
                Dim qrytb4 = From param In Dataset1.Tables(4)
                             Where param.Item("paramname") = sapcc
                             Select param

                For Each result In qrytb4
                    Dim mypct = result.Item("nvalue")
                    myamount = result.Item("ivalue")
                    Dim grosssalary = getGrossSalary(PersonSalaryDict, personjoindatecategoryid)
                    Dim mylist() = result.Item("cvalue").ToString.Split(",")
                    Dim myresult = (grosssalary * mypct) + myamount
                    'Dim mc = From myrecord In Dataset1.Tables(14)
                    '               Where myrecord.Item("category") = mycategory And myrecord.Item("categorytype") = expensesnature
                    '               Select myrecord

                    For Each mr In mylist
                        Dim mymonth = MonthToIntDict(mr.Trim)
                        Dim currentdate As Date = CDate(myyear & "-" & mymonth & "-1")
                        Dim check As Boolean = True
                        If p.Item("enddate").ToString <> "" Then
                            check = currentdate < p.Item("enddate")
                        End If
                        'If currentdate >= p.Item("joindate") And check Then
                        If ValidJoinDate(p.Item("joindate"), mymonth, myyear) And check Then
                            createrecord(stringBuilder1, personexpensesid, myresult * p.Item("headcount"), myverid, DateFormatyyyyMMdd(CDate(currentdate)), p.Item("headcount"), p.Item("enddate"))
                            'createrecord(stringBuilder1, personexpensesid, myresult, myverid, DateFormatyyyyMMdd(CDate(currentdate)))
                        End If

                    Next
                Next
            End If
        Next
    End Sub

    Private Function getvalidmonthnewcommer(ByVal joindate As Date) As Integer
        Dim validmonth As Integer = 12
        If joindate >= CDate(myyear & "-01-01") And joindate <= CDate(myyear & "-12-31") Then
            'new commer
            validmonth = 12 - (joindate.Month - 1)
        End If
        Return validmonth
    End Function

    Private Function getvalidmonthnewcommer(ByVal joindate As Date, ByVal enddate As Date) As Integer
        Dim validmonth As Integer = 12
        If joindate >= CDate(myyear & "-01-01") And joindate <= CDate(myyear & "-12-31") Then
            'new commer
            validmonth = enddate.Month - (joindate.Month)
            If validmonth = 0 Then validmonth = 12
        End If
        Return validmonth
    End Function

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        loadcombobox()
    End Sub

    Private Function ValidJoinDate(ByVal joindate As Date, ByVal M As Integer, ByVal myyear As Integer) As Boolean
        Dim myret = True
        If Year(joindate) = myyear Then
            If M < Month(joindate) Then
                myret = False
            End If
        End If

        Return myret

    End Function
    Private Function ValidateJoinDate(ByVal joindate As Date, ByVal enddate As Object, ByVal amount As Double) As Double
        Dim myret = True
        If Not IsDBNull(enddate) Then
            amount = 0
            Return amount
        End If
        If Year(joindate) = myyear Then
            'amount = amount * (CDate(myyear & "-12-31").Subtract(joindate).Days + 1) / 365 / (13 - joindate.Month)
            amount = amount / (13 - joindate.Month)
        Else
            amount = amount / 12
        End If

        Return amount

    End Function
End Class

Public Class persondata
    Public Property salaryDict As New Dictionary(Of Integer, Double)
    Public Property doublepay As Double
    Public Property commision As New Dictionary(Of Integer, Double)
    Public Property housing As New Dictionary(Of Integer, Double)
    Public Property bonus As Double

    Sub New()
        doublepay = 0
        bonus = 0
        For i = 1 To 12
            commision(i) = 0
            housing(i) = 0
        Next
    End Sub

End Class

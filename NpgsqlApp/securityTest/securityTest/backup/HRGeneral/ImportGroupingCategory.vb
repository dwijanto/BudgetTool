Imports System.ComponentModel
Imports HR.HelperClass
Imports System.IO
Imports System.Text

Imports DJLib.Dbtools

Public Class ImportGroupingCategory

    Private WithEvents BackgroundWorker1 As New BackgroundWorker
    Dim FileName As String = String.Empty
    Dim Status As Boolean = False
    Dim Dataset1 As DataSet
    Dim ConnectionString As String = dbtools1.getConnectionString

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not (BackgroundWorker1.IsBusy) Then
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "Txt files (*.txt)|*.txt|All files (*.*)|*.*"
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

    End Sub

    Private Function ImportData(ByVal FileName As String, ByRef errMsg As String) As Boolean
        Dim myReturn As Boolean = False
        Dim i As Integer = 0
        Dim sb As New StringBuilder
        Dim stopwatch As New Stopwatch
        stopwatch.Start()
        Try

            Dim objreader As New StreamReader(FileName)
            Dim sline As String = ""
            Dim allLine As String = ""
            'Dim arrtext As New ArrayList()

            Do
                'allLine = objreader.ReadToEnd
                sline = objreader.ReadLine()
                Dim arrtext As New ArrayList()
                If Not sline Is Nothing And i > 0 Then
                    arrtext.Add(sline.Split(vbTab))
                    sb.Append(CType(arrtext(0), String())(0).ToString & vbTab)
                    sb.Append(CType(arrtext(0), String())(1).ToString & vbTab)
                    sb.Append(CType(arrtext(0), String())(2).ToString & vbCrLf)


                    'sb.Append(CType(arrtext(0), String())(0).ToString & vbTab & CType(arrtext(0), String())(1).ToString & vbTab & CType(arrtext(0), String())(2).ToString & vbTab & vbCrLf)
                End If
                i += 1
            Loop Until sline Is Nothing
            objreader.Close()

            'For Each sline In arrtext
            '    'Debug.WriteLine(sline)

            'Next

            'If arrtext.Count > 0 Then

            'End If

            Dim sqlstr = "delete from groupingtable;select setval('groupingtable_groupingtableid_seq',1,false);copy groupingtable(category,sapaccname,sapaccount) from stdin;"
            BackgroundWorker1.ReportProgress(2, "Copy To Db (GroupingCategory)")
            If sb.ToString <> "" Then
                errMsg = dbtools1.copy(sqlstr, sb.ToString, myReturn)
                BackgroundWorker1.ReportProgress(2, "Copy To Db.")
            Else
                BackgroundWorker1.ReportProgress(2, "Nothing to Copy.")
                myReturn = True
            End If
            Stopwatch.Stop()
            BackgroundWorker1.ReportProgress(3, "Elapsed Time: " & Format(stopwatch.Elapsed.Minutes, "00") & ":" & Format(stopwatch.Elapsed.Seconds, "00") & "." & stopwatch.Elapsed.Milliseconds.ToString)

            'BackgroundWorker1.ReportProgress(3, "")


            myReturn = True
        Catch ex As Exception
            errMsg = ex.Message
        End Try

        Return myReturn
    End Function

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

    End Sub
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub
    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class
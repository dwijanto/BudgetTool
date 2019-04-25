Imports System.Threading

Delegate Sub ReportProgress(ByVal myint As Integer, ByVal myString As String)
Delegate Sub EndThread(ByVal myint As Integer, ByVal myString As String)
Public Class Threading2
    Dim t As Thread

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        't = New Thread(New ThreadStart(AddressOf WorkerThread))
        't.Start()


        Dim aisdone As New AutoResetEvent(False)
        Dim bisdone As New AutoResetEvent(False)
        'Dim cisdone As New AutoResetEvent(False)
        For i = 0 To 4
            'ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf WorkerThread), i)
            ThreadPool.QueueUserWorkItem(AddressOf WorkerThread, i)
        Next i

        'ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf WorkerThread), cisdone)

        Dim threadcount, threadports As Integer
        ThreadPool.GetAvailableThreads(threadcount, threadports)
        'aisdone.WaitOne()
        'bisdone.WaitOne()
        'cisdone.WaitOne()
    End Sub

    Private Sub WorkerThread(ByVal ThreadStateData As Object)
        For i = 0 To 1000
            Thread.Sleep(1)
            myProgress(ThreadStateData, i & "Hi")
            Application.DoEvents()
        Next
        EndWorker(ThreadStateData, ".Done")

    End Sub

    Private Sub myProgress(ByVal myint As Integer, ByVal mystring As String)
        If Me.TextBox1.InvokeRequired Then
            Me.Invoke(New ReportProgress(AddressOf myProgress), New Object() {myint, mystring})
        Else
            Select Case myint
                Case 0
                    TextBox1.Text = myint & " " & mystring
                Case 1
                    TextBox2.Text = myint & " " & mystring
                Case 2
                    TextBox3.Text = myint & " " & mystring
                Case 3
                    TextBox4.Text = myint & " " & mystring
                Case 4
                    TextBox5.Text = myint & " " & mystring
            End Select

        End If

    End Sub

    Private Sub EndWorker(ByVal myint As Integer, ByVal myString As String)
        If Me.TextBox1.InvokeRequired Then
            Me.Invoke(New EndThread(AddressOf EndWorker), New Object() {myint, myString})
        Else

            Select Case myint
                Case 0
                    TextBox1.Text &= ". Done"
                Case 1
                    TextBox2.Text &= ". Done"
                Case 2
                    TextBox3.Text &= ". Done"
                Case 3
                    TextBox4.Text &= ". Done"
                Case 4
                    TextBox5.Text &= ". Done"
            End Select
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'WorkerThread()
    End Sub
End Class
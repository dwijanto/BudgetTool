Imports System.ComponentModel
Imports System.Net
Imports System.Net.Sockets

Delegate Sub UpdateForm(ByVal mystring As String)
Public Class Threading1
    Dim WithEvents myWorker As New BackgroundWorker

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        myWorker.WorkerReportsProgress = True
        If Not myWorker.IsBusy Then
            myWorker.RunWorkerAsync()
        End If
    End Sub

    Private Sub myWorker_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles myWorker.Disposed

    End Sub

    Private Sub myWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles myWorker.DoWork
        'Me.Invoke(New UpdateForm(AddressOf UpdateForm), New Object() {"One"})
        'Threading.Thread.Sleep(New TimeSpan(0, 0, 0, 5, 0))
        myWorker.ReportProgress(1, "hello")
        Threading.Thread.Sleep(New TimeSpan(0, 0, 0, 2, 0))

        'MsgBox("hello")
        For i = 0 To 1000
            Threading.Thread.Sleep(1)
            myWorker.ReportProgress(1, i.ToString)
            'Me.Invoke(New UpdateForm(AddressOf UpdateForm), New Object() {i.ToString})
        Next

        'Threading.Thread.Sleep(New TimeSpan(0, 0, 0, 5, 0))
        myWorker.ReportProgress(2, "juga")
    End Sub

    Private Sub myWorker_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles myWorker.ProgressChanged
        Select Case e.ProgressPercentage
            Case 1
                Me.Invoke(New UpdateForm(AddressOf UpdateForm), New Object() {e.UserState})
            Case 2
                Me.Invoke(New UpdateForm(AddressOf UpdateForm), New Object() {e.UserState})
        End Select

    End Sub

    Private Sub myWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles myWorker.RunWorkerCompleted
        MsgBox("Complete")
    End Sub

    Private Sub UpdateForm(ByVal mystring As String)
        TextBox1.Text = mystring
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        FormMenu.setBubbleMessage("Message Title", "Message Description")
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If (Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'If (args.Length = 0) Then
        '    ' Print a message and exit.
        '    Console.WriteLine("You must specify the name of a host computer.")
        '    End
        'End If
        ' Start the asynchronous request for DNS information.
        'Dim result As IAsyncResult = Dns.BeginGetHostEntry(args(0), Nothing, Nothing)
        Dim result As IAsyncResult = Dns.BeginGetHostEntry("www.yahoo.com", Nothing, Nothing)
        Dim result1 As IAsyncResult = Dns.BeginGetHostAddresses("www.yahoo.com", Nothing, Nothing)

        Debug.Write("Processing request for information")
        ' Wait until the operation completes.
        While result.IsCompleted <> True AndAlso result1.IsCompleted <> True
            Debug.Write(".")
        End While
        Debug.WriteLine("")
        result.AsyncWaitHandle.WaitOne()
        result1.AsyncWaitHandle.WaitOne()
        ' The operation completed. Process the results.
        Try
            ' Get the results.
            Dim abc As IPAddress() = Dns.EndGetHostAddresses(result1)
            For Each a As Object In abc
                Debug.WriteLine("IpAddress {0}", a)
            Next
            Dim host As IPHostEntry = Dns.EndGetHostEntry(result)
            Dim aliases() As String = host.Aliases
            Dim addresses() As IPAddress = host.AddressList
            Dim i As Integer
            If aliases.Length > 0 Then
                'Console.WriteLine("Aliases")
                Debug.WriteLine("Aliases")
                For i = 0 To aliases.Length - 1
                    'Console.WriteLine("{0}", aliases(i))
                    Debug.WriteLine("{0}", aliases(i))
                Next i
            End If
            If addresses.Length > 0 Then
                Debug.WriteLine("Addresses")
                For i = 0 To addresses.Length - 1
                    Debug.WriteLine("{0}", addresses(i).ToString())
                Next i
            End If
        Catch ex As SocketException
            Debug.WriteLine("An exception occurred while processing the request: {0}" _
              , ex.Message)
        End Try

    End Sub
End Class
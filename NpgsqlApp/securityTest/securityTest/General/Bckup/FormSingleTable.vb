Imports DJLib.AppConfig
Imports DJLib.ExcelStuff
Public Class FormSingleTable
    Dim sqlstr As String = String.Empty
    Dim DataSet As DataSet

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        sqlstr = "Select * from cmmfpricemonth"
        DataSet = New DataSet
        Dim stopwatch As New Stopwatch
        stopwatch.Start()
        'If ExportToExcelFullPath("c:\00\test.xlsx", sqlstr, dbTools) Then
        '    stopwatch.Stop()
        '    If MsgBox(stopwatch.Elapsed.ToString & "File name: " & "c:\00\test.xlsx" & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
        '        Process.Start("c:\00\test.xlsx")
        '    End If
        'End If
        If dbTools.getDataSet(sqlstr, DataSet) Then
            'DJLib.ExcelStuff.ExportToExcelAskDirectory("test.xlsx", DataSet)
            ExportToExcel("test.xlsx", DataSet)
        End If

    End Sub
End Class
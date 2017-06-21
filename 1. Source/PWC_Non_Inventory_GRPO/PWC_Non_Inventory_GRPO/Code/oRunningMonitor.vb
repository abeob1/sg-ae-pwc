Public Class oRunningMonitor
    Public Sub UpdateMonitor(TransType As String, HeaderID As Integer)
        Dim cn As New Connection
        Dim dt As DataTable = cn.Integration_RunQuery("insert into RunningMonitor(TransType,HeaderID) values ('" + TransType + "'," + CStr(HeaderID) + ") ")
    End Sub

    Public Function GetLastRunning() As String
        Exit Function
        Dim str As String = ""
        Dim cn As New Connection
        Dim dt As DataTable = cn.Integration_RunQuery("select * from RunningMonitor with(Nolock) where id =(select max(ID) from RunningMonitor) ")
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0).Item("TransType").ToString + " - " + dt.Rows(0).Item("HeaderID").ToString
        End If
        Return str
    End Function
End Class

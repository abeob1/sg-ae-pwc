Public Class frmMapping

    Private Sub btnStop_Click(sender As System.Object, e As System.EventArgs) Handles btnStop.Click
        Me.Close()
    End Sub

    Private Sub frmMapping_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim cn As New Connection
        Dim dt As DataTable = cn.Integration_RunQuery("exec sp_Mapping_Load")
        grMonitor.DataSource = dt
    End Sub

    Private Sub btnStart_Click(sender As System.Object, e As System.EventArgs) Handles btnStart.Click
        Dim cn As New Connection
        Dim strQuery As String = ""

        For i As Integer = 0 To grMonitor.RowCount - 1
            strQuery = "exec sp_Mapping_Update '" + grMonitor.Rows(i).Cells("Code").Value.ToString + "','"
            strQuery = strQuery + grMonitor.Rows(i).Cells("Value").Value.ToString + "','"
            strQuery = strQuery + grMonitor.Rows(i).Cells("Description").Value.ToString + "'"
            cn.Integration_RunQuery(strQuery)
        Next
        MessageBox.Show("Update completed!")
    End Sub
End Class
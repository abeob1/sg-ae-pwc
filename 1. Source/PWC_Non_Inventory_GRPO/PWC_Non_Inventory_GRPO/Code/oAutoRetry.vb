Public Class oAutoRetry
    Public Sub RetryAll()
        Try
            Dim cn As New Connection

            Dim str As String = ""
            Dim strEnd As String = " set ReceiveDate=null, ErrMsg=null where ReceiveDate is not null and isnull(errMsg,'')<>''"
            str = "Update POHeader" + strEnd
            cn.Integration_RunQuery(str)
            str = "Update GRPOHeader" + strEnd
            cn.Integration_RunQuery(str)
            str = "Update GoodsReturnHeader" + strEnd
            cn.Integration_RunQuery(str)
            str = "Update TransferHeader" + strEnd
            cn.Integration_RunQuery(str)
            str = "Update InvoiceHeader" + strEnd
            cn.Integration_RunQuery(str)
            str = "Update SendEmailPOHeader" + strEnd
            cn.Integration_RunQuery(str)
        Catch ex As Exception
            Functions.WriteLog("RetryAll:" + ex.ToString())
        End Try
    End Sub
End Class

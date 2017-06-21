Public Class oMapping
    Public Sub LoadMapping()
        Dim cn As New Connection
        Dim dt As DataTable = cn.Integration_RunQuery("sp_Mapping_Load")
        For Each dr As DataRow In dt.Rows
            Select Case dr("Code").ToString()
                Case "ToEmail"
                    PublicVariable.ToEmail = dr("Value").ToString
                Case "ToEmailName"
                    PublicVariable.ToEmailName = dr("Value").ToString
                Case "smtpServer"
                    PublicVariable.smtpServer = dr("Value").ToString
                Case "smtpPort"
                    PublicVariable.smtpPort = dr("Value").ToString
                Case "smtpSenderEmail"
                    PublicVariable.smtpSenderEmail = dr("Value").ToString
                Case "smtpPwd"
                    PublicVariable.smtpPwd = dr("Value").ToString
                Case "EmailSub"
                    PublicVariable.EmailSub = dr("Value").ToString

                Case "ToErrorEmail"
                    PublicVariable.ToErrorEmail = dr("Value").ToString

                Case "pmCash"
                    PublicVariable.pmCashAcct = dr("Value").ToString
                Case "pmTransfer"
                    PublicVariable.pmTransferAcct = dr("Value").ToString
                Case "GST"
                    PublicVariable.GSTCode = dr("Value").ToString
                Case "NGST"
                    PublicVariable.NonGSTCode = dr("Value").ToString
                Case "TransitWhs"
                    PublicVariable.TransitWhs = dr("Value").ToString
            End Select
        Next
    End Sub
End Class

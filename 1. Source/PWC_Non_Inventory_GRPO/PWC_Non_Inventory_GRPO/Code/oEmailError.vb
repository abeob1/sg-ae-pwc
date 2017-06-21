Imports System.Web.Mail
Imports System.Net
Imports System.Globalization
Imports System.IO

Public Class oEmailError
    Public Sub SendErrorEmail()
        Try
            Dim cn As New Connection
            Dim dt As DataTable = cn.Integration_RunQuery("IntegrationMonitor_Error")

            Dim ret As String = ""
            ret = SendMailByDS(dt, PublicVariable.ToErrorEmail, "Integration Error Notice", Application.StartupPath + "\EmailError.htm")
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
        End Try
    End Sub
    Public Function SendMailByDS(dt As DataTable, ToEmailList As String, Subject As String, TemplatePath As String) As String

        Dim l_SenderEmail As String = PublicVariable.smtpSenderEmail
        If ToEmailList.Trim().Equals(String.Empty) Then
            ToEmailList = l_SenderEmail
        End If

        Dim mail As New MailMessage()
        mail.To = ToEmailList
        mail.From = l_SenderEmail
        mail.Subject = Subject
        mail.Body = GetTemplateforDS(dt, TemplatePath)
        mail.BodyFormat = MailFormat.Html
        mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1") 'basic authentication
        mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", l_SenderEmail) 'set your username here
        mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", PublicVariable.smtpPwd) 'set your password here

        Try
            SmtpMail.SmtpServer = PublicVariable.smtpServer 'your real server goes here
            SmtpMail.Send(mail)

        Catch ex As System.Web.HttpException
            Return ex.Message
        End Try

        Return ""
    End Function
    Private Function GetTemplateforDS(dt As DataTable, TemplatePath As String) As String
        Dim l_Rs As String = ""
        Try
            Dim l_PathTemplate As String = String.Empty

            If TemplatePath.Trim().Equals(String.Empty) Then
                l_Rs = String.Format("Template is empty")
            Else
                l_Rs = File.ReadAllText(TemplatePath)

                '-----------line-----------------
                Dim str As String = ""
                For i As Integer = 0 To dt.Rows.Count - 1
                    str = str & "<tr>"
                    str = str & "<td style=""border: thin solid #008080;""><@Type" & i.ToString() & "></td>"
                    str = str & "<td style=""border: thin solid #008080;""><@ID" & i.ToString() & "></td>"
                    str = str & "<td style=""border: thin solid #008080;""><@ErrMsg" & i.ToString() & "></td>"
                    str = str & "</tr>"
                Next
                l_Rs = l_Rs.Replace("<@ITEMLINEHERE>", str)
                Dim j As Integer = 0
                For Each dr1 As DataRow In dt.Rows
                    l_Rs = l_Rs.Replace("<@Type" & j.ToString() & ">", dr1("Type").ToString())
                    l_Rs = l_Rs.Replace("<@ID" & j.ToString() & ">", dr1("ID").ToString())
                    l_Rs = l_Rs.Replace("<@ErrMsg" & j.ToString() & ">", dr1("ErrMsg"))
                    j += 1
                Next
            End If
        Catch ex As Exception
            Functions.WriteLog(ex.ToString())
            Return ""
        End Try
        Return l_Rs
    End Function
End Class

Imports System.Web.Mail
Imports System.Net
Imports System.Globalization
Imports System.IO
'Imports System.Net.Mail

Public Class oEmailWeb
    Public Sub SendPOEmail()
        Try
            Dim cn As New Connection
            Dim dt As DataTable = cn.Integration_RunQuery("sp_SendEmailPO_LoadHeader")
            For Each dr As DataRow In dt.Rows
                Dim ds As New DataSet
                Dim dtRow As DataTable = cn.Integration_RunQuery("sp_SendEmailPO_LoadLineByID " + dr("ID").ToString)

                Dim dtHeader As DataTable = dt.Clone
                dtHeader.ImportRow(dr)

                ds.Tables.Add(dtHeader.Copy)
                ds.Tables.Add(dtRow.Copy)
                Dim ret As String = ""
                ret = SendMailByDS(ds, PublicVariable.ToEmail, PublicVariable.EmailSub, Application.StartupPath + "\EmailTemplate.htm")
                cn.Integration_RunQuery("sp_SendEmailPO_UpdateReceived " + dr("ID").ToString + ",'" + ret + "'")
            Next
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
        End Try
    End Sub
    Public Function SendMailByDS(ds As DataSet, ToEmailList As String, Subject As String, TemplatePath As String) As String

        Dim l_SenderEmail As String = PublicVariable.smtpSenderEmail
        If ToEmailList.Trim().Equals(String.Empty) Then
            ToEmailList = l_SenderEmail
        End If

        Dim mail As New MailMessage()
        mail.To = ToEmailList
        mail.From = l_SenderEmail
        mail.Subject = Subject
        mail.Body = GetTemplateforDS(ds, TemplatePath)
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



        '-----------------
        'Dim msg As New System.Net.Mail.MailMessage()
        'msg.From = New MailAddress(l_SenderEmail)
        'msg.[To].Add(ToEmailList)
        'msg.Subject = Subject
        'msg.Body = GetTemplateforDS(ds, TemplatePath)
        'msg.IsBodyHtml = True
        'Try
        '    Dim client As New SmtpClient(PublicVariable.smtpServer, PublicVariable.smtpPort)
        '    client.EnableSsl = True
        '    'client.Timeout = 0
        '    client.Credentials = New NetworkCredential(l_SenderEmail, PublicVariable.smtpPwd)
        '    client.Send(msg)
        'Catch ex As SmtpException
        '    Return ex.Message
        'End Try

        Return ""
    End Function
    Private Function GetTemplateforDS(Ds As DataSet, TemplatePath As String) As String
        Dim l_Rs As String = ""
        Try
            Dim l_PathTemplate As String = String.Empty

            If TemplatePath.Trim().Equals(String.Empty) Then
                l_Rs = String.Format("Template is empty")
            Else
                If Ds.Tables.Count < 2 OrElse Ds.Tables(0).Rows.Count < 1 OrElse Ds.Tables(1).Rows.Count < 1 Then
                    l_Rs = String.Format("No Data")
                Else
                    l_Rs = File.ReadAllText(TemplatePath)

                    '-----------header-----------------
                    Dim dr As DataRow = Ds.Tables(0).Rows(0)
                    l_Rs = l_Rs.Replace("<@Code>", dr("CardCode").ToString())
                    l_Rs = l_Rs.Replace("<@Name>", dr("CardName").ToString())
                    Dim ivC As CultureInfo = New System.Globalization.CultureInfo("es-US")
                    l_Rs = l_Rs.Replace("<@Date>", [String].Format("{0:MM/dd/yyyy}", Convert.ToDateTime(dr("DocDate"), ivC)))

                    '-----------line-----------------
                    Dim str As String = ""
                    For i As Integer = 0 To Ds.Tables(1).Rows.Count - 1
                        str = str & "<tr>"
                        str = str & "<td style=""border: thin solid #008080;""><@ItemCode" & i.ToString() & "></td>"
                        str = str & "<td style=""border: thin solid #008080;""><@ItemName" & i.ToString() & "></td>"
                        str = str & "<td align=""right"" style=""border: thin solid #008080;""><@Quantity" & i.ToString() & "></td>"
                        str = str & "<td align=""right"" style=""border: thin solid #008080;""><@Price" & i.ToString() & "></td>"
                        str = str & "<td align=""right"" style=""border: thin solid #008080;""><@LineTotal" & i.ToString() & "></td>"
                        str = str & "<td align=""right"" style=""border: thin solid #008080;""><@Location" & i.ToString() & "></td>"
                        str = str & "</tr>"
                    Next
                    l_Rs = l_Rs.Replace("<@ITEMLINEHERE>", str)
                    Dim j As Integer = 0
                    For Each dr1 As DataRow In Ds.Tables(1).Rows
                        l_Rs = l_Rs.Replace("<@ItemCode" & j.ToString() & ">", dr1("ItemCode").ToString())
                        l_Rs = l_Rs.Replace("<@ItemName" & j.ToString() & ">", dr1("Dscription").ToString())
                        l_Rs = l_Rs.Replace("<@Quantity" & j.ToString() & ">", String.Format("{0:n0}", dr1("Quantity")))
                        l_Rs = l_Rs.Replace("<@Price" & j.ToString() & ">", String.Format("{0:n2}", dr1("Price")))
                        l_Rs = l_Rs.Replace("<@LineTotal" & j.ToString() & ">", String.Format("{0:n2}", dr1("LineTotal")))
                        l_Rs = l_Rs.Replace("<@Location" & j.ToString() & ">", dr1("Location").ToString())
                        j += 1
                    Next
                    '-----------footer-----------------
                    l_Rs = l_Rs.Replace("<@AmountDue>", String.Format("{0:n2}", dr("GrandTotal")))
                    l_Rs = l_Rs.Replace("<@SubTotal>", String.Format("{0:n2}", dr("SubTotal")))
                    l_Rs = l_Rs.Replace("<@Tax>", String.Format("{0:n2}", dr("GST")))
                End If
            End If
        Catch ex As Exception
            Functions.WriteLog(ex.ToString())
            Return ""
        End Try
        Return l_Rs
    End Function
End Class

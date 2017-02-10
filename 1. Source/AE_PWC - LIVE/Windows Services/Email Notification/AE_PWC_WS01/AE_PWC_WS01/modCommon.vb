Option Explicit On
Imports System.Xml
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Net.Mail
Imports System.Configuration
Imports System.Data.Sql
Imports System.Data.SqlClient



Module modCommon

    Public DateConversion As New System.Globalization.CultureInfo("fr-FR", True)




    ''Public Sub UpdateXML(ByVal oDICompany As SAPbobsCOM.Company, ByVal oDITargetComp As SAPbobsCOM.Company, _
    ''                         ByVal sNode As String, ByVal sTblName As String, ByVal sField1 As String, ByVal sField2 As String, _
    ''                         ByVal bIsNumeric As Boolean, ByRef oXMLDoc As XmlDocument, ByRef sXMLFile As String)

    ''    Dim oNode As XmlNode
    ''    Dim sFuncName As String = String.Empty
    ''    Dim sSQL As String = String.Empty
    ''    Dim oRs As SAPbobsCOM.Recordset
    ''    Dim iCode As Integer
    ''    Dim sCode As String = String.Empty

    ''    Try
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating " & sField1 & " in XML file..", sFuncName)
    ''        oRs = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    ''        oNode = oXMLDoc.SelectSingleNode(sNode)

    ''        If Not IsNothing(oNode) Then
    ''            If Not oNode.InnerText = String.Empty Then
    ''                If bIsNumeric Then
    ''                    iCode = CInt(oNode.InnerText)

    ''                    If sTblName = "OLGT" Then
    ''                        If CInt(oNode.InnerText) = 0 Then iCode = 1
    ''                    End If


    ''                    sSQL = " SELECT " & sField1 & " from  [" & oDITargetComp.CompanyDB.ToString & "].[dbo]." & sTblName & _
    ''                           " WHERE " & sField2 & " in (select " & sField2 & " from " & sTblName & " WHERE " & sField1 & "=" & iCode & ")"
    ''                Else
    ''                    sCode = oNode.InnerText
    ''                    sSQL = " SELECT " & sField1 & " from  [" & oDITargetComp.CompanyDB.ToString & "].[dbo]." & sTblName & _
    ''                           " WHERE " & sField2 & " in (select " & sField2 & " from " & sTblName & " WHERE " & sField1 & "='" & sCode & "')"
    ''                End If

    ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL Query" & sSQL, sFuncName)
    ''                oRs.DoQuery(sSQL)
    ''                If Not oRs.EoF Then
    ''                    oNode.InnerText = oRs.Fields.Item(0).Value
    ''                Else
    ''                    oNode.ParentNode.RemoveChild(oNode)
    ''                    oXMLDoc.Save(sXMLFile)
    ''                End If
    ''                oXMLDoc.Save(sXMLFile)
    ''            Else
    ''                oNode.ParentNode.RemoveChild(oNode)
    ''                oXMLDoc.Save(sXMLFile)
    ''            End If
    ''        End If

    ''    Catch ex As Exception

    ''    End Try

    ''End Sub

    ''Public Function GetDate(ByVal sDate As String, ByRef oCompany As SAPbobsCOM.Company) As String

    ''    Dim dateValue As DateTime
    ''    Dim DateString As String = String.Empty
    ''    Dim sSQL As String = String.Empty
    ''    Dim oRs As SAPbobsCOM.Recordset
    ''    Dim sDatesep As String

    ''    oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    ''    sSQL = "SELECT DateFormat,DateSep FROM OADM"

    ''    oRs.DoQuery(sSQL)

    ''    If Not oRs.EoF Then
    ''        sDatesep = oRs.Fields.Item("DateSep").Value

    ''        Select Case oRs.Fields.Item("DateFormat").Value
    ''            Case 0
    ''                If Date.TryParseExact(sDate, "dd" & sDatesep & "MM" & sDatesep & "yy", _
    ''                   New CultureInfo("en-US"), _
    ''                   DateTimeStyles.None, _
    ''                   dateValue) Then
    ''                    DateString = dateValue.ToString("yyyyMMdd")
    ''                End If
    ''            Case 1
    ''                If Date.TryParseExact(sDate, "dd" & sDatesep & "MM" & sDatesep & "yyyy", _
    ''                   New CultureInfo("en-US"), _
    ''                   DateTimeStyles.None, _
    ''                   dateValue) Then
    ''                    DateString = dateValue.ToString("yyyyMMdd")
    ''                End If
    ''            Case 2
    ''                If Date.TryParseExact(sDate, "MM" & sDatesep & "dd" & sDatesep & "yy", _
    ''                    New CultureInfo("en-US"), _
    ''                    DateTimeStyles.None, _
    ''                    dateValue) Then
    ''                    DateString = dateValue.ToString("yyyyMMdd")
    ''                End If
    ''            Case 3
    ''                If Date.TryParseExact(sDate, "MM" & sDatesep & "dd" & sDatesep & "yyyy", _
    ''                    New CultureInfo("en-US"), _
    ''                    DateTimeStyles.None, _
    ''                    dateValue) Then
    ''                    DateString = dateValue.ToString("yyyyMMdd")
    ''                End If
    ''            Case 4
    ''                If Date.TryParseExact(sDate, "yyyy" & sDatesep & "MM" & sDatesep & "dd", _
    ''                    New CultureInfo("en-US"), _
    ''                    DateTimeStyles.None, _
    ''                    dateValue) Then
    ''                    DateString = dateValue.ToString("yyyyMMdd")
    ''                End If
    ''            Case 5
    ''                If Date.TryParseExact(sDate, "dd" & sDatesep & "MMMM" & sDatesep & "yyyy", _
    ''                    New CultureInfo("en-US"), _
    ''                    DateTimeStyles.None, _
    ''                    dateValue) Then
    ''                    DateString = dateValue.ToString("yyyyMMdd")
    ''                End If
    ''            Case 6
    ''                If Date.TryParseExact(sDate, "yy" & sDatesep & "MM" & sDatesep & "dd", _
    ''                    New CultureInfo("en-US"), _
    ''                    DateTimeStyles.None, _
    ''                    dateValue) Then
    ''                    DateString = dateValue.ToString("yyyyMMdd")
    ''                End If
    ''            Case Else
    ''                DateString = dateValue.ToString("yyyyMMdd")
    ''        End Select

    ''    End If

    ''    Return DateString

    ''End Function

    Public Sub Write_TextFile_Account(ByVal sAccount() As String)
        Try
            Dim irow As Integer
            Dim sPath As String = System.Windows.Forms.Application.StartupPath & "\"
            Dim sFileName As String = "AccountCode_NotMap.txt"
            Dim sbuffer As String = String.Empty

            If File.Exists(sPath & sFileName) Then
                Try
                    File.Delete(sPath & sFileName)
                Catch ex As Exception
                End Try
            End If

            Dim sw As StreamWriter = New StreamWriter(sPath & sFileName)
            ' Add some text to the file.
            sw.WriteLine("")
            sw.WriteLine("Error!  The following AccNumbers do not have a corresponding SAP G/L Account in the mapping table! ")
            sw.WriteLine("")
            sw.WriteLine("Account Code                       ")
            sw.WriteLine("=============================================================")
            sw.WriteLine(" ")

            For irow = 0 To sAccount.Length
                If Not String.IsNullOrEmpty(sAccount(irow)) Then
                    sw.WriteLine(sAccount(irow).ToString.PadRight(40, " "c))
                Else
                    Exit For
                End If
            Next irow

            sw.WriteLine(" ")
            sw.WriteLine("===============================================================")
            sw.WriteLine("Please create an entry for each of these invalid AccNumbers.")
            sw.Close()
            Process.Start(sPath & sFileName)


        Catch ex As Exception

        End Try

    End Sub

    Public Sub Write_TextFile_ActiveAccount(ByVal sAccount() As String)
        Try
            Dim irow As Integer
            Dim sPath As String = System.Windows.Forms.Application.StartupPath & "\"
            Dim sFileName As String = "AccountCode_ExistorInactive.txt"
            Dim sbuffer As String = String.Empty

            If File.Exists(sPath & sFileName) Then
                Try
                    File.Delete(sPath & sFileName)
                Catch ex As Exception
                End Try
            End If

            Dim sw As StreamWriter = New StreamWriter(sPath & sFileName)
            ' Add some text to the file.
            sw.WriteLine("")
            sw.WriteLine("Error!The following SAP G/L Accounts are not found in the Chart of Accounts or the Account is not an Active ! ")
            sw.WriteLine("")
            sw.WriteLine("Account Code                       ")
            sw.WriteLine("=============================================================")
            sw.WriteLine(" ")

            For irow = 0 To sAccount.Length
                If Not String.IsNullOrEmpty(sAccount(irow)) Then
                    sw.WriteLine(sAccount(irow).ToString.PadRight(40, " "c))
                Else
                    Exit For
                End If
            Next irow

            sw.WriteLine(" ")
            sw.WriteLine("===============================================================")
            sw.WriteLine("Please create an entry for each of these invalid Account Numbers in Chart of Accounts or make sure these accounts are Active.")
            sw.Close()
            Process.Start(sPath & sFileName)


        Catch ex As Exception

        End Try

    End Sub


    Public Sub Write_TextFile_Amount(ByVal sAmount(,) As String)
        Try
            Dim irow As Integer
            Dim sPath As String = System.Windows.Forms.Application.StartupPath & "\"
            Dim sFileName As String = "AccountCode_NotMap.txt"
            Dim sbuffer As String = String.Empty

            If File.Exists(sPath & sFileName) Then
                Try
                    File.Delete(sPath & sFileName)
                Catch ex As Exception
                End Try
            End If

            Dim sw As StreamWriter = New StreamWriter(sPath & sFileName)
            ' Add some text to the file.
            sw.WriteLine("")
            sw.WriteLine("Error!  The Total Debit is not equal to the Total Credit for the following group(RefNo)")
            sw.WriteLine("")
            sw.WriteLine("Debit Amount                 Credit Amount                Difference                       RefNo")
            sw.WriteLine("================================================================================================")
            sw.WriteLine(" ")

            For irow = 0 To UBound(sAmount, 1) - 1
                If Not String.IsNullOrEmpty(sAmount(irow, 0)) Then
                    sw.WriteLine(sAmount(irow, 0).ToString.PadRight(30, " "c) & sAmount(irow, 1).ToString.PadRight(30, " "c) & " " & sAmount(irow, 2).ToString.PadRight(30, " "c) & " " & sAmount(irow, 3))
                Else
                    Exit For
                End If
            Next irow

            sw.WriteLine(" ")
            sw.WriteLine("================================================================================================")
            sw.WriteLine("Please check the grouping of entries in the CSV file")
            sw.Close()
            Process.Start(sPath & sFileName)

        Catch ex As Exception

        End Try

    End Sub

    Public Function SendEmailNotification(ByVal sBody As String, ByVal sSubject As String, ByVal sSenderEmail As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oSmtpServer As New SmtpClient()
        Dim oMail As New MailMessage
        Dim p_SyncDateTime As String = String.Empty

        Try
            sFuncName = "SendEmailNotification()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            ''p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
            ' ''--------- Message Content in HTML tags
            ''Dim sBody As String = String.Empty

            ''sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
            ''sBody = sBody & " Dear Sir/Madam,<br /><br />"
            ''sBody = sBody & p_SyncDateTime & " <br /><br />"
            ''sBody = sBody & " " & " Request for your " & sDocType & " approval in Sap . <br /><br />"
            ''sBody = sBody & " " & " Please login to " & sEntity & "  to approve the document."
            ''sBody = sBody & "<br /><br />"
            ''sBody = sBody & "<br /><br />"
            ''sBody = sBody & " Please do not reply to this email. <div/>"


            oSmtpServer.Credentials = New Net.NetworkCredential(p_oCompDef.sSMTPUser, p_oCompDef.sSMTPPassword)
            oSmtpServer.Port = p_oCompDef.sSMTPPort '587
            oSmtpServer.Host = p_oCompDef.sSMTPServer '"smtp.gmail.com"
            If p_oCompDef.sSSL = "ON" Then
                oSmtpServer.EnableSsl = True
            Else
                oSmtpServer.EnableSsl = False
            End If
            '
            oMail.From = New MailAddress(p_oCompDef.sEmailFrom) '("sapb1.abeoelectra@gmail.com")
            oMail.To.Add(sSenderEmail)
            ' oMail.Attachments.Add(New Attachment(sfileName192.168.1.4
            oMail.Subject = sSubject 'sDocType & " approval notification (" & sEntity & ")"
            oMail.Body = sBody
            oMail.IsBodyHtml = True
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email Notification Sending to " & sSenderEmail, sFuncName)
            oSmtpServer.Send(oMail)
            oMail.Dispose()


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SendEmailNotification = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            '' oMail.Dispose()
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            'Console.WriteLine("Completed with Error " & sFuncName)
            SendEmailNotification = RTN_ERROR
        Finally
            oMail.Dispose()

        End Try

    End Function

    Public Function ExecuteSQLQuery_DT(ByVal sQuery As String) As DataTable

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : JOHN
        ' Date          : MAY 2014 20
        ' Change        :
        '**************************************************************
        Dim oDs As New DataSet
        If sQuery <> "" Then
            Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

            Dim oCon As New SqlConnection(sConstr)
            Dim oCmd As New SqlCommand

            Dim sFuncName As String = String.Empty

            'Dim sConstr As String = "DRIVER={HDBODBC32};SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"

            Try
                sFuncName = "ExecExecuteSQLQuery_DT()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connection Details " & sConstr, sFuncName)
                'oCon.ConnectionString = "DRIVER={HDBODBC};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & " ;SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName & ""
                ' oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName

                oCon.Open()
                oCmd.CommandType = CommandType.Text
                oCmd.CommandText = sQuery
                oCmd.Connection = oCon
                oCmd.CommandTimeout = 0
                Dim da As New SqlDataAdapter(oCmd)
                da.Fill(oDs)
                'Console.WriteLine("Completed Successfully. ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Successfully.", sFuncName)

            Catch ex As Exception
                WriteToLogFile(ex.Message, sFuncName)
                'Console.WriteLine("Completed with ERROR ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Throw New Exception(ex.Message)
            Finally
                oCon.Dispose()
            End Try
            Return oDs.Tables(0)
        Else
            Return Nothing
        End If
    End Function

    Public Function ExecuteSQLInsertQuery(ByVal sQuery As String)

        '**************************************************************
        ' Function      : ExecuteSQLInsertQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : JOHN
        ' Date          : APRIL 2015 20
        ' Change        :
        '**************************************************************
        Dim oDs As New DataSet
        If sQuery <> "" Then
            Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

            Dim oCon As New SqlConnection(sConstr)
            Dim oCmd As New SqlCommand

            Dim sFuncName As String = String.Empty

            'Dim sConstr As String = "DRIVER={HDBODBC32};SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"

            Try
                sFuncName = "ExecuteSQLQuery()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connection Details " & sConstr, sFuncName)
                'oCon.ConnectionString = "DRIVER={HDBODBC};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & " ;SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName & ""
                ' oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName

                oCon.Open()
                oCmd.CommandType = CommandType.Text
                oCmd.CommandText = sQuery
                oCmd.Connection = oCon
                oCmd.CommandTimeout = 0
                oCmd.ExecuteNonQuery()
                ''Dim da As New SqlDataAdapter(oCmd)
                ''Try
                ''    da.Fill(oDs)
                ''Catch ex As Exception
                ''End Try

                'Console.WriteLine("Completed Successfully. ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Successfully.", sFuncName)

            Catch ex As Exception
                WriteToLogFile(ex.Message, sFuncName)
                'Console.WriteLine("Completed with ERROR ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Throw New Exception(ex.Message)
            Finally
                oCon.Dispose()
            End Try
            Return Nothing
        Else
            Return Nothing
        End If
    End Function

    Public Function GetSystemIntializeInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetSystemIntializeInfo()
        '   Purpose     :   This function will be providing information about the initialing variables
        '               
        '   Parameters  :   ByRef oCompDef As CompanyDefault
        '                       oCompDef =  set the Company Default structure
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2014
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty
        Dim sSqlstr As String = String.Empty
        Try

            sFuncName = "GetSystemIntializeInfo()"
            ' Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oCompDef.sDBName = String.Empty
            oCompDef.sServer = String.Empty
            oCompDef.sLicenseServer = String.Empty
            '' oCompDef.iServerLanguage = 3
            'oCompDef.iServerType = 7
            ' oCompDef.sSAPUser = String.Empty
            ' oCompDef.sSAPPwd = String.Empty
            oCompDef.sSAPDBName = String.Empty

            oCompDef.sDebug = String.Empty

            'Email Credentials
            oCompDef.sSMTPServer = String.Empty
            oCompDef.sSMTPPort = String.Empty
            oCompDef.sSMTPUser = String.Empty
            oCompDef.sSMTPPassword = String.Empty
            oCompDef.sToEmailID = String.Empty
            oCompDef.sEmailFrom = String.Empty
            oCompDef.sSSL = String.Empty


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.sServer = ConfigurationManager.AppSettings("Server")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ServerType")) Then
                oCompDef.sServerType = ConfigurationManager.AppSettings("ServerType")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LicenseServer")) Then
                oCompDef.sLicenseServer = ConfigurationManager.AppSettings("LicenseServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPDBName")) Then
                oCompDef.sSAPDBName = ConfigurationManager.AppSettings("SAPDBName")
            End If

            'If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPUserName")) Then
            '    oCompDef.sSAPUser = ConfigurationManager.AppSettings("SAPUserName")
            'End If

            'If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPPassword")) Then
            '    oCompDef.sSAPPwd = ConfigurationManager.AppSettings("SAPPassword")
            'End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.sDBUser = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.sDBPwd = ConfigurationManager.AppSettings("DBPwd")
            End If

            ' folder

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Debug")) Then
                oCompDef.sDebug = ConfigurationManager.AppSettings("Debug")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogPath")) Then
                oCompDef.sFilepath = ConfigurationManager.AppSettings("LogPath")
            End If


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sSMTPServer")) Then
                oCompDef.sSMTPServer = ConfigurationManager.AppSettings("sSMTPServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sSMTPPort")) Then
                oCompDef.sSMTPPort = ConfigurationManager.AppSettings("sSMTPPort")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sSMTPUser")) Then
                oCompDef.sSMTPUser = ConfigurationManager.AppSettings("sSMTPUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sSMTPPassword")) Then
                oCompDef.sSMTPPassword = ConfigurationManager.AppSettings("sSMTPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sAEmailID")) Then
                oCompDef.sToEmailID = ConfigurationManager.AppSettings("sAEmailID")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sEmailFrom")) Then
                oCompDef.sEmailFrom = ConfigurationManager.AppSettings("sEmailFrom")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sSSL")) Then
                oCompDef.sSSL = ConfigurationManager.AppSettings("sSSL")
            End If

            'Console.WriteLine("Completed with SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            'Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function


End Module

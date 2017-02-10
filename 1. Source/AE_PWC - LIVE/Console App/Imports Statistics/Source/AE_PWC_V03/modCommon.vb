Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.IO

Module modCommon


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
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oCompDef.sDBName = String.Empty
            oCompDef.sServer = String.Empty
            oCompDef.sLicenseServer = String.Empty
            oCompDef.iServerLanguage = 3
            'oCompDef.iServerType = 7
            oCompDef.sSAPUser = String.Empty
            oCompDef.sSAPPwd = String.Empty
            oCompDef.sSAPDBName = String.Empty

            oCompDef.sInboxDir = String.Empty
            oCompDef.sSuccessDir = String.Empty
            oCompDef.sFailDir = String.Empty
            oCompDef.sLogPath = String.Empty
            oCompDef.sDebug = String.Empty

            'Email Credentials
            oCompDef.sSMTPServer = String.Empty
            oCompDef.sSMTPPort = String.Empty
            oCompDef.sSMTPUser = String.Empty
            oCompDef.sSMTPPassword = String.Empty
            oCompDef.sToEmailID = String.Empty
            oCompDef.sEmailFrom = String.Empty



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

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPUserName")) Then
                oCompDef.sSAPUser = ConfigurationManager.AppSettings("SAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPPassword")) Then
                oCompDef.sSAPPwd = ConfigurationManager.AppSettings("SAPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.sDBUser = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.sDBPwd = ConfigurationManager.AppSettings("DBPwd")
            End If

            ' folder
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("InboxDir")) Then
                oCompDef.sInboxDir = ConfigurationManager.AppSettings("InboxDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SuccessDir")) Then
                oCompDef.sSuccessDir = ConfigurationManager.AppSettings("SuccessDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FailDir")) Then
                oCompDef.sFailDir = ConfigurationManager.AppSettings("FailDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogPath")) Then
                oCompDef.sLogPath = ConfigurationManager.AppSettings("LogPath")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Debug")) Then
                oCompDef.sDebug = ConfigurationManager.AppSettings("Debug")
            End If

            ' ''Email Credentials:
            ''If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sSMTPServer")) Then
            ''    oCompDef.sSMTPServer = ConfigurationManager.AppSettings("sSMTPServer")
            ''End If

            ''If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sSMTPPort")) Then
            ''    oCompDef.sSMTPPort = ConfigurationManager.AppSettings("sSMTPPort")
            ''End If

            ''If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sSMTPUser")) Then
            ''    oCompDef.sSMTPUser = ConfigurationManager.AppSettings("sSMTPUser")
            ''End If

            ''If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sSMTPPassword")) Then
            ''    oCompDef.sSMTPPassword = ConfigurationManager.AppSettings("sSMTPPassword")
            ''End If

            ''If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sToEmailID")) Then
            ''    oCompDef.sToEmailID = ConfigurationManager.AppSettings("sToEmailID")
            ''End If

            ''If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("sEmailFrom")) Then
            ''    oCompDef.sEmailFrom = ConfigurationManager.AppSettings("sEmailFrom")
            ''End If

            Console.WriteLine("Completed with SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function

    Public Function ExecuteSQLQuery_DT(ByVal sQuery As String, ByVal sConnString As String, ByRef sErrDesc As String) As DataTable

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : JOHN
        ' Date          : MAY 2014 20
        ' Change        :
        '**************************************************************

        ''Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & sCompanyDB & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

        Dim oCon As New SqlConnection(sConnString)
        Dim oCmd As New SqlCommand
        Dim oDs As New DataSet
        Dim oDT As New DataTable
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "ExecExecuteSQLQuery_DT()"
            Console.WriteLine("Starting Function.. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing Query : " & sQuery, sFuncName)

            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            da.Fill(oDT)
            Console.WriteLine("Completed Successfully. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Successfully.", sFuncName)

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return Nothing
        Finally
            If oCon.State = ConnectionState.Open Then
                oCon.Dispose()
            End If
        End Try
        Return oDT
    End Function

    Public Function GetEntitiesDetails(ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetEntitiesDetails()
        '   Purpose     :   This function will be providing information about the Entities, SAP username, SAP Password, Banking Details
        '               
        '   Parameters  :   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************


        Dim sSqlstr As String = String.Empty
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "GetEntitiesDetails()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            Console.WriteLine("Starting Function", sFuncName)


            sSqlstr = "SELECT T0.[PrcCode] [OUCode], T0.[PrcName] [OU Name], T0.[U_AB_ENTITY] [Entity],T0.[U_AB_REPORTCODE] [BU Code], " & _
                "T2.[U_AB_REPORTCODE] [LOS Code], T3.[U_AB_USERCODE] [User], T3.[U_AB_PASSWORD] [Pass] " & _
                "FROM OPRC T0  INNER JOIN ODIM T1 ON T0.[DimCode] = T1.[DimCode] left outer join OPRC T2 " & _
                "on T2.[PrcCode] = T0.[U_AB_REPORTCODE] left outer join [@AB_COMPANYDATA] T3 on T0.[U_AB_ENTITY] = T3.[Name] WHERE T1.[DimCode] = 3"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

            p_oEntitesDetails = ExecuteSQLQuery_DT(sSqlstr, sConstr, sErrDesc)

            sSqlstr = "SELECT T0.[Code], T0.[Name], T0.[U_AB_STDESCRIPTION], T0.[U_AB_STNEWCODE] FROM [dbo].[@AB_STOLDCODE]  T0"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            p_oSTOLDCODE = ExecuteSQLQuery_DT(sSqlstr, sConstr, sErrDesc)

            sSqlstr = "SELECT max(cast(T0.[Code] as integer))+ 1 [RowCount] FROM [dbo].[@AB_STOLDCODE]  T0"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AB_STOLDCODE Count " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            p_oDTPWCRowCount = ExecuteSQLQuery_DT(sSqlstr, sConstr, sErrDesc)


            Console.WriteLine("Completed With SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)
            GetEntitiesDetails = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed With ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR", sFuncName)
            GetEntitiesDetails = RTN_ERROR
        End Try

    End Function

    Public Function IdentifyExcelFile(ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   IdentifyExcelFile()
        '   Purpose     :   This function will identify the Excel file of Journal Entry
        '                    Upload the file into Dataview and provide the information to post transaction in SAP.
        '                     Transaction Success : Move the Excel file to SUCESS folder
        '                     Transaction Fail :    Move the Excel file to FAIL folder and send Error notification to concern person
        '               
        '   Parameters  :   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************


        Dim sSqlstr As String = String.Empty
        Dim bJEFileExist As Boolean
        Dim sFileType As String = String.Empty
        Dim oDTDistinct As DataTable = Nothing
        Dim oDTRowFilter As DataTable = Nothing
        Dim oDSJE As DataSet = Nothing
        Dim oDICompany() As SAPbobsCOM.Company = Nothing

        Dim sFuncName As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim oDVLineTable As DataView = Nothing
        Dim oDTHeader As DataTable = Nothing
        Dim sCompanyDB As String = String.Empty
        Dim sConnString As String = String.Empty
        Dim oFileInfo As System.IO.FileInfo = Nothing

        Try
            sFuncName = "IdentifyExcelFile()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
            Dim files() As System.IO.FileInfo

            files = DirInfo.GetFiles("*.*")

            For Each File As System.IO.FileInfo In files
                bJEFileExist = True
                Console.WriteLine("Attempting File Name - " & File.Name, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting File Name - " & File.Name, sFuncName)
                'sFileType = Replace(File.Name, ".txt", "").Trim
                'upload the CSV to Dataview

                oFileInfo = File

                Console.WriteLine("Calling GetDataViewFromExcel() ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetDataViewFromExcel() ", sFuncName)
                oDVLineTable = GetDataViewFromExcel(File.FullName, File.Name)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("oDVLineTable.count " & oDVLineTable.Count, sFuncName)

                For Each dr As DataRowView In oDVLineTable
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("oDVLineTable Entity - " & dr.Item("CompanyDB").ToString.Trim, sFuncName)
                Next

                If oDT_OUCODE.Rows.Count > 0 Then
                    Console.WriteLine("Calling Validation Error ...... ! ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Validation Error ", sFuncName)
                    Write_TextFile(oDT_OUCODE, sErrDesc)
                    IdentifyExcelFile = RTN_ERROR

                    Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                    FileMoveToArchive(File, File.FullName, RTN_ERROR, "")

                    sErrDesc = "Validation Error ..... ! [OU Code are not Exists in Company DB]"
                    Exit Function
                End If

                Console.WriteLine("Getting Distinct of Entity", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Distinct of Entity ", sFuncName)
               

                oDTDistinct = oDVLineTable.Table.DefaultView.ToTable(True, "CompanyDB")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("oDTDistinct Count " & oDTDistinct.Rows.Count, sFuncName)
                For Each dr As DataRow In oDTDistinct.Rows
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("oDTDistinct Entity " & dr.Item("CompanyDB").ToString.Trim, sFuncName)
                Next

                P_sQueryString = String.Empty

                For imjs As Integer = 0 To oDTDistinct.Rows.Count - 1

                    sCompanyDB = oDTDistinct.Rows(imjs).Item(0).ToString.Trim()
                    ''If sCompanyDB = "" Then Continue For
                    Console.WriteLine("Filtering Data with respective DataBase -  " & sCompanyDB, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Filtering data with respective Entity -  " & sCompanyDB, sFuncName)
                    oDVLineTable.RowFilter = "CompanyDB = '" & sCompanyDB & "'"
                    Console.WriteLine("Calling Function JournalEntry_Posting() ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function ImportStatistics() ", sFuncName)
                    If ImportStatistics(oDVLineTable, sCompanyDB, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Next imjs

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("P_sQueryString " & P_sQueryString, sFuncName)
                If P_sQueryString <> String.Empty Then
                    Console.WriteLine("Calling ExecuteSQLQuery_DT() for Inserting the Data Into the Statistics Table", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() for Inserting thee Data Into the Statistics Table", sFuncName)
                    sConnString = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd
                    ExecuteInsertSQLQuery(sConnString, P_sQueryString, sErrDesc)
                End If
                Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                FileMoveToArchive(File, File.FullName, RTN_SUCCESS, "")
            Next

            If bJEFileExist = False Then
                Console.WriteLine("No input file found  ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No input file found ", sFuncName)
            End If

            Console.WriteLine("Completed With SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)
            IdentifyExcelFile = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(ex.Message, sFuncName)

            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Try", sFuncName)
            Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)

            FileMoveToArchive(oFileInfo, oFileInfo.FullName, RTN_ERROR, "")
            IdentifyExcelFile = RTN_ERROR

            Console.WriteLine("Completed With ERROR", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR", sFuncName)
            IdentifyExcelFile = RTN_ERROR

        End Try

    End Function

    Public Function ImportStatistics(ByRef oDVLineDetails As DataView, ByVal sCompanyDB As String, ByRef sErrDesc As String) As Long

        'Function   :   ImportStatistics()
        'Purpose    :   Import Text File Data Into UDT
        'Parameters :   ByVal oForm As SAPbouiCOM.Form
        '                   oForm=Form Type
        '               ByRef sErrDesc As String
        '                   sErrDesc=Error Description to be returned to calling function
        '               
        'Return     :   0 - FAILURE
        '               1 - SUCCESS
        'Author     :   SAI
        'Date       :   22/1/2015
        'Change     :

        Dim sFuncName As String = String.Empty

        Dim oDt As DataTable
        Dim iCode As Integer
        Dim sConnString As String = String.Empty
        Dim oDTCode As DataTable = Nothing

        Dim sSql As String

        Try
            sFuncName = "ImportStatistics()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If oDVLineDetails Is Nothing Then
                sErrDesc = "No Datas in the TXT file"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Datas in the TXT file", sFuncName)

            End If

            sConnString = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & sCompanyDB & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

            sSql = "SELECT isnull(Max(CAST( CODE as int)),0)+1 AS CODE FROM [@AB_STATITISTICSDATA]"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() for getting the latest code", sFuncName)
            oDTCode = ExecuteSQLQuery_DT(sSql, sConnString, sErrDesc)

            iCode = oDTCode.Rows(0)(0).ToString().Trim()

            oDt = oDVLineDetails.ToTable
            sSql = String.Empty

            For Each row As DataRow In oDt.Rows
                ' write insert statement
                If row.Item(0).ToString.ToUpper.Trim() = "AB_GLCODE" Then Continue For

                ''sSql += " Insert Into " & sCompanyDB & "..[@AB_STATITISTICSDATA] ( [Code],  [Name], [U_AB_PERIOD],[U_AB_OPER_UNIT], " & _
                ''    " [U_AB_ENTITY], [U_AB_DEBIT_CREDIT],[U_AB_AMOUNT], [U_AB_GLCODE], [U_AB_DESCRIPTION],[U_AB_TRANSDATE]) " & _
                ''    " Values ('" & iCode & "','" & iCode & "','" & row.Item(1).ToString & "','" & row.Item(5).ToString & "','" & row.Item(6).ToString & "', " & _
                ''    "'" & row.Item(3).ToString & "'," & CDbl(row.Item(2).ToString) & ",'" & row.Item(0).ToString & "', '" & row.Item(4).ToString & "','" & row.Item(8).ToString & "')"
                sSql += " Insert Into " & sCompanyDB & "..[@AB_STATITISTICSDATA] ( [Code],  [Name], [U_AB_PERIOD],[U_AB_OPER_UNIT], " & _
                   " [U_AB_ENTITY], [U_AB_DEBIT_CREDIT],[U_AB_AMOUNT], [U_AB_GLCODE], [U_AB_DESCRIPTION]) " & _
                   " Values ('" & iCode & "','" & iCode & "','" & row.Item(1).ToString & "','" & row.Item(5).ToString & "','" & row.Item(6).ToString & "', " & _
                   "'" & row.Item(3).ToString & "'," & CDbl(row.Item(2).ToString) & ",'" & row.Item(0).ToString & "', '" & row.Item(4).ToString & "')"

                iCode = iCode + 1
            Next

            P_sQueryString += sSql

            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() for Inserting the data to the Table.", sFuncName)

            'ExecuteSQLQuery_DT(sSql, sConnString, sErrDesc)

            ImportStatistics = RTN_SUCCESS

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch exc As Exception
            ImportStatistics = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
        End Try
    End Function

    Public Function SendEmailNotification(ByVal CurrFileToUpload As String, ByVal sCompanyCode As String, _
                                          ByVal sCompanyName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oSmtpServer As New SmtpClient()
        Dim oMail As New MailMessage
        Dim p_SyncDateTime As String = String.Empty

        Try
            sFuncName = "SendEmailNotification()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)

            Console.WriteLine("Sending Mail To : " & p_oCompDef.sToEmailID)


            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
            '--------- Message Content in HTML tags
            Dim sBody As String = String.Empty

            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
            sBody = sBody & " Dear Sir/Madam,<br /><br />"
            sBody = sBody & p_SyncDateTime & " <br /><br />"
            sBody = sBody & " " & "Please find the attached FAILED document in SAP and followed by the ERROR.<br /><br /> "
            sBody = sBody & " " & " Company Code : " & sCompanyCode & "<br /> "
            sBody = sBody & " " & " Company Name : " & sCompanyName & " <br /> "
            sBody = sBody & "<br /> <font color=""red""> Error Message : " & sErrDesc & "</font><br />"
            sBody = sBody & "<br /><br />"
            sBody = sBody & " Please do not reply to this email. <div/>"


            ''<font size="3" color="red">This is some text!</font>

            Dim attachment As System.Net.Mail.Attachment
            attachment = New System.Net.Mail.Attachment(CurrFileToUpload)
            oMail.Attachments.Add(attachment)


            oSmtpServer.Credentials = New Net.NetworkCredential(p_oCompDef.sSMTPUser, p_oCompDef.sSMTPPassword)
            oSmtpServer.Port = p_oCompDef.sSMTPPort '587
            oSmtpServer.Host = p_oCompDef.sSMTPServer '"smtp.gmail.com"
            oSmtpServer.EnableSsl = True
            oMail.From = New MailAddress(p_oCompDef.sEmailFrom) '("sapb1.abeoelectra@gmail.com")
            oMail.To.Add(p_oCompDef.sToEmailID)
            ' oMail.Attachments.Add(New Attachment(sfileName192.168.1.4
            oMail.Subject = "Reg., Error While Uploading Journal Entry. "
            oMail.Body = sBody
            oMail.IsBodyHtml = True

            oSmtpServer.Send(oMail)
            oMail.Dispose()
            Console.WriteLine("Sending Mail Completed Successfully to this EmailID : " & p_oCompDef.sToEmailID)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email Notification Sent to " & p_oCompDef.sToEmailID, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SendEmailNotification = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            oMail.Dispose()
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            SendEmailNotification = RTN_ERROR
        Finally
            oMail.Dispose()

        End Try

    End Function

    Public Function GetDataViewFromExcel(ByVal CurrFileToUpload As String, ByVal Filename As String) As DataView

        ' **********************************************************************************
        '   Function    :   GetDataViewFromExcel()
        '   Purpose     :   This function will upload the data from Excel file to Dataview
        '   Parameters  :   ByRef CurrFileToUpload AS String 
        '                       CurrFileToUpload = File Name
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************

        Dim oDTHeader As New DataTable

        Dim sGLCode As String = String.Empty
        Dim sDescription As String = String.Empty
        Dim dAmount As Double = 0
        Dim sType As String = String.Empty
        Dim sPeriod As String = String.Empty
        Dim sOperUnit As String = String.Empty
        Dim sEntity As String = String.Empty
        Dim sChTransDate As String = String.Empty
        Dim sCompanyDB As String = String.Empty
        Dim oDTInsert As New DataTable
        Dim sInsertString As String = String.Empty
        Dim sConnString As String = String.Empty
        Dim sErrDesc As String = String.Empty

        Dim dvEntiry As DataView = New DataView(p_oEntitesDetails)
        Dim oDvNewGlCode As DataView = New DataView(p_oSTOLDCODE)
        oDT_OUCODE = New DataTable()

        Dim iCount As Integer = 1

        Dim sFuncName As String = String.Empty

        Dim ExcelApp As New Microsoft.Office.Interop.Excel.Application
        Dim ExcelWorkbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ExcelWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim excelRng As Microsoft.Office.Interop.Excel.Range


        Try
            sFuncName = "GetDataViewFromExcel"

            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)


            ExcelWorkbook = ExcelApp.Workbooks.Open(CurrFileToUpload)
            ExcelWorkSheet = ExcelWorkbook.ActiveSheet
            excelRng = ExcelWorkSheet.Range("A1")
            Dim RowIndex As Integer = 6

            oDTInsert.Columns.Add("OUCode", GetType(String))
            oDTInsert.Columns.Add("U_AB_STDESCRIPTION", GetType(String))
            oDTInsert.Columns.Add("U_AB_STNEWCODE", GetType(String))

            oDTHeader.Columns.Add("GLCode", GetType(String))
            oDTHeader.Columns.Add("Period", GetType(String))
            oDTHeader.Columns.Add("Amount", GetType(Double))
            oDTHeader.Columns.Add("Type", GetType(String))
            oDTHeader.Columns.Add("Description", GetType(String))
            oDTHeader.Columns.Add("OperUnit", GetType(String))
            oDTHeader.Columns.Add("Entity", GetType(String))
            oDTHeader.Columns.Add("CompanyDB", GetType(String))
            oDTHeader.Columns.Add("ChTransDate", GetType(String))

            oDT_OUCODE.Columns.Add("Sno", GetType(String))
            oDT_OUCODE.Columns.Add("OUCode", GetType(String))
            oDT_OUCODE.Columns.Add("Msg", GetType(String))

            While excelRng.Range("A" & RowIndex & "").Text <> "" And excelRng.Range("B" & RowIndex & "").Text <> ""
                RowIndex = RowIndex + 1
            End While


            Dim i As Integer = 1
            For i = 2 To RowIndex - 1

                If String.IsNullOrEmpty(excelRng.Range("A" & i & "").Text) Then Continue For

                sGLCode = excelRng.Range("A" & i & "").Text
                sPeriod = excelRng.Range("B" & i & "").Text
                sChTransDate = excelRng.Range("C" & i & "").Text
                dAmount = excelRng.Range("E" & i & "").Text
                sType = excelRng.Range("F" & i & "").Text
                sDescription = excelRng.Range("J" & i & "").Text
                sOperUnit = excelRng.Range("P" & i & "").Text
                sEntity = excelRng.Range("Q" & i & "").Text

                dvEntiry.RowFilter = "OUCode='" & sOperUnit & "'"
                If dvEntiry.Count > 0 Then
                    sCompanyDB = dvEntiry.Item(0)(2).ToString
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entity " & sCompanyDB, sFuncName)
                    If String.IsNullOrEmpty(sCompanyDB) Then
                        oDT_OUCODE.Rows.Add(i, sOperUnit, "No Entity tagged with this OU Code in Company DB :- " & p_oCompDef.sSAPDBName)
                    End If
                Else
                    oDT_OUCODE.Rows.Add(i, sOperUnit, "OU Code not Exists in Company DB :- " & p_oCompDef.sSAPDBName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("OU Code " & sOperUnit & " Entity is Empty in Line " & iCount, sFuncName)
                    sCompanyDB = ""
                End If

                'MsgBox(excelRng.Range("A" & i & "").Text.ToString.Trim)

                If Left(sGLCode, 2) = "ST" Then
                    oDvNewGlCode.RowFilter = "U_AB_STNEWCODE = '" & sGLCode & "'"
                Else
                    oDvNewGlCode.RowFilter = "Name = '" & sGLCode & "'"
                End If


                If oDvNewGlCode.Count > 0 Then
                    sGLCode = oDvNewGlCode.Item(0)("U_AB_STNEWCODE").ToString
                Else

                    If Left(sGLCode, 2) = "ST" Then
                        oDTInsert.Rows.Add(sDescription, sGLCode, sGLCode) ' oDTInsert.Rows.Add(sDescription, "NULL", sGLCode)
                    Else
                        oDTInsert.Rows.Add(sDescription, sGLCode, sGLCode) '  oDTInsert.Rows.Add(sDescription, sGLCode, "NULL")
                    End If
                End If

                oDTHeader.Rows.Add(sGLCode, sPeriod, dAmount, sType, sDescription, sOperUnit, sEntity, sCompanyDB, sChTransDate)

            Next


            If oDTInsert Is Nothing Then
            Else
                If oDTInsert.Rows.Count > 0 Then
                    Dim iCode As String = p_oDTPWCRowCount.Rows(0).Item(0).ToString.Trim
                    For Each row As DataRow In oDTInsert.Rows

                        sInsertString += " Insert Into [@AB_STOLDCODE]  ( [Code],  [Name], [U_AB_STDESCRIPTION],[U_AB_STNEWCODE])" & _
                                     " Values ('" & iCode & "','" & row.Item(1).ToString & "','" & row.Item(0).ToString & "','" & row.Item(2).ToString & "')"
                        iCode += 1
                    Next
                End If
            End If

            ExcelWorkbook.Close()
            ExcelWorkbook = Nothing
            ExcelApp.Quit()
            ExcelApp = Nothing
            ExcelWorkSheet = Nothing
            excelRng = Nothing

            If sInsertString.Length > 0 Then
                Console.WriteLine("Calling ExecuteSQLQuery_DT() for Inserting the Data Into the Statistics OLD Code", sFuncName)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() for Inserting thee Data Into the Statistics Table", sFuncName)
                sConnString = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd
                ExecuteInsertSQLQuery(sConnString, sInsertString, sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)

            Return oDTHeader.DefaultView

        Catch ex As Exception
            ExcelWorkbook.Close()
            ExcelWorkbook = Nothing
            ExcelApp.Quit()
            ExcelApp = Nothing
            ExcelWorkSheet = Nothing
            excelRng = Nothing
            Return Nothing
        End Try
    End Function

    Public Function GetCompanyDetails(ByVal sEntity As String, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetBankingDetails()
        '   Purpose     :   This function will get the relavent Banking informations with respective Entities 
        '   Parameters  :   ByRef sEntity AS String 
        '                       sEntity = Entity Name
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************
        Dim sFuncName As String = String.Empty
        sFuncName = "GetCompanyDetails()"

        Try
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            Dim Findatarow() As DataRow = p_oEntitesDetails.Select("Entity = '" & sEntity.ToString.Trim & "'")

            For Each row As DataRow In Findatarow
                p_sSAPEntityName = row(2)
                p_sSAPUName = row(5)
                p_sSAPUPass = row(6)

            Next

            Console.WriteLine("Completed With SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)
            GetCompanyDetails = RTN_SUCCESS

        Catch ex As Exception
            Console.WriteLine("Completed With ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            GetCompanyDetails = RTN_ERROR
        End Try

    End Function

    Public Function ConnectToTargetCompany(ByRef oCompany As SAPbobsCOM.Company, _
                                          ByVal sEntity As String, _
                                          ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   ConnectToTargetCompany()
        '   Purpose     :   This function will be providing to proceed the connectivity of 
        '                   using SAP DIAPI function
        '               
        '   Parameters  :   ByRef oCompany As SAPbobsCOM.Company
        '                       oCompany =  set the SAP DI Company Object
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2013 21
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Dim sSQL As String = String.Empty
        Dim oDs As New DataSet

        Try
            sFuncName = "ConnectToTargetCompany()"
            Console.WriteLine("Starting function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Calling GetCompanyDetails ", sFuncName)
            Console.WriteLine("Calling GetCompanyDetails ", sFuncName)
            If GetCompanyDetails(sEntity, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If String.IsNullOrEmpty(p_sSAPUName) Then
                sErrDesc = "No Database login information found in COMPANYDATA Table. Please check"
                Console.WriteLine("No Database login information found in COMPANYDATA Table. Please check ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Database login information found in COMPANYDATA Table. Please check", sFuncName)
                Throw New ArgumentException(sErrDesc)
            Else

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)
                Console.WriteLine("Initializing the Company Object ", sFuncName)
                oCompany = New SAPbobsCOM.Company

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)
                Console.WriteLine("Assigning the representing database name ", sFuncName)
                oCompany.Server = p_oCompDef.sServer

                If p_oCompDef.sServerType = "2008" Then
                    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
                ElseIf p_oCompDef.sServerType = "2012" Then
                    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
                ElseIf p_oCompDef.sServerType = "2014" Then
                    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
                End If


                oCompany.LicenseServer = p_oCompDef.sLicenseServer
                oCompany.CompanyDB = p_sSAPEntityName
                oCompany.UserName = p_sSAPUName
                oCompany.Password = p_sSAPUPass

                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

                oCompany.UseTrusted = False

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)
                Console.WriteLine("Connecting to the Company Database. ", sFuncName)
                iRetValue = oCompany.Connect()

                If iRetValue <> 0 Then
                    oCompany.GetLastError(iErrCode, sErrDesc)

                    sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                        oCompany.CompanyDB, System.Environment.NewLine, _
                                    vbTab, sErrDesc)

                    Throw New ArgumentException(sErrDesc)
                End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS ", sFuncName)
            ConnectToTargetCompany = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            ConnectToTargetCompany = RTN_ERROR
        End Try
    End Function

    Public Sub FileMoveToArchive(ByVal oFile As System.IO.FileInfo, ByVal CurrFileToUpload As String, ByVal iStatus As Integer, ByVal sErrDesc As String)

        'Event      :   FileMoveToArchive
        'Purpose    :   For Renaming the file with current time stamp & moving to archive folder
        'Author     :   JOHN 
        'Date       :   21 MAY 2014

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "FileMoveToArchive"
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            'Dim RenameCurrFileToUpload = Replace(CurrFileToUpload.ToUpper, ".CSV", "") & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
            Dim RenameCurrFileToUpload As String = Mid(oFile.Name, 1, oFile.Name.Length - 4) & "_" & Now.ToString("yyyyMMddhhmmss") & ".xls"

            If iStatus = RTN_SUCCESS Then
                Console.WriteLine("Moving CSV file to success folder ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving CSV file to success folder", sFuncName)
                oFile.MoveTo(p_oCompDef.sSuccessDir & "\" & RenameCurrFileToUpload)
            Else
                Console.WriteLine("Moving CSV file to Fail folder ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving CSV file to Fail folder", sFuncName)
                oFile.MoveTo(p_oCompDef.sFailDir & "\" & RenameCurrFileToUpload)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in renaming/copying/moving ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in renaming/copying/moving", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Public Function Del_schema(ByVal csvFileFolder As String) As Long

        ' ***********************************************************************************
        '   Function   :    Del_schema()
        '   Purpose    :    This function is handles - Delete the Schema file
        '   Parameters :    ByVal csvFileFolder As String
        '                       csvFileFolder = Passing file name
        '   Author     :    JOHN
        '   Date       :    26/06/2014 
        '   Change     :   
        '                   
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "Del_schema()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            Console.WriteLine("Starting Function... " & sFuncName)

            Dim FileToDelete As String
            FileToDelete = csvFileFolder & "\\schema.ini"
            If System.IO.File.Exists(FileToDelete) = True Then
                System.IO.File.Delete(FileToDelete)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS " & sFuncName)
            Del_schema = RTN_SUCCESS
        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            Del_schema = RTN_ERROR
        End Try
    End Function

    Public Function Create_schema(ByVal csvFileFolder As String, ByVal FileName As String) As Long

        ' ***********************************************************************************
        '   Function   :    Create_schema()
        '   Purpose    :    This function is handles - Create the Schema file
        '   Parameters :    ByVal csvFileFolder As String
        '                       csvFileFolder = Passing file name
        '   Author     :    JOHN
        '   Date       :    26/06/2014 
        '   Change     :   
        '                   
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "Create_schema()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            Console.WriteLine("Starting Function... " & sFuncName)

            Dim csvFileName As String = FileName
            Dim fsOutput As FileStream = New FileStream(csvFileFolder & "\\schema.ini", FileMode.Create, FileAccess.Write)
            Dim srOutput As StreamWriter = New StreamWriter(fsOutput)
            'Dim s1, s2, s3, s4, s5 As String

            srOutput.WriteLine("[" & csvFileName & "]")
            srOutput.WriteLine("ColNameHeader=False")
            srOutput.WriteLine("Format=CSVDelimited")
            srOutput.WriteLine("Col1=F1 Text")
            srOutput.WriteLine("Col2=F2 Text")
            srOutput.WriteLine("Col3=F3 Text")
            srOutput.WriteLine("Col4=F4 Text")
            srOutput.WriteLine("Col5=F5 Text")
            srOutput.WriteLine("Col6=F6 Text")
            srOutput.WriteLine("Col7=F7 Text")
            srOutput.WriteLine("Col8=F8 Text")
            srOutput.WriteLine("Col9=F9 Text")
            srOutput.WriteLine("Col10=F10 Double")
            srOutput.WriteLine("Col11=F11 Text")
            srOutput.WriteLine("Col12=F12 Double")
            srOutput.WriteLine("Col13=F13 Text")
            srOutput.WriteLine("Col14=F14 Text")
            srOutput.WriteLine("Col15=F15 Text")
            srOutput.WriteLine("MaxScanRows=0")
            srOutput.WriteLine("CharacterSet=OEM")
            'srOutput.WriteLine(s1.ToString() + ControlChars.Lf + s2.ToString() + ControlChars.Lf + s3.ToString() + ControlChars.Lf + s4.ToString() + ControlChars.Lf)
            srOutput.Close()
            fsOutput.Close()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS " & sFuncName)
            Create_schema = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            Create_schema = RTN_ERROR
        End Try

    End Function

    Public Function Write_TextFile(ByVal oDTDisplay As DataTable, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = String.Empty

        Try
            Dim irow As Integer
            Dim sPath As String = System.Windows.Forms.Application.StartupPath & "\"
            Dim sFileName As String = "Validationims.txt"
            Dim sbuffer As String = String.Empty

            sFuncName = "Write_TextFile()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            If File.Exists(sPath & sFileName) Then
                Try
                    File.Delete(sPath & sFileName)
                Catch ex As Exception
                End Try
            End If

            Dim sw As StreamWriter = New StreamWriter(sPath & sFileName)
            ' Add some text to the file.
            sw.WriteLine("")
            sw.WriteLine("Validation Error!  The following OU Codes are not Existing / No Entities Tagged for this OU Codes  ")
            sw.WriteLine("")
            sw.WriteLine("Line No.  OU Code        Message                                                       ")
            sw.WriteLine("=======================================================================================")
            sw.WriteLine(" ")

            For irow = 0 To oDTDisplay.Rows.Count - 1
                If Not String.IsNullOrEmpty(oDTDisplay.Rows(irow).Item(0).ToString) Then
                    sw.WriteLine(oDTDisplay.Rows(irow).Item(0).ToString.PadRight(10, " "c) + oDTDisplay.Rows(irow).Item(1).ToString.PadRight(17, " "c) _
                                 + oDTDisplay.Rows(irow).Item(2).ToString.PadRight(57, " "c))
                Else
                    Exit For
                End If
            Next irow

            sw.WriteLine(" ")
            sw.WriteLine("========================================================================================")
            sw.WriteLine("Please Check.")
            sw.Close()
            Process.Start(sPath & sFileName)

            Write_TextFile = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)

        Catch ex As Exception
            Write_TextFile = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

    Public Function ExecuteInsertSQLQuery(ByVal sConstr As String, ByVal sQuery As String, ByRef sErrDesc As String) As Long

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : JOHN
        ' Date          : MAY 2014 20
        ' Change        :
        '**************************************************************

        '' Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oDs As New DataSet
        Dim sFuncName As String = String.Empty

        'Dim sConstr As String = "DRIVER={HDBODBC32};SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"

        Try
            sFuncName = "ExecuteInsertSQLQuery()"
            Console.WriteLine("Starting Function.. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            'oCon.ConnectionString = "DRIVER={HDBODBC};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & " ;SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName & ""
            ' oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL " & sQuery, sFuncName)
            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            oCmd.ExecuteNonQuery()

            Console.WriteLine("Completed Successfully. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Successfully.", sFuncName)

            Return RTN_SUCCESS
        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return RTN_ERROR
        Finally
            oCon.Dispose()
        End Try

    End Function

End Module
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

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("JESeries")) Then
                oCompDef.sSeries = ConfigurationManager.AppSettings("JESeries")
            End If

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


    Public Function ExecuteSQLQuery_DT(ByVal sQuery As String) As DataTable

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : JOHN
        ' Date          : MAY 2014 20
        ' Change        :
        '**************************************************************

        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oDs As New DataSet
        Dim sFuncName As String = String.Empty

        'Dim sConstr As String = "DRIVER={HDBODBC32};SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"

        Try
            sFuncName = "ExecExecuteSQLQuery_DT()"
            ' Console.WriteLine("Starting Function.. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            'oCon.ConnectionString = "DRIVER={HDBODBC};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & " ;SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName & ""
            ' oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName

            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            da.Fill(oDs)
            '  Console.WriteLine("Completed Successfully. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Successfully.", sFuncName)

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            oCon.Dispose()
        End Try
        Return oDs.Tables(0)
    End Function

    Public Function ExecuteSQLQuery(ByVal sEntity As String, ByRef sErrDesc As String) As Long

        '**************************************************************
        ' Function      : ExecuteSQLQuery_DT
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : JOHN
        ' Date          : MAY 2014 20
        ' Change        :
        '**************************************************************

        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oDt As New DataTable
        Dim sFuncName As String = String.Empty
        Dim sQuery As String = String.Empty

        'Dim sConstr As String = "DRIVER={HDBODBC32};SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"

        Try
            sFuncName = "ExecExecuteSQLQuery_DT()"
            ' Console.WriteLine("Starting Function.. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            'oCon.ConnectionString = "DRIVER={HDBODBC};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & " ;SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName & ""
            ' oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName
            sQuery = "SELECT isnull(Max(CAST( CODE as int)),0)+1 AS CODE FROM " & sEntity & ".. [@AB_STATITISTICSDATA]"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Statistics Row Count SQL " & sQuery, sFuncName)
            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            da.Fill(oDt)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entity." & sEntity & " Row Count " & oDt.Rows(0).Item(0).ToString.Trim, sFuncName)
            oDT_StatisticsRowCount.Rows.Add(sEntity, oDt.Rows(0).Item(0).ToString.Trim)
            '  Console.WriteLine("Completed Successfully. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Successfully.", sFuncName)

            ExecuteSQLQuery = RTN_SUCCESS

        Catch ex As Exception
            ExecuteSQLQuery = RTN_ERROR
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            sErrDesc = ex.Message
        Finally
            oCon.Dispose()
        End Try

    End Function


    Public Function ExecuteInsertSQLQuery(ByVal sQuery As String, ByRef sErrDesc As String) As Long

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : JOHN
        ' Date          : MAY 2014 20
        ' Change        :
        '**************************************************************

        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

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
            Dim da As New SqlDataAdapter(oCmd)
            Try
                da.Fill(oDs)
            Catch ex As Exception
            End Try

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
            ' Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            Console.WriteLine("Starting Function " & sFuncName)
            ' Getting the details of Entity, SAP User name, Password and Banking from the COMPANYDATA Table
            'sSqlstr = "SELECT T0.[PrcCode] [Center Code], T0.[PrcName] [Center Name], T1.[Name] [DB Name], T1.[U_AE_UPass] [Pass], T1.[U_AE_UName] [User Name] FROM OPRC T0 " & _
            '    "inner join [dbo].[@AE_COMPANYDATA]  T1 on T0.[U_AE_DBName] = T1.Name"

            sSqlstr = "SELECT T0.[PrcCode] [OUCode], T0.[PrcName] [OU Name], T0.[U_AB_ENTITY] [Entity],T0.[U_AB_REPORTCODE] [BU Code], " & _
                "T2.[U_AB_REPORTCODE] [LOS Code], T3.[U_AB_USERCODE] [User], T3.[U_AB_PASSWORD] [Pass], T0.[U_AB_OUCOMMON] [EntityFlag], T3.[U_AB_IPOWERCODE] [EntityCode] " & _
                "FROM OPRC T0  INNER JOIN ODIM T1 ON T0.[DimCode] = T1.[DimCode] left outer join OPRC T2 " & _
                "on T2.[PrcCode] = T0.[U_AB_REPORTCODE] left outer join [@AB_COMPANYDATA] T3 on T0.[U_AB_ENTITY] = T3.[Name] WHERE T1.[DimCode] = 3"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            p_oEntitesDetails = ExecuteSQLQuery_DT(sSqlstr)

            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '----------------------- GL Account
            sSqlstr = "SELECT T0.[AcctCode], T0.[AcctName], T0.FrgnName [ExportCode] FROM OACT T0"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() ", sFuncName)
            p_oGLAccount = ExecuteSQLQuery_DT(sSqlstr)
            ' SELECT T0.[Code], T0.[Name], T0.[U_AB_STDESCRIPTION], T0.[U_AB_STNEWCODE] FROM [dbo].[@AB_IPOWERSTCODE]  T0
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '----------------------- AB_IPOWERSTCODE
            sSqlstr = "SELECT T0.[Code], T0.[Name], T0.[U_AB_STDESCRIPTION], T0.[U_AB_STNEWCODE] , case when left(T0.[U_AB_STNEWCODE],2) = 'ST' then 'IMS' else 'IP' end [Cat] FROM [dbo].[@AB_IPOWERSTCODE]  T0"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            p_oSTOLDCODE = ExecuteSQLQuery_DT(sSqlstr)

            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '----------------------- Company Data
            sSqlstr = "SELECT T0.[U_AB_IPOWERCODE] [ipEntityCode], T0.[U_AB_COMCODE]  [EntityCode], T0.[U_AB_COMPANYNAME], T0.[U_AB_USERCODE], T0.[U_AB_PASSWORD] FROM [dbo].[@AB_COMPANYDATA]  T0"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            p_oDTCompanyData = ExecuteSQLQuery_DT(sSqlstr)

            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '----------------------- iPower Period
            sSqlstr = "select Code, Name,left(LEFT(name, CHARINDEX(' ', Name  )),3) + ' ' + CAST(YEAR(getdate()) as varchar)[Month Name]," & _
"cast(month(left(LEFT(name, CHARINDEX(' ', Name  )),3) + ' 1 2015') as varchar) + ' ' + CAST(YEAR(getdate()) as varchar) [Month Number]," & _
"cast(month(left(LEFT(name, CHARINDEX(' ', Name  )),3) + ' 1 2015') as varchar) [Month] , CAST(YEAR(getdate()) as varchar) [Year]" & _
" from [@AB_IPOWERPERIOD] "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            p_oDTiPowerPeriod = ExecuteSQLQuery_DT(sSqlstr)

            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '----------------------- SAP Period
            sSqlstr = "select Code, Name, cast(MONTH(F_RefDate ) as varchar) + ' ' + cast(year(F_RefDate ) as varchar), " & _
                "month(F_RefDate ) [F_Month], MONTH(T_RefDate ) [T_Month]," & _
"YEAR(F_RefDate ) [Year], F_RefDate [RefDate_F], T_RefDate [RefDate_T] from OFPR  "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            p_oDTSAPPeriod = ExecuteSQLQuery_DT(sSqlstr)



            Console.WriteLine("Completed With SUCCESS " & sFuncName)
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

    Public Function IdentifyTXTFile_JournalEntry(ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   IdentifyTXTFile_JournalEntry()
        '   Purpose     :   This function will identify the TXT file of Journal Entry
        '                    Upload the file into Dataview and provide the information to post transaction in SAP.
        '                     Transaction Success : Move the TXT file to SUCESS folder
        '                     Transaction Fail :    Move the TXT file to FAIL folder and send Error notification to concern person
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
        Dim oDVJE As DataView = Nothing
        Dim oDVIMPSTS As DataView = Nothing
        Dim oDICompany() As SAPbobsCOM.Company = Nothing
        Dim sCompanyDB As String = String.Empty
        Dim oDT_Entity As DataTable = Nothing
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "IdentifyTXTFile_JournalEntry()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oDT_StatisticsRowCount = New DataTable
            oDT_StatisticsRowCount.Columns.Add("Entity", GetType(String))
            oDT_StatisticsRowCount.Columns.Add("Count", GetType(Integer))

            oDT_Entity = New DataTable()
            oDT_Entity.Columns.Add("Entity", GetType(String))

            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
            Dim files() As System.IO.FileInfo

            files = DirInfo.GetFiles("*.txt")

            For Each File As System.IO.FileInfo In files
                bJEFileExist = True
                Console.WriteLine("Attempting File Name - " & File.Name, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting File Name - " & File.Name, sFuncName)
                'sFileType = Replace(File.Name, ".txt", "").Trim
                'upload the CSV to Dataview

                Console.WriteLine("GetDataViewFromTXT() ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("GetDataViewFromTXT() ", sFuncName)
                oDVJE = GetDataViewFromTXT(File.FullName, File.Name, sErrDesc)
                ''  oDTDistinct = oDVJE.Table.DefaultView.ToTable(True, "Entity")
                If sErrDesc.Length > 1 Then
                    Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                    FileMoveToArchive(File, File.FullName, RTN_ERROR, "")
                    Write_TextFile_I("Invalid File Format , preferable format is Txt {Tab} Delimiter ", sErrDesc)
                    IdentifyTXTFile_JournalEntry = RTN_ERROR
                    Exit Function
                End If
                oDVIMPSTS = New DataView(p_oDTImportStatistics)
                '' oDTDistinct = oDVIMPSTS.Table.DefaultView.ToTable(True, "CompanyDB")

                If oDT_OUCODE.Rows.Count > 0 Then
                    Console.WriteLine("Calling Validation Error ...... ! ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Validation Error ", sFuncName)
                    Write_TextFile(oDT_OUCODE, sErrDesc)

                    Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                    FileMoveToArchive(File, File.FullName, RTN_ERROR, "")

                    sErrDesc = "Validation Error ..... ! [OU Code are not Exists in Company DB]"
                    IdentifyTXTFile_JournalEntry = RTN_ERROR
                    Exit Function
                End If
                Console.WriteLine("Merging Entity from iPower & Statistics", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Merging Entity from iPower & Statistics", sFuncName)

                For Each odr As DataRowView In oDVJE
                    oDT_Entity.Rows.Add(odr("Entity").ToString.Trim())
                Next
                For Each odr As DataRowView In oDVIMPSTS
                    oDT_Entity.Rows.Add(odr("CompanyDB").ToString.Trim())
                Next

                Console.WriteLine("Getting Distinct of Entity", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Distinct of Entity ", sFuncName)
                ' oDTDistinct = oDVJE.Table.DefaultView.ToTable(True, "Entity")
                oDTDistinct = oDT_Entity.DefaultView.ToTable(True, "Entity")
                ReDim oDICompany(oDTDistinct.Rows.Count)

                oDT_StatisticsRowCount.Clear()
                For imjs As Integer = 0 To oDTDistinct.Rows.Count - 1
                    If ExecuteSQLQuery(oDTDistinct.Rows(imjs).Item(0).ToString.Trim, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Next

                For imjs As Integer = 0 To oDTDistinct.Rows.Count - 1

                    If String.IsNullOrEmpty(oDTDistinct.Rows(imjs).Item(0).ToString) Then Exit For
                    oDICompany(imjs) = New SAPbobsCOM.Company


                    ''If oDTDistinct.Rows(imjs).Item(0).ToString <> "MYCO" Then Continue For

                    Console.WriteLine("Calling ConnectToTargetCompany()", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany() " & oDICompany(imjs).CompanyDB, sFuncName)
                    If ConnectToTargetCompany(oDICompany(imjs), oDTDistinct.Rows(imjs).Item(0).ToString, sErrDesc) <> RTN_SUCCESS Then
                        Throw New ArgumentException(sErrDesc)
                    End If

                    Console.WriteLine("Starting transaction on company database " & oDICompany(imjs).CompanyDB, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDICompany(imjs).CompanyDB, sFuncName)
                    oDICompany(imjs).StartTransaction()


                    Console.WriteLine("Filtering data with respective Entity -  " & oDTDistinct.Rows(imjs).Item(0).ToString, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Filtering data with respective Entity -  " & oDTDistinct.Rows(imjs).Item(0).ToString, sFuncName)
                    oDVJE.RowFilter = "Entity = '" & oDTDistinct.Rows(imjs).Item(0).ToString & "'"

                    Console.WriteLine("Calling Function JournalEntry_Posting() ", sFuncName)


                    If oDVJE.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function JournalEntry_Posting() ", sFuncName)
                        If JournalEntry_Posting(oDVJE, oDICompany(imjs), File.Name, sErrDesc) <> RTN_SUCCESS Then
                            For lCounter As Integer = 0 To UBound(oDICompany)
                                If Not oDICompany(lCounter) Is Nothing Then
                                    If oDICompany(lCounter).Connected = True Then
                                        If oDICompany(lCounter).InTransaction = True Then
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                            oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        End If
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                        oDICompany(lCounter).Disconnect()
                                        oDICompany(lCounter) = Nothing
                                    End If
                                End If
                            Next
                            Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder ", sFuncName)
                            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                            'AddDataToTable(p_oDtError, File.Name, "Error", sErrDesc)
                            FileMoveToArchive(File, File.FullName, RTN_ERROR, "")

                            'Console.WriteLine("Error in updation. RollBack executed for ", sFuncName)
                            'If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Error in updation. RollBack executed for " & File.FullName, sFuncName)
                            IdentifyTXTFile_JournalEntry = RTN_ERROR
                            Exit Function
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data Found ", sFuncName)
                        Console.WriteLine("No Data Found ", sFuncName)
                    End If


                    '------------------------------------------------------------------------------------------------------------------------------------------------
                    '----------------------- Import Statistics
                    '------------------------------------------------------------------------------------------------------------------------------------------------

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Import Statistics Functionality ", sFuncName)
                    Console.WriteLine(" Attempting Import Statistics Functionality ..", sFuncName)


                    oDVIMPSTS.RowFilter = "CompanyDB = '" & oDTDistinct.Rows(imjs).Item(0).ToString & "'"

                    Console.WriteLine("Calling Function ImportStatistics() ", sFuncName)
                    If oDVIMPSTS.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function ImportStatistics() ", sFuncName)
                        If ImportStatistics(oDVIMPSTS, oDICompany(imjs), sErrDesc) <> RTN_SUCCESS Then
                            For lCounter As Integer = 0 To UBound(oDICompany)
                                If Not oDICompany(lCounter) Is Nothing Then
                                    If oDICompany(lCounter).Connected = True Then
                                        If oDICompany(lCounter).InTransaction = True Then
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                            oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        End If
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                        oDICompany(lCounter).Disconnect()
                                        oDICompany(lCounter) = Nothing
                                    End If
                                End If
                            Next

                            Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder ", sFuncName)
                            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                            'AddDataToTable(p_oDtError, File.Name, "Error", sErrDesc)
                            FileMoveToArchive(File, File.FullName, RTN_ERROR, "")
                            Throw New ArgumentException(sErrDesc)
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function ImportStatistics() ", sFuncName)
                        Console.WriteLine("No Data Found ", sFuncName)
                    End If

                Next imjs

                For lCounter As Integer = 0 To UBound(oDICompany)
                    If Not oDICompany(lCounter) Is Nothing Then
                        If oDICompany(lCounter).Connected = True Then
                            If oDICompany(lCounter).InTransaction = True Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            End If
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                            oDICompany(lCounter).Disconnect()
                            oDICompany(lCounter) = Nothing
                        End If
                    End If
                Next
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
            IdentifyTXTFile_JournalEntry = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed With ERROR", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR", sFuncName)
            IdentifyTXTFile_JournalEntry = RTN_ERROR
        End Try

    End Function

    Public Function GetDataViewFromTXT_OLD(ByVal CurrFileToUpload As String, ByVal Filename As String) As DataView

        ' **********************************************************************************
        '   Function    :   GetDataViewFromCSV()
        '   Purpose     :   This function will upload the data from CSV file to Dataview
        '   Parameters  :   ByRef CurrFileToUpload AS String 
        '                       CurrFileToUpload = File Name
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************

        Dim dv As DataView

        Dim sFuncName As String = String.Empty
        Dim dvEntiry As DataView = New DataView(p_oEntitesDetails)
        Dim dvGLAcccount As DataView = New DataView(p_oGLAccount)
        Dim sEntity As String = String.Empty
        Dim sBUCode As String = String.Empty
        Dim sLOS As String = String.Empty
        Dim oSR As StreamReader
        Dim iCount As Integer = 1
        Dim sGLAccount As String = String.Empty

        Try
            sFuncName = "GetDataViewFromTXT"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            'Console.WriteLine("Create_schema() ", sFuncName)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Create_schema() ", sFuncName)
            'Create_schema(p_oCompDef.sInboxDir, Filename)

            'The Datatable to Return
            Dim oipower As New DataTable()
            Dim oDT_ImportStatistics As New DataTable()

            oipower.Columns.Add("GL_Code", GetType(String))
            oipower.Columns.Add("Date", GetType(DateTime)) ' Date
            oipower.Columns.Add("Col3", GetType(String))
            oipower.Columns.Add("Amount", GetType(String)) ' Amount
            oipower.Columns.Add("Ref1", GetType(String))
            oipower.Columns.Add("Col6", GetType(String))
            oipower.Columns.Add("Col7", GetType(String))
            oipower.Columns.Add("Description", GetType(String))
            oipower.Columns.Add("Col9", GetType(String))
            oipower.Columns.Add("Col10", GetType(String))
            oipower.Columns.Add("OU", GetType(String))
            oipower.Columns.Add("EntityCode", GetType(String))
            oipower.Columns.Add("Col13", GetType(String))
            oipower.Columns.Add("GST", GetType(String))
            oipower.Columns.Add("Col15", GetType(String))
            oipower.Columns.Add("Voucher", GetType(String))
            oipower.Columns.Add("Entity", GetType(String))
            oipower.Columns.Add("BUCode", GetType(String))
            oipower.Columns.Add("LOSCode", GetType(String))
            oipower.Columns.Add("Year", GetType(String))
            oipower.Columns.Add("Code", GetType(String))

            oDT_ImportStatistics.Columns.Add("GL_Code", GetType(String))
            oDT_ImportStatistics.Columns.Add("Date", GetType(DateTime)) ' Date
            oDT_ImportStatistics.Columns.Add("Col3", GetType(String))
            oDT_ImportStatistics.Columns.Add("Amount", GetType(String)) ' Amount
            oDT_ImportStatistics.Columns.Add("Ref1", GetType(String))
            oDT_ImportStatistics.Columns.Add("Col6", GetType(String))
            oDT_ImportStatistics.Columns.Add("Col7", GetType(String))
            oDT_ImportStatistics.Columns.Add("Description", GetType(String))
            oDT_ImportStatistics.Columns.Add("Col9", GetType(String))
            oDT_ImportStatistics.Columns.Add("Col10", GetType(String))
            oDT_ImportStatistics.Columns.Add("OU", GetType(String))
            oDT_ImportStatistics.Columns.Add("EntityCode", GetType(String))
            oDT_ImportStatistics.Columns.Add("Col13", GetType(String))
            oDT_ImportStatistics.Columns.Add("GST", GetType(String))
            oDT_ImportStatistics.Columns.Add("Col15", GetType(String))
            oDT_ImportStatistics.Columns.Add("Voucher", GetType(String))
            oDT_ImportStatistics.Columns.Add("Entity", GetType(String))
            oDT_ImportStatistics.Columns.Add("BUCode", GetType(String))
            oDT_ImportStatistics.Columns.Add("LOSCode", GetType(String))
            oDT_ImportStatistics.Columns.Add("Year", GetType(String))
            oDT_ImportStatistics.Columns.Add("Code", GetType(String))

            'Open the file in a stream reader.
            oSR = New StreamReader(CurrFileToUpload)

            Dim sText As String
            Dim sString(-1) As String
            Dim sDelimiter As String() = {vbTab}

            While oSR.Peek <> -1
                sText = oSR.ReadLine()
                If sText.Length > 1 Then
                    sString = sText.Split(sDelimiter, StringSplitOptions.None) ' "RemoveEmptyEntrie" I am also using the option to remove empty entries a
                    ' dtEntiry = p_oEntitesDetails.DefaultView.ToTable(True, sString(10))
                    dvEntiry.RowFilter = "OUCode='" & sString(10) & "'"

                    If dvEntiry.Count > 0 Then
                        If dvEntiry.Item(0)("EntityFlag").ToString.ToUpper = "YES" Then
                            dvEntiry.RowFilter = "EntityCode= '" & sString(11) & "'"
                            If dvEntiry.Count > 0 Then
                                sEntity = dvEntiry.Item(0)(2).ToString
                                sBUCode = dvEntiry.Item(0)(3).ToString
                                sLOS = dvEntiry.Item(0)(4).ToString
                            Else
                                sEntity = String.Empty
                                sBUCode = String.Empty
                                sLOS = String.Empty
                            End If

                        Else
                            sEntity = dvEntiry.Item(0)(2).ToString
                            sBUCode = dvEntiry.Item(0)(3).ToString
                            sLOS = dvEntiry.Item(0)(4).ToString
                        End If

                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("OU Code " & sString(10) & " Entity is Empty in Line " & iCount, sFuncName)
                        sEntity = ""
                        sBUCode = ""
                        sLOS = ""
                    End If
                    dvGLAcccount.RowFilter = "ExportCode='" & sString(0) & "'"
                    If dvGLAcccount.Count = 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No GL Account Found  " & sString(0) & "  Line No " & iCount, sFuncName)
                        sGLAccount = ""
                    Else
                        sGLAccount = dvGLAcccount.Item(0)(0).ToString
                    End If
                    oipower.Rows.Add(sGLAccount, DateTime.ParseExact(Right(sString(1), 8), "yyyyMMdd", Nothing), sString(2), sString(3), sString(4), sString(5), sString(6), sString(7), sString(8), _
                                     sString(9), sString(10), sString(11), sString(12), sString(13), sString(14), sString(15), sEntity, sBUCode, sLOS, Left(sString(1), 4), Right(Left(sString(1), 7), 3))

                    iCount += 1
                End If
            End While

            'Console.WriteLine("Del_schema() ", sFuncName)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Del_schema() ", sFuncName)
            'Del_schema(p_oCompDef.sInboxDir)

            dv = New DataView(oipower)
            Return dv

        Catch ex As Exception

            Console.WriteLine("Error occured while reading content of  " & ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            Return Nothing
        Finally
            oSR.Close()
            oSR = Nothing
        End Try

    End Function

    Public Function ImportStatistics(ByRef oDVLineDetails As DataView, ByRef ocompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long

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
        Dim oRset As SAPbobsCOM.Recordset = ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sSql As String
        Dim oDVStatisticsrowcount As DataView = New DataView(oDT_StatisticsRowCount)
        Dim asql(100) As String
        Dim iloop As Integer = 0

        Try
            sFuncName = "ImportStatistics()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Filtering Entity " & ocompany.CompanyDB, sFuncName)
            oDVStatisticsrowcount.RowFilter = "Entity='" & ocompany.CompanyDB & "'"

            oDt = oDVLineDetails.ToTable
            sSql = String.Empty
            iCode = oDVStatisticsrowcount.Item(0)(1)
            ReDim asql(25)
            Dim icount As Integer = 0

            For Each row As DataRow In oDt.Rows
                ' write insert statement
                If icount = 5000 Then
                    iloop += 1
                    icount = 0
                End If

                asql(iloop) += " Insert Into [@AB_STATITISTICSDATA] ( [Code],  [Name], [U_AB_PERIOD],[U_AB_OPER_UNIT], " & _
                    " [U_AB_ENTITY], [U_AB_DEBIT_CREDIT],[U_AB_AMOUNT], [U_AB_GLCODE], [U_AB_DESCRIPTION]) " & _
                    " Values ('" & iCode & "','" & iCode & "','" & row.Item(1).ToString & "','" & row.Item(5).ToString & "','" & row.Item(6).ToString & "', " & _
                    "'" & row.Item(3).ToString & "'," & CDbl(row.Item(2).ToString) & ",'" & row.Item(0).ToString & "', '" & row.Item(4).ToString & "')"

                iCode = iCode + 1
                icount = icount + 1
            Next

            If asql.Length > 1 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Insert Data in " & ocompany.CompanyDB, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(ocompany.CompanyDB & " - Import Statistics Count " & icount, sFuncName)

                For Each element As String In asql
                    If Not String.IsNullOrEmpty(element) Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(element, sFuncName)
                        oRset.DoQuery(element)
                    Else
                        Exit For
                    End If
                Next
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserted Successful", sFuncName)
            End If

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
            Throw New Exception(ex.Message)
            Return Nothing
        Finally
            If oCon.State = ConnectionState.Open Then
                oCon.Dispose()
            End If
        End Try
        Return oDT
    End Function

    Public Function GetDataViewFromTXT(ByVal CurrFileToUpload As String, ByVal Filename As String, ByRef sErrDesc As String) As DataView

        ' **********************************************************************************
        '   Function    :   GetDataViewFromCSV()
        '   Purpose     :   This function will upload the data from CSV file to Dataview
        '   Parameters  :   ByRef CurrFileToUpload AS String 
        '                       CurrFileToUpload = File Name
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************

        Dim dv As DataView

        Dim sFuncName As String = String.Empty
        sErrDesc = String.Empty
        Dim dvEntiry As DataView = New DataView(p_oEntitesDetails)
        ' Dim dvGLAcccount As DataView = New DataView(p_oGLAccount)
        Dim oDvNewGlCode As DataView = New DataView(p_oSTOLDCODE)
        Dim oDVCompanyData As DataView = New DataView(p_oDTCompanyData)
        Dim oDVipowerPeriod As DataView = New DataView(p_oDTiPowerPeriod)
        Dim oDVsapPeriod As DataView = New DataView(p_oDTSAPPeriod)
        Dim dperioddate As Date
        Dim sEntity As String = String.Empty
        Dim sBUCode As String = String.Empty
        Dim sLOS As String = String.Empty
        Dim oSR As StreamReader
        Dim iCount As Integer = 1
        Dim sGLAccount As String = String.Empty
        Dim sSAPPeriodCode As String = String.Empty
        oDT_OUCODE = New DataTable()


        Try
            sFuncName = "GetDataViewFromTXT"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            'Console.WriteLine("Create_schema() ", sFuncName)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Create_schema() ", sFuncName)
            'Create_schema(p_oCompDef.sInboxDir, Filename)

            'The Datatable to Return
            Dim oipower As New DataTable()
            p_oDTImportStatistics = New DataTable()

            oipower.Columns.Add("GL_Code", GetType(String))
            oipower.Columns.Add("Date", GetType(DateTime)) ' Date
            oipower.Columns.Add("Col3", GetType(String))
            oipower.Columns.Add("Amount", GetType(String)) ' Amount
            oipower.Columns.Add("Ref1", GetType(String))
            oipower.Columns.Add("Col6", GetType(String))
            oipower.Columns.Add("Col7", GetType(String))
            oipower.Columns.Add("Description", GetType(String))
            oipower.Columns.Add("Col9", GetType(String))
            oipower.Columns.Add("Col10", GetType(String))
            oipower.Columns.Add("OU", GetType(String))
            oipower.Columns.Add("EntityCode", GetType(String))
            oipower.Columns.Add("Col13", GetType(String))
            oipower.Columns.Add("GST", GetType(String))
            oipower.Columns.Add("Col15", GetType(String))
            oipower.Columns.Add("Voucher", GetType(String))
            oipower.Columns.Add("Entity", GetType(String))
            oipower.Columns.Add("BUCode", GetType(String))
            oipower.Columns.Add("LOSCode", GetType(String))
            oipower.Columns.Add("Year", GetType(String))
            oipower.Columns.Add("Code", GetType(String))

            p_oDTImportStatistics.Columns.Add("GL_Code", GetType(String))
            p_oDTImportStatistics.Columns.Add("Period", GetType(String))
            p_oDTImportStatistics.Columns.Add("Amount", GetType(Double))
            p_oDTImportStatistics.Columns.Add("Type", GetType(String))
            p_oDTImportStatistics.Columns.Add("Description", GetType(String))
            p_oDTImportStatistics.Columns.Add("OperUnit", GetType(String))
            p_oDTImportStatistics.Columns.Add("Entity", GetType(String))
            p_oDTImportStatistics.Columns.Add("CompanyDB", GetType(String))

            oDT_OUCODE.Columns.Add("Sno", GetType(String))
            oDT_OUCODE.Columns.Add("OUCode", GetType(String))
            oDT_OUCODE.Columns.Add("Msg", GetType(String))

            'Open the file in a stream reader.
            oSR = New StreamReader(CurrFileToUpload)

            Dim sText As String
            Dim sString(-1) As String
            ''  Dim sDelimiter As String() = {vbTab}
            Dim sDelimiter As String() = {vbTab}

            While oSR.Peek <> -1
                sText = oSR.ReadLine()
                If sText.Length > 1 Then
                    sString = sText.Split(sDelimiter, StringSplitOptions.None) ' "RemoveEmptyEntrie" I am also using the option to remove empty entries a
                    'sString = sText.Split(" ")
                    ' dtEntiry = p_oEntitesDetails.DefaultView.ToTable(True, sString(10))
                    If sString.Length = "1" Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invalid File Format , preferable format is Txt {Tab} Delimiter  ", sFuncName)
                        Console.WriteLine("Invalid File Format , preferable format is Txt {Tab} Delimiter ")
                        sErrDesc = "Invalid File Format , preferable format is Txt {Tab} Delimiter  "
                        Exit While
                    End If
                    If sString(0) = "5725" Or sString(0) = "3172" Then Continue While

                    oDvNewGlCode.RowFilter = "Name='" & sString(0) & "'"
                    If oDvNewGlCode.Count = 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No data found for Export Code [@AB_IPOWERSTCODE] " & sString(0) & "  Line No " & iCount, sFuncName)
                        oDT_OUCODE.Rows.Add(iCount, sString(0), "No data found for Export Code in [@AB_IPOWERSTCODE]  :- " & p_oCompDef.sSAPDBName)
                    Else
                        sGLAccount = oDvNewGlCode.Item(0)("U_AB_STNEWCODE").ToString
                        If oDvNewGlCode.Item(0)("Cat").ToString = "IMS" Then
                            dvEntiry.RowFilter = "OUCode='" & sString(10) & "'"
                            If dvEntiry.Count > 0 Then
                                ''sEntity = dvEntiry.Item(0)(2).ToString
                                ''sBUCode = dvEntiry.Item(0)(3).ToString
                                ''sLOS = dvEntiry.Item(0)(4).ToString
                                ''If String.IsNullOrEmpty(sEntity) Then
                                ''    oDT_OUCODE.Rows.Add(iCount, sString(10), "No Entity tagged with this OU Code in Company DB :- " & p_oCompDef.sSAPDBName)
                                ''End If
                                If dvEntiry.Item(0)("EntityFlag").ToString.ToUpper = "YES" Then
                                    'If sString(10) = "52000" Then
                                    '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entity Flag is YES - Entity Code  " & sString(11) & " --  " & iCount, sFuncName)
                                    'End If
                                    oDVCompanyData.RowFilter = "ipEntityCode ='" & sString(11) & "'"
                                    ' dvEntiry.RowFilter = "EntityCode= '" & sString(11) & "'"
                                    If oDVCompanyData.Count > 0 Then
                                        'If sString(10) = "52000" Then
                                        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Count > 0   " & oDVCompanyData.Count & " --  " & iCount, sFuncName)
                                        'End If
                                        sEntity = oDVCompanyData.Item(0)(1).ToString
                                        sBUCode = dvEntiry.Item(0)(3).ToString
                                        sLOS = dvEntiry.Item(0)(4).ToString
                                        'If sString(10) = "52000" Then
                                        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entity Name  " & sEntity & " --  " & iCount, sFuncName)
                                        'End If
                                        If String.IsNullOrEmpty(sEntity) Then
                                            oDT_OUCODE.Rows.Add(iCount, sString(10), "No Entity tagged with this OU Code in Company DB :- " & p_oCompDef.sSAPDBName)
                                        End If
                                    Else
                                        sEntity = String.Empty
                                        sBUCode = String.Empty
                                        sLOS = String.Empty
                                        oDT_OUCODE.Rows.Add(iCount, sString(10), "No Entity tagged with this OU Code in Company DB :- " & p_oCompDef.sSAPDBName)
                                    End If

                                Else
                                    sEntity = dvEntiry.Item(0)(2).ToString
                                    sBUCode = dvEntiry.Item(0)(3).ToString
                                    sLOS = dvEntiry.Item(0)(4).ToString
                                    If String.IsNullOrEmpty(sEntity) Then
                                        oDT_OUCODE.Rows.Add(iCount, sString(10), "No Entity tagged with this OU Code in Company DB :- " & p_oCompDef.sSAPDBName)
                                    End If
                                End If
                            Else
                                oDT_OUCODE.Rows.Add(iCount, sString(10), "OU Code not Exists in Company DB :- " & p_oCompDef.sSAPDBName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("OU Code " & sString(10) & " Entity is Empty in Line " & iCount, sFuncName)
                                sEntity = ""
                                sBUCode = ""
                                sLOS = ""
                            End If
                            oDVipowerPeriod.RowFilter = "Code='" & Right(Left(sString(1), 7), 3) & "'"
                            If oDVipowerPeriod.Count > 0 Then

                                dperioddate = DateTime.ParseExact(Left(Left(sString(1), 7), 4) & oDVipowerPeriod.Item(0)("Month").ToString.PadLeft(2, "0"c) & "01", "yyyyMMdd", Nothing)
                                ' MsgBox("RefDate_F >= '#" & dperioddate & "#' and RefDate_T <= '#" & dperioddate & "#'")
                                oDVsapPeriod.RowFilter = "RefDate_F <= '#" & dperioddate & "#' and RefDate_T >= '#" & dperioddate & "#'" '"RefDate_F >= " & dperioddate & " and RefDate_T <= " & dperioddate & ""
                                If oDVsapPeriod.Count > 0 Then
                                    sSAPPeriodCode = oDVsapPeriod.Item(0)("Code").ToString
                                    If sSAPPeriodCode.Length = 6 Then
                                        sSAPPeriodCode = Right(sSAPPeriodCode, 4) & "0" & Left(sSAPPeriodCode, 2)
                                    End If
                                Else
                                    oDT_OUCODE.Rows.Add(iCount, sString(10), "No Related Period in SAP period table for Month " & MonthName(Month(dperioddate)) & "  Year " & Year(dperioddate) & " in Company DB :- " & p_oCompDef.sSAPDBName)
                                End If

                            Else
                                oDT_OUCODE.Rows.Add(iCount, sString(10), "No Related Period in ipower period table for " & Right(Left(sString(1), 7), 3) & " in Company DB :- " & p_oCompDef.sSAPDBName)
                            End If


                            p_oDTImportStatistics.Rows.Add(sGLAccount, sSAPPeriodCode, Left(sString(3), Len(sString(3)) - 1), Right(sString(3), 1), _
                                                         Replace(sString(7), "'", "''"), sString(10), sString(11), sEntity)
                        ElseIf oDvNewGlCode.Item(0)("Cat").ToString = "IP" Then

                            dvEntiry.RowFilter = "OUCode='" & sString(10) & "'"
                            If dvEntiry.Count > 0 Then
                                If dvEntiry.Item(0)("EntityFlag").ToString.ToUpper = "YES" Then
                                    'If sString(10) = "52000" Then
                                    '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entity Flag is YES - Entity Code  " & sString(11) & " --  " & iCount, sFuncName)
                                    'End If
                                    oDVCompanyData.RowFilter = "ipEntityCode ='" & sString(11) & "'"
                                    ' dvEntiry.RowFilter = "EntityCode= '" & sString(11) & "'"
                                    If oDVCompanyData.Count > 0 Then
                                        'If sString(10) = "52000" Then
                                        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Count > 0   " & oDVCompanyData.Count & " --  " & iCount, sFuncName)
                                        'End If
                                        sEntity = oDVCompanyData.Item(0)(1).ToString
                                        sBUCode = dvEntiry.Item(0)(3).ToString
                                        sLOS = dvEntiry.Item(0)(4).ToString
                                        'If sString(10) = "52000" Then
                                        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entity Name  " & sEntity & " --  " & iCount, sFuncName)
                                        'End If
                                        If String.IsNullOrEmpty(sEntity) Then
                                            oDT_OUCODE.Rows.Add(iCount, sString(10), "No Entity tagged with this OU Code in Company DB :- " & p_oCompDef.sSAPDBName)
                                        End If
                                    Else
                                        sEntity = String.Empty
                                        sBUCode = String.Empty
                                        sLOS = String.Empty
                                        oDT_OUCODE.Rows.Add(iCount, sString(10), "No Entity tagged with this OU Code in Company DB :- " & p_oCompDef.sSAPDBName)
                                    End If

                                Else
                                    sEntity = dvEntiry.Item(0)(2).ToString
                                    sBUCode = dvEntiry.Item(0)(3).ToString
                                    sLOS = dvEntiry.Item(0)(4).ToString
                                    If String.IsNullOrEmpty(sEntity) Then
                                        oDT_OUCODE.Rows.Add(iCount, sString(10), "No Entity tagged with this OU Code in Company DB :- " & p_oCompDef.sSAPDBName)
                                    End If
                                End If

                            Else
                                oDT_OUCODE.Rows.Add(iCount, sString(10), "OU Code not Exists in Company DB :- " & p_oCompDef.sSAPDBName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("OU Code " & sString(10) & " Entity is Empty in Line " & iCount, sFuncName)
                                sEntity = ""
                                sBUCode = ""
                                sLOS = ""
                            End If

                            oipower.Rows.Add(sGLAccount, DateTime.ParseExact(Right(sString(1), 8), "yyyyMMdd", Nothing), sString(2), sString(3), sString(4), sString(5), sString(6), sString(7), sString(8), _
                                       sString(9), sString(10), sString(11), sString(12), sString(13), sString(14), sString(15), sEntity, sBUCode, sLOS, Left(sString(1), 4), Right(Left(sString(1), 7), 3))
                        End If
                    End If
                End If
                iCount += 1
            End While

            'Console.WriteLine("Del_schema() ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
            'Del_schema(p_oCompDef.sInboxDir)

            dv = New DataView(oipower)
            Return dv

        Catch ex As Exception

            Console.WriteLine("Error occured while reading content of  " & ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            Return Nothing
        Finally
            oSR.Close()
            oSR = Nothing
        End Try

    End Function


    Public Function GetCompanyDetails(ByVal sEntity As String, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetCompanyDetails()
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

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL Server : " & oCompany.Server & " SQL Type " & p_oCompDef.sServerType _
                  & " License Server " & p_oCompDef.sLicenseServer & "  CompanyDB " & p_sSAPEntityName & " User Name " & p_sSAPUName & " pass " & p_sSAPUPass, sFuncName)

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
            Dim RenameCurrFileToUpload As String = Mid(oFile.Name, 1, oFile.Name.Length - 4) & "_" & Now.ToString("yyyyMMddhhmmss") & ".txt"

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
            Dim sFileName As String = "Validationip.txt"
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

    Public Function Write_TextFile_I(ByVal sString As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = String.Empty

        Try
            Dim sPath As String = System.Windows.Forms.Application.StartupPath & "\"
            Dim sFileName As String = "Validationip.txt"
            Dim sbuffer As String = String.Empty

            sFuncName = "Write_TextFile_I()"
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
            sw.WriteLine("")
            sw.WriteLine("Validation Error!    " & sString)
            sw.WriteLine(" ")
            sw.WriteLine(" ")
            sw.WriteLine("========================================================================================")
            sw.WriteLine("Please Check.")
            sw.Close()
            Process.Start(sPath & sFileName)

            Write_TextFile_I = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)

        Catch ex As Exception
            Write_TextFile_I = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

    Public Function MergeAutoNumberedToDataTable(ByVal SourceTable As DataTable, ByVal sErrDesc As String) As DataTable

        'Function   :   MergeAutoNumberedToDataTable()
        'Purpose    :   
        'Parameters :   ByVal SourceTable As DataTable
        '                   SourceTable= Source Datatable
        '               ByRef sErrDesc As String
        '                   sErrDesc=Error Description to be returned to calling function
        '               
        '                   =
        'Return     :   0 - FAILURE
        '               1 - SUCCESS
        'Author     :   John
        'Date       :   19-05-2015
        'Change     :

        Dim sFuncName As String

        Try
            sFuncName = "MergeAutoNumberedToDataTable()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Dim ResultTable As DataTable = New DataTable()
            Dim AutoNumberColumn As DataColumn = New DataColumn()
            AutoNumberColumn.ColumnName = "SNo"
            AutoNumberColumn.DataType = GetType(Integer)
            AutoNumberColumn.AutoIncrement = True
            AutoNumberColumn.AutoIncrementSeed = 1
            AutoNumberColumn.AutoIncrementStep = 1
            ResultTable.Columns.Add(AutoNumberColumn)
            ResultTable.Merge(SourceTable)
            ResultTable.Columns(0).SetOrdinal(ResultTable.Columns.Count - 1)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Return ResultTable
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try


    End Function

End Module
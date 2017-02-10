Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Namespace AE_PWC_AO06
    Module modCommon

        Public Function ConnectDICompSSO(ByRef objCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
            ' ***********************************************************************************
            '   Function   :    ConnectDICompSSO()
            '   Purpose    :    Connect To DI Company Object
            '
            '   Parameters :    ByRef objCompany As SAPbobsCOM.Company
            '                       objCompany = set the SAP Company Object
            '                   ByRef sErrDesc As String
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :    Sri
            '   Date       :    29 April 2013
            '   Change     :
            ' ***********************************************************************************
            Dim sCookie As String = String.Empty
            Dim sConnStr As String = String.Empty
            Dim sFuncName As String = String.Empty
            Dim lRetval As Long
            Dim iErrCode As Int32
            Try
                sFuncName = "ConnectDICompSSO()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                objCompany = New SAPbobsCOM.Company

                sCookie = objCompany.GetContextCookie
                sConnStr = p_oUICompany.GetConnectionContext(sCookie)
                'sConnStr = p_oSBOApplication.Company.GetConnectionContext(sCookie)
                lRetval = objCompany.SetSboLoginContext(sConnStr)

                If Not lRetval = 0 Then
                    Throw New ArgumentException("SetSboLoginContext of Single SignOn Failed.")
                End If
                p_oSBOApplication.StatusBar.SetText("Please Wait While Company Connecting... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                lRetval = objCompany.Connect
                If lRetval <> 0 Then
                    objCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException("Connect of Single SignOn failed : " & sErrDesc)
                Else
                    p_oSBOApplication.StatusBar.SetText("Company Connection Has Established with the " & objCompany.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                End If
                ConnectDICompSSO = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                ConnectDICompSSO = RTN_ERROR
            End Try
        End Function

        Public Function ConnectToTargetCompany(ByRef oCompany As SAPbobsCOM.Company, _
                                            ByVal sDBCode As String, _
                                            ByVal sApproverCode As String, _
                                            ByRef sErrDesc As String) As Long

            Dim sFuncName As String = String.Empty
            Dim iRetValue As Integer = -1
            Dim iErrCode As Integer = -1
            Dim sSQL As String = String.Empty
            Dim sSAPUser As String = String.Empty
            Dim sSAPPWd As String = String.Empty
            Dim sTrgtDBName As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset

            Try
                sFuncName = "ConnectToTargetCompany()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                Dim sConnectedCompany As String = p_oDICompany.CompanyDB

                '' sSQL = "SELECT * FROM [@AE_USERDETAILS] WHERE U_ENTITYCODE = '" & sDBCode & "' AND U_SAPUSERID = '" & sApproverCode & "' "
                sSQL = "SELECT U_PASSWORD FROM " & sHoldingDB & "..OUSR WHERE USER_CODE = '" & p_oDICompany.UserName & "' "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSQL, sFuncName)
                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(sSQL)
                If oRecordSet.RecordCount > 0 Then
                    'sTrgtDBName = oRecordSet.Fields.Item("U_ENTITYCODE").Value
                    'sSAPUser = oRecordSet.Fields.Item("U_SAPUSERID").Value
                    'sSAPPWd = oRecordSet.Fields.Item("U_PASSWORD").Value

                    sTrgtDBName = sDBCode
                    sSAPUser = p_oDICompany.UserName
                    sSAPPWd = oRecordSet.Fields.Item("U_PASSWORD").Value

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)
                    oCompany = New SAPbobsCOM.Company

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name : " & sTrgtDBName, sFuncName)
                    oCompany.Server = p_oDICompany.Server

                    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012

                    oCompany.LicenseServer = p_oDICompany.LicenseServer
                    oCompany.CompanyDB = sTrgtDBName
                    oCompany.UserName = sSAPUser
                    oCompany.Password = sSAPPWd

                    oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

                    oCompany.UseTrusted = False

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)
                    iRetValue = oCompany.Connect()

                    If iRetValue <> 0 Then
                        oCompany.GetLastError(iErrCode, sErrDesc)

                        sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                            oCompany.CompanyDB, System.Environment.NewLine, _
                                        vbTab, sErrDesc)

                        Throw New ArgumentException(sErrDesc)
                    End If
                Else
                    sErrDesc = "No Database login information found in COMPANYDATA Table. Please check"
                    Throw New ArgumentException(sErrDesc)
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connection established with " & oCompany.CompanyName, sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                ConnectToTargetCompany = RTN_SUCCESS
            Catch ex As Exception
                sErrDesc = ex.Message
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                ConnectToTargetCompany = RTN_ERROR
            End Try
        End Function

        Public Sub ShowErr(ByVal sErrMsg As String)
            ' ***********************************************************************************
            '   Function   :    ShowErr()
            '   Purpose    :    Show Error Message
            '   Parameters :  
            '                   ByVal sErrDesc As String
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :    Dev
            '   Date       :    23 Jan 2007
            '   Change     :
            ' ***********************************************************************************
            Try
                If sErrMsg <> "" Then
                    If Not p_oSBOApplication Is Nothing Then
                        If p_iErrDispMethod = ERR_DISPLAY_STATUS Then

                            p_oSBOApplication.SetStatusBarMessage("Error : " & sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short)
                        ElseIf p_iErrDispMethod = ERR_DISPLAY_DIALOGUE Then
                            p_oSBOApplication.MessageBox("Error : " & sErrMsg)
                        End If
                    End If
                End If
            Catch exc As Exception
                WriteToLogFile(exc.Message, "ShowErr()")
            End Try
        End Sub

        Public Sub LoadFromXML(ByVal FileName As String, ByVal Sbo_application As SAPbouiCOM.Application)
            Try
                Dim oXmlDoc As New Xml.XmlDocument
                Dim sPath As String
                ''sPath = IO.Directory.GetParent(Application.StartupPath).ToString
                sPath = System.Windows.Forms.Application.StartupPath.ToString
                'oXmlDoc.Load(sPath & "\AE_FleetMangement\" & FileName)
                oXmlDoc.Load(sPath & "\" & FileName)
                ' MsgBox(Application.StartupPath)

                Sbo_application.LoadBatchActions(oXmlDoc.InnerXml)
            Catch ex As Exception
                MsgBox(ex)
            End Try

        End Sub

        Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
        End Function

        Function ExecuteSQLQuery_DT(ByVal sQuery As String, ByVal sDBName As String, ByVal sSQLUser As String, ByVal sSQLPwd As String) As DataTable

            Dim oDT_INTDBInformations As DataTable
            Dim sFuncName As String = String.Empty
            Dim oConnection As SqlConnection = Nothing
            Dim oSQLCommand As SqlCommand = Nothing
            Dim oSQLAdapter As SqlDataAdapter = New SqlDataAdapter
            Dim sConnectionString As String

            Try
                sFuncName = "ExecuteSQLQuery_DT()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                sConnectionString = "Data Source=" & p_oDICompany.Server & ";Initial Catalog=" & sDBName & ";User ID=" & sSQLUser & "; Password=" & sSQLPwd

                oConnection = New SqlConnection(sConnectionString)

                If (oConnection.State = ConnectionState.Closed) Then
                    oConnection.Open()
                End If

                oDT_INTDBInformations = New DataTable
                oSQLCommand = New SqlCommand(sQuery, oConnection)
                oSQLAdapter.SelectCommand = oSQLCommand
                oSQLCommand.CommandTimeout = 0
                oSQLAdapter.Fill(oDT_INTDBInformations)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                Return oDT_INTDBInformations

            Catch ex As Exception
                Call WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Return Nothing
            Finally
                oSQLAdapter.Dispose()
                oSQLCommand.Dispose()
                oConnection.Close()
            End Try
        End Function

        Public Function ExecuteNonQuery(ByVal sQuery As String) As DataSet
            Dim sFuncName As String = "ExecuteNonQuery()"
            Dim oCmd As New SqlCommand
            Dim oDs As New DataSet
            Dim sSQLUser, sSQLPWd As String

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                sSQLUser = ConfigurationManager.AppSettings("DBUser")
            End If
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                sSQLPWd = ConfigurationManager.AppSettings("DBPwd")
            End If

            Dim sconstr As String = "Data Source=" & p_oDICompany.Server & ";Initial Catalog=" & p_oDICompany.CompanyDB & ";User ID=" & sSQLUser & "; Password=" & sSQLPWd
            Dim oCon As New SqlConnection(sconstr)

            Try
                oCon.ConnectionString = sconstr
                oCon.Open()
                oCmd.CommandType = CommandType.Text
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                oCmd.CommandText = sQuery
                oCmd.Connection = oCon
                oCmd.CommandTimeout = 0
                oCmd.ExecuteNonQuery()
                oCon.Close()
                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

            Catch ex As Exception
                Call WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while executing query", sFuncName)
                Throw New Exception(ex.Message)
            Finally
                oCon.Dispose()
            End Try
            Return oDs
        End Function

        Public Function CreateUDOTable(ByVal TableName As String, ByVal TableDescription As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
            Dim sFuncName As String = "CreateUDOTable"
            Dim sErrDesc As String = String.Empty
            Dim intRetCode As Integer
            Dim objUserTableMD As SAPbobsCOM.UserTablesMD
            objUserTableMD = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            Try
                If (Not objUserTableMD.GetByKey(TableName)) Then
                    objUserTableMD.TableName = TableName
                    objUserTableMD.TableDescription = TableDescription
                    objUserTableMD.TableType = TableType
                    intRetCode = objUserTableMD.Add()
                    If (intRetCode = 0) Then
                        Return True
                    End If
                Else
                    Return False
                End If
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFuncName)
                Throw New ArgumentException(sErrDesc)
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD)
                GC.Collect()
            End Try
        End Function

        Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
            Dim sFuncName As String = "addField"
            Dim sErrDesc As String = String.Empty
            Dim intLoop As Integer
            Dim strValue, strDesc As Array
            Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
            Try

                strValue = ValidValues.Split(Convert.ToChar(","))
                strDesc = ValidDescriptions.Split(Convert.ToChar(","))
                If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                    Throw New Exception("Invalid Valid Values")
                End If

                objUserFieldMD = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                If (Not isColumnExist(TableName, ColumnName)) Then
                    objUserFieldMD.TableName = TableName
                    objUserFieldMD.Name = ColumnName
                    objUserFieldMD.Description = ColDescription
                    objUserFieldMD.Type = FieldType
                    If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                        objUserFieldMD.Size = Size
                    Else
                        objUserFieldMD.EditSize = Size
                    End If
                    objUserFieldMD.SubType = SubType
                    objUserFieldMD.DefaultValue = SetValidValue
                    For intLoop = 0 To strValue.GetLength(0) - 1
                        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                        objUserFieldMD.ValidValues.Add()
                    Next
                    If (objUserFieldMD.Add() <> 0) Then
                        sErrDesc = p_oDICompany.GetLastErrorCode() & ":" & p_oDICompany.GetLastErrorDescription()
                        Throw New ArgumentException(sErrDesc)
                    End If
                End If

            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFuncName)
                Throw New ArgumentException(sErrDesc)
            Finally

                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
                GC.Collect()

            End Try


        End Sub

        Private Function isColumnExist(ByVal TableName As String, ByVal ColumnName As String) As Boolean
            Dim objRecordSet As SAPbobsCOM.Recordset
            objRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                objRecordSet.DoQuery("SELECT COUNT(*) FROM CUFD WHERE TableID = '" & TableName & "' AND AliasID = '" & ColumnName & "'")
                If (Convert.ToInt16(objRecordSet.Fields.Item(0).Value) <> 0) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet)
                GC.Collect()
            End Try

        End Function

        Public Function StartTransaction(ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
            Dim sFuncName As String = "StartTransaction"
            Try
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Transaction", sFuncName)

                If oCompany.InTransaction Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback hanging transactions", sFuncName)
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End If

                oCompany.StartTransaction()

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Trancation Started Successfully", sFuncName)
                StartTransaction = RTN_SUCCESS

            Catch ex As Exception
                Call WriteToLogFile_Debug(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while starting Trancation", sFuncName)
                StartTransaction = RTN_ERROR
            End Try

        End Function

        Public Function CommitTransaction(ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
            Dim sFuncName As String = "CommitTransaction"
            Try
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
                If oCompany.InTransaction Then
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Transaction is Active", sFuncName)
                End If

                CommitTransaction = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit Transaction Complete", sFuncName)
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while committing Transaciton", sFuncName)
                CommitTransaction = RTN_ERROR
            End Try
        End Function

        Public Function RollbackTransaction(ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "RollbackTransaction()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If oCompany.InTransaction Then
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No transaction is active", sFuncName)
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Success", sFuncName)
                RollbackTransaction = RTN_SUCCESS
            Catch ex As Exception
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFuncName)
                RollbackTransaction = RTN_ERROR
            End Try

        End Function

    End Module
End Namespace


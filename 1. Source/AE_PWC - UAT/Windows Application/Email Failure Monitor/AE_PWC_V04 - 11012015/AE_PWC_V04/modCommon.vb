Imports System.Configuration
Imports System.Data.SqlClient



Module modCommon

#Region "Variable Declaration"
    Public frmEmailMonitorF As frmEmailMonitor
    ' Return Value Variable Control
    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0
    ' Debug Value Variable Control
    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    ' Global variables group
    Public p_iDebugMode As Int16 = DEBUG_ON
    Public p_iDeleteDebugLog As Int16

    Public Structure CompanyDefault

        Public sServer As String
        Public sDBName As String
        Public sDBUser As String
        Public sDBPwd As String
    End Structure

    Public p_oCompDef As CompanyDefault

#End Region



    Public Sub EmailMonitor_ShowDialog()

        ' If CompanyDetails_F IsNot Nothing AndAlso Not CompanyDetails_F.IsDisposed Then Exit Sub
        Try
            Dim CloseApp As Boolean = False
            frmEmailMonitorF = New frmEmailMonitor
            frmEmailMonitorF.Show()
            CloseApp = (frmEmailMonitorF.DialogResult = DialogResult.Abort)
            frmEmailMonitorF = Nothing

            If CloseApp Then Application.Exit()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Public Function Load_Emailstatusfails(ByRef sErrDesc As String, ByVal Dgv_Emailmonitor As DataGridView) As Long

        ' **********************************************************************************
        '   Function    :   Load_Emailstatusfails()
        '   Purpose     :   This function will provide the list of Email status fails 
        '               
        '   Parameters  :   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   July 2015
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDsBPList As DataSet = Nothing
        Try

            sFuncName = "Load_Emailstatusfails()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sSQL = "Select T0.DocType , RIGHT( T0.EmailSub, LEN( T0.EmailSub) - (CHARINDEX('Draft No.', T0.EmailSub) + 9)) [Draftkey] ,T0.Entity , T0.EmailID  , " & _
                "T0.ErrMsg , T0.Sno  from [AB_EmailStatus] T0 where T0.Status = 'Fail'"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Execute SQL" & sSQL, sFuncName)
            oDsBPList = ExecuteSQLQuery(sSQL)
            Dgv_Emailmonitor.Rows.Clear()
            For imjs As Integer = 0 To oDsBPList.Tables(0).Rows.Count - 1
                Dgv_Emailmonitor.Rows.Add(1)
                '' MsgBox(oDsBPList.Tables(0).Rows(imjs)("CardCode").ToString)
                Dgv_Emailmonitor.Rows.Item(imjs).Cells.Item("DocType").Value = oDsBPList.Tables(0).Rows(imjs)("DocType").ToString
                Dgv_Emailmonitor.Rows.Item(imjs).Cells.Item("Draftkey").Value = oDsBPList.Tables(0).Rows(imjs)("Draftkey").ToString
                Dgv_Emailmonitor.Rows.Item(imjs).Cells.Item("entity").Value = oDsBPList.Tables(0).Rows(imjs)("Entity").ToString
                Dgv_Emailmonitor.Rows.Item(imjs).Cells.Item("emailid").Value = oDsBPList.Tables(0).Rows(imjs)("EmailID").ToString
                Dgv_Emailmonitor.Rows.Item(imjs).Cells.Item("errmsg").Value = oDsBPList.Tables(0).Rows(imjs)("ErrMsg").ToString
                Dgv_Emailmonitor.Rows.Item(imjs).Cells.Item("Refno").Value = oDsBPList.Tables(0).Rows(imjs)("Sno").ToString

            Next
            ''   Me.Dgv_BPList.Rows(0).Cells(2).Selected = True

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Load_Emailstatusfails = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Load_Emailstatusfails = RTN_ERROR
        End Try
    End Function

    Public Function ExecuteSQLQuery(ByVal sQuery As String) As DataSet

        '**************************************************************
        ' Function      : ExecuteSQLQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : Sri
        ' Date          : 
        ' Change        :
        '**************************************************************

        Dim sFuncName As String = String.Empty

        ' Dim sConstr As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName
        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd
        ''Dim oCon As New Odbc.OdbcConnection(sConstr)
        ''Dim oCmd As New Odbc.OdbcCommand
        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oDs As New DataSet

        Try
            sFuncName = "ExecuteQuery()"
            oCon.ConnectionString = sConstr
            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            da.Fill(oDs)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            oCon.Dispose()
        End Try
        Return oDs
    End Function

    Public Function ExecuteNonSQLQuery(ByVal sQuery As String) As DataSet

        '**************************************************************
        ' Function      : ExecuteNonSQLQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : Sri
        ' Date          : 
        ' Change        :
        '**************************************************************

        Dim sFuncName As String = String.Empty

        ' Dim sConstr As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName
        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd
        ''Dim oCon As New Odbc.OdbcConnection(sConstr)
        ''Dim oCmd As New Odbc.OdbcCommand
        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oDs As New DataSet

        Try
            sFuncName = "ExecuteNonSQLQuery()"
            oCon.ConnectionString = sConstr
            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            oCmd.ExecuteNonQuery()
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            oCon.Dispose()
        End Try
        Return oDs
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
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oCompDef.sDBName = String.Empty
            oCompDef.sServer = String.Empty
            oCompDef.sDBUser = String.Empty
            oCompDef.sDBPwd = String.Empty


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.sServer = ConfigurationManager.AppSettings("Server")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBName")) Then
                oCompDef.sDBName = ConfigurationManager.AppSettings("DBName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.sDBUser = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.sDBPwd = ConfigurationManager.AppSettings("DBPwd")
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

    Public Function Update_Emailstatusfails(ByRef sErrDesc As String, ByVal Dgv_Emailmonitor As DataGridView) As Long

        ' **********************************************************************************
        '   Function    :   Update_Emailstatusfails()
        '   Purpose     :   This function will provide the list of Email status fails 
        '               
        '   Parameters  :   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   July 2015
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim iCount As Integer = 1
        Try

            sFuncName = "Update_Emailstatusfails()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            For imjs As Integer = 0 To Dgv_Emailmonitor.Rows.Count - 2 'oDsBPList.Tables(0).Rows.Count - 1
                If Convert.ToBoolean(Dgv_Emailmonitor.Rows(imjs).Cells(0).Value) = True Then
                    sSQL = sSQL + "UPDATE [AB_EmailStatus] SET EmailID = '" & Dgv_Emailmonitor.Rows.Item(imjs).Cells.Item("emailid").Value & "', " & _
                                            "Status = 'Open', ErrMsg = '', EmailDate = NULL, EmailTime = NULL WHERE SNO = '" & Dgv_Emailmonitor.Rows.Item(imjs).Cells.Item("Refno").Value & "'"
                    iCount += 1
                    If iCount = 500 Then
                        If sSQL.Length > 1 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteNonSQLQuery()", sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Update SQL " & sSQL, sFuncName)
                            ExecuteNonSQLQuery(sSQL)
                            sSQL = String.Empty
                        End If
                    End If
                End If
            Next

            If sSQL.Length > 1 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteNonSQLQuery()", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Update SQL " & sSQL, sFuncName)
                ExecuteNonSQLQuery(sSQL)
                sSQL = String.Empty
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Update_Emailstatusfails = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Update_Emailstatusfails = RTN_ERROR
        End Try
    End Function

End Module
